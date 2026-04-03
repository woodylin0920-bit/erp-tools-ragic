#!/usr/bin/env python3
"""
Ragic 銷貨單自動化上傳腳本
用法：
  python ragic_upload.py client/LE/0324T221.xlsx (檔案名稱)
  python ragic_upload.py --dry-run client/LE/0324T221.xlsx
"""

import argparse
import json
import os
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

import questionary
import requests

QSTYLE = questionary.Style([
    ("question",                        "bold"),
    ("answer",                          "fg:#00aa00 bold"),
    ("pointer",                         "bold"),
    ("highlighted",                     "fg:#000000 bold"),
    ("text",                            "fg:#1a1a1a"),
    ("instruction",                     "fg:#555555"),
    # autocomplete 下拉選單
    ("completion-menu.completion",          "bg:#eeeeee fg:#000000"),
    ("completion-menu.completion.current",  "bg:#0066cc fg:#ffffff bold"),
])
from rich.console import Console
from rich.table import Table

# ============================================================
# ★ CONFIG ★
# ============================================================

RAGIC_BASE    = os.getenv("RAGIC_BASE",    "https://ap12.ragic.com")
RAGIC_ACCOUNT = os.getenv("RAGIC_ACCOUNT", "toybebop")

PRODUCT_PRICE_SHEET = os.getenv("PRODUCT_PRICE_SHEET", "ragicsales-order-management/20006")   # 商品單價管理
CUSTOMER_SHEET      = os.getenv("CUSTOMER_SHEET",      "ragicsales-order-management/20004")   # 客戶
SALES_ORDER_SHEET   = os.getenv("SALES_ORDER_SHEET",   "ragicsales-order-management/20001")   # 銷貨單

ORDER_ITEMS_SUBTABLE_KEY = os.getenv("ORDER_ITEMS_SUBTABLE_KEY", "_subtable_3000842")   # 訂購項目子表

# 客戶尚未建檔時使用的預留客戶
UNREGISTERED_CUSTOMER = {"code": "C-00000", "name": "000尚未建檔", "address": ""}

# 上傳記錄檔（防重複）
_UPLOAD_LOG = Path(__file__).resolve().parent.parent / "upload_log.json"


def _load_upload_log() -> dict:
    if _UPLOAD_LOG.exists():
        try:
            return json.loads(_UPLOAD_LOG.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_upload_log(log: dict):
    _UPLOAD_LOG.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")

_CID_LABELS = {
    "3000812": "訂單單別",    "3000813": "訂單日期",      "3000814": "訂單狀態",
    "3000815": "客戶編號",    "3000836": "課稅別",        "3000838": "稅率",
    "3001498": "訂單運費",    "3001684": "國貿條規",      "3000835": "小計",
    "3000837": "稅額",        "3000839": "總金額(含稅)",  "3000840": "備註",
    "1000074": "內部備注",    "3000845": "建檔日期時間",  "3000847": "最後修改日期時間",
    "3000830": "商品販售代號","3000832": "單價",          "3000833": "數量",
    "3000834": "金額",
}


def _humanize_payload(payload: dict) -> dict:
    result = {}
    for k, v in payload.items():
        if k == ORDER_ITEMS_SUBTABLE_KEY:
            rows = {rk: {_CID_LABELS.get(ck, ck): cv for ck, cv in rv.items()}
                    for rk, rv in v.items()}
            result["訂購項目"] = rows
        else:
            result[_CID_LABELS.get(k, k)] = v
    return result


# ============================================================

console = Console()

_KEY_FILE = Path.home() / ".boptoys-ai_key"

# ── Ragic API ────────────────────────────────────────────────

def _auth_header() -> dict:
    """Ragic API key 已是 Base64 格式，直接作為 Basic auth token。"""
    api_key = os.environ.get("RAGIC_API_KEY", "")
    if not api_key and _KEY_FILE.exists():
        api_key = _KEY_FILE.read_text().strip()
        os.environ["RAGIC_API_KEY"] = api_key
    if not api_key:
        console.print("[yellow]尚未設定 RAGIC_API_KEY[/yellow]")
        api_key = questionary.password("請輸入 Ragic API Key：").ask() or ""
        if not api_key:
            console.print("[red]未輸入 API Key，結束[/red]")
            sys.exit(1)
        _KEY_FILE.write_text(api_key)
        os.environ["RAGIC_API_KEY"] = api_key
        console.print(f"[green]✓ API Key 已儲存至 {_KEY_FILE}，下次不需再輸入[/green]")
    return {"Authorization": f"Basic {api_key}"}


def ragic_get(sheet_path: str, limit: int = 2000) -> dict:
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}?api&limit={limit}"
    r = requests.get(url, headers=_auth_header(), timeout=30)
    r.raise_for_status()
    data = r.json()
    return {k: v for k, v in data.items() if not k.startswith("_") and k != "info"}


def ragic_post(sheet_path: str, payload: dict) -> dict:
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}?api&doLinkLoad=true"
    r = requests.post(url, headers=_auth_header(), json=payload, timeout=30)
    r.raise_for_status()
    return r.json()


# ── 快取載入 ─────────────────────────────────────────────────

def load_price_index() -> Dict[str, list]:
    """載入商品單價管理，建立 {條碼: [商品...]} 索引。"""
    console.print("[cyan]載入商品單價資料（Ragic API）...[/cyan]")
    records = ragic_get(PRODUCT_PRICE_SHEET)
    index: Dict[str, list] = {}
    for rec in records.values():
        barcode = str(rec.get("國際條碼", "")).strip()
        if len(barcode) < 12:
            continue
        entry = {
            "product_code": str(rec.get("商品單價代號", "")),
            "product_name": str(rec.get("商品名稱", "")),
            "spec":         rec.get("規格", 1),
            "unit":         str(rec.get("單位", "")),
            "price":        float(rec.get("價格", 0) or 0),
        }
        index.setdefault(barcode, []).append(entry)
    console.print(f"[green]✓ 載入 {len(index)} 種條碼的商品[/green]")
    return index


def load_customers() -> list:
    """載入客戶資料表。"""
    console.print("[cyan]載入客戶資料（Ragic API）...[/cyan]")
    records = ragic_get(CUSTOMER_SHEET)
    customers = []
    for rec in records.values():
        customers.append({
            "code":    str(rec.get("客戶編號", "")),
            "name":    str(rec.get("客戶名稱", "")),
            "address": str(rec.get("送貨完整地址", "")),
        })
    console.print(f"[green]✓ 載入 {len(customers)} 筆客戶[/green]")
    return customers


# ── 客戶比對 ─────────────────────────────────────────────────

def find_customer(customers: list, store_code: str, client_code: str = "") -> Optional[dict]:
    # 若有 client_code（如 TRU），優先在該客群中搜尋
    if client_code:
        narrowed = [c for c in customers if store_code in c["name"] and client_code in c["name"]]
        if narrowed:
            matches = narrowed
        else:
            matches = [c for c in customers if store_code in c["name"]]
    else:
        matches = [c for c in customers if store_code in c["name"]]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        choices = [f"{c['code']}  {c['name']}" for c in matches]
        sel = questionary.select(f"找到多個含「{store_code}」的客戶，請選擇：", choices=choices).ask()
        return matches[choices.index(sel)]
    # 找不到 → 先詢問是否暫用尚未建檔，再讓使用者搜尋
    console.print(f"[yellow]⚠ 找不到含「{store_code}」的客戶[/yellow]")
    use_placeholder = questionary.confirm(
        "是否暫用「C-00000 000尚未建檔」代替？（之後在 Ragic 補填客戶）",
        default=True,
    ).ask()
    if use_placeholder:
        return UNREGISTERED_CUSTOMER
    all_choices = [f"{c['code']}  {c['name']}" for c in customers]
    sel = questionary.autocomplete(
        "搜尋客戶（輸入代碼或名稱片段）：",
        choices=all_choices,
        validate=lambda v: v in all_choices or "請從清單中選擇",
        style=QSTYLE,
    ).ask()
    return customers[all_choices.index(sel)]


# ── 商品比對 ─────────────────────────────────────────────────

def resolve_items(order_items, price_index: dict, auto_unit_spec: bool = False) -> list:
    """
    auto_unit_spec=True：數量單位為「個/盒」時自動選 spec=1（LE 格式適用）
    """
    resolved = []
    for item in order_items:
        matches = price_index.get(item.barcode, [])
        if not matches:
            console.print(f"[yellow]⚠ 條碼 {item.barcode}（{item.le_name}）不在商品單價表，已跳過[/yellow]")
            continue

        # TRU 檔已帶單價，優先使用；否則從 Ragic 商品表取預設價
        override_price = item.unit_price if item.unit_price > 0 else None

        if len(matches) == 1:
            product = matches[0]
            final_qty = int(item.quantity)
        else:
            # 同條碼多規格：找出能整除數量的規格
            viable = []
            for m in matches:
                spec_qty = int(float(m["spec"])) if m["spec"] else 1
                if spec_qty <= 1 or item.quantity % spec_qty == 0:
                    n = int(item.quantity / spec_qty) if spec_qty > 1 else int(item.quantity)
                    viable.append((m, n))

            if not viable:
                viable = [(m, int(item.quantity)) for m in matches]

            # 按規格數值升冪排列：單盒(1) → 中盒 → 整箱
            viable.sort(key=lambda x: int(float(x[0]["spec"]) if x[0]["spec"] else 1))

            if len(viable) == 1:
                product, final_qty = viable[0]
            elif auto_unit_spec:
                # LE 格式：數量為個/盒單位，自動選 spec=1（最小單位）
                unit_options = [(m, n) for m, n in viable if int(float(m["spec"]) if m["spec"] else 1) == 1]
                if unit_options:
                    product, final_qty = unit_options[0]
                else:
                    product, final_qty = viable[0]
            else:
                choices = [
                    f"{m['unit']} × {n}"
                    + (f"  ({int(float(m['spec']))}pcs/盒)" if int(float(m['spec']) if m['spec'] else 1) > 1 else "")
                    + f"  ({m['product_code']} @ {m['price']:.2f} = {m['price']*n:,.2f})"
                    for m, n in viable
                ]
                name_hint = item.le_name or item.barcode
                sel = questionary.select(
                    f"{name_hint}（數量: {int(item.quantity)}）- 請選擇規格",
                    choices=choices,
                ).ask()
                product, final_qty = viable[choices.index(sel)]

        resolved.append({
            "product_code": product["product_code"],
            "product_name": product["product_name"],
            "spec":         product["spec"],
            "unit":         product["unit"],
            "unit_price":   override_price if override_price else product["price"],
            "quantity":     final_qty,
            "amount":       round((override_price if override_price else product["price"]) * final_qty, 2),
        })
    return resolved


# ── 互動 UI ──────────────────────────────────────────────────

def show_items_table(customer: dict, store_code: str, po_number: str, resolved: list):
    console.print()
    console.rule(f"[bold]門市: {store_code}  PO: {po_number}  客戶: {customer['code']} {customer['name']}[/bold]")
    table = Table(show_header=True, header_style="bold cyan", box=None)
    table.add_column("#",       width=3)
    table.add_column("商品名稱", min_width=22)
    table.add_column("規格",    width=5,  justify="right")
    table.add_column("數量",    width=6,  justify="right")
    table.add_column("單價",    width=9,  justify="right")
    table.add_column("金額",    width=11, justify="right")
    subtotal = 0.0
    for i, it in enumerate(resolved, 1):
        table.add_row(str(i), it["product_name"], str(it["spec"]),
                      str(it["quantity"]), f"{it['unit_price']:,.2f}", f"{it['amount']:,.2f}")
        subtotal += it["amount"]
    console.print(table)
    console.print(f"[bold]小計: {subtotal:,.2f}[/bold]")
    console.print()


def ask_order_options(is_unregistered: bool = False) -> tuple:
    order_type = questionary.select(
        "請選擇訂單單別",
        choices=["一般訂單", "公關品", "樣品", "蝦皮", "官網"],
    ).ask()

    order_status = questionary.select(
        "請選擇訂單狀態",
        choices=["未出貨", "預接單", "已收款未出貨", "已出貨未收款", "尚未建檔"],
        default="尚未建檔" if is_unregistered else "未出貨",
    ).ask()

    tax_choice = questionary.select(
        "請選擇稅率",
        choices=["5%（含稅/外加）", "(5%)（內含/不計稅）"],
    ).ask()
    tax_rate = "5%" if "5%（" in tax_choice else "(5%)"

    shipping_str = questionary.text("運費（預設 0）", default="0").ask()
    shipping_fee = float(shipping_str or "60")

    notes = questionary.text("備註（留空直接按 Enter）", default="").ask()
    internal_notes = questionary.text("內部備注（留空直接按 Enter）", default="").ask()
    return order_type, order_status, tax_rate, shipping_fee, notes or "", internal_notes or ""


def show_confirmation(customer: dict, resolved: list, order_type: str, order_status: str,
                      tax_rate: str, shipping_fee: float, notes: str, internal_notes: str) -> tuple:
    subtotal = sum(it["amount"] for it in resolved)
    tax_amount = round(subtotal * 0.05, 2) if tax_rate == "5%" else 0.0
    total = round(subtotal + shipping_fee + tax_amount, 2)

    console.print()
    console.rule("[bold red]最終確認[/bold red]")
    console.print(f"訂單單別: [cyan]{order_type}[/cyan]  狀態: [cyan]{order_status}[/cyan]  客戶: [cyan]{customer['code']}  {customer['name']}[/cyan]")
    console.print(f"課稅別: 營業稅              稅率: [cyan]{tax_rate}[/cyan]")
    console.print(f"小計: {subtotal:>12,.2f}     稅額: {tax_amount:,.2f}")
    console.print(f"運費: {shipping_fee:>12.2f}     總計: [bold]{total:,.2f}[/bold]")
    if notes:
        console.print(f"備註: {notes}")
    if internal_notes:
        console.print(f"內部備注: [dim]{internal_notes}[/dim]")
    console.rule()
    return subtotal, tax_amount, total


# ── Payload 組裝 ─────────────────────────────────────────────

def build_payload(customer: dict, resolved: list, order_type: str, order_status: str,
                  tax_rate: str, shipping_fee: float, notes: str, internal_notes: str) -> dict:
    now   = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    today = date.today().strftime("%Y/%m/%d")

    # 計算各項金額
    subtable = {}
    subtotal = 0.0
    for i, it in enumerate(resolved):
        amount = round(it["unit_price"] * it["quantity"], 2)
        subtotal += amount
        subtable[str(-(i + 1))] = {
            "3000830": it["product_code"],   # 商品販售代號
            "3000832": it["unit_price"],      # 單價
            "3000833": it["quantity"],        # 數量
            "3000834": amount,               # 金額（單價×數量）
        }

    subtotal    = round(subtotal, 2)
    tax_amount  = round(subtotal * 0.05, 2) if tax_rate == "5%" else 0.0
    total       = round(subtotal + tax_amount + shipping_fee, 2)

    return {
        "3000812": order_type,               # 訂單單別
        "3000813": today,                    # 訂單日期
        "3000814": order_status,             # 訂單狀態
        "3000815": customer["code"],         # 客戶編號
        "3000836": "營業稅",                 # 課稅別
        "3000838": tax_rate,                 # 稅率
        "3001498": int(shipping_fee),        # 訂單運費
        "3001684": "DDP",                    # 國貿條規（預設）
        "3000835": subtotal,                 # 小計
        "3000837": tax_amount,               # 稅額
        "3000839": total,                    # 總金額(含稅)
        "3000840": notes,                    # 備註
        "1000074": f"【AI建單】 {internal_notes}".strip() if internal_notes else "【AI建單】",  # 內部備注
        "3000845": now,                      # 建檔日期時間
        "3000847": now,                      # 最後修改日期時間
        ORDER_ITEMS_SUBTABLE_KEY: subtable,
    }


BASE_CLIENT_ORDER = Path(__file__).resolve().parent.parent / "client_order"


def find_pending_files(base_dir: Path) -> list:
    files = []
    for client_dir in sorted(base_dir.iterdir()):
        if client_dir.is_dir():
            files.extend(sorted(client_dir.glob("*.xlsx")))
    return files


def process_file(excel_path: Path, args, price_index: dict, customers: list):
    from parsers import PARSERS
    client_code = excel_path.parent.name.upper()
    if client_code not in PARSERS:
        console.print(f"[red]不支援的客戶代碼：{client_code}（支援：{', '.join(PARSERS)}）[/red]")
        return 0, 0

    console.print(f"\n[cyan]解析 {excel_path.name}（{client_code} 格式）...[/cyan]")
    try:
        orders = PARSERS[client_code](str(excel_path)).parse()
    except Exception as e:
        console.print(f"[red]無法讀取 Excel 檔案：{e}[/red]")
        return 0, 0
    if not orders:
        console.print("[red]無法解析任何訂單，請確認檔案格式[/red]")
        return 0, 0
    console.print(f"[green]✓ 偵測到 {len(orders)} 張訂單[/green]")

    upload_log = _load_upload_log()
    success_count = 0
    for i, order in enumerate(orders, 1):
        console.print(f"\n{'═'*58}")
        console.print(f"[bold]訂單 {i}/{len(orders)}  門市: {order.store_code}  PO: {order.po_number}[/bold]")

        # 防重複：檢查是否已上傳過
        log_key = f"{client_code}_{order.store_code}_{order.po_number}"
        if log_key in upload_log and not args.dry_run:
            rec = upload_log[log_key]
            console.print(f"[yellow]⚠ 此訂單已於 {rec['uploaded_at']} 上傳（Ragic ID: {rec['ragic_id']}）[/yellow]")
            skip = questionary.confirm("是否跳過（建議跳過以避免重複）？", default=True).ask()
            if skip:
                console.print("[yellow]已跳過[/yellow]")
                continue

        customer = find_customer(customers, order.store_code, client_code)
        if not customer:
            console.print("[red]無法確認客戶，跳過[/red]")
            continue

        resolved = resolve_items(order.items, price_index, auto_unit_spec=(client_code == "LE"))
        if not resolved:
            console.print("[red]無有效商品，跳過[/red]")
            continue

        show_items_table(customer, order.store_code, order.po_number, resolved)

        is_unregistered = customer["code"] == UNREGISTERED_CUSTOMER["code"]
        order_type, order_status, tax_rate, shipping_fee, notes, internal_notes = ask_order_options(is_unregistered)

        show_confirmation(customer, resolved, order_type, order_status, tax_rate, shipping_fee, notes, internal_notes)

        confirmed = questionary.confirm("確認送出此訂單？", default=True).ask()
        if not confirmed:
            action = questionary.select(
                "請選擇：",
                choices=["跳過此單，繼續下一張", "放棄整個檔案，回到選單（不移至 done）"],
            ).ask()
            if "放棄" in action:
                if success_count > 0:
                    console.print(f"[yellow]已放棄。前 {success_count} 張已送出至 Ragic，請自行確認是否需要刪除。[/yellow]")
                else:
                    console.print("[yellow]已放棄，未送出任何訂單。[/yellow]")
                return success_count, len(orders), True
            console.print("[yellow]已跳過[/yellow]")
            continue

        payload = build_payload(customer, resolved, order_type, order_status, tax_rate, shipping_fee, notes, internal_notes)
        console.print(json.dumps(_humanize_payload(payload), ensure_ascii=False, indent=2))

        if args.dry_run:
            console.print("[yellow]★ DRY-RUN，未實際送出[/yellow]")
            success_count += 1
        else:
            try:
                result = ragic_post(SALES_ORDER_SHEET, payload)
                if result.get("status") == "SUCCESS" or result.get("ragicId"):
                    ragic_id = result.get("ragicId", "")
                    console.print(f"[green]✓ 訂單建立成功！Ragic ID: {ragic_id}[/green]")
                    success_count += 1
                    upload_log[log_key] = {
                        "ragic_id":    str(ragic_id),
                        "uploaded_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
                        "file":        excel_path.name,
                    }
                    _save_upload_log(upload_log)
                else:
                    console.print(f"[red]✗ Ragic 回傳異常：{result}[/red]")
            except Exception as e:
                console.print(f"[red]✗ 送出失敗：{e}[/red]")

    if success_count > 0 and not args.dry_run:
        done_dir = excel_path.parent / "done"
        done_dir.mkdir(exist_ok=True)
        dest = done_dir / excel_path.name
        shutil.move(str(excel_path), str(dest))
        console.print(f"[green]✓ 已移至 {dest.parent.name}/done/{dest.name}[/green]")

    return success_count, len(orders), False


# ── 主程式 ───────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Ragic 銷貨單自動化上傳")
    parser.add_argument("excel", nargs="?", default=None,
        help="採購單路徑（省略時自動掃描 client_order/ 下所有待處理檔案）")
    parser.add_argument("--dry-run", action="store_true", help="預覽模式，不實際送出 Ragic")
    parser.add_argument("--reset-key", action="store_true", help="重設 Ragic API Key")
    args = parser.parse_args()

    if args.reset_key:
        if _KEY_FILE.exists():
            _KEY_FILE.unlink()
            console.print("[yellow]已清除舊的 API Key[/yellow]")
        _auth_header()  # 觸發重新輸入並儲存
        return

    # 指定單一檔案模式（命令列傳入路徑）
    if args.excel:
        excel_path = Path(args.excel).expanduser().resolve()
        if not excel_path.exists():
            console.print(f"[red]找不到檔案：{excel_path}[/red]")
            sys.exit(1)
        if args.dry_run:
            console.print("[yellow bold]★ DRY-RUN 模式：不會實際送出，也不會移動檔案[/yellow bold]")
        price_index = load_price_index()
        customers   = load_customers()
        s, o, _ = process_file(excel_path, args, price_index, customers)
        console.print(f"\n[bold green]完成！成功處理 {s}/{o} 張訂單[/bold green]")
        return

    # 互動選單模式：完成一個→回到選單，直到使用者不再選擇
    # 詢問 dry-run（只問一次）
    if not args.dry_run:
        mode_input = questionary.text(
            "按 Enter 開始正式執行（輸入 debug 進入測試模式）：",
            default="",
        ).ask() or ""
        if mode_input.strip().lower() == "debug":
            args.dry_run = True
    if args.dry_run:
        console.print("[yellow bold]★ DRY-RUN 模式：不會實際送出，也不會移動檔案[/yellow bold]")

    # 載入 Ragic 快取（一次性）
    price_index = load_price_index()
    customers   = load_customers()

    total_success = total_orders = 0
    while True:
        all_files = find_pending_files(BASE_CLIENT_ORDER)
        if not all_files:
            console.print("[yellow]沒有待處理的 Excel 檔案了[/yellow]")
            break

        labels = [f"{f.parent.name}/{f.name}" for f in all_files]
        selected = questionary.checkbox(
            "請選擇要處理的採購單（空白鍵勾選，Enter 確認；不選直接 Enter 結束）：",
            choices=[questionary.Choice(label, checked=False) for label in labels],
        ).ask()
        if not selected:
            console.print("[yellow]結束[/yellow]")
            break

        # 每次只處理第一個選中的檔案，處理完回到選單重新掃描
        excel_path = all_files[labels.index(selected[0])]
        s, o, _ = process_file(excel_path, args, price_index, customers)
        total_success += s
        total_orders  += o

        console.print(f"\n[bold cyan]{'─'*58}[/bold cyan]")

    console.print(f"\n[bold green]完成！成功處理 {total_success}/{total_orders} 張訂單[/bold green]")


if __name__ == "__main__":
    main()
