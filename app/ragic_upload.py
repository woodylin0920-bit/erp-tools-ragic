#!/usr/bin/env python3
"""
Ragic 銷貨單自動化上傳腳本
用法：
  python ragic_upload.py client/LE/0324T221.xlsx (檔案名稱)
  python ragic_upload.py --dry-run client/LE/0324T221.xlsx
"""

import argparse
import hashlib
import json
import logging
import os
import shutil
import sys
import time
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Optional

from dotenv import load_dotenv
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

import questionary
import requests

QSTYLE = questionary.Style([
    ("question",                        "bold #D4C9B0"),
    ("answer",                          "fg:#5A9A4A bold"),
    ("pointer",                         "fg:#FF7700 bold"),
    ("highlighted",                     "bg:#C5A059 fg:#1A1A1A bold"),
    ("text",                            "fg:#D4C9B0"),
    ("instruction",                     "fg:#666666"),
    ("checkbox",                        "fg:#C5A059"),
    ("checkbox-selected",               "fg:#FF7700 bold"),
    # autocomplete 下拉選單
    ("completion-menu.completion",          "bg:#2a2a2a fg:#D4C9B0"),
    ("completion-menu.completion.current",  "bg:#C5A059 fg:#1A1A1A bold"),
])

# 全域套用 QSTYLE，所有 questionary 呼叫自動帶入樣式
def _q_styled(fn):
    def wrapper(*args, **kwargs):
        kwargs.setdefault("style", QSTYLE)
        return fn(*args, **kwargs)
    return wrapper

questionary.select      = _q_styled(questionary.select)
questionary.checkbox    = _q_styled(questionary.checkbox)
questionary.confirm     = _q_styled(questionary.confirm)
questionary.text        = _q_styled(questionary.text)
questionary.password    = _q_styled(questionary.password)
questionary.autocomplete = _q_styled(questionary.autocomplete)
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.rule import Rule
from rich.text import Text

# ============================================================
# ★ CONFIG ★
# ============================================================

RAGIC_BASE    = os.getenv("RAGIC_BASE",    "https://ap12.ragic.com")
RAGIC_ACCOUNT = os.getenv("RAGIC_ACCOUNT", "toybebop")

PRODUCT_PRICE_SHEET  = os.getenv("PRODUCT_PRICE_SHEET",  "ragicsales-order-management/20006")  # 商品單價管理
CUSTOMER_SHEET       = os.getenv("CUSTOMER_SHEET",       "ragicsales-order-management/20004")  # 客戶
SALES_ORDER_SHEET    = os.getenv("SALES_ORDER_SHEET",    "ragicsales-order-management/20001")  # 銷貨單
DELIVERY_ORDER_SHEET = os.getenv("DELIVERY_ORDER_SHEET", "ragicsales-order-management/20002")  # 出貨單
OUTBOUND_ORDER_SHEET = os.getenv("OUTBOUND_ORDER_SHEET", "ragicinventory/20009")               # 出庫單
INVENTORY_SHEET      = os.getenv("INVENTORY_SHEET",      "ragicinventory/20008")               # 倉庫庫存

ORDER_ITEMS_SUBTABLE_KEY    = os.getenv("ORDER_ITEMS_SUBTABLE_KEY",    "_subtable_3000842")  # 銷貨單訂購項目子表
OUTBOUND_ITEMS_SUBTABLE_KEY = os.getenv("OUTBOUND_ITEMS_SUBTABLE_KEY", "_subtable_3001132")  # 出庫單項目子表

# 客戶尚未建檔時使用的預留客戶
UNREGISTERED_CUSTOMER = {"code": "C-00000", "name": "000尚未建檔", "address": ""}

# 上傳記錄檔（防重複）
_UPLOAD_LOG = Path(__file__).resolve().parent.parent / "upload_log.json"

# 操作日誌資料夾
_LOGS_DIR = Path(__file__).resolve().parent.parent / "logs"


def _load_upload_log() -> dict:
    if _UPLOAD_LOG.exists():
        try:
            return json.loads(_UPLOAD_LOG.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_upload_log(log: dict):
    _UPLOAD_LOG.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")


def _setup_logging():
    _LOGS_DIR.mkdir(exist_ok=True)
    log_file = _LOGS_DIR / f"{date.today()}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")],
    )

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
        console.print("[#FF7700]尚未設定 RAGIC_API_KEY[/#FF7700]")
        api_key = questionary.password("請輸入 Ragic API Key：").ask() or ""
        if not api_key:
            console.print("[red]未輸入 API Key，結束[/red]")
            sys.exit(1)
        _KEY_FILE.write_text(api_key, encoding="utf-8")
        os.environ["RAGIC_API_KEY"] = api_key
        console.print(f"[#5A9A4A]✓ API Key 已儲存至 {_KEY_FILE}，下次不需再輸入[/#5A9A4A]")
    return {"Authorization": f"Basic {api_key}"}


def _ragic_request(method: str, url: str, **kwargs) -> requests.Response:
    """帶自動重試的 HTTP 請求（最多 3 次，指數退避 1s/2s/4s）。"""
    retryable_errors = (requests.exceptions.ConnectionError, requests.exceptions.Timeout)
    last_exc = None
    for attempt in range(3):
        try:
            r = requests.request(method, url, **kwargs)
            if r.status_code >= 500 and attempt < 2:
                wait = 2 ** attempt
                console.print(f"[#FF7700]⚠ 伺服器錯誤（{r.status_code}），{wait} 秒後重試...[/#FF7700]")
                logging.warning("HTTP %s on %s, retrying in %ss (attempt %d)", r.status_code, url, wait, attempt + 1)
                time.sleep(wait)
                continue
            r.raise_for_status()
            return r
        except retryable_errors as e:
            last_exc = e
            if attempt < 2:
                wait = 2 ** attempt
                console.print(f"[#FF7700]⚠ 網路錯誤，{wait} 秒後重試...[/#FF7700]")
                logging.warning("Network error on %s: %s, retrying in %ss (attempt %d)", url, e, wait, attempt + 1)
                time.sleep(wait)
    raise last_exc


def ragic_get(sheet_path: str, limit: int = 2000) -> dict:
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}?api&limit={limit}"
    r = _ragic_request("GET", url, headers=_auth_header(), timeout=30)
    data = r.json()
    return {k: v for k, v in data.items() if not k.startswith("_") and k != "info"}


def ragic_post(sheet_path: str, payload: dict) -> dict:
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}?api&doLinkLoad=true"
    r = _ragic_request("POST", url, headers=_auth_header(), json=payload, timeout=30)
    return r.json()


def ragic_patch(sheet_path: str, record_id: str, payload: dict) -> dict:
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}/{record_id}?api&doLinkLoad=true"
    r = _ragic_request("PATCH", url, headers=_auth_header(), json=payload, timeout=30)
    return r.json()


def ragic_get_action_button_id(sheet_path: str, button_name: str) -> Optional[int]:
    """從 Ragic metadata 取得指定名稱的 action button ID（massOperation 類別）。"""
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}/metadata/actionButton?api&category=massOperation"
    r = _ragic_request("GET", url, headers=_auth_header(), timeout=30)
    for btn in r.json().get("actionButtons", []):
        if btn.get("name") == button_name:
            return btn["id"]
    return None


def ragic_trigger_button(sheet_path: str, record_id: str, button_id) -> dict:
    """對單筆記錄觸發 Ragic action button。"""
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet_path}/{record_id}?api&bId={button_id}"
    r = _ragic_request("POST", url, headers=_auth_header(), timeout=60)
    return r.json()


# ── 快取載入 ─────────────────────────────────────────────────

def load_price_index() -> Dict[str, list]:
    """載入商品單價管理，建立 {條碼: [商品...]} 索引。"""
    console.print("[#B0A898]載入商品單價資料（Ragic API）...[/#B0A898]")
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
    console.print(f"[#5A9A4A]✓ 載入 {len(index)} 種條碼的商品[/#5A9A4A]")
    return index


def load_customers() -> list:
    """載入客戶資料表。"""
    console.print("[#B0A898]載入客戶資料（Ragic API）...[/#B0A898]")
    records = ragic_get(CUSTOMER_SHEET)
    customers = []
    for rec in records.values():
        customers.append({
            "code":    str(rec.get("客戶編號", "")),
            "name":    str(rec.get("客戶名稱", "")),
            "address": str(rec.get("送貨完整地址", "")),
        })
    console.print(f"[#5A9A4A]✓ 載入 {len(customers)} 筆客戶[/#5A9A4A]")
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
    console.print(f"[#FF7700]⚠ 找不到含「{store_code}」的客戶[/#FF7700]")
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
            console.print(f"[#FF7700]⚠ 條碼 {item.barcode}（{item.le_name}）不在商品單價表，已跳過[/#FF7700]")
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
            "amount":       float((Decimal(str(override_price if override_price else product["price"])) * Decimal(str(final_qty))).quantize(Decimal("0.01"), ROUND_HALF_UP)),
        })
    return resolved


# ── 互動 UI ──────────────────────────────────────────────────

def show_items_table(customer: dict, store_code: str, po_number: str, resolved: list):
    console.print()
    console.rule(f"[bold]門市: {store_code}  PO: {po_number}  客戶: {customer['code']} {customer['name']}[/bold]")
    table = Table(show_header=True, header_style="bold #C5A059", box=None)
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

    commission = questionary.select(
        "業務分潤",
        choices=["8%", "2%", "（無）"],
        default="（無）",
    ).ask()
    commission = "" if commission == "（無）" else commission

    notes = questionary.text("備註（留空直接按 Enter）", default="").ask()
    internal_notes = questionary.text("內部備注（留空直接按 Enter）", default="").ask()
    return order_type, order_status, tax_rate, shipping_fee, notes or "", internal_notes or "", commission


def show_confirmation(customer: dict, resolved: list, order_type: str, order_status: str,
                      tax_rate: str, shipping_fee: float, notes: str, internal_notes: str,
                      commission: str = "") -> tuple:
    subtotal = sum(Decimal(str(it["amount"])) for it in resolved)
    tax_amount = (subtotal * Decimal("0.05")).quantize(Decimal("0.01"), ROUND_HALF_UP) if tax_rate == "5%" else Decimal("0")
    total = subtotal + tax_amount + Decimal(str(shipping_fee))

    console.print()
    console.rule("[bold red]最終確認[/bold red]")
    console.print(f"訂單單別: [#B0A898]{order_type}[/#B0A898]  狀態: [#B0A898]{order_status}[/#B0A898]  客戶: [#B0A898]{customer['code']}  {customer['name']}[/#B0A898]")
    console.print(f"課稅別: 營業稅              稅率: [#B0A898]{tax_rate}[/#B0A898]")
    console.print(f"小計: {subtotal:>12,.2f}     稅額: {tax_amount:,.2f}")
    console.print(f"運費: {shipping_fee:>12.2f}     總計: [bold]{total:,.2f}[/bold]")
    if commission:
        console.print(f"業務分潤: [#B0A898]{commission}[/#B0A898]")
    if notes:
        console.print(f"備註: {notes}")
    if internal_notes:
        console.print(f"內部備注: [dim]{internal_notes}[/dim]")
    console.rule()
    return subtotal, tax_amount, total


# ── Payload 組裝 ─────────────────────────────────────────────

def build_payload(customer: dict, resolved: list, order_type: str, order_status: str,
                  tax_rate: str, shipping_fee: float, notes: str, internal_notes: str,
                  commission: str = "") -> dict:
    now   = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    today = date.today().strftime("%Y/%m/%d")

    # 計算各項金額
    subtable = {}
    subtotal = Decimal("0")
    total_items = len(resolved)
    for i, it in enumerate(resolved):
        amount = (Decimal(str(it["unit_price"])) * Decimal(str(it["quantity"]))).quantize(Decimal("0.01"), ROUND_HALF_UP)
        subtotal += amount
        subtable[str(-(total_items - i))] = {
            "3000829": i + 1,                # 項次
            "3000830": it["product_code"],   # 商品販售代號
            "3000832": it["unit_price"],      # 單價
            "3000833": it["quantity"],        # 數量
            "3000834": float(amount),        # 金額（單價×數量）
        }

    tax_amount  = (subtotal * Decimal("0.05")).quantize(Decimal("0.01"), ROUND_HALF_UP) if tax_rate == "5%" else Decimal("0")
    total       = subtotal + tax_amount + Decimal(str(shipping_fee))

    return {
        "3000812": order_type,               # 訂單單別
        "3000813": today,                    # 訂單日期
        "3000814": order_status,             # 訂單狀態
        "3000815": customer["code"],         # 客戶編號
        "3000836": "營業稅",                 # 課稅別
        "3000838": tax_rate,                 # 稅率
        "3001498": int(shipping_fee),        # 訂單運費
        "3001684": "DDP",                    # 國貿條規（預設）
        "3000835": float(subtotal),          # 小計
        "3000837": float(tax_amount),        # 稅額
        "3000839": float(total),             # 總金額(含稅)
        "3000840": notes,                    # 備註
        "1000065": commission,               # 業務分潤
        "1000074": f"【AI建單】 {internal_notes}".strip() if internal_notes else "【AI建單】",  # 內部備注
        "3000845": now,                      # 建檔日期時間
        "3000847": now,                      # 最後修改日期時間
        ORDER_ITEMS_SUBTABLE_KEY: subtable,
    }


BASE_CLIENT_ORDER = Path(__file__).resolve().parent.parent / "client_order"
BASE_TEMPLATES    = Path(__file__).resolve().parent.parent / "templates"
BASE_OUTPUT       = Path(__file__).resolve().parent.parent / "exports"


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

    console.print(f"\n[#B0A898]解析 {excel_path.name}（{client_code} 格式）...[/#B0A898]")
    try:
        orders = PARSERS[client_code](str(excel_path)).parse()
    except Exception as e:
        console.print(f"[red]無法讀取 Excel 檔案：{e}[/red]")
        return 0, 0
    if not orders:
        console.print("[red]無法解析任何訂單，請確認檔案格式[/red]")
        return 0, 0
    console.print(f"[#5A9A4A]✓ 偵測到 {len(orders)} 張訂單[/#5A9A4A]")

    upload_log = _load_upload_log()
    file_hash = hashlib.md5(excel_path.read_bytes()).hexdigest()
    success_count = 0
    for i, order in enumerate(orders, 1):
        console.print(f"\n{'═'*58}")
        console.print(f"[bold]訂單 {i}/{len(orders)}  門市: {order.store_code}  PO: {order.po_number}[/bold]")

        # 防重複：以 PO 為單位判斷是否已上傳過
        log_key = f"{client_code}_{order.store_code}_{order.po_number}"
        if log_key in upload_log and not args.dry_run:
            rec = upload_log[log_key]
            console.print(f"[#FF7700]⚠ 此訂單已於 {rec['uploaded_at']} 上傳（Ragic ID: {rec['ragic_id']}）[/#FF7700]")
            logging.info("重複跳過 log_key=%s ragic_id=%s", log_key, rec['ragic_id'])
            skip = questionary.confirm("是否跳過（建議跳過以避免重複）？", default=True).ask()
            if skip:
                console.print("[#FF7700]已跳過[/#FF7700]")
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
        order_type, order_status, tax_rate, shipping_fee, notes, internal_notes, commission = ask_order_options(is_unregistered)

        show_confirmation(customer, resolved, order_type, order_status, tax_rate, shipping_fee, notes, internal_notes, commission)

        confirmed = questionary.confirm("確認送出此訂單？", default=True).ask()
        if not confirmed:
            action = questionary.select(
                "請選擇：",
                choices=["跳過此單，繼續下一張", "放棄整個檔案，回到選單（不移至 done）"],
            ).ask()
            if "放棄" in action:
                if success_count > 0:
                    console.print(f"[#FF7700]已放棄。前 {success_count} 張已送出至 Ragic，請自行確認是否需要刪除。[/#FF7700]")
                else:
                    console.print("[#FF7700]已放棄，未送出任何訂單。[/#FF7700]")
                return success_count, len(orders), True
            console.print("[#FF7700]已跳過[/#FF7700]")
            continue

        payload = build_payload(customer, resolved, order_type, order_status, tax_rate, shipping_fee, notes, internal_notes, commission)
        console.print(json.dumps(_humanize_payload(payload), ensure_ascii=False, indent=2))

        if args.dry_run:
            console.print("[#FF7700]★ DRY-RUN，未實際送出[/#FF7700]")
            success_count += 1
        else:
            try:
                result = ragic_post(SALES_ORDER_SHEET, payload)
                if result.get("status") == "SUCCESS" or result.get("ragicId"):
                    ragic_id = result.get("ragicId", "")
                    console.print(f"[#5A9A4A]✓ 訂單建立成功！Ragic ID: {ragic_id}[/#5A9A4A]")
                    logging.info("銷貨單建立成功 ragic_id=%s file=%s log_key=%s", ragic_id, excel_path.name, log_key)
                    success_count += 1
                    upload_log[log_key] = {
                        "ragic_id":    str(ragic_id),
                        "uploaded_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
                        "file":        excel_path.name,
                        "file_hash":   file_hash,
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
        console.print(f"[#5A9A4A]✓ 已移至 {dest.parent.name}/done/{dest.name}[/#5A9A4A]")
        logging.info("檔案移至 done: %s", dest)

    return success_count, len(orders), False


# ── 主選單流程 ───────────────────────────────────────────────

def run_new_sales_order(args, price_index: dict, customers: list):
    """新建銷售單（原 main while 迴圈，沒有 xlsx 時回主選單）。"""
    total_success = total_orders = 0
    while True:
        all_files = find_pending_files(BASE_CLIENT_ORDER)
        if not all_files:
            console.print("[#FF7700]沒有待處理的 Excel 檔案了，返回主選單[/#FF7700]")
            break

        labels = [f"{f.parent.name}/{f.name}" for f in all_files]
        selected = questionary.checkbox(
            "請選擇要處理的採購單（空白鍵勾選，Enter 確認；不選直接 Enter 返回）：",
            choices=[questionary.Choice(label, checked=False) for label in labels],
        ).ask()
        if not selected:
            console.print("[#FF7700]返回主選單[/#FF7700]")
            break

        excel_path = all_files[labels.index(selected[0])]

        console.print(f"[#B0A898]── 即將處理：{excel_path.parent.name}/{excel_path.name} ──[/#B0A898]")
        ok = questionary.confirm("確認執行？", default=True).ask()
        if not ok:
            continue

        s, o, _ = process_file(excel_path, args, price_index, customers)
        total_success += s
        total_orders  += o
        console.print(f"\n[bold cyan]{'─'*58}[/bold cyan]")

    if total_orders > 0:
        console.print(f"[bold #5A9A4A]本次共處理 {total_success}/{total_orders} 張訂單[/bold #5A9A4A]")


def run_create_delivery_order(args):
    """銷貨單批量拋轉建立出貨單（訂單狀態：未出貨 / 預接單 / 已收款未出貨）。"""
    TARGET_STATUSES = {"未出貨", "預接單", "已收款未出貨"}

    console.print("[#B0A898]載入銷貨單資料（Ragic API）...[/#B0A898]")
    records = ragic_get(SALES_ORDER_SHEET)

    candidates = []
    for rid, rec in records.items():
        status = str(rec.get("訂單狀態", ""))
        if status in TARGET_STATUSES:
            candidates.append({
                "id":    rid,
                "label": f"{rec.get('訂單編號','?')}  {rec.get('客戶名稱','?')}  {rec.get('訂單日期','?')}  [{status}]",
            })

    if not candidates:
        console.print("[#FF7700]沒有待拋轉的銷貨單（未出貨 / 預接單 / 已收款未出貨）[/#FF7700]")
        return

    console.print(f"[#5A9A4A]✓ 找到 {len(candidates)} 筆待拋轉銷貨單[/#5A9A4A]")

    record_ids = None
    while True:
        selected = questionary.checkbox(
            "請選擇要建立出貨單的銷貨單（空白鍵勾選，Enter 確認）：",
            choices=[questionary.Choice(c["label"], checked=False) for c in candidates],
        ).ask()
        if not selected:
            console.print("[#FF7700]返回主選單[/#FF7700]")
            return

        record_ids = [c["id"] for c in candidates if c["label"] in selected]

        console.print("[#B0A898]── 即將執行：建立出貨單 ──[/#B0A898]")
        for label in selected:
            console.print(f"  {label}")
        ok = questionary.confirm("確認執行？", default=True).ask()
        if ok:
            break

    console.print("[#B0A898]取得 Ragic 按鈕設定...[/#B0A898]")
    button_id = ragic_get_action_button_id(SALES_ORDER_SHEET, "建立出貨單")
    if button_id is None:
        console.print("[red]找不到「建立出貨單」按鈕，請確認 Ragic 表單設定[/red]")
        return

    if args.dry_run:
        console.print(f"[#FF7700]★ DRY-RUN：buttonId={button_id}，對象 {record_ids}[/#FF7700]")
        return

    success = 0
    for rid in record_ids:
        try:
            result = ragic_trigger_button(SALES_ORDER_SHEET, rid, button_id)
            if result.get("status") == "SUCCESS":
                urls = result.get("urls", [])
                console.print(f"[#5A9A4A]✓ {rid} 拋轉成功[/#5A9A4A]" + (f"  → {urls[0]}" if urls else ""))
                logging.info("出貨單建立成功 sales_id=%s", rid)
                success += 1
            else:
                console.print(f"[red]✗ {rid} 拋轉失敗：{result.get('msg', result)}[/red]")
                logging.warning("出貨單建立失敗 sales_id=%s msg=%s", rid, result.get('msg', result))
        except Exception as e:
            console.print(f"[red]✗ {rid} 發生錯誤：{e}[/red]")
            logging.error("出貨單建立錯誤 sales_id=%s error=%s", rid, e)
    console.print(f"[bold #5A9A4A]完成！{success}/{len(record_ids)} 筆出貨單建立成功[/bold #5A9A4A]")
    console.print("[dim]請至 Ragic 出貨單頁面確認[/dim]")


def run_create_outbound_order(args):
    """出貨單拋轉建立出庫單，並自動補填子表的倉庫代碼和庫存編號。"""
    # 載入出貨單
    console.print("[#B0A898]載入出貨單資料（Ragic API）...[/#B0A898]")
    records = ragic_get(DELIVERY_ORDER_SHEET)

    candidates = []
    for rid, rec in records.items():
        candidates.append({
            "id":    rid,
            "label": f"{rec.get('出貨單號','?')}  {rec.get('客戶名稱','?')}  {rec.get('訂單日期','?')}",
        })

    if not candidates:
        console.print("[#FF7700]沒有出貨單資料[/#FF7700]")
        return

    console.print(f"[#5A9A4A]✓ 找到 {len(candidates)} 筆出貨單[/#5A9A4A]")

    # 載入倉庫庫存（一次性，不隨步驟重複）
    console.print("[#B0A898]載入倉庫庫存資料（Ragic API）...[/#B0A898]")
    inventory = ragic_get(INVENTORY_SHEET)

    warehouses: dict = {}
    inv_by_wh_prod: Dict[tuple, list] = {}
    for rec in inventory.values():
        wh_code  = str(rec.get("倉庫代碼", "")).strip()
        wh_name  = str(rec.get("倉庫名稱", "")).strip()
        prod     = str(rec.get("商品編號", "")).strip()
        inv_code = str(rec.get("庫存編號", "")).strip()
        if wh_code:
            warehouses[wh_code] = wh_name
        if wh_code and prod and inv_code:
            inv_by_wh_prod.setdefault((wh_code, prod), []).append(inv_code)

    if not warehouses:
        console.print("[red]無法載入倉庫資料[/red]")
        return

    DEFAULT_WH = "TW01"
    BACK = "← 返回"
    sorted_wh = sorted(warehouses.items(), key=lambda x: (0 if x[0] == DEFAULT_WH else 1, x[0]))
    wh_choices = [f"{code}  {name}" for code, name in sorted_wh]
    DELIVERY_SUBTABLE = "_subtable_3000886"

    step = 1
    selected_records = record_ids = None
    warehouse_code = warehouse_name = None
    prod_inv_map = None

    while True:
        if step == 1:
            selected = questionary.checkbox(
                "請選擇要建立出庫單的出貨單（空白鍵勾選，Enter 確認）：",
                choices=[questionary.Choice(c["label"], checked=False) for c in candidates],
            ).ask()
            if not selected:
                console.print("[#FF7700]返回主選單[/#FF7700]")
                return
            selected_records = [c for c in candidates if c["label"] in selected]
            record_ids = [c["id"] for c in selected_records]
            step = 2

        elif step == 2:
            wh_sel = questionary.select("請選擇倉庫：", choices=[BACK] + wh_choices).ask()
            if not wh_sel or wh_sel == BACK:
                step = 1
                continue
            warehouse_code = wh_sel.split("  ")[0].strip()
            warehouse_name = warehouses.get(warehouse_code, "")
            step = 3

        elif step == 3:
            products_needed: list = []
            seen_prods: set = set()
            for c in selected_records:
                sub = records[c["id"]].get(DELIVERY_SUBTABLE, {})
                for row in sub.values():
                    prod = str(row.get("商品編號*", "") or row.get("商品編號", "")).strip()
                    if prod and prod not in seen_prods:
                        seen_prods.add(prod)
                        products_needed.append({"prod": prod, "name": str(row.get("商品名稱", "")).strip()})

            prod_inv_map = {}
            cancelled = False
            for item in products_needed:
                prod = item["prod"]
                options = inv_by_wh_prod.get((warehouse_code, prod), [])
                if not options:
                    console.print(f"[#FF7700]⚠ {prod} 在 {warehouse_code} 無庫存紀錄，跳過[/#FF7700]")
                    prod_inv_map[prod] = ""
                    continue
                if len(options) == 1:
                    prod_inv_map[prod] = options[0]
                    console.print(f"[dim]{prod} {item['name']} → {options[0]}（唯一選項，自動帶入）[/dim]")
                else:
                    inv_sel = questionary.select(
                        f"請選擇 {prod} {item['name']} 的庫存編號：",
                        choices=[BACK] + options,
                    ).ask()
                    if not inv_sel or inv_sel == BACK:
                        cancelled = True
                        break
                    prod_inv_map[prod] = inv_sel
            if cancelled:
                step = 2
                continue
            step = 4

        elif step == 4:
            console.print("[#B0A898]── 即將執行：建立出庫單 ──[/#B0A898]")
            for c in selected_records:
                console.print(f"  {c['label']}")
            console.print(f"  倉庫：{warehouse_code}  {warehouse_name}")
            for prod, inv in prod_inv_map.items():
                if inv:
                    console.print(f"  {prod} → {inv}")
            ok = questionary.confirm("確認執行？", default=True).ask()
            if not ok:
                step = 1
                continue
            break

    console.print("[#B0A898]取得 Ragic 按鈕設定...[/#B0A898]")
    button_id = ragic_get_action_button_id(DELIVERY_ORDER_SHEET, "建立出庫單")
    if button_id is None:
        console.print("[red]找不到「建立出庫單」按鈕，請確認 Ragic 表單設定[/red]")
        return

    if args.dry_run:
        console.print(f"[#FF7700]★ DRY-RUN：buttonId={button_id}，倉庫={warehouse_code}，對象 {record_ids}[/#FF7700]")
        return

    # 記錄觸發前的出庫單 ID
    console.print("[#B0A898]記錄現有出庫單...[/#B0A898]")
    before_ids = set(ragic_get(OUTBOUND_ORDER_SHEET).keys())

    console.print(f"[#B0A898]逐筆觸發建立出庫單（{len(record_ids)} 筆）...[/#B0A898]")
    for rid in record_ids:
        try:
            result = ragic_trigger_button(DELIVERY_ORDER_SHEET, rid, button_id)
            if result.get("status") == "SUCCESS":
                console.print(f"[#5A9A4A]✓ {rid} 拋轉成功[/#5A9A4A]")
                logging.info("出庫單觸發成功 delivery_id=%s", rid)
            else:
                console.print(f"[red]✗ {rid} 拋轉失敗：{result.get('msg', result)}[/red]")
                logging.warning("出庫單觸發失敗 delivery_id=%s msg=%s", rid, result.get('msg', result))
        except Exception as e:
            console.print(f"[red]✗ {rid} 發生錯誤：{e}[/red]")
            logging.error("出庫單觸發錯誤 delivery_id=%s error=%s", rid, e)

    console.print("[dim]等待 Ragic 建立出庫單（3 秒）...[/dim]")
    time.sleep(3)

    after_records = ragic_get(OUTBOUND_ORDER_SHEET)
    new_ids = set(after_records.keys()) - before_ids
    if not new_ids:
        console.print("[#FF7700]⚠ 未偵測到新建立的出庫單（可能已被 Ragic 擋掉重複拋轉）[/#FF7700]")
        return

    console.print(f"[#5A9A4A]✓ 偵測到 {len(new_ids)} 筆新出庫單，開始補填倉庫資料...[/#5A9A4A]")

    patched = 0
    for oid in new_ids:
        rec = after_records[oid]
        subtable = rec.get(OUTBOUND_ITEMS_SUBTABLE_KEY, {})
        if not subtable:
            console.print(f"[#FF7700]⚠ 出庫單 {oid} 沒有子表項目，跳過[/#FF7700]")
            continue

        # 填倉庫代碼、庫存編號（用 CID，必填欄位用欄位名稱會被 Ragic validation 擋掉）
        # 注意：倉庫名稱(3001125)為唯讀，填入倉庫代碼後 Ragic 自動帶入，不需手動寫
        updated_rows = {}
        for row_id, row in subtable.items():
            if str(row_id).startswith("_"):
                continue
            prod = str(row.get("商品編號", "")).strip()
            inv_code = prod_inv_map.get(prod, "")
            if not inv_code:
                console.print(f"[#FF7700]⚠ 出庫單 {oid} 商品 {prod} 無庫存編號，該列倉庫欄位略過[/#FF7700]")
                logging.warning("出庫單 %s 商品 %s 無庫存編號，略過", oid, prod)
                continue
            updated_rows[str(row_id)] = {
                "3001124": warehouse_code,  # 倉庫代碼
                "3001126": inv_code,        # 庫存編號
            }

        try:
            ragic_patch(OUTBOUND_ORDER_SHEET, oid, {OUTBOUND_ITEMS_SUBTABLE_KEY: updated_rows})
            patched += 1
            console.print(f"[#5A9A4A]✓ 出庫單 {oid} 補填完成（{warehouse_code}）[/#5A9A4A]")
            logging.info("出庫單補填成功 outbound_id=%s warehouse=%s", oid, warehouse_code)
        except Exception as e:
            console.print(f"[red]⚠ 出庫單 {oid} 補填失敗：{e}[/red]")
            logging.error("出庫單補填失敗 outbound_id=%s error=%s", oid, e)

    console.print(f"[bold #5A9A4A]完成！{patched}/{len(new_ids)} 筆出庫單已補填倉庫資料[/bold #5A9A4A]")
    console.print("[dim]請至 Ragic 出庫單頁面確認[/dim]")


def run_export_inventory(args, price_index: dict):
    """從 Ragic 倉庫庫存匯出 Excel，自動換算 PCS 填入客戶模板的現貨欄位。"""
    import copy
    import openpyxl

    BASE_OUTPUT.mkdir(exist_ok=True)

    # ── 倉庫選擇 ─────────────────────────────────────────────
    console.print("[#B0A898]載入倉庫庫存資料（Ragic API）...[/#B0A898]")
    inventory_all = ragic_get(INVENTORY_SHEET)

    warehouses: dict = {}
    for rec in inventory_all.values():
        wh_code = str(rec.get("倉庫代碼", "")).strip()
        wh_name = str(rec.get("倉庫名稱", "")).strip()
        if wh_code:
            warehouses[wh_code] = wh_name

    if not warehouses:
        console.print("[red]無法載入倉庫資料[/red]")
        return

    DEFAULT_WH = "TW01"
    BACK = "← 返回"
    sorted_wh = sorted(warehouses.items(), key=lambda x: (0 if x[0] == DEFAULT_WH else 1, x[0]))
    wh_choices = [f"{code}  {name}" for code, name in sorted_wh]

    wh_sel = questionary.select("請選擇倉庫：", choices=[BACK] + wh_choices).ask()
    if not wh_sel or wh_sel == BACK:
        console.print("[#FF7700]返回主選單[/#FF7700]")
        return
    warehouse_code = wh_sel.split("  ")[0].strip()
    warehouse_name = warehouses.get(warehouse_code, "")

    # ── 模板選擇 ─────────────────────────────────────────────
    BASE_TEMPLATES.mkdir(exist_ok=True)
    templates = sorted(BASE_TEMPLATES.glob("*.xlsx"), reverse=True)
    if not templates:
        console.print(f"[red]找不到模板，請將 .xlsx 模板放入 {BASE_TEMPLATES}[/red]")
        return

    TPL_DISPLAY = {
        "quote-template.xlsx":     "quote-template.xlsx（報價單）",
        "inventory-template.xlsx": "inventory-template.xlsx（庫存總覽）",
    }
    tpl_map = {TPL_DISPLAY.get(t.name, t.name): t for t in templates}
    selected = questionary.checkbox(
        "請選擇模板（空白鍵勾選，Enter 確認；不選直接 Enter 返回）：",
        choices=[questionary.Choice(label, checked=False) for label in tpl_map],
    ).ask()
    if not selected:
        console.print("[#FF7700]返回主選單[/#FF7700]")
        return

    tpl_path = tpl_map[selected[0]]

    # ── 確認 ─────────────────────────────────────────────────
    console.print(f"[#B0A898]── 即將執行：匯出庫存報表 ──[/#B0A898]")
    console.print(f"  倉庫：{warehouse_code}  {warehouse_name}")
    console.print(f"  模板：{tpl_path.name}")
    ok = questionary.confirm("確認執行？", default=True).ask()
    if not ok:
        return

    # ── 建立 product_code → barcode 反向索引 ─────────────────
    # 商品單價代號格式為 BBB042-1（有尾綴），庫存商品編號為 BBB042（無尾綴）
    # 去掉 -數字 尾綴後建立反向索引
    import re as _re
    code_to_barcode: Dict[str, str] = {}
    for barcode, entries in price_index.items():
        for entry in entries:
            base = _re.sub(r'-\d+$', '', entry["product_code"])
            code_to_barcode[base] = barcode

    # ── 計算各條碼的 PCS（只算 spec > 1 的）────────────────
    # 報客戶的單位（中盒/箱類），其餘（單盒/個/袋等）跳過
    BULK_UNITS = {"中盒", "箱", "整箱", "端盒"}

    inventory_pcs: Dict[str, int] = {}
    skipped_single = 0
    for rec in inventory_all.values():
        if str(rec.get("倉庫代碼", "")).strip() != warehouse_code:
            continue
        unit = str(rec.get("單位", "")).strip()
        if unit not in BULK_UNITS:
            skipped_single += 1
            continue
        prod_code = str(rec.get("商品編號", "")).strip()
        qty_raw = rec.get("數量", 0)
        spec_raw = rec.get("規格", "1")
        try:
            qty = int(float(qty_raw or 0))
        except (ValueError, TypeError):
            qty = 0
        try:
            spec = int(float(spec_raw or 1))
        except (ValueError, TypeError):
            spec = 1

        barcode = code_to_barcode.get(prod_code)
        if not barcode:
            continue

        pcs = qty * spec
        inventory_pcs[barcode] = inventory_pcs.get(barcode, 0) + pcs

    console.print(f"[#5A9A4A]✓ 計算完成：{len(inventory_pcs)} 種條碼有庫存（略過 {skipped_single} 筆單盒項目）[/#5A9A4A]")

    # ── 填入模板 ─────────────────────────────────────────────
    wb = openpyxl.load_workbook(tpl_path)
    ws = wb.active

    # 自動偵測現貨欄位置（從 row 2 或 row 3 找「現貨」）
    inv_col_idx = None
    for check_row in (2, 3):
        for cell in ws[check_row]:
            if str(cell.value or '').strip() == '現貨':
                inv_col_idx = cell.column - 1  # 轉為 0-indexed
                break
        if inv_col_idx is not None:
            break
    if inv_col_idx is None:
        console.print("[red]✗ 此模板找不到「現貨」欄位，無法匯出庫存。請選擇 inventory 或 quote 模板。[/red]")
        return

    filled = 0
    for row in ws.iter_rows(min_row=4):
        d_cell = row[3]  # D 欄 index=3
        if d_cell.value is None:
            continue
        try:
            barcode = str(int(float(d_cell.value)))
        except (ValueError, TypeError):
            continue
        if barcode in inventory_pcs and inv_col_idx < len(row):
            row[inv_col_idx].value = inventory_pcs[barcode]
            filled += 1

    # ── 儲存輸出 ─────────────────────────────────────────────
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    tpl_prefix = tpl_path.stem.replace("-template", "")
    out_path = BASE_OUTPUT / f"{tpl_prefix}_{warehouse_code}_{ts}.xlsx"
    wb.save(out_path)

    console.print(f"[bold #5A9A4A]✓ 完成！填入 {filled} 筆，輸出至：{out_path}[/bold #5A9A4A]")
    logging.info("庫存報表匯出成功 warehouse=%s filled=%d path=%s", warehouse_code, filled, out_path)


# ── 歡迎畫面 ─────────────────────────────────────────────────

def _get_current_user() -> str:
    """取得用戶名稱：優先 Ragic API，失敗則用系統登入名。"""
    try:
        if _KEY_FILE.exists():
            api_key = _KEY_FILE.read_text(encoding="utf-8").strip()
            url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}?api&getUserInfo=true"
            resp = requests.get(url, headers={"Authorization": f"Basic {api_key}"}, timeout=3)
            data = resp.json()
            name = (data.get("name") or data.get("fullName") or
                    data.get("userName") or data.get("user", {}).get("name", ""))
            if name:
                return name
    except Exception:
        pass
    try:
        import os
        return os.getlogin()
    except Exception:
        return ""


def _calc_revenue(data: dict, date_from: str, date_to: str) -> float:
    """加總指定日期範圍內銷貨單總計。"""
    total = 0.0
    for rec in data.values():
        order_date = rec.get("訂單日期", rec.get("日期", ""))
        if not order_date:
            continue
        d = order_date[:10].replace("-", "/")
        if date_from <= d <= date_to:
            val = rec.get("總金額(含稅)", rec.get("小計", rec.get("總計", "0"))) or "0"
            try:
                total += float(str(val).replace(",", ""))
            except ValueError:
                pass
    return total


def _get_revenue_summary() -> list:
    """回傳 [(label, amount_str), ...] 上月、本季、本年，失敗回傳空列表。"""
    try:
        import datetime as dt
        today = date.today()

        # 上個月
        first_this_month = today.replace(day=1)
        lm_end = first_this_month - dt.timedelta(days=1)
        lm_start = lm_end.replace(day=1)

        # 本季
        q = (today.month - 1) // 3
        q_start = date(today.year, q * 3 + 1, 1)
        q_end = today

        # 本年
        y_start = date(today.year, 1, 1)
        y_end = today

        data = ragic_get(SALES_ORDER_SHEET, limit=2000)

        results = []
        lm_total = _calc_revenue(data, lm_start.strftime("%Y/%m/%d"), lm_end.strftime("%Y/%m/%d"))
        if lm_total:
            results.append((f"上月 ({lm_start.strftime('%Y/%m')})", f"NT$ {lm_total:,.0f}"))

        q_total = _calc_revenue(data, q_start.strftime("%Y/%m/%d"), q_end.strftime("%Y/%m/%d"))
        if q_total:
            results.append((f"本季 (Q{q + 1})", f"NT$ {q_total:,.0f}"))

        y_total = _calc_revenue(data, y_start.strftime("%Y/%m/%d"), y_end.strftime("%Y/%m/%d"))
        if y_total:
            results.append((f"本年 ({today.year})", f"NT$ {y_total:,.0f}"))

        return results
    except Exception:
        return []


def _get_recent_activity() -> list:
    """從 upload_log.json 讀取最近操作，回傳 [(日期, 描述), ...] 最多 5 筆。"""
    from collections import defaultdict
    log = _load_upload_log()
    if not log:
        return []
    date_counts: dict = defaultdict(int)
    for v in log.values():
        date = v.get("uploaded_at", "")[:10]
        if date:
            date_counts[date] += 1
    sorted_dates = sorted(date_counts.items(), key=lambda x: x[0], reverse=True)[:5]
    return [(d, f"銷貨單 × {c} 筆") for d, c in sorted_dates]


def _show_welcome():
    """顯示仿 Claude Code 風格的歡迎畫面。"""
    username = _get_current_user()
    welcome_line = f"歡迎回來，{username}！" if username else "歡迎回來！"

    revenue_rows = _get_revenue_summary()

    # 左欄
    left = Table.grid(padding=(0, 2))
    left.add_column()
    left.add_row(Text(welcome_line, style="bold #D4C9B0"))
    left.add_row("")
    left.add_row(Text("Boptoys", style="bold #C5A059"))
    left.add_row(Text("潮玩波普國際有限公司", style="#C5A059"))
    left.add_row(Text("統一編號 82906411", style="dim"))
    if revenue_rows:
        left.add_row("")
        rev_table = Table.grid(padding=(0, 2))
        rev_table.add_column(style="dim", no_wrap=True)
        rev_table.add_column(style="bold #FF7700", no_wrap=True)
        for label, amount in revenue_rows:
            rev_table.add_row(label, amount)
        left.add_row(rev_table)

    # 右欄：最近操作
    activity = _get_recent_activity()
    right = Table.grid(padding=(0, 2))
    right.add_column(style="dim", no_wrap=True)
    right.add_column()
    if activity:
        for d, desc in activity:
            right.add_row(d, desc)
    else:
        right.add_row("尚無操作紀錄", "")

    # 組合成雙欄
    layout = Table.grid(expand=True, padding=(0, 1))
    layout.add_column(ratio=1)
    layout.add_column(ratio=1)
    layout.add_row(left, right)

    console.print(Panel(layout, title="[bold #C5A059]Ragic ERP Tools[/bold #C5A059]", border_style="#C5A059"))
    console.print(Rule(style="#C5A059"))


# ── 主程式 ───────────────────────────────────────────────────

def main():
    _setup_logging()
    parser = argparse.ArgumentParser(description="Ragic 銷貨單自動化上傳")
    parser.add_argument("excel", nargs="?", default=None,
        help="採購單路徑（省略時自動掃描 client_order/ 下所有待處理檔案）")
    parser.add_argument("--dry-run", action="store_true", help="預覽模式，不實際送出 Ragic")
    parser.add_argument("--reset-key", action="store_true", help="重設 Ragic API Key")
    args = parser.parse_args()

    if args.reset_key:
        if _KEY_FILE.exists():
            _KEY_FILE.unlink()
            console.print("[#FF7700]已清除舊的 API Key[/#FF7700]")
        _auth_header()  # 觸發重新輸入並儲存
        return

    # 指定單一檔案模式（命令列傳入路徑）
    if args.excel:
        excel_path = Path(args.excel).expanduser().resolve()
        if not excel_path.exists():
            console.print(f"[red]找不到檔案：{excel_path}[/red]")
            sys.exit(1)
        if args.dry_run:
            console.print("[bold #FF7700]★ DRY-RUN 模式：不會實際送出，也不會移動檔案[/bold #FF7700]")
        price_index = load_price_index()
        customers   = load_customers()
        s, o, _ = process_file(excel_path, args, price_index, customers)
        console.print(f"\n[bold #5A9A4A]完成！成功處理 {s}/{o} 張訂單[/bold #5A9A4A]")
        return

    console.clear()
    _show_welcome()

    # DRY-RUN 提示（頂層，一次即可）
    if not args.dry_run:
        mode_input = questionary.text(
            "按 Enter 開始正式執行（輸入 debug 進入測試模式）：",
            default="",
        ).ask() or ""
        if mode_input.strip().lower() == "debug":
            args.dry_run = True
    if args.dry_run:
        console.print("[bold #FF7700]★ DRY-RUN 模式：不會實際送出，也不會移動檔案[/bold #FF7700]")

    # 快取懶載入（進入新建銷售單時才 API 一次）
    price_index = customers = None

    # ── 主選單 ─────────────────────────────────────────────────
    while True:
        console.print(Rule(style="#C5A059"))
        choice = questionary.select(
            "請選擇功能：",
            choices=[
                "新建銷售單",
                "建立出貨單（銷貨單拋轉）",
                "建立出庫單（出貨單拋轉）",
                "匯出庫存報表（Excel）",
                "新竹物流建單",
                "Agent mode",
                "退出 (Esc)",
            ],
        ).ask()

        if not choice or choice == "退出 (Esc)":
            break
        elif choice == "新建銷售單":
            if price_index is None:
                price_index = load_price_index()
                customers   = load_customers()
            run_new_sales_order(args, price_index, customers)
        elif choice == "建立出貨單（銷貨單拋轉）":
            run_create_delivery_order(args)
        elif choice == "建立出庫單（出貨單拋轉）":
            run_create_outbound_order(args)
        elif choice == "匯出庫存報表（Excel）":
            if price_index is None:
                price_index = load_price_index()
                customers   = load_customers()
            run_export_inventory(args, price_index)
        elif choice == "新竹物流建單":
            console.print("[#FF7700]功能開發中，敬請期待[/#FF7700]")
        elif choice == "Agent mode":
            from ai_assistant import run_agent_mode
            run_agent_mode()

    console.print("[bold #5A9A4A]再見！[/bold #5A9A4A]")


if __name__ == "__main__":
    main()
