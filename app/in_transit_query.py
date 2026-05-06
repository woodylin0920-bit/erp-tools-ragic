"""在途查詢：從採購單彙總「尚未進貨數量 > 0」的商品，避免重複下單。

從主選單呼叫 query()；CLI 單獨執行：
    python -m app.in_transit_query                # 列出全部在途商品
    python -m app.in_transit_query BBB042         # 篩選關鍵字（商品編號或名稱）
"""
from __future__ import annotations

import sys
from collections import defaultdict

try:
    from app.ragic_upload import console, ragic_get
except ImportError:
    from ragic_upload import console, ragic_get

PURCHASING_SHEET = "ragicpurchasing/20003"
SUBTABLE_KEY = "_subtable_3001286"


def _to_int(v) -> int:
    try:
        return int(float(v or 0))
    except (TypeError, ValueError):
        return 0


def collect_in_transit() -> list:
    """回傳採購單清單，每張單一個 dict：{po_no, vendor, date, items: [...]}。"""
    records = ragic_get(PURCHASING_SHEET)
    pos: list[dict] = []
    for rec in records.values():
        po_no   = str(rec.get("採購單號", "")).strip()
        vendor  = str(rec.get("廠商名稱", "")).strip()
        po_date = str(rec.get("採購日期", "")).strip()
        sub = rec.get(SUBTABLE_KEY, {}) or {}
        items = []
        for row in sub.values():
            remain = _to_int(row.get("尚未進貨數量", 0))
            if remain <= 0:
                continue
            prod = str(row.get("商品編號*", "") or row.get("商品編號", "")).strip()
            if not prod:
                continue
            price = _to_int(row.get("單價", 0))
            items.append({
                "prod": prod,
                "name": str(row.get("商品名稱", "")).strip(),
                "spec": str(row.get("規格", "")).strip(),
                "unit": str(row.get("單位*", "") or row.get("單位", "")).strip(),
                "qty": remain,
                "price": price,
                "amount": remain * price,
            })
        if items:
            pos.append({
                "po_no": po_no,
                "vendor": vendor,
                "date": po_date,
                "items": items,
            })
    pos.sort(key=lambda p: p["date"], reverse=True)
    return pos


def query(keyword: str | None = None) -> None:
    """以「一張採購單一個分類」呈現在途資料。
    keyword：可為商品編號、商品名稱、採購單號、廠商名稱片段（留空=全部）。"""
    from rich.table import Table

    console.print("[#B0A898]載入採購單資料...[/#B0A898]")
    pos = collect_in_transit()
    if not pos:
        console.print("[#5A9A4A]✓ 目前沒有任何在途採購單[/#5A9A4A]")
        return

    kw = (keyword or "").strip().upper()
    if kw:
        filtered_pos = []
        for po in pos:
            po_match = (kw in po["po_no"].upper() or kw in po["vendor"].upper())
            matched_items = [
                it for it in po["items"]
                if po_match or kw in it["prod"].upper() or kw in it["name"].upper()
            ]
            if matched_items:
                filtered_pos.append({**po, "items": matched_items})
    else:
        filtered_pos = pos

    if not filtered_pos:
        console.print(f"[#FF7700]找不到符合「{keyword}」的在途資料[/#FF7700]")
        return

    grand_qty = grand_amount = 0
    for po in filtered_pos:
        po_qty = sum(it["qty"] for it in po["items"])
        po_amount = sum(it["amount"] for it in po["items"])
        grand_qty += po_qty
        grand_amount += po_amount

        console.print(
            f"\n[bold #C5A059]▸ {po['po_no']}[/bold #C5A059]  "
            f"[#B0A898]{po['vendor']}  ({po['date']})[/#B0A898]  "
            f"[dim]{len(po['items'])} 項 / {po_qty:,} 件 / 合計 {po_amount:,}[/dim]"
        )

        table = Table(show_header=True, header_style="bold #C5A059")
        table.add_column("商品編號", style="cyan")
        table.add_column("商品名稱")
        table.add_column("規格", justify="right")
        table.add_column("單位")
        table.add_column("在途量", justify="right", style="bold")
        table.add_column("單價", justify="right")
        table.add_column("在途金額", justify="right", style="bold #C5A059")
        table.add_column("備註", style="dim")
        for it in po["items"]:
            note = "樣品" if it["price"] == 0 else ""
            table.add_row(
                it["prod"], it["name"][:32], str(it["spec"]), it["unit"],
                f"{it['qty']:,}", f"{it['price']:,}", f"{it['amount']:,}", note,
            )
        console.print(table)

    console.print(
        f"\n[#5A9A4A]✓ 共 {len(filtered_pos)} 張採購單在途，"
        f"總件數 {grand_qty:,}，在途金額合計 {grand_amount:,}[/#5A9A4A]"
    )


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    query(arg)
