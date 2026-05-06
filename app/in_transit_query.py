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


def collect_in_transit() -> dict:
    """回傳 {商品編號: {"name": str, "qty": int, "orders": [(採購單號, 廠商, 在途量), ...]}}"""
    records = ragic_get(PURCHASING_SHEET)
    agg: dict[str, dict] = {}
    for rec in records.values():
        po_no   = str(rec.get("採購單號", "")).strip()
        vendor  = str(rec.get("廠商名稱", "")).strip()
        po_date = str(rec.get("採購日期", "")).strip()
        sub = rec.get(SUBTABLE_KEY, {}) or {}
        for row in sub.values():
            remain = _to_int(row.get("尚未進貨數量", 0))
            if remain <= 0:
                continue
            prod = str(row.get("商品編號*", "") or row.get("商品編號", "")).strip()
            name = str(row.get("商品名稱", "")).strip()
            unit = str(row.get("單位*", "") or row.get("單位", "")).strip()
            spec = str(row.get("規格", "")).strip()
            if not prod:
                continue
            entry = agg.setdefault(prod, {"name": name, "unit": unit, "spec": spec, "qty": 0, "orders": []})
            entry["qty"] += remain
            entry["orders"].append({
                "po_no": po_no,
                "vendor": vendor,
                "date": po_date,
                "remain": remain,
            })
    return agg


def query(keyword: str | None = None) -> None:
    """互動式查詢；keyword 可為 None（列全部）、商品編號或名稱片段。"""
    from rich.table import Table

    console.print("[#B0A898]載入採購單資料...[/#B0A898]")
    in_transit = collect_in_transit()
    if not in_transit:
        console.print("[#5A9A4A]✓ 目前沒有任何在途商品[/#5A9A4A]")
        return

    if keyword:
        kw = keyword.strip().upper()
        filtered = {
            prod: info for prod, info in in_transit.items()
            if kw in prod.upper() or kw in info["name"].upper()
        }
    else:
        filtered = in_transit

    if not filtered:
        console.print(f"[#FF7700]找不到符合「{keyword}」的在途商品[/#FF7700]")
        return

    table = Table(show_header=True, header_style="bold #C5A059")
    table.add_column("商品編號", style="cyan")
    table.add_column("商品名稱")
    table.add_column("規格", justify="right")
    table.add_column("單位")
    table.add_column("在途量", justify="right", style="bold")
    table.add_column("採購單", style="dim")

    for prod in sorted(filtered.keys()):
        info = filtered[prod]
        po_str = ", ".join(
            f"{o['po_no']}({o['remain']})" for o in info["orders"]
        )
        table.add_row(
            prod,
            info["name"][:32],
            str(info["spec"]),
            info["unit"],
            f"{info['qty']:,}",
            po_str,
        )
    console.print(table)
    total_skus = len(filtered)
    total_qty = sum(i["qty"] for i in filtered.values())
    console.print(f"[#5A9A4A]✓ {total_skus} 個 SKU 在途，總數量 {total_qty:,}[/#5A9A4A]")


if __name__ == "__main__":
    arg = sys.argv[1] if len(sys.argv) > 1 else None
    query(arg)
