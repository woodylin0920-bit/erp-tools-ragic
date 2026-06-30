"""建立出庫單的純邏輯層（無互動）。供 GUI 使用；CLI 仍走 ragic_upload 內既有實作。

重用 ragic_upload 既有的純函式（compute_break_plan / build_code_to_barcode /
extract_po / ragic_* ），這裡只負責「組流程」：載入情境、算拆盒、拋轉+補欄位。

★ commit=False 一律不寫入（不拆盒、不拋轉、不補欄位）。★
"""
import time
from typing import Callable, Optional

import ragic_upload as R

DELIVERY_SUBTABLE = "_subtable_3000886"
DEFAULT_WH = "TW01"


def load_context() -> dict:
    """載入出貨單、倉庫、(倉庫,商品)→庫存編號清單。回 dict。"""
    records = R.ragic_get(R.DELIVERY_ORDER_SHEET)
    inventory = R.ragic_get(R.INVENTORY_SHEET)
    warehouses, inv_by_wh_prod = {}, {}
    for rec in inventory.values():
        wh = str(rec.get("倉庫代碼", "")).strip()
        prod = str(rec.get("商品編號", "")).strip()
        inv_code = str(rec.get("庫存編號", "")).strip()
        if wh:
            warehouses[wh] = str(rec.get("倉庫名稱", "")).strip()
        if wh and prod and inv_code:
            inv_by_wh_prod.setdefault((wh, prod), []).append(inv_code)
    candidates = [{"id": rid,
                   "label": f"{rec.get('出貨單號','?')}  {rec.get('客戶名稱','?')}  {rec.get('訂單日期','?')}"}
                  for rid, rec in records.items()]
    return {"records": records, "inventory": inventory, "warehouses": warehouses,
            "inv_by_wh_prod": inv_by_wh_prod, "candidates": candidates}


def products_of(records: dict, record_ids: list) -> list:
    """選定出貨單需要的商品（去重）。回 [{prod, name}]。"""
    seen, out = set(), []
    for rid in record_ids:
        for row in records[rid].get(DELIVERY_SUBTABLE, {}).values():
            prod = str(row.get("商品編號*", "") or row.get("商品編號", "")).strip()
            if prod and prod not in seen:
                seen.add(prod)
                out.append({"prod": prod, "name": str(row.get("商品名稱", "")).strip()})
    return out


def break_plan(records: dict, record_ids: list, inventory: dict, warehouse_code: str) -> list:
    """逐出貨單明細算拆盒計畫（重用 compute_break_plan）。"""
    line_needs = []
    for rid in record_ids:
        rec = records[rid]
        label = f"{rec.get('出貨單號', rid)} {rec.get('客戶名稱', '')}".strip()
        agg = {}
        for row in rec.get(DELIVERY_SUBTABLE, {}).values():
            prod = str(row.get("商品編號*", "") or row.get("商品編號", "")).strip()
            if prod:
                agg[prod] = agg.get(prod, 0) + R._to_int(row.get("數量", 0), 0)
        for prod, q in agg.items():
            line_needs.append((label, prod, q))
    return R.compute_break_plan(line_needs, inventory, warehouse_code)


def merge_breakbox(plan: list) -> dict:
    """把 status=ok 的拆盒依中盒彙整（跨單同中盒只 PATCH 一次）。"""
    merged = {}
    for p in plan:
        if p["status"] != "ok":
            continue
        m = merged.setdefault(p["parent"], {
            "parent_rid": p["parent_rid"], "parent_qty": p["parent_qty"],
            "unit_rid": p["unit_rid"], "unit_qty": p["unit_qty"],
            "boxes": 0, "gain": 0, "unit": p["prod"]})
        m["boxes"] += p["boxes"]
        m["gain"] += p["gain"]
    return merged


def apply_breakbox(merged: dict) -> list:
    """實際改 20008 數量（中盒-、單盒+）。回 [(中盒, ok, msg)]。"""
    results = []
    for pc, m in merged.items():
        try:
            R.ragic_patch(R.INVENTORY_SHEET, m["parent_rid"], {R.INVENTORY_QTY_CID: m["parent_qty"] - m["boxes"]})
            R.ragic_patch(R.INVENTORY_SHEET, m["unit_rid"], {R.INVENTORY_QTY_CID: m["unit_qty"] + m["gain"]})
            results.append((pc, True, f"中盒 -{m['boxes']}、{m['unit']} 單盒 +{m['gain']}"))
        except Exception as e:
            results.append((pc, False, str(e)))
    return results


def create_outbound(records: dict, record_ids: list, warehouse_code: str,
                    prod_inv_map: dict, progress: Optional[Callable] = None) -> dict:
    """拋轉建立出庫單 + 自動補（倉庫代碼/庫存編號/單據備註=客戶[/PO#]/明細備註=EAN）。
    這是寫入動作；呼叫前請先讓使用者確認。回 {triggered, new, patched, msgs}。"""
    bid = R.ragic_get_action_button_id(R.DELIVERY_ORDER_SHEET, "建立出庫單")
    if bid is None:
        raise RuntimeError("找不到「建立出庫單」按鈕")
    before = set(R.ragic_get(R.OUTBOUND_ORDER_SHEET).keys())
    triggered = 0
    trigger_errs = []
    for rid in record_ids:
        # 單張觸發失敗(逾時/HTTP錯)不可中斷：否則前面已建好的出庫單會漏補欄位。
        try:
            res = R.ragic_trigger_button(R.DELIVERY_ORDER_SHEET, rid, bid)
            if res.get("status") == "SUCCESS":
                triggered += 1
            else:
                trigger_errs.append(f"{rid}: {res.get('msg', res)}")
        except Exception as e:
            trigger_errs.append(f"{rid}: {e}")
        if progress:
            progress("trigger", triggered, len(record_ids))
    time.sleep(3)
    after = R.ragic_get(R.OUTBOUND_ORDER_SHEET)
    new_ids = set(after.keys()) - before
    if not new_ids:
        return {"triggered": triggered, "new": 0, "patched": 0,
                "msgs": ["未偵測到新出庫單（可能被擋重複拋轉）"] + trigger_errs}

    shipno_cust = {str(records[r].get("出貨單號", "")).strip(): str(records[r].get("客戶名稱", "")).strip()
                   for r in record_ids}
    shipno_order = {str(records[r].get("出貨單號", "")).strip(): str(records[r].get("訂單編號", "")).strip()
                    for r in record_ids}
    order_po = {}
    try:
        for so in R.ragic_get(R.SALES_ORDER_SHEET).values():
            on = str(so.get("訂單編號", "")).strip()
            po = R.extract_po(so.get("備註", ""))
            if on and po:
                order_po[on] = po
    except Exception:
        pass
    code_to_barcode = R.build_code_to_barcode(R.load_price_index())

    patched, msgs = 0, []
    for oid in new_ids:
        rec = after[oid]
        sub = rec.get(R.OUTBOUND_ITEMS_SUBTABLE_KEY, {})
        if not sub:
            continue
        ship_no = str(rec.get("出貨單號", "")).strip()
        customer = shipno_cust.get(ship_no, "")
        po = order_po.get(shipno_order.get(ship_no, ""), "")
        doc_note = f"{customer} / PO#{po}".strip(" /") if po else customer
        rows = {}
        for row_id, row in sub.items():
            if str(row_id).startswith("_"):
                continue
            prod = str(row.get("商品編號", "")).strip()
            cell = {}
            inv_code = prod_inv_map.get(prod, "")
            if inv_code:
                cell["3001124"] = warehouse_code
                cell["3001126"] = inv_code
            bars = code_to_barcode.get(prod, [])
            if bars:
                cell[R.OUTBOUND_ROW_NOTE_CID] = f"【EAN】{bars[0]}"
            if cell:
                rows[str(row_id)] = cell
        body = {R.OUTBOUND_ITEMS_SUBTABLE_KEY: rows}
        if doc_note:
            body[R.OUTBOUND_DOC_NOTE_CID] = doc_note
        try:
            R.ragic_patch(R.OUTBOUND_ORDER_SHEET, oid, body)
            patched += 1
        except Exception as e:
            msgs.append(f"出庫單 {oid} 補填失敗：{e}")
        if progress:
            progress("patch", patched, len(new_ids))
    if trigger_errs:
        msgs = [f"部分拋轉失敗（已建好的仍補欄位）：{'; '.join(trigger_errs[:5])}"] + msgs
    return {"triggered": triggered, "new": len(new_ids), "patched": patched, "msgs": msgs}
