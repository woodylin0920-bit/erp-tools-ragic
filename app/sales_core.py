"""新建銷售單的純邏輯（無互動）。供 GUI 使用；CLI 仍走 ragic_upload 既有實作。

GUI 與 CLI 差異（安全考量）：
- 客戶只「比對」不自動建檔；對不到 → 標 customer_missing，請人工先建檔。
- 同條碼多規格不互動詢問 → 自動取最小規格並標 ambiguous 提醒複核。
★ commit=False 不寫入。★
"""
from typing import Optional

import ragic_upload as R


def list_pending() -> list:
    return R.find_pending_files(R.BASE_CLIENT_ORDER)


def parse_file(path) -> tuple:
    """回 (client_code, orders)；不支援的客群或解析失敗會 raise。"""
    from parsers import PARSERS
    client = path.parent.name.upper()
    if client not in PARSERS:
        raise ValueError(f"不支援的客戶代碼：{client}（支援：{', '.join(PARSERS)}）")
    orders = PARSERS[client](str(path)).parse()
    if not orders:
        raise ValueError("無法解析任何訂單，請確認檔案格式")
    return client, orders


def match_customer(customers: list, store_code: str, client_code: str) -> Optional[dict]:
    """非互動比對（不建檔）。比對不到或多筆相符 → None。"""
    narrowed = [c for c in customers if store_code in c["name"] and client_code in c["name"]]
    matches = narrowed or [c for c in customers if store_code in c["name"]]
    return matches[0] if len(matches) == 1 else None


def preview_file(path, price_index: dict, customers: list) -> list:
    """解析整檔，回每張訂單的預覽（不寫入）。
    [{store, po, customer, customer_missing, items, box_notes, ambiguous, subtotal}]"""
    client, orders = parse_file(path)
    out = []
    for order in orders:
        cust = match_customer(customers, order.store_code, client)
        resolved = R.resolve_items(order.items, price_index,
                                   auto_unit_spec=(client == "LE"), auto_pick_ambiguous=True)
        box_notes = [it["box_note"] for it in resolved if it.get("box_note")]
        ambiguous = any(getattr(it, "_ambiguous", False) for it in order.items)
        out.append({
            "client": client, "store": order.store_code, "po": order.po_number,
            "customer": cust, "customer_missing": cust is None,
            "items": resolved, "box_notes": box_notes, "ambiguous": ambiguous,
            "subtotal": sum(it["amount"] for it in resolved),
        })
    return out


def create_order(customer: dict, resolved: list, order_type: str, order_status: str,
                 tax_rate: str, po_number: str = "", client: str = "", store: str = "",
                 commit: bool = False) -> dict:
    """開一張銷貨單。commit=False 只回 payload 不寫入。回 {ok, msg, ragic_id, payload, log_key}。

    防重複：以 client_store_PO 為鍵查/寫 upload_log（與 CLI process_file 一致）。
    PO#：若有 po_number，寫進備註「PO#xxxx」，讓出庫流程下游能帶到。
    """
    log_key = f"{client}_{store}_{po_number}" if (client and store and po_number) else ""
    if commit and log_key:
        log = R._load_upload_log()
        if log_key in log:
            return {"ok": None, "msg": f"已開過（{log[log_key].get('uploaded_at','?')}），略過防重複",
                    "ragic_id": str(log[log_key].get("ragic_id", "")), "payload": None,
                    "log_key": log_key, "dup": True}
    notes = f"PO#{po_number}" if po_number else ""
    payload = R.build_payload(customer, resolved, order_type, order_status,
                              tax_rate=tax_rate, shipping_fee=0, notes=notes,
                              internal_notes="【程式建單·GUI】")
    if not commit:
        return {"ok": None, "msg": "dry-run", "ragic_id": "", "payload": payload, "log_key": log_key}
    try:
        res = R.ragic_post(R.SALES_ORDER_SHEET, payload)
        ok = res.get("status") == "SUCCESS" or bool(res.get("ragicId"))
        ragic_id = str(res.get("ragicId", ""))
        if ok and log_key:
            from datetime import datetime
            log = R._load_upload_log()
            log[log_key] = {"ragic_id": ragic_id,
                            "uploaded_at": datetime.now().strftime("%Y/%m/%d %H:%M"), "file": "GUI"}
            R._save_upload_log(log)
        return {"ok": ok, "msg": "" if ok else str(res.get("msg", res)),
                "ragic_id": ragic_id, "payload": payload, "log_key": log_key}
    except Exception as e:
        return {"ok": False, "msg": str(e), "ragic_id": "", "payload": payload, "log_key": log_key}
