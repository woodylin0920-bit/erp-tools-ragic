"""批次發樣的純邏輯層（無互動、無 UI）。

CLI（ragic_upload.run_sample_orders）與 GUI（gui.py）共用這層，
確保「組合計算、開單 payload、寫入 Ragic」只有一份實作、可單獨測試。

★ commit=False 一律只回 payload 預覽、不寫入 Ragic。★
"""
from typing import Callable, List, Optional

import ragic_upload as R   # 重用 Ragic 讀寫、build_payload、JSON 存取（不形成循環：
                            # ragic_upload 只在 run_sample_orders 內延遲 import 本模組）

ORDER_TYPES = R.SAMPLE_ORDER_TYPES   # ["樣品申請", "公關品", "活動贈品"]


# ── 資料載入（給 GUI 一致入口）─────────────────────────────
def load_products() -> list:
    return R._load_sample_products()


def search_products(products: list, keyword: str) -> list:
    return R._search_products(products, keyword)


def load_customers() -> list:
    return R.load_customers()


# ── 組合範本 CRUD ─────────────────────────────────────────
def load_combos() -> dict:
    return R._load_json_file(R.SAMPLE_COMBOS_FILE)


def save_combo(name: str, items: list) -> None:
    combos = load_combos()
    combos[name] = items
    R._save_json_file(R.SAMPLE_COMBOS_FILE, combos)


def delete_combo(name: str) -> None:
    combos = load_combos()
    combos.pop(name, None)
    R._save_json_file(R.SAMPLE_COMBOS_FILE, combos)


# ── 客戶名單 CRUD ─────────────────────────────────────────
def load_cust_lists() -> dict:
    return R._load_json_file(R.SAMPLE_CUSTLIST_FILE)


def save_cust_list(name: str, codes: list) -> None:
    lists = load_cust_lists()
    lists[name] = codes
    R._save_json_file(R.SAMPLE_CUSTLIST_FILE, lists)


def delete_cust_list(name: str) -> None:
    lists = load_cust_lists()
    lists.pop(name, None)
    R._save_json_file(R.SAMPLE_CUSTLIST_FILE, lists)


# ── 開單 ──────────────────────────────────────────────────
def build_sample_payload(customer: dict, combo_items: list, order_type: str) -> dict:
    """customer={code,...}、combo_items=[{code,qty}] → Ragic 銷貨單 payload。
    樣品：單價/總額全 0、狀態未出貨、內部備注標『批次發樣』。"""
    resolved = [{"product_code": it["code"], "unit_price": 0, "quantity": it["qty"]}
                for it in combo_items]
    return R.build_payload(customer, resolved, order_type, "未出貨",
                           tax_rate="", shipping_fee=0, notes="", internal_notes="批次發樣")


def create_sample_orders(order_type: str, combo_items: list, customers: list,
                         commit: bool = False,
                         progress: Optional[Callable[[int, int], None]] = None) -> List[dict]:
    """為每個客戶各開一張樣品單。

    commit=False（預設）：只組 payload 回傳預覽，**不寫入 Ragic**。
    commit=True：逐張 POST，單張失敗不影響其他。
    回 [{customer, ok, msg, payload}]；ok=None 表示 dry-run 未送出。
    progress(done, total) 可選，供 GUI 顯示進度。
    """
    results = []
    total = len(customers)
    for c in customers:
        payload = build_sample_payload(c, combo_items, order_type)
        if not commit:
            results.append({"customer": c, "ok": None, "msg": "dry-run", "payload": payload})
        else:
            try:
                res = R.ragic_post(R.SALES_ORDER_SHEET, payload)
                ok = res.get("status") == "SUCCESS"
                results.append({"customer": c, "ok": ok,
                                "msg": "" if ok else str(res.get("msg", res)), "payload": payload})
            except Exception as e:
                results.append({"customer": c, "ok": False, "msg": str(e), "payload": payload})
        if progress:
            progress(len(results), total)
    return results
