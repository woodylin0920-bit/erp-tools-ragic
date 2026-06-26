#!/usr/bin/env python3
"""
Parser 回歸測試（純 Python，不需 pytest）。

執行：  python tests/test_parsers.py

涵蓋：
1. TRU 新版面（含 現貨/訂單總數/各店 欄）— 不可有幽靈門市、PO 取自 PO號碼欄、一店一單
2. TRU 舊版面（DC 等文字門市在數字店號左側）— 文字門市須正確納入
3. #3 非整中盒 → resolve_items 標記、build_payload 寫入內部備注
4. （選用）本機真實檔回歸：放在 tests/fixtures/TRU|LE/ 才會跑，否則略過

合成版面不含任何客戶實際資料，可安全進 CI。
"""
import io
import os
import sys
from types import SimpleNamespace

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "app"))

from parsers.tru_parser import TRUParser  # noqa: E402

_failures = []


def check(cond, msg):
    mark = "✅" if cond else "❌"
    print(f"  {mark} {msg}")
    if not cond:
        _failures.append(msg)


def _save_tmp(wb, name):
    path = os.path.join(
        os.environ.get("TMPDIR", "/tmp"), f"_tru_test_{name}.xlsx"
    )
    wb.save(path)
    return path


def _new_layout_workbook():
    """新版面：…SKN, PO號碼, 現貨, 訂單總數, 各店/PO#, <門市…>, TTL。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    # 0   1   2   3      4   5   6    7   8   9    10  11  12     13   14     15    16      17      18
    # IP 系列 英文 國際條碼 參圖 規格 MSRP 定價 折數 單價 備註 SKN PO號碼 現貨 訂單總數 各店  台北新生 林口三井 TTL
    ws.append(["潮玩波普-報價單"])
    ws.append(["IP", "系列", "英文", "國際條碼", "參考圖", "規格", "MSRP", "定價",
               "折數", "單價", "備註", "SKN", "PO號碼", "現貨", "訂單總數", "各店",
               "台北新生", "林口三井", "TTL"])
    ws.append(["盲盒產品", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
               "PO#", "4402", "4463", "TTL"])
    # 兩個產品；PO號碼=90001；現貨/訂單總數欄故意有值（測試不被當門市/PO）
    ws.append(["A", "系列A", "EN", "6975931340001", "", "規格", 390, 399, 0.5, 195,
               "", "55001", 90001, 36, 30, "", 10, 18, 28])
    ws.append(["B", "系列B", "EN", "6975931340002", "", "規格", 190, 199, 0.5, 95,
               "", "55002", 90001, 120, 44, "", 20, 24, 44])
    return wb


def _old_layout_workbook():
    """舊版面：…SKN, PO號碼, DC(文字門市), <數字店號/文字店號交錯>。無 現貨/各店 欄。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    # 0  1   2   3      4   5   6       7      8   9   10  11  12  13     14  15      16       17
    # IP 系列 英文 國際條碼 參圖 規格 建議零售價 TRU定價 折數 單價 備註 數量 SKN PO號碼 DC 台北新生 板橋大遠百 中壢家福
    ws.append(["玩具反斗城-潮玩波普-價格表"])
    ws.append(["IP", "系列", "英文", "國際條碼", "參考圖", "規格", "建議零售價",
               "TRU定價", "折數", "單價", "備註", "數量", "SKN", "PO號碼",
               "DC", "台北新生", "板橋大遠百", "中壢家福"])
    # 店號列：數字店號在 15、17；文字門市「板橋大遠百」在 16；DC 在上一列(14)、本列空
    ws.append(["盲盒產品", "", "", "", "", "", "", "", "", "", "", "", "", "",
               "", "4402", "板橋大遠百", "4410"])
    ws.append(["A", "系列A", "EN", "6975931340001", "", "規格", 190, 199, 0.5, 95,
               "", 600, "55001", 521493, 360, 20, 20, 10])
    ws.append(["B", "系列B", "EN", "6975931340002", "", "規格", 390, 399, 0.5, 195,
               "", 100, "55002", 521493, 40, 8, 0, 16])
    return wb


def test_new_layout():
    print("[1] TRU 新版面（現貨/訂單總數/各店）")
    path = _save_tmp(_new_layout_workbook(), "new")
    orders = TRUParser(path).parse()
    stores = sorted(str(o.store_code) for o in orders)
    check("訂單總數" not in stores and "各店" not in stores, "無幽靈門市（訂單總數/各店/現貨）")
    check(stores == ["4402", "4463"], f"門市正確 = {stores}")
    check(all(str(o.po_number) == "90001" for o in orders), "PO 取自 PO號碼欄(90001)，非現貨值")
    by = {o.store_code: o for o in orders}
    check(len(by["4463"].items) == 2, "4463 兩品項歸同一張單（未被拆）")
    tot = sum(i.quantity for o in orders for i in o.items)
    check(tot == 10 + 18 + 20 + 24, f"總量正確 = {tot}")


def test_old_layout():
    print("[2] TRU 舊版面（DC 在數字店號左側）")
    path = _save_tmp(_old_layout_workbook(), "old")
    orders = TRUParser(path).parse()
    stores = sorted(str(o.store_code) for o in orders)
    check("DC" in stores, "DC（文字門市，位於店號左側）有納入")
    check("板橋大遠百" in stores, "板橋大遠百（交錯文字門市）有納入")
    check(set(stores) == {"DC", "4402", "板橋大遠百", "4410"}, f"門市完整 = {stores}")
    check(all(str(o.po_number) == "521493" for o in orders), "PO=521493")


def test_mid_box_note():
    print("[3] 非整中盒 → 標記 + 寫入內部備注")
    import questionary
    import ragic_upload as R

    price_index = {
        "6975931340001": [
            {"product_code": "P1", "product_name": "丫丫-單盒", "spec": 1, "unit": "單盒", "price": 195},
            {"product_code": "P2", "product_name": "丫丫-中盒", "spec": 8, "unit": "中盒", "price": 1560},
        ],
        "6975931340002": [
            {"product_code": "Q1", "product_name": "薇薇-單盒", "spec": 1, "unit": "單盒", "price": 95},
            {"product_code": "Q2", "product_name": "薇薇-中盒", "spec": 10, "unit": "中盒", "price": 950},
        ],
    }

    class _Sel:
        def __init__(self, c): self._c = c
        def ask(self): return self._c

    questionary.select = lambda msg, choices, **k: _Sel(
        next((c for c in choices if "中盒" in c), choices[0]))

    def it(bc, q): return SimpleNamespace(barcode=bc, quantity=q, unit_price=0, le_name="")

    resolved = R.resolve_items([it("6975931340001", 98), it("6975931340002", 100)],
                               price_index, auto_unit_spec=False)
    notes = {r["product_code"]: r["box_note"] for r in resolved}
    check(bool(notes.get("P1")), "丫丫 98pcs（中盒8）被標記非整中盒")
    check(not notes.get("Q2"), "薇薇 100pcs（中盒10）未被標記")

    pay = R.build_payload({"code": "C-1", "name": "TRU-X"}, resolved,
                          "一般訂單", "未出貨", "5%", 0, "", "", commission="")
    check("非整中盒" in pay["1000074"], "非整中盒提醒已寫入內部備注(1000074)")

    le = R.resolve_items([it("6975931340001", 98)], price_index, auto_unit_spec=True)
    check(le[0]["box_note"] == "", "LE(auto_unit_spec) 不誤標")


def test_stock_check():
    print("[3b] 開單前庫存把關（離線，預灌庫存快取）")
    import ragic_upload as R
    R._stock_cache["TW01"] = {"BBK012": 4, "BMC012": 59}  # 預灌，避免打 API
    resolved = [
        {"product_code": "BMC012-1", "product_name": "小汐-中盒", "spec": 8,
         "unit": "中盒", "quantity": 3, "unit_price": 0, "amount": 0, "box_note": ""},
        {"product_code": "BBK012-1", "product_name": "貝琪-中盒", "spec": 8,
         "unit": "中盒", "quantity": 10, "unit_price": 0, "amount": 0, "box_note": ""},
        {"product_code": "ZZZ999-1", "product_name": "未追蹤品", "spec": 1,
         "unit": "個", "quantity": 2, "unit_price": 0, "amount": 0, "box_note": ""},
    ]
    short = R.check_stock(resolved, "TW01")
    check(len(short) == 1 and "貝琪" in short[0], "貝琪超量(10>4)入不足清單；小汐足、未追蹤不誤判")
    check(R._strip_code_suffix("BMC012-1") == "BMC012", "代號後綴剝除 BMC012-1 → BMC012")


def test_create_customer():
    print("[3c] 互動式新建客戶（離線，攔截 API）")
    import questionary
    import ragic_upload as R

    class _Seq:
        def __init__(self, vals): self.vals = list(vals)
        def __call__(self, *a, **k):
            v = self.vals.pop(0)
            return type("R", (), {"ask": lambda s, _v=v: _v})()

    class _Const:
        def __init__(self, v): self.v = v
        def __call__(self, *a, **k):
            return type("R", (), {"ask": lambda s: self.v})()

    answers = ["TRU-4465", "竹北大遠百", "何小姐", "0912-000-111", "新北市XX路1號", "Woody"]

    # dry-run 經 find_customer → 建立、入快取
    questionary.select = _Const("➕ 建立新客戶")
    questionary.confirm = _Const(True)
    questionary.text = _Seq(list(answers))
    customers = []
    cust = R.find_customer(customers, "4465", "TRU", dry_run=True)
    check(cust["name"] == "TRU-4465" and customers and customers[-1]["name"] == "TRU-4465",
          "dry-run 建立 TRU-4465 並加入快取")

    # 非 dry-run：攔截 ragic_post 驗證 payload CID
    captured = {}
    R.ragic_post = lambda sheet, payload: (
        captured.update(sheet=sheet, payload=payload)
        or {"status": "SUCCESS", "ragicId": 999, "data": {"3000666": "C-00246"}})
    questionary.text = _Seq(list(answers))
    cust2 = R.create_customer_interactive("4465", "TRU", [], dry_run=False)
    p = captured["payload"]
    check(p.get("3000479") == "TRU-4465" and "3000666" not in p,
          "payload 客戶名稱用 CID、客戶編號不填（自動）")
    check(p.get("3000909") == "0912-000-111" and p.get("3000483") == "0912-000-111",
          "電話寫入手機+電話兩欄")
    check(cust2["code"] == "C-00246", "回讀 Ragic 自動編號 C-00246")


def test_real_fixtures():
    """選用：本機 tests/fixtures/ 有真實檔才跑（檔案不進 git）。"""
    from parsers.le_parser import LEParser
    fx = os.path.join(ROOT, "tests", "fixtures")
    if not os.path.isdir(fx):
        print("[4] 真實檔回歸：略過（無 tests/fixtures/）")
        return
    print("[4] 真實檔回歸（本機 fixtures）")
    import glob
    for f in sorted(glob.glob(os.path.join(fx, "TRU", "*.xlsx"))):
        n = len(TRUParser(f).parse())
        check(n > 0, f"TRU {os.path.basename(f)[:34]} → {n} 單")
    for f in sorted(glob.glob(os.path.join(fx, "LE", "*.xlsx"))):
        n = len(LEParser(f).parse())
        check(n > 0, f"LE  {os.path.basename(f)[:34]} → {n} 單")


if __name__ == "__main__":
    test_new_layout()
    test_old_layout()
    test_mid_box_note()
    test_stock_check()
    test_create_customer()
    test_real_fixtures()
    print()
    if _failures:
        print(f"❌ {len(_failures)} 項失敗")
        sys.exit(1)
    print("✅ 全部通過")
