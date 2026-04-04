"""
TemplateParser - 通用客戶訂單模板格式
格式特徵：
  - 使用 order-template.xlsx 模板
  - Row 1：標題
  - Row 2：欄位名稱，N 欄起為店號（客戶自填，如 S001）
  - Row 3：A3:M3 合併為分類標籤，N 欄起為店名（客戶自填）
  - Row 4+：商品資料
  - D 欄：國際條碼（13碼）
  - J 欄：單價（含稅）
  - L 欄：SKN（客戶自填）
  - M 欄：PO 號碼（客戶自填）
  - N 欄起：各門市數量（客戶自填）
"""

import re
from typing import Dict, List, Tuple

import openpyxl

from .base import BaseParser, Order, OrderItem

BARCODE_COL     = 3   # Column D (0-indexed)
UNIT_PRICE_COL  = 9   # Column J
PO_COL          = 12  # Column M
STORE_START_COL = 13  # Column N (0-indexed)

BARCODE_RE = re.compile(r'^\d{12,14}$')

# 不當作店號的保留字
SKIP_VALS = {"TTL", "TOTAL", "合計", "小計", "總數", "現貨", "在途",
             "PO號碼", "PO", "PO#", "SUBTOTAL", ""}


def _val(cell) -> str:
    return str(cell).strip() if cell is not None else ""


def _float(cell) -> float:
    try:
        return float(cell) if cell is not None else 0.0
    except (ValueError, TypeError):
        return 0.0


class TemplateParser(BaseParser):

    def parse(self) -> List[Order]:
        wb = openpyxl.load_workbook(self.filepath, data_only=True)
        orders = []
        for sheet in wb.worksheets:
            orders.extend(self._parse_sheet(sheet))
        return orders

    def _parse_sheet(self, sheet) -> List[Order]:
        all_rows = list(sheet.iter_rows(values_only=True))
        if len(all_rows) < 3:
            return []

        # 找店號列（Row 2，index 1）
        header_row = all_rows[1]
        store_cols: Dict[str, int] = {}
        for col_idx in range(STORE_START_COL, len(header_row)):
            val = _val(header_row[col_idx])
            if val and val.upper() not in SKIP_VALS:
                store_cols[val] = col_idx

        if not store_cols:
            return []

        store_orders: Dict[Tuple[str, str], Order] = {}

        for row in all_rows[3:]:  # Row 4 起
            if not row or len(row) <= BARCODE_COL:
                continue
            barcode = _val(row[BARCODE_COL])
            if not BARCODE_RE.match(barcode):
                continue

            po_number  = _val(row[PO_COL]) if PO_COL < len(row) else ""
            unit_price = _float(row[UNIT_PRICE_COL]) if UNIT_PRICE_COL < len(row) else 0.0
            le_name    = _val(row[1]) if len(row) > 1 else ""

            for store_code, col_idx in store_cols.items():
                if col_idx >= len(row):
                    continue
                qty = _float(row[col_idx])
                if qty <= 0:
                    continue

                po = po_number or sheet.title
                key = (store_code, po)
                if key not in store_orders:
                    store_orders[key] = Order(
                        client_code="TEMPLATE",
                        store_code=store_code,
                        po_number=po,
                        source_file=self.filepath,
                    )
                store_orders[key].items.append(OrderItem(
                    barcode=barcode,
                    quantity=qty,
                    le_name=le_name,
                    unit_price=unit_price,
                ))

        return list(store_orders.values())
