"""
TRUParser - 玩具反斗城報價表格式
格式特徵：
  - 一張 Excel = 多張 Ragic 訂單（每個門市各一張）
  - Row 2-3：表頭，門市代碼（4402, 4418 ...）在 Column P (index 15) 之後
  - 商品列（Column D = 13碼條碼）從 Row 4 起
  - Column J = 單價（不含稅）
  - PO 號碼：Column N (index 13)
  - 每個門市欄值 > 0 的商品組成該門市的訂單
  - 多 PO 支援：若有 PO# 欄位，段落標題列會在該欄標記新 PO 號，
    後續商品列沿用此 PO，每個 (門市, PO) 組合建立一張獨立訂單
"""

import io
import re
import zipfile
from typing import Dict, List, Optional, Tuple

import openpyxl

from .base import BaseParser, Order, OrderItem

def _strip_autofilter(xml: str) -> str:
    # Remove block form: <autoFilter ...> ... </autoFilter>
    xml = re.sub(r'<autoFilter\b.*?</autoFilter>', '', xml, flags=re.DOTALL | re.IGNORECASE)
    # Remove self-closing form: <autoFilter ... />
    xml = re.sub(r'<autoFilter\b[^>]*/>', '', xml, flags=re.IGNORECASE)
    return xml


def _load_wb(filepath: str):
    """Load workbook, stripping broken <autoFilter> XML on first failure."""
    try:
        return openpyxl.load_workbook(filepath, data_only=True)
    except Exception:
        buf = io.BytesIO()
        with zipfile.ZipFile(filepath, 'r') as zin, \
             zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                    data = _strip_autofilter(data.decode('utf-8')).encode('utf-8')
                zout.writestr(item, data)
        buf.seek(0)
        return openpyxl.load_workbook(buf, data_only=True)

BARCODE_COL    = 3   # Column D (0-indexed)
UNIT_PRICE_COL = 9   # Column J
PO_COL         = 13  # Column N
STORE_START_COL = 15 # Column P (first store column)

BARCODE_RE = re.compile(r'^\d{12,14}$')
PO_RE      = re.compile(r'^\d{5,8}$')   # PO 號格式：5-8 位數字


def _val(cell) -> str:
    return str(cell).strip() if cell is not None else ""


def _float(cell) -> float:
    try:
        return float(cell) if cell is not None else 0.0
    except (ValueError, TypeError):
        return 0.0


class TRUParser(BaseParser):

    def parse(self) -> List[Order]:
        wb = _load_wb(self.filepath)
        orders = []

        for sheet in wb.worksheets:
            orders.extend(self._parse_sheet(sheet))

        return orders

    def _parse_sheet(self, sheet) -> List[Order]:
        all_rows = list(sheet.iter_rows(values_only=True))
        if len(all_rows) < 3:
            return []

        # Find the header row that contains store codes (4-digit numbers like 4402)
        header_row_idx, store_cols, po_override_col = self._find_store_columns(all_rows)
        if not store_cols:
            return []

        # Build orders keyed by (store_code, po_number) to support multi-PO files
        store_orders: Dict[Tuple[str, str], Order] = {}

        for row in all_rows[header_row_idx + 1:]:
            if len(row) <= BARCODE_COL:
                continue

            barcode = _val(row[BARCODE_COL])
            if not BARCODE_RE.match(barcode):
                continue

            # Determine PO per column:
            #   - stores LEFT  of po_override_col → PO from col13
            #   - stores RIGHT of po_override_col → PO from po_override_col cell
            po_left  = _val(row[PO_COL]) if PO_COL < len(row) else ""
            po_right = (
                _val(row[po_override_col])
                if po_override_col is not None and po_override_col < len(row)
                else ""
            )
            unit_price = _float(row[UNIT_PRICE_COL]) if UNIT_PRICE_COL < len(row) else 0.0
            le_name    = _val(row[1]) if len(row) > 1 else ""

            for store_code, col_idx in store_cols.items():
                if col_idx >= len(row):
                    continue
                qty = _float(row[col_idx])
                if qty <= 0:
                    continue

                # Pick PO based on which side of po_override_col this store is
                if po_override_col is not None and col_idx > po_override_col:
                    po_number = po_right or po_left or sheet.title
                else:
                    po_number = po_left or sheet.title

                key = (store_code, po_number)
                if key not in store_orders:
                    store_orders[key] = Order(
                        client_code="TRU",
                        store_code=store_code,
                        po_number=po_number,
                        source_file=self.filepath,
                    )

                store_orders[key].items.append(OrderItem(
                    barcode=barcode,
                    quantity=qty,
                    le_name=le_name,
                    unit_price=unit_price,
                ))

        # Sort by Excel column order: left-of-PO# stores first, then right-of-PO# stores
        col_order = {code: i for i, code in enumerate(store_cols.keys())}

        def _sort_key(order):
            col = col_order.get(order.store_code, 999)
            group = 1 if (po_override_col is not None and col > po_override_col) else 0
            return (group, col)

        return sorted(store_orders.values(), key=_sort_key)

    def _find_store_columns(self, all_rows) -> tuple:
        """
        Scan rows 0-4 for the row containing 4-digit TRU store codes (e.g. 4402).
        Also capture text-named stores from the same row AND the row above
        (e.g. DC in row above, 統領/板橋大遠百 in same row).
        Returns (header_row_index, {store_code: col_index}, po_override_col).
        po_override_col: column index of the 'PO#' special column, or None.
        """
        store_code_re = re.compile(r'^4\d{3}$')
        skip_vals = {"TTL", "TOTAL", "合計", "小計", "PO號碼", "PO", "PO#", ""}

        for row_idx, row in enumerate(all_rows[:5]):
            # Check if this row contains any 4-digit store codes
            has_numeric = any(
                store_code_re.match(_val(cell))
                for col_idx, cell in enumerate(row)
                if col_idx >= STORE_START_COL - 1
            )
            if not has_numeric:
                continue

            # Collect all stores in LEFT-TO-RIGHT column order
            # Check both this row and the row above for each column
            prev_row = all_rows[row_idx - 1] if row_idx > 0 else []
            store_cols = {}
            po_override_col = None

            for col_idx in range(STORE_START_COL - 1, len(row)):
                val = _val(row[col_idx]) if col_idx < len(row) else ""
                if not val and col_idx < len(prev_row):
                    val = _val(prev_row[col_idx])

                if not val:
                    continue
                if val.upper() == "PO#":
                    po_override_col = col_idx   # 記住 PO# 欄位置，不加入門市
                    continue
                if val.upper() not in skip_vals:
                    store_cols[val] = col_idx

            return row_idx, store_cols, po_override_col

        return 0, {}, None
