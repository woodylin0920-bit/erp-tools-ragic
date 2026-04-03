from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class OrderItem:
    barcode: str
    quantity: float
    le_code: str = ""       # client's own product code
    le_name: str = ""       # client's product name
    unit_price: float = 0.0 # pre-filled if known from client file (TRU)


@dataclass
class Order:
    client_code: str        # e.g. "LE" or "TRU"
    store_code: str         # e.g. "AD227" or "4402"
    po_number: str
    items: List[OrderItem] = field(default_factory=list)
    source_file: str = ""


class BaseParser:
    """Subclass this for each client's Excel format."""

    def __init__(self, filepath: str):
        self.filepath = filepath

    def parse(self) -> List[Order]:
        raise NotImplementedError
