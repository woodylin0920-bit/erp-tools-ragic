"""平台介面：每個電商平台實作這些屬性與方法。

新增平台 = 新增一個 platforms/<name>.py 實作 BasePlatform，
並在 platforms/__init__.py 註冊；共用的 core.py 不需更動。
"""
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class EItem:
    title: str          # 平台商品標題（原文）
    qty: int            # 數量
    price: float        # 平台售價（單價）


@dataclass
class EOrder:
    platform: str
    order_no: str                    # 平台訂單號
    date: str                        # 訂購時間
    items: List[EItem] = field(default_factory=list)
    buyer: str = ""                  # 顧客姓名/帳號
    pay_method: str = ""             # 付款方式（如 超商取貨付款(7-11)、信用卡）
    pay_status: str = ""             # 付款狀態（未付款/已付款）
    ship_method: str = ""            # 送貨方式
    fee: float = 0.0                 # 運費

    @property
    def is_cod_pending(self) -> bool:
        """超商取貨付款且尚未付款 → 待取貨，有未領風險。"""
        return ("取貨付款" in self.pay_method or "貨到付款" in self.pay_method) \
            and "未" in self.pay_status


class BasePlatform:
    name = ""                                   # 平台代號（對應 product_map.json 的 key）
    mailbox_user = "info@boptoys.com.tw"        # 訂單通知信箱
    mailbox_pw_file = "~/.boptoys-info_app_password"
    sender_query = ""                           # Gmail 搜尋（限定「新訂單」信）
    customer = ""                               # Ragic 客戶名稱（單一通路客戶）
    order_type = ""                             # Ragic 訂單單別
    has_detail = True                           # 訂單信是否含商品明細

    def parse_order(self, body: str) -> Optional[EOrder]:
        """從信件內文解析出一張訂單。解析不出回 None。"""
        raise NotImplementedError

    def dedup_note(self, order: EOrder) -> str:
        """回傳用來和 Ragic『備註』比對是否已開單的鍵。
        ShopStore/Pinkoi：平台訂單號；蝦皮：另以買家+日期+金額，於子類覆寫。"""
        return order.order_no
