"""ShopStore（官網）平台解析。

訂單通知信寄件者 service@shopstore.tw，主旨含「有一筆新訂單」。
信中「訂單商品內容」區塊為：商品標題  N × NT$價格。
單位由標題判定：(整中收藏)/(中盒)→中盒、(整箱)→整箱、否則單盒。
Ragic 客戶=ShopStore、單別=官網、防重複鍵=平台訂單號（存在 Ragic 備註）。
"""
import re

from .base import BasePlatform, EItem, EOrder


class ShopStore(BasePlatform):
    name = "shopstore"
    sender_query = "from:shopstore.tw 新訂單"
    customer = "ShopStore"
    customer_code = "C-00094"
    order_type = "官網"
    has_detail = True

    def parse_order(self, subject: str, body: str):
        m_no = re.search(r"訂單編號[:：]\s*(\S+)", body)
        if not m_no:
            return None
        # 取「訂單商品內容」到「商品合計」之間整段，逐一抓出所有品項（含加購品）。
        blk = re.search(r"訂單商品內容\s*(.+?)\s*(?:商品合計|$)", body, re.S)
        block = blk.group(1) if blk else ""
        items = []
        # 每項格式：商品名  N × NT$單價  NT$小計
        for m in re.finditer(r"(.+?)\s+(\d+)\s*[×x]\s*NT\$?\s*([\d,]+)\s*NT\$?\s*[\d,]+",
                             block, re.S):
            title = re.sub(r"\s+", " ", m.group(1)).strip()
            if title:
                items.append(EItem(title, int(m.group(2)),
                                   float(m.group(3).replace(",", ""))))
        if not items:
            return None

        # 各欄位值夾在「本標籤」與「下一個已知標籤」之間
        labels = (r"顧客姓名|顧客帳號|顧客電話|收件者姓名|收件者電話|付款方式|付款狀態|"
                  r"送貨方式|運費|取貨地址|出貨狀態|訂單備註|訂單商品內容|前往店家")

        def grab(label):
            m = re.search(label + r"\s*(.+?)\s*(?=" + labels + r"|$)", body, re.S)
            return re.sub(r"\s+", " ", m.group(1)).strip() if m else ""

        m_date = re.search(r"訂購時間\s*([\d\-: ]+)", body)
        m_fee = re.search(r"運費\s*NT\$?\s*([\d,]+)", body)
        return EOrder(
            self.name, m_no.group(1).strip(),
            m_date.group(1).strip() if m_date else "", items,
            buyer=grab("顧客姓名"),
            pay_method=grab("付款方式"),
            pay_status=grab("付款狀態"),
            ship_method=grab("送貨方式"),
            fee=float(m_fee.group(1).replace(",", "")) if m_fee else 0.0,
        )
