"""蝦皮平台解析。

訂單通知信寄件者 info@mail.shopee.tw。
主旨：「來自{買家}的貨到付款訂單#{訂單號}已被確認」（訂單號/買家在主旨）。
內文明細：「N. 商品標題 / 數量: X / 價格: NT$ Y」，消費者多為單盒、貨到付款。
Ragic 客戶=蝦皮、單別=蝦皮、備註=買家帳號 → 防重複用 買家+日期。

信箱：歷史訂單在 toybebop@gmail.com（253 封）；新單已改寄 info@（之後可加讀）。
"""
import re

from .base import BasePlatform, EItem, EOrder


class Shopee(BasePlatform):
    name = "shopee"
    mailbox_user = "toybebop@gmail.com"
    mailbox_pw_file = "~/.boptoys-gmail_app_password"
    sender_query = "from:shopee.tw 貨到付款 訂單"
    customer = "蝦皮"
    customer_code = "C-00038"
    order_type = "蝦皮"
    has_detail = True

    def ragic_note(self, order):
        """蝦皮備註存買家帳號（與現有人工單一致）。"""
        return order.buyer

    def parse_order(self, subject: str, body: str):
        ms = re.search(r"來自(.+?)的.*?訂單\s*#?([A-Z0-9]+)", subject)
        buyer = ms.group(1).strip() if ms else ""
        m_no = re.search(r"訂單單號[:：]\s*#?(\S+)", body)
        order_no = (m_no.group(1) if m_no else (ms.group(2) if ms else "")).strip("#")
        if not order_no:
            return None
        items = []
        for m in re.finditer(
            r"\n\s*\d+\.\s*(.+?)\n.*?數量[:：]\s*(\d+).*?價格[:：]\s*NT\$?\s*([\d,]+)",
            body, re.S
        ):
            items.append(EItem(m.group(1).strip(), int(m.group(2)),
                               float(m.group(3).replace(",", ""))))
        if not items:
            return None
        m_date = re.search(r"訂單日期[:：]\s*([\d\-: ]+)", body)
        return EOrder(self.name, order_no,
                      m_date.group(1).strip() if m_date else "", items,
                      buyer=buyer, pay_method="貨到付款", pay_status="未付款")

    def is_existing(self, order, ragic_orders):
        """蝦皮備註=買家帳號，無訂單號 → 以 買家 + 同日 判斷是否已開。"""
        od = re.sub(r"\D", "", order.date)[:8]   # YYYYMMDD
        for r in ragic_orders:
            if r["customer"] == self.customer and r["note"] == order.buyer:
                if re.sub(r"\D", "", r["date"])[:8] == od:
                    return True
        return False
