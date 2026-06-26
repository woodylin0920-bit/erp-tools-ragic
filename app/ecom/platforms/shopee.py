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
    # 涵蓋兩種帶明細的訂單信：貨到付款確認、以及（一般付款的明細在）出貨提醒。
    # 同一訂單可能兩封都收到 → 以訂單號去重（reconcile 內處理）。
    sender_query = "from:shopee.tw (貨到付款訂單 OR 準時出貨)"
    customer = "蝦皮"
    customer_code = "C-00038"
    order_type = "蝦皮"
    has_detail = True

    def ragic_note(self, order):
        """蝦皮備註存買家帳號（與現有人工單一致）。"""
        return order.buyer

    def parse_order(self, subject: str, body: str):
        # 訂單號優先從內文（兩種信都有），退而從主旨
        m_no = re.search(r"訂單單號[:：]\s*#?(\S+)", body)
        order_no = m_no.group(1).strip("#") if m_no else ""
        if not order_no:
            ms = re.search(r"訂單\s*#?([A-Z0-9]{8,})", subject)
            order_no = ms.group(1) if ms else ""
        if not order_no:
            return None
        # 買家：貨到付款信在主旨「來自X的」；出貨提醒在內文「買家X」
        mb = re.search(r"來自(.+?)的", subject) or re.search(r"買家\s*([A-Za-z0-9_.]{4,})", body)
        buyer = mb.group(1).strip() if mb else ""
        items = []
        for m in re.finditer(
            r"\n?\s*\d+\.\s*(.+?)\s*(?:選項[:：].*?)?數量[:：]\s*(\d+)\s*價格[:：]\s*NT\$?\s*([\d,]+)",
            body, re.S
        ):
            items.append(EItem(re.sub(r"\s+", " ", m.group(1)).strip(),
                               int(m.group(2)), float(m.group(3).replace(",", ""))))
        if not items:
            return None
        cod = ("貨到付款" in subject) or ("貨到付款" in body)
        # 日期正規化：兼容「2026-06-24」與「2026年3月12日」→ 統一 YYYY-MM-DD（補零）
        dm = re.search(r"訂單日期[:：]\s*(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", body)
        date = f"{dm.group(1)}-{int(dm.group(2)):02d}-{int(dm.group(3)):02d}" if dm else ""
        return EOrder(self.name, order_no, date, items,
                      buyer=buyer,
                      pay_method="貨到付款" if cod else "一般付款",
                      pay_status="未付款" if cod else "已付款")

    def is_existing(self, order, ragic_orders):
        """蝦皮備註=買家帳號，無訂單號 → 以 買家 + 同日 判斷是否已開。"""
        od = re.sub(r"\D", "", order.date)[:8]   # YYYYMMDD
        for r in ragic_orders:
            if r["customer"] == self.customer and r["note"] == order.buyer:
                if re.sub(r"\D", "", r["date"])[:8] == od:
                    return True
        return False
