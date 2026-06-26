"""共用核心：商品比對、Ragic 讀取、對帳補漏。

★ 全部唯讀（GET）★ —— dry-run 不會寫入 Ragic、不更動信箱。

商品比對策略：先查 product_map.json（由歷史訂單回推、可靠），
新品才用模糊比對（IP 核心詞 + 單位）暫對，對不到就標記待補。
對帳補漏：讀平台新訂單信 → 比對 Ragic 銷貨單『備註』是否已開 → 列出漏開。
"""
import json
import os
import re
import urllib.request
from pathlib import Path

from . import mailbox

RAGIC_BASE = "https://ap12.ragic.com"
RAGIC_ACCOUNT = "toybebop"
SALES_ORDER_SHEET = "ragicsales-order-management/20001"
PRODUCT_PRICE_SHEET = "ragicsales-order-management/20006"
_KEY = os.path.expanduser("~/.boptoys-ai_key")
_MAP = Path(__file__).parent / "product_map.json"

# 平台標題的雜訊詞（比對前去除），保留可辨識 IP 的核心
_NOISE = ["此為預購商品", "預購", "系列", "毛絨盲盒", "毛茸茸", "小盲盒", "盲盒",
          "手辦", "公仔", "萌粒", "國際版", "正版", "官方", "台灣現貨", "收藏",
          "送禮", "BOPTOYS", "Boptoys", "整中", "整箱", "單盒", "中盒", "盒",
          "款", "限定", "透色版", "透色", "mini"]


def _ragic_get(sheet: str, extra: str = "") -> dict:
    key = open(_KEY).read().strip()
    url = f"{RAGIC_BASE}/{RAGIC_ACCOUNT}/{sheet}?api{extra}"
    req = urllib.request.Request(url, headers={"Authorization": "Basic " + key})
    return json.load(urllib.request.urlopen(req, timeout=90))


def load_product_map() -> dict:
    return json.load(open(_MAP, encoding="utf-8"))


_price_cache = None


def _price_index() -> list:
    global _price_cache
    if _price_cache is None:
        _price_cache = list(_ragic_get(PRODUCT_PRICE_SHEET, "&listing=true").values())
    return _price_cache


def _norm(s: str) -> str:
    return re.sub(r"[\s・·\-_（）()｜|]", "", str(s))


def _core(title: str) -> str:
    t = title.split("｜")[0].split("|")[0]
    t = re.sub(r"[（(].*?[)）]", "", t)
    t = re.sub(r"^[\*※\s]+", "", t)
    for n in _NOISE:
        t = t.replace(n, "")
    return re.sub(r"[\s・·\-_]", "", t)


def _unit(title: str) -> str:
    if "整中" in title or "中盒" in title:
        return "中盒"
    if "整箱" in title:
        return "整箱"
    return "單盒"


def _product_by_code(code: str):
    for r in _price_index():
        if str(r.get("商品單價代號", "")).strip() == code:
            return r
    return None


def match_product(platform: str, title: str):
    """回 (code, product_dict|None, source)。
    source: map(對照表) / fuzzy(模糊命中) / ambiguous(多候選) / none(對不到)。"""
    pmap = load_product_map().get(platform, {})
    if title in pmap:
        code = pmap[title]
        return code, _product_by_code(code), "map"
    key, unit = _core(title), _unit(title)
    hits = [r for r in _price_index()
            if key and key in _norm(r.get("商品名稱", "")) and unit in str(r.get("單位", ""))]
    if len(hits) == 1:
        return str(hits[0].get("商品單價代號", "")), hits[0], "fuzzy"
    return None, None, ("ambiguous" if hits else "none")


def existing_orders() -> list:
    """Ragic 銷貨單清單（備註/日期/金額/客戶），供各平台判斷是否已開。"""
    recs = list(_ragic_get(SALES_ORDER_SHEET, "&listing=true").values())
    return [{
        "note": str(r.get("備註", "")).strip(),
        "date": str(r.get("訂單日期", "")).strip(),
        "total": str(r.get("總金額(含稅)", "")).strip(),
        "customer": str(r.get("客戶名稱", "")).strip(),
    } for r in recs]


def reconcile(platform_obj, limit=None):
    """對帳補漏（唯讀）：讀平台新訂單信 vs Ragic 已開 → 回 (done, missing)。
    done/missing 皆為 EOrder list。"""
    existing = existing_orders()
    M = mailbox.connect(platform_obj.mailbox_user, platform_obj.mailbox_pw_file)
    try:
        uids = mailbox.search(M, platform_obj.sender_query)
        if limit:
            uids = uids[-limit:]
        done, missing, seen = [], [], set()
        for u in uids:
            subject, body = mailbox.fetch(M, u)
            order = platform_obj.parse_order(subject, body)
            if not order or order.order_no in seen:
                continue
            seen.add(order.order_no)
            (done if platform_obj.is_existing(order, existing) else missing).append(order)
    finally:
        M.logout()
    return done, missing
