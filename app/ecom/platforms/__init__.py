"""平台註冊表。新增平台在此加入即可。"""
from .shopee import Shopee
from .shopstore import ShopStore

PLATFORMS = {p.name: p for p in [ShopStore(), Shopee()]}
