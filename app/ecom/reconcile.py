"""電商對帳補漏 dry-run（唯讀，不寫入 Ragic、不更動信箱）。

用法：
    python3 app/ecom/reconcile.py shopstore [--limit N]

流程：讀該平台「新訂單」信 → 比對 Ragic 銷貨單『備註』是否已開
      → 列出「漏開」訂單，並對應到 Ragic 商品（先對照表、新品才模糊）。
此為唯讀預覽；確認後實際開單為後續功能。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # app/

from ecom import core                       # noqa: E402
from ecom.platforms import PLATFORMS        # noqa: E402

_SRC_LABEL = {"map": "對照表", "fuzzy": "模糊", "ambiguous": "多候選需確認", "none": "對不到需補"}


def main():
    args = sys.argv[1:]
    name = args[0] if args and not args[0].startswith("-") else "shopstore"
    limit = None
    if "--limit" in args:
        limit = int(args[args.index("--limit") + 1])

    plat = PLATFORMS.get(name)
    if not plat:
        print(f"不支援的平台：{name}（可用：{', '.join(PLATFORMS)}）")
        return

    print(f"=== {name} 對帳補漏（dry-run，唯讀，未寫入 Ragic）===")
    done, missing = core.reconcile(plat, limit=limit)
    total = len(done) + len(missing)
    print(f"Email 新訂單 {total} 張 ｜ ✅ Ragic 已開 {len(done)} ｜ 📥 漏開 {len(missing)}\n")

    hist = core.historical_prices()   # 參考歷史訂單單價
    need_review = 0
    for o in missing:
        pay = f"{o.pay_method or '?'}/{o.pay_status or '?'}"
        flag = "  🔴待取貨(未領風險)" if o.is_cod_pending else ""
        print(f"📥 訂單 {o.order_no}  {o.date}  → 客戶「{plat.customer}」單別「{plat.order_type}」")
        print(f"    買家:{o.buyer or '?'} ｜ 付款:{pay} ｜ 運費:{o.fee:g}{flag}")
        for it in o.items:
            code, prod, src = core.match_product(name, it.title)
            if code:
                pname = (prod or {}).get("商品名稱", "?")
                hp = hist.get(code)
                note = ""
                if hp and str(int(float(hp))) != str(int(it.price)):
                    note = f"  ⚠售價≠歷史({hp})"
                print(f"    ✅ {code:<10} {pname[:22]:<22} ×{it.qty} @ {it.price:g}  [{_SRC_LABEL.get(src, src)}]{note}")
            else:
                need_review += 1
                print(f"    ❓ 對不到「{it.title[:28]}」×{it.qty} @ {it.price:g}  [{_SRC_LABEL.get(src, src)}]")
    if need_review:
        print(f"\n⚠ {need_review} 個品項需補對照表（product_map.json）後才能開單")
    print("\n（dry-run 完成：未寫入 Ragic、未更動信箱）")


if __name__ == "__main__":
    main()
