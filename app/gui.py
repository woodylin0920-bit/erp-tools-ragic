"""潮玩波普 ERP — 桌面 GUI（customtkinter）。

多畫面架構：左側功能列點選 → 右側切換畫面。各畫面邏輯走既有純函式
（sample_core / ragic_upload / in_transit_query / ecom），與 CLI 共用。

執行：python3 app/gui.py
需求：Python 含 tkinter（macOS: brew install python-tk）+ pip 裝 customtkinter。
安全：寫入動作一律跳確認框；批次發樣有「預覽模式」開關。
"""
import os
import queue
import sys
import threading
import tkinter.messagebox as mbox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk   # noqa: E402
import sample_core as SC      # noqa: E402
import ragic_upload as R      # noqa: E402

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BLUE = "#007AFF"
GREEN = "#34C759"
RED = "#FF3B30"
ORANGE = "#FF9500"
GRAY = "#6C6C70"
CARD_BORDER = "#E5E5EA"

NAV = [
    ("開單", ["新建銷售單", "批次發樣"]),
    ("拋轉", ["建立出貨單", "建立出庫單"]),
    ("查詢與對帳", ["匯出庫存報表", "在途查詢", "電商對帳"]),
]


# ════════════════════════════════════════════════════════════
#  基底畫面：提供非同步載入（背景取資料、主執行緒更新 UI）
# ════════════════════════════════════════════════════════════
class Screen(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master, fg_color="#FFFFFF", corner_radius=0)

    def toolbar(self, title, right=None):
        bar = ctk.CTkFrame(self, height=52, fg_color="transparent")
        bar.pack(fill="x", padx=22, pady=(14, 0))
        ctk.CTkLabel(bar, text=title, font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        if right is not None:
            right(bar)
        return bar

    def run_async(self, work, done, status_label=None):
        """背景執行 work()（回傳值傳給 done()）；Tk 不可跨執行緒，故走 queue 輪詢。"""
        q = queue.Queue()

        def w():
            try:
                q.put(("ok", work()))
            except Exception as e:
                q.put(("err", e))
        threading.Thread(target=w, daemon=True).start()

        def poll():
            try:
                kind, val = q.get_nowait()
            except queue.Empty:
                self.after(120, poll)
                return
            if kind == "ok":
                done(val)
            else:
                if status_label is not None:
                    status_label.configure(text=f"載入失敗：{val}")
                else:
                    mbox.showerror("載入失敗", str(val))
        poll()


# ════════════════════════════════════════════════════════════
#  批次發樣
# ════════════════════════════════════════════════════════════
class SampleOrderScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        self.all_customers = []
        self.cust_vars = {}
        self.chosen = set()
        self.combos = {}

        def dry(bar):
            self.dry_switch = ctk.CTkSwitch(bar, text="預覽模式（不寫入）", progress_color=GREEN,
                                            font=ctk.CTkFont(size=12))
            self.dry_switch.select()
            self.dry_switch.pack(side="right")
        self.toolbar("批次發樣", right=dry)

        self.seg = ctk.CTkSegmentedButton(self, values=SC.ORDER_TYPES, font=ctk.CTkFont(size=13))
        self.seg.set(SC.ORDER_TYPES[0])
        self.seg.pack(anchor="w", padx=24, pady=(16, 8))

        cols = ctk.CTkFrame(self, fg_color="transparent")
        cols.pack(fill="both", expand=True, padx=24, pady=8)
        cols.grid_columnconfigure((0, 1), weight=1, uniform="c")
        cols.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(cols, border_width=1, border_color=CARD_BORDER, corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(left, text="樣品組合", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=GRAY).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        self.combo_menu = ctk.CTkOptionMenu(left, values=["（讀取中…）"], command=self._on_combo,
                                            fg_color=BLUE)
        self.combo_menu.grid(row=1, column=0, sticky="ew", padx=16, pady=4)
        self.combo_box = ctk.CTkTextbox(left, height=180, font=ctk.CTkFont(size=13))
        self.combo_box.grid(row=2, column=0, sticky="nsew", padx=16, pady=(8, 14))

        right = ctk.CTkFrame(cols, border_width=1, border_color=CARD_BORDER, corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=1)
        self.cust_hd = ctk.CTkLabel(right, text="發給客戶 · 已選 0",
                                    font=ctk.CTkFont(size=13, weight="bold"), text_color=GRAY)
        self.cust_hd.grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        self.search = ctk.CTkEntry(right, placeholder_text="搜尋客戶（中文 / 代號）")
        self.search.grid(row=1, column=0, sticky="ew", padx=16, pady=4)
        self.search.bind("<KeyRelease>", lambda e: self._refresh_custs())
        self.cust_scroll = ctk.CTkScrollableFrame(right, fg_color="transparent")
        self.cust_scroll.grid(row=2, column=0, sticky="nsew", padx=8, pady=(6, 12))
        self.cust_scroll.grid_columnconfigure(0, weight=1)

        bar = ctk.CTkFrame(self, height=64, fg_color="#FAFAFB")
        bar.pack(fill="x", side="bottom")
        bar.grid_columnconfigure(0, weight=1)
        self.summary = ctk.CTkLabel(bar, text="", text_color=GRAY, font=ctk.CTkFont(size=13))
        self.summary.grid(row=0, column=0, sticky="w", padx=26, pady=14)
        self.go_btn = ctk.CTkButton(bar, text="預覽 0 張單", width=160, height=38, corner_radius=9,
                                    fg_color=BLUE, font=ctk.CTkFont(size=14, weight="bold"),
                                    command=self._go)
        self.go_btn.grid(row=0, column=1, sticky="e", padx=26, pady=12)

        self.seg.configure(command=lambda v: self._update_summary())
        self.dry_switch.configure(command=lambda: self._update_summary())
        self._update_summary()
        self.run_async(lambda: (SC.load_customers(), SC.load_combos()), self._loaded)

    def _loaded(self, data):
        custs, combos = data
        self.all_customers = [c for c in custs if c.get("code")]
        self.combos = combos
        names = list(combos.keys()) or ["（尚無組合，請先用 CLI 建立）"]
        self.combo_menu.configure(values=names)
        self.combo_menu.set(names[0])
        self._on_combo(names[0])
        self._refresh_custs()

    def _on_combo(self, name):
        items = self.combos.get(name, [])
        self.combo_box.delete("1.0", "end")
        self.combo_box.insert("end", "\n".join(
            f"{it['code']:<12} {it['name'][:18]}  ×{it['qty']}" for it in items) or "（此組合無品項）")
        self._update_summary()

    def _refresh_custs(self):
        for w in self.cust_scroll.winfo_children():
            w.destroy()
        kw = self.search.get().strip().lower()
        subset = [c for c in self.all_customers
                  if not kw or kw in c["name"].lower() or kw in c["code"].lower()][:200]
        for c in subset:
            var = self.cust_vars.get(c["code"])
            if var is None:
                var = ctk.BooleanVar(value=c["code"] in self.chosen)
                self.cust_vars[c["code"]] = var
            ctk.CTkCheckBox(self.cust_scroll, text=f"{c['name']}｜{c['code']}", variable=var,
                            font=ctk.CTkFont(size=13),
                            command=lambda code=c["code"]: self._toggle(code)).pack(anchor="w", padx=6, pady=2)
        self.cust_hd.configure(text=f"發給客戶 · 已選 {len(self.chosen)}")

    def _toggle(self, code):
        (self.chosen.add if self.cust_vars[code].get() else self.chosen.discard)(code)
        self.cust_hd.configure(text=f"發給客戶 · 已選 {len(self.chosen)}")
        self._update_summary()

    def _sel_custs(self):
        by = {c["code"]: c for c in self.all_customers}
        return [by[cc] for cc in self.chosen if cc in by]

    def _combo_items(self):
        return self.combos.get(self.combo_menu.get(), [])

    def _update_summary(self):
        n = len(self.chosen)
        self.summary.configure(text=f"將開立 {n} 張「{self.seg.get()}」單 · 單價全 0 · 狀態未出貨")
        self.go_btn.configure(text=f"{'預覽' if self.dry_switch.get() else '開立'} {n} 張單")

    def _go(self):
        combo, custs, ot = self._combo_items(), self._sel_custs(), self.seg.get()
        if not combo:
            mbox.showwarning("提醒", "請先選一個有品項的組合"); return
        if not custs:
            mbox.showwarning("提醒", "請至少勾選一個客戶"); return
        if self.dry_switch.get():
            res = SC.create_sample_orders(ot, combo, custs, commit=False)
            lines = "\n".join(f"  {r['customer']['name']}（{r['customer']['code']}）" for r in res)
            mbox.showinfo("預覽（未寫入）",
                          f"將開立 {len(res)} 張「{ot}」單，單價全 0、狀態未出貨：\n\n{lines}\n\n"
                          "關閉「預覽模式」後再按開立才會實際寫入。")
            return
        if not mbox.askyesno("確認開立", f"確定開立 {len(custs)} 張「{ot}」單到 Ragic？"):
            return
        res = SC.create_sample_orders(ot, combo, custs, commit=True)
        ok = sum(1 for r in res if r["ok"])
        fail = [f"{r['customer']['name']}：{r['msg']}" for r in res if not r["ok"]]
        mbox.showinfo("結果", f"完成！{ok}/{len(res)} 張已建立。" +
                      ("\n\n失敗：\n" + "\n".join(fail) if fail else ""))


# ════════════════════════════════════════════════════════════
#  在途查詢（唯讀）
# ════════════════════════════════════════════════════════════
class InTransitScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        import in_transit_query as IT
        self.IT = IT
        self.pos = []
        self.toolbar("在途查詢（採購單未到貨）")
        self.search = ctk.CTkEntry(self, placeholder_text="篩選商品（編號 / 名稱關鍵字）")
        self.search.pack(fill="x", padx=24, pady=(14, 6))
        self.search.bind("<KeyRelease>", lambda e: self._render())
        self.status = ctk.CTkLabel(self, text="讀取採購單中…", text_color=GRAY, font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", padx=26)
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=18, pady=10)
        self.run_async(lambda: IT.collect_in_transit(), self._loaded, self.status)

    def _loaded(self, pos):
        self.pos = pos
        self._render()

    def _render(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        kw = self.search.get().strip().lower()
        shown = 0
        for po in self.pos:
            items = [it for it in po["items"]
                     if not kw or kw in it["prod"].lower() or kw in it["name"].lower()]
            if not items:
                continue
            shown += 1
            card = ctk.CTkFrame(self.scroll, border_width=1, border_color=CARD_BORDER, corner_radius=10)
            card.pack(fill="x", padx=6, pady=5)
            ctk.CTkLabel(card, text=f"{po['po_no']} · {po['vendor']} · {po['date']}",
                         font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(9, 2))
            for it in items:
                ctk.CTkLabel(card, text=f"   {it['prod']:<10} {it['name'][:22]}  規格{it['spec']}  未到 {it['qty']} {it['unit']}",
                             font=ctk.CTkFont(size=12), text_color="#1C1C1E").pack(anchor="w", padx=14)
            ctk.CTkLabel(card, text="", height=4).pack()
        self.status.configure(text=f"共 {shown} 張採購單有未到貨商品" if self.pos else "目前無在途採購單")


# ════════════════════════════════════════════════════════════
#  電商對帳（唯讀）
# ════════════════════════════════════════════════════════════
class EcomScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        try:
            from ecom import core as ecore
            from ecom.platforms import PLATFORMS
        except ImportError:
            from app.ecom import core as ecore
            from app.ecom.platforms import PLATFORMS
        self.ecore, self.PLATFORMS = ecore, PLATFORMS
        self.toolbar("電商訂單對帳（唯讀，不寫入）")

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=24, pady=(14, 6))
        self.plat = ctk.CTkOptionMenu(top, values=list(PLATFORMS.keys()), fg_color=BLUE, width=160)
        self.plat.pack(side="left")
        self.run_btn = ctk.CTkButton(top, text="開始對帳", fg_color=BLUE, width=120, command=self._run)
        self.run_btn.pack(side="left", padx=10)
        self.status = ctk.CTkLabel(self, text="選平台後按「開始對帳」（讀信比對，需數秒）",
                                   text_color=GRAY, font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", padx=26)
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=18, pady=10)

    def _run(self):
        name = self.plat.get()
        plat = self.PLATFORMS.get(name)
        if not plat:
            return
        self.run_btn.configure(state="disabled", text="對帳中…")
        self.status.configure(text=f"讀取 {name} 訂單信並比對 Ragic…")
        self.run_async(lambda: self.ecore.reconcile(plat), lambda r: self._done(name, r), self.status)

    def _done(self, name, result):
        self.run_btn.configure(state="normal", text="開始對帳")
        done, missing = result
        for w in self.scroll.winfo_children():
            w.destroy()
        self.status.configure(text=f"{name}：Email 訂單 {len(done)+len(missing)} 張 · Ragic 已開 {len(done)} · 漏開 {len(missing)}")
        for o in missing:
            card = ctk.CTkFrame(self.scroll, border_width=1, border_color=CARD_BORDER, corner_radius=10)
            card.pack(fill="x", padx=6, pady=5)
            flag = "  ⚠待取貨" if getattr(o, "is_cod_pending", False) else ""
            ctk.CTkLabel(card, text=f"漏開 · 訂單 {o.order_no} · {o.date} · 買家 {o.buyer or '-'}{flag}",
                         font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=14, pady=(9, 2))
            for it in o.items:
                ctk.CTkLabel(card, text=f"   {it.title[:30]} ×{it.qty} @ {it.price:g}",
                             font=ctk.CTkFont(size=12)).pack(anchor="w", padx=14)
            ctk.CTkLabel(card, text="", height=4).pack()
        if not missing:
            ctk.CTkLabel(self.scroll, text="沒有漏開的訂單 ✓", text_color=GREEN,
                         font=ctk.CTkFont(size=14)).pack(pady=20)


# ════════════════════════════════════════════════════════════
#  建立出貨單（銷貨單拋轉，寫入）
# ════════════════════════════════════════════════════════════
class DeliveryScreen(Screen):
    TARGET = {"未出貨", "預接單", "已收款未出貨"}

    def __init__(self, master):
        super().__init__(master)
        self.rows = []          # [(rid, label, var)]
        self.toolbar("建立出貨單（銷貨單拋轉）")
        self.status = ctk.CTkLabel(self, text="讀取待拋轉銷貨單…", text_color=GRAY, font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", padx=26, pady=(12, 4))
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll.pack(fill="both", expand=True, padx=18, pady=8)
        bar = ctk.CTkFrame(self, height=60, fg_color="#FAFAFB")
        bar.pack(fill="x", side="bottom")
        self.go = ctk.CTkButton(bar, text="建立出貨單", fg_color=BLUE, height=36, width=150,
                                font=ctk.CTkFont(size=14, weight="bold"), command=self._go)
        self.go.pack(side="right", padx=24, pady=12)
        self.run_async(self._load, self._loaded, self.status)

    def _load(self):
        recs = R.ragic_get(R.SALES_ORDER_SHEET)
        out = []
        for rid, rec in recs.items():
            if str(rec.get("訂單狀態", "")).strip() in self.TARGET:
                out.append((rid, f"{rec.get('訂單編號','?')}  {rec.get('客戶名稱','?')}  "
                                 f"{rec.get('訂單日期','?')}  [{rec.get('訂單狀態','')}]"))
        return out

    def _loaded(self, items):
        for rid, label in items:
            var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(self.scroll, text=label, variable=var,
                            font=ctk.CTkFont(size=13)).pack(anchor="w", padx=8, pady=2)
            self.rows.append((rid, label, var))
        self.status.configure(text=f"共 {len(items)} 張待拋轉（未出貨/預接單/已收款未出貨），勾選後按建立")

    def _go(self):
        sel = [(rid, lbl) for rid, lbl, v in self.rows if v.get()]
        if not sel:
            mbox.showwarning("提醒", "請至少勾選一張銷貨單"); return
        if not mbox.askyesno("確認", f"確定把 {len(sel)} 張銷貨單拋轉成出貨單？"):
            return
        self.go.configure(state="disabled", text="拋轉中…")

        def work():
            bid = R.ragic_get_action_button_id(R.SALES_ORDER_SHEET, "建立出貨單")
            if bid is None:
                raise RuntimeError("找不到「建立出貨單」按鈕")
            ok = 0
            for rid, _ in sel:
                res = R.ragic_trigger_button(R.SALES_ORDER_SHEET, rid, bid)
                if res.get("status") == "SUCCESS":
                    ok += 1
            return ok
        self.run_async(work, lambda ok: self._done(ok, len(sel)))

    def _done(self, ok, total):
        self.go.configure(state="normal", text="建立出貨單")
        mbox.showinfo("結果", f"完成！{ok}/{total} 張出貨單已建立。\n請至 Ragic 確認。")


# ════════════════════════════════════════════════════════════
#  尚未搬入的功能（佔位）
# ════════════════════════════════════════════════════════════
class PlaceholderScreen(Screen):
    def __init__(self, master, name):
        super().__init__(master)
        self.toolbar(name)
        ctk.CTkLabel(self, text=f"「{name}」尚未搬入 GUI，目前請用 CLI（./start.command）。",
                     text_color=GRAY, font=ctk.CTkFont(size=14)).pack(pady=40)


SCREENS = {
    "批次發樣": SampleOrderScreen,
    "在途查詢": InTransitScreen,
    "電商對帳": EcomScreen,
    "建立出貨單": DeliveryScreen,
}


# ════════════════════════════════════════════════════════════
#  主視窗
# ════════════════════════════════════════════════════════════
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("潮玩波普 ERP")
        self.geometry("1080x700")
        self.minsize(940, 600)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.nav_btns = {}
        self.current = None
        self._build_sidebar()
        self.content_holder = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=0)
        self.content_holder.grid(row=0, column=1, sticky="nsew")
        self.content_holder.grid_columnconfigure(0, weight=1)
        self.content_holder.grid_rowconfigure(0, weight=1)
        self.show("批次發樣")

    def _build_sidebar(self):
        bar = ctk.CTkFrame(self, width=248, corner_radius=0, fg_color="#F4F4F6")
        bar.grid(row=0, column=0, sticky="nsew")
        bar.grid_propagate(False)
        ctk.CTkLabel(bar, text="潮玩波普 ERP", font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=GRAY).pack(anchor="w", padx=20, pady=(22, 12))
        for sec, items in NAV:
            ctk.CTkLabel(bar, text=sec, font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="#B0B0B5").pack(anchor="w", padx=20, pady=(10, 2))
            for label in items:
                btn = ctk.CTkButton(bar, text=label, anchor="w", height=34, corner_radius=7,
                                    fg_color="transparent", text_color="#1C1C1E",
                                    hover_color="#E5E5EA", font=ctk.CTkFont(size=14),
                                    command=lambda l=label: self.show(l))
                btn.pack(fill="x", padx=10, pady=1)
                self.nav_btns[label] = btn

    def show(self, label):
        for lb, btn in self.nav_btns.items():
            active = lb == label
            btn.configure(fg_color=BLUE if active else "transparent",
                          text_color="#FFFFFF" if active else "#1C1C1E",
                          hover_color=BLUE if active else "#E5E5EA")
        if self.current is not None:
            self.current.destroy()
        cls = SCREENS.get(label)
        self.current = cls(self.content_holder) if cls else PlaceholderScreen(self.content_holder, label)
        self.current.grid(row=0, column=0, sticky="nsew")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
