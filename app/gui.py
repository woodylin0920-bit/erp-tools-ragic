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

try:                          # 拖放（選用）；沒裝 tkinterdnd2 也能跑，只是不能拖
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except Exception:
    HAS_DND = False
    DND_FILES = None

R.NONINTERACTIVE = True   # GUI：金鑰缺失/失效時 raise（不跳 CLI 問答卡死背景緒）

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
    ("系統", ["設定"]),
]


def has_api_key() -> bool:
    return bool(os.environ.get("RAGIC_API_KEY") or R._KEY_FILE.exists())


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

    def run_async(self, work, done, status_label=None, on_error=None):
        """背景執行 work()（回傳值傳給 done()）；Tk 不可跨執行緒，故走 queue 輪詢。
        出錯：呼叫 on_error(e)（若有），否則顯示錯誤；金鑰失效等會走這條。"""
        q = queue.Queue()

        def w():
            try:
                q.put(("ok", work()))
            except BaseException as e:   # 連 SystemExit 都接住（金鑰失效時 CLI 層可能 sys.exit）
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
                if on_error is not None:
                    on_error(val)
                elif status_label is not None:
                    status_label.configure(text=f"載入失敗：{val}")
                else:
                    mbox.showerror("操作失敗", str(val))
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
#  建立出庫單（出貨單拋轉，含拆盒；寫入）
# ════════════════════════════════════════════════════════════
class OutboundScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        import outbound_core as OC
        self.OC = OC
        self.ctx = None
        self.rows = []     # [(rid, label, var)]

        def wh_box(bar):
            self.wh = ctk.CTkOptionMenu(bar, values=["TW01"], width=180, fg_color=BLUE)
            self.wh.pack(side="right")
            ctk.CTkLabel(bar, text="倉庫", text_color=GRAY, font=ctk.CTkFont(size=12)).pack(side="right", padx=6)
        self.toolbar("建立出庫單（出貨單拋轉）", right=wh_box)

        self.status = ctk.CTkLabel(self, text="讀取出貨單與庫存…", text_color=GRAY, font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", padx=26, pady=(12, 4))
        self.search = ctk.CTkEntry(self, placeholder_text="篩選出貨單（單號 / 客戶）")
        self.search.pack(fill="x", padx=24, pady=4)
        self.search.bind("<KeyRelease>", lambda e: self._render_orders())
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", height=240)
        self.scroll.pack(fill="both", expand=True, padx=18, pady=8)

        bar = ctk.CTkFrame(self, height=60, fg_color="#FAFAFB")
        bar.pack(fill="x", side="bottom")
        ctk.CTkButton(bar, text="預覽拆盒", fg_color="#8E8E93", height=36, width=120,
                      command=self._preview).pack(side="right", padx=(8, 24), pady=12)
        self.exec_btn = ctk.CTkButton(bar, text="拋轉並補資料", fg_color=BLUE, height=36, width=150,
                                      font=ctk.CTkFont(size=14, weight="bold"), command=self._execute)
        self.exec_btn.pack(side="right", pady=12)
        self._busy = False

        self.run_async(OC.load_context, self._loaded, self.status)

    def _loaded(self, ctx):
        self.ctx = ctx
        whs = sorted(ctx["warehouses"].keys(), key=lambda w: (0 if w == "TW01" else 1, w))
        self.wh.configure(values=whs)
        self.wh.set("TW01" if "TW01" in whs else (whs[0] if whs else "TW01"))
        self._render_orders()

    def _render_orders(self):
        for w in self.scroll.winfo_children():
            w.destroy()
        self.rows = []
        kw = self.search.get().strip().lower()
        shown = 0
        for c in self.ctx["candidates"]:
            if kw and kw not in c["label"].lower():
                continue
            var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(self.scroll, text=c["label"], variable=var,
                            font=ctk.CTkFont(size=13)).pack(anchor="w", padx=8, pady=2)
            self.rows.append((c["id"], c["label"], var))
            shown += 1
            if shown >= 200:
                break
        self.status.configure(text=f"共 {len(self.ctx['candidates'])} 張出貨單，勾選後可預覽拆盒或拋轉")

    def _selected_ids(self):
        return [rid for rid, _, v in self.rows if v.get()]

    def _plan(self):
        ids = self._selected_ids()
        if not ids:
            mbox.showwarning("提醒", "請至少勾選一張出貨單"); return None, None
        wh = self.wh.get()
        plan = self.OC.break_plan(self.ctx["records"], ids, self.ctx["inventory"], wh)
        return ids, plan

    def _preview(self):
        ids, plan = self._plan()
        if ids is None:
            return
        auto = [p for p in plan if p["status"] == "ok"]
        manual = [p for p in plan if p["status"] == "manual" and max(0, p["need"] - p["have"]) > 0]
        issues = [p for p in plan if p["status"] in ("parent_short", "no_parent", "no_stock")]
        lines = []
        for p in auto:
            lines.append(f"✓ {p['prod']} 客戶{p['need']} → 拆 {p['parent']} 中盒×{p['boxes']}（中盒 {p['parent_qty']}→{p['parent_qty']-p['boxes']}）")
        for p in manual:
            lines.append(f"⚠ {p['prod']} 客戶{p['need']}（非整中盒）→ 用散盒/拆實體 {p['boxes']} 盒")
        for p in issues:
            lines.append(f"⛔ {p['prod']} 客戶{p['need']} 中盒不足/查無 → 人工處理")
        body = "\n".join(lines) if lines else "（無需拆盒）"
        mbox.showinfo("拆盒預覽（未改任何庫存）", f"倉庫 {self.wh.get()}：\n\n{body}")

    def _execute(self):
        ids, plan = self._plan()
        if ids is None:
            return
        wh = self.wh.get()
        merged = self.OC.merge_breakbox(plan)
        nbox = sum(m["boxes"] for m in merged.values())
        # 把拆盒 blocker（中盒不足/查無/非整中盒缺口）攤在確認框，避免被無聲略過
        blockers = [p for p in plan if p["status"] in ("parent_short", "no_parent", "no_stock")]
        manual = [p for p in plan if p["status"] == "manual" and max(0, p["need"] - p["have"]) > 0]
        msg = (f"確定執行？\n\n倉庫：{wh}\n出貨單：{len(ids)} 張\n"
               f"拆盒：{len(merged)} 種中盒、共 {nbox} 盒（會改 20008 庫存）\n"
               f"接著拋轉建立出庫單並自動補欄位。")
        if blockers:
            msg += "\n\n⛔ 下列無法自動處理（中盒不足/查無），仍會繼續拋轉但這些不會拆盒：\n" \
                   + "\n".join(f"  {p['prod']} 客戶{p['need']}" for p in blockers[:6])
        if manual:
            msg += "\n\n⚠ 非整中盒（零頭，需人工拆實體）：\n" \
                   + "\n".join(f"  {p['prod']} 客戶{p['need']}" for p in manual[:6])
        # 多庫存編號：GUI 取第一筆，但明確告知（影響扣哪一批庫存），讓使用者決定要不要改用 CLI
        multi = [(p["prod"], self.ctx["inv_by_wh_prod"].get((wh, p["prod"]), []))
                 for p in self.OC.products_of(self.ctx["records"], ids)]
        multi = [(prod, opts) for prod, opts in multi if len(opts) > 1]
        if multi:
            msg += ("\n\n⚠ 下列商品在此倉有多個庫存編號，GUI 會取第一筆（影響扣哪批）。\n"
                    "  若需指定批號請改用 CLI：\n"
                    + "\n".join(f"  {prod}：{opts}" for prod, opts in multi[:5]))
        if self._busy:
            return
        if not mbox.askyesno("確認執行（會寫入 ERP）", msg):
            return
        # 防連點 + 避免用過期庫存快照覆蓋：執行時禁用按鈕、寫入前重載庫存重算拆盒。
        self._busy = True
        self.exec_btn.configure(state="disabled", text="執行中…")
        self.status.configure(text="執行中（重載庫存→拆盒→拋轉→補欄位）…")

        def work():
            ctx = self.OC.load_context()          # 重載，拿到最新絕對庫存值
            # Preflight：動庫存「之前」先確認出庫單按鈕在。否則拆盒已扣庫存卻建不出出庫單，
            # 重試會重複扣。按鈕缺失就在這裡 raise，庫存完全沒被動。
            if R.ragic_get_action_button_id(R.DELIVERY_ORDER_SHEET, "建立出庫單") is None:
                raise RuntimeError("找不到「建立出庫單」按鈕，未改動任何庫存。請確認 Ragic 表單設定。")
            fresh_plan = self.OC.break_plan(ctx["records"], ids, ctx["inventory"], wh)
            fresh_merged = self.OC.merge_breakbox(fresh_plan)
            prod_inv = {}
            for p in self.OC.products_of(ctx["records"], ids):
                opts = ctx["inv_by_wh_prod"].get((wh, p["prod"]), [])
                if opts:
                    prod_inv[p["prod"]] = opts[0]
            br = self.OC.apply_breakbox(fresh_merged) if fresh_merged else []
            out = self.OC.create_outbound(ctx["records"], ids, wh, prod_inv)
            self.ctx = ctx                          # 更新快照供後續操作
            return br, out
        self.run_async(work, self._executed, on_error=self._exec_error)

    def _exec_error(self, e):
        self._busy = False
        self.exec_btn.configure(state="normal", text="拋轉並補資料")
        self.status.configure(text="執行失敗")
        mbox.showerror("執行失敗", str(e))

    def _executed(self, data):
        self._busy = False
        self.exec_btn.configure(state="normal", text="拋轉並補資料")
        br, out = data
        bk_ok = sum(1 for _, ok, _ in br if ok)
        lines = [f"拆盒：{bk_ok}/{len(br)} 種中盒已改",
                 f"拋轉：觸發 {out['triggered']} 張、新出庫單 {out['new']} 張、補欄位 {out['patched']} 張"]
        if out["msgs"]:
            lines += ["", "提醒："] + out["msgs"]
        self.status.configure(text="完成，請至 Ragic 出庫單頁面確認")
        mbox.showinfo("結果", "\n".join(lines))


# ════════════════════════════════════════════════════════════
#  匯出庫存報表（讀 Ragic、輸出本機 Excel；不寫 ERP）
# ════════════════════════════════════════════════════════════
class ExportScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        import export_core as EC
        self.EC = EC
        self.toolbar("匯出庫存報表（Excel）")

        self.rtype = ctk.CTkSegmentedButton(
            self, values=["客戶現貨報表", "月度庫存金額統計"],
            command=lambda v: self._switch_type(), font=ctk.CTkFont(size=13))
        self.rtype.set("客戶現貨報表")
        self.rtype.pack(anchor="w", padx=26, pady=(16, 6))

        # 現貨報表：倉庫 + 模板
        self.form = ctk.CTkFrame(self, fg_color="transparent")
        self.form.pack(fill="x", padx=26, pady=6)
        ctk.CTkLabel(self.form, text="倉庫", text_color=GRAY, font=ctk.CTkFont(size=13)).grid(row=0, column=0, sticky="w", pady=6)
        self.wh = ctk.CTkOptionMenu(self.form, values=["（讀取中…）"], width=240, fg_color=BLUE)
        self.wh.grid(row=0, column=1, sticky="w", padx=12, pady=6)
        ctk.CTkLabel(self.form, text="模板", text_color=GRAY, font=ctk.CTkFont(size=13)).grid(row=1, column=0, sticky="w", pady=6)
        self.tpls = {t.name: t for t in EC.list_templates()}
        self.tpl = ctk.CTkOptionMenu(self.form, values=list(self.tpls) or ["（無模板，請放入 templates/）"], width=360)
        self.tpl.grid(row=1, column=1, sticky="w", padx=12, pady=6)

        # 月度金額：月份輸入
        self.mform = ctk.CTkFrame(self, fg_color="transparent")
        ctk.CTkLabel(self.mform, text="月份", text_color=GRAY, font=ctk.CTkFont(size=13)).grid(row=0, column=0, sticky="w", pady=6)
        self.month = ctk.CTkEntry(self.mform, width=160, placeholder_text="YYYY-MM（留空＝上個月）")
        self.month.grid(row=0, column=1, sticky="w", padx=12, pady=6)
        ctk.CTkLabel(self.mform, text="全倉分頁、含成本與庫存現金，給會計月結用",
                     text_color=GRAY, font=ctk.CTkFont(size=12)).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))

        ctk.CTkButton(self, text="匯出", fg_color=BLUE, width=120, height=38,
                      font=ctk.CTkFont(size=14, weight="bold"), command=self._export).pack(anchor="w", padx=26, pady=(6, 4))
        self.status = ctk.CTkLabel(self, text="選報表類型與參數後按匯出（輸出到 exports/）",
                                   text_color=GRAY, font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", padx=26, pady=8)
        self.run_async(EC.load_warehouses, self._loaded, self.status)

    def _switch_type(self):
        self.form.pack_forget()
        self.mform.pack_forget()
        target = self.form if self.rtype.get() == "客戶現貨報表" else self.mform
        target.pack(fill="x", padx=26, pady=6, after=self.rtype)

    def _loaded(self, whs):
        order = sorted(whs.keys(), key=lambda w: (0 if w == "TW01" else 1, w))
        self.wh.configure(values=[f"{w}  {whs[w]}" for w in order] or ["（無倉庫）"])
        if order:
            self.wh.set(f"{order[0]}  {whs[order[0]]}")

    def _export(self):
        if self.rtype.get() == "月度庫存金額統計":
            month = self.month.get().strip()
            self.status.configure(text="產生月度庫存金額統計中（全倉、含成本）…")

            def work():
                try:
                    from export_inventory_value import export as ev
                except ImportError:
                    from app.export_inventory_value import export as ev
                return ev(month or None)
            self.run_async(work, lambda p: self._done((p, None, None)), self.status,
                           on_error=lambda e: (self.status.configure(text="匯出失敗"), mbox.showerror("匯出失敗", str(e))))
            return
        if not self.tpls:
            mbox.showwarning("提醒", "templates/ 內沒有模板"); return
        wh = self.wh.get().split("  ")[0].strip()
        tpl = self.tpls.get(self.tpl.get())
        if not tpl:
            mbox.showwarning("提醒", "請選模板"); return
        self.status.configure(text="匯出中（讀庫存、填模板）…")

        def work():
            return self.EC.export_to_template(wh, tpl, R.load_price_index())
        self.run_async(work, self._done, self.status)

    def _done(self, result):
        out_path, filled, skipped = result
        detail = f"填入 {filled} 筆（略過 {skipped} 單盒項目）" if filled is not None else "月度庫存金額統計"
        self.status.configure(text=f"✓ 完成！{detail}")
        if mbox.askyesno("完成", f"已輸出：\n{out_path}\n\n{detail}。要打開資料夾嗎？"):
            import subprocess
            import platform
            folder = str(out_path.parent)
            try:
                if platform.system() == "Darwin":
                    subprocess.run(["open", folder])
                elif platform.system() == "Windows":
                    os.startfile(folder)   # noqa
                else:
                    subprocess.run(["xdg-open", folder])
            except Exception:
                pass


# ════════════════════════════════════════════════════════════
#  新建銷售單（讀客戶 Excel → 開銷貨單；寫入）
# ════════════════════════════════════════════════════════════
class NewSalesScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        import sales_core as SLC
        self.SLC = SLC
        self.preview = []
        self.customers = []
        self.price_index = {}

        def dry(bar):
            self.dry_switch = ctk.CTkSwitch(bar, text="預覽模式（不寫入）", progress_color=GREEN,
                                            font=ctk.CTkFont(size=12))
            self.dry_switch.select()
            self.dry_switch.pack(side="right")
        self.toolbar("新建銷售單（讀客戶 Excel）", right=dry)

        self._picked = None     # 透過「選擇檔案」丟進來的檔（不在 client_order/）

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=24, pady=(14, 2))
        ctk.CTkLabel(top, text="待處理檔案", text_color=GRAY, font=ctk.CTkFont(size=13)).pack(side="left")
        self.file_menu = ctk.CTkOptionMenu(top, values=["（讀取中…）"], width=300,
                                           command=lambda v: self._on_dropdown())
        self.file_menu.pack(side="left", padx=10)
        ctk.CTkButton(top, text="選擇檔案…", fg_color="#8E8E93", width=100,
                      command=self._pick_file).pack(side="left", padx=(0, 6))
        ctk.CTkButton(top, text="解析預覽", fg_color=BLUE, width=100, command=self._do_preview).pack(side="left")

        # 選檔列：顯示選到的檔 + 格式（檔案不在 client_order/ 時需指定）
        pickrow = ctk.CTkFrame(self, fg_color="transparent")
        pickrow.pack(fill="x", padx=24, pady=(0, 4))
        hint = "（或按「選擇檔案」／把 Excel 拖到下方區域）" if HAS_DND else "（或按「選擇檔案」從電腦任意位置選 Excel）"
        self.picked_label = ctk.CTkLabel(pickrow, text=hint, text_color=GRAY, font=ctk.CTkFont(size=12))
        self.picked_label.pack(side="left")
        ctk.CTkLabel(pickrow, text="格式", text_color=GRAY, font=ctk.CTkFont(size=12)).pack(side="left", padx=(14, 2))
        self.fmt = ctk.CTkOptionMenu(pickrow, values=["TRU", "LE", "TEMPLATE"], width=110)
        self.fmt.pack(side="left")

        self.opts = ctk.CTkFrame(self, fg_color="transparent")
        self.opts.pack(fill="x", padx=24, pady=4)
        ctk.CTkLabel(self.opts, text="單別", text_color=GRAY, font=ctk.CTkFont(size=12)).pack(side="left")
        self.otype = ctk.CTkOptionMenu(self.opts, values=["一般訂單", "經銷商", "寄賣訂單"], width=130)
        self.otype.pack(side="left", padx=(4, 14))
        ctk.CTkLabel(self.opts, text="狀態", text_color=GRAY, font=ctk.CTkFont(size=12)).pack(side="left")
        self.ostat = ctk.CTkOptionMenu(self.opts, values=["未出貨", "預接單", "已收款未出貨"], width=130)
        self.ostat.pack(side="left", padx=(4, 14))
        ctk.CTkLabel(self.opts, text="稅率", text_color=GRAY, font=ctk.CTkFont(size=12)).pack(side="left")
        # 與 CLI 一致的有效值：5%(含稅/外加)、(5%)(內含/不計稅)。不可用 "0"（Ragic 不認）
        self.otax = ctk.CTkOptionMenu(self.opts, values=["5%", "(5%)"], width=80)
        self.otax.pack(side="left", padx=4)

        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent",
                                             label_text=("把 Excel 拖到這裡" if HAS_DND else ""))
        self.scroll.pack(fill="both", expand=True, padx=18, pady=8)
        if HAS_DND:
            try:
                self.scroll.drop_target_register(DND_FILES)
                self.scroll.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass
        bar = ctk.CTkFrame(self, height=60, fg_color="#FAFAFB")
        bar.pack(fill="x", side="bottom")
        self.status = ctk.CTkLabel(bar, text="", text_color=GRAY, font=ctk.CTkFont(size=12))
        self.status.pack(side="left", padx=24)
        self.go = ctk.CTkButton(bar, text="開立", fg_color=BLUE, height=36, width=140,
                                font=ctk.CTkFont(size=14, weight="bold"), command=self._go)
        self.go.pack(side="right", padx=24, pady=12)

        self.run_async(lambda: (R.load_price_index(), R.load_customers(), SLC.list_pending()), self._loaded)

    def _loaded(self, data):
        self.price_index, self.customers, files = data
        self.files = {f"{f.parent.name}/{f.name}": f for f in files}
        names = list(self.files) or ["（client_order/ 內無待處理檔案）"]
        self.file_menu.configure(values=names)
        self.file_menu.set(names[0])
        self.status.configure(text=f"找到 {len(files)} 個待處理檔案")

    def _clear_preview(self):
        self.preview = []
        self._preview_file = None
        for w in self.scroll.winfo_children():
            w.destroy()

    def _on_dropdown(self):
        """選了 client_order/ 內的檔 → 清掉「選擇檔案」來源。"""
        self._picked = None
        self.picked_label.configure(text="（或按「選擇檔案」從電腦任意位置丟一個 Excel 進來）")
        self._clear_preview()
        self.status.configure(text="已切換檔案，請按「解析預覽」")

    def _pick_file(self):
        import tkinter.filedialog as fd
        path = fd.askopenfilename(title="選擇客戶訂單 Excel",
                                  filetypes=[("Excel", "*.xlsx"), ("所有檔案", "*.*")])
        if path:
            self._use_picked(path)

    def _on_drop(self, event):
        # 拖放可能一次多檔/含空白路徑，用 tk.splitlist 解析，取第一個 .xlsx
        try:
            paths = self.scroll.tk.splitlist(event.data)
        except Exception:
            paths = [event.data]
        xlsx = [p for p in paths if str(p).lower().endswith(".xlsx")]
        if not xlsx:
            mbox.showwarning("提醒", "請拖入 .xlsx 檔"); return
        self._use_picked(xlsx[0])

    def _use_picked(self, path):
        import pathlib
        p = pathlib.Path(str(path).strip().strip("{}"))   # 去掉 DnD 可能的大括號
        self._picked = p
        self.picked_label.configure(text=f"已選：{p.name}", text_color="#1C1C1E")
        parent = p.parent.name.upper()
        if parent in ("TRU", "LE", "TEMPLATE"):
            self.fmt.set(parent)
        self._clear_preview()
        self.status.configure(text=f"已選檔案，格式={self.fmt.get()}，請按「解析預覽」")

    def _active_source(self):
        """回 (檔路徑, 格式 or None, 是否來自 client_order)。優先用選擇的檔。"""
        if self._picked:
            return self._picked, self.fmt.get(), False
        f = self.files.get(self.file_menu.get())
        return f, None, True   # client_order/ 內的檔，格式由父資料夾推斷

    def _do_preview(self):
        f, fmt, _ = self._active_source()
        if not f:
            mbox.showwarning("提醒", "沒有可解析的檔案"); return
        self._preview_file = f
        self.status.configure(text="解析中…")

        def work():
            return self.SLC.preview_file(f, self.price_index, self.customers, client=fmt)
        self.run_async(work, self._previewed, self.status, on_error=lambda e: (
            self.status.configure(text="解析失敗"), mbox.showerror("解析失敗", str(e))))

    def _previewed(self, preview):
        self.preview = preview
        for w in self.scroll.winfo_children():
            w.destroy()
        for p in preview:
            card = ctk.CTkFrame(self.scroll, border_width=1, border_color=CARD_BORDER, corner_radius=10)
            card.pack(fill="x", padx=6, pady=5)
            cust = p["customer"]["name"] if p["customer"] else "⚠ 對不到客戶（需先建檔）"
            color = ORANGE if p["customer_missing"] else "#1C1C1E"
            ctk.CTkLabel(card, text=f"門市 {p['store']} · PO {p['po']} · {cust}",
                         font=ctk.CTkFont(size=13, weight="bold"), text_color=color).pack(anchor="w", padx=14, pady=(9, 2))
            for it in p["items"]:
                ctk.CTkLabel(card, text=f"   {it['product_code']:<10} {it['product_name'][:18]} {it['unit']} ×{it['quantity']} @ {it['unit_price']:g}",
                             font=ctk.CTkFont(size=12)).pack(anchor="w", padx=14)
            if p["box_notes"]:
                ctk.CTkLabel(card, text="   ⚠ " + "；".join(p["box_notes"]), text_color=ORANGE,
                             font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14)
            if p["ambiguous"]:
                ctk.CTkLabel(card, text="   ⛔ 規格需人工選（多規格），此單請改用 CLI 開立", text_color=RED,
                             font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14)
            if p.get("skipped"):
                ctk.CTkLabel(card, text=f"   ⛔ 有 {p['skipped']} 項商品不在單價表（被略過），此單請改用 CLI 處理", text_color=RED,
                             font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14)
            if not p["items"]:
                ctk.CTkLabel(card, text="   ⛔ 無有效商品，跳過", text_color=RED,
                             font=ctk.CTkFont(size=11)).pack(anchor="w", padx=14)
            ctk.CTkLabel(card, text="", height=4).pack()
        ok = sum(1 for p in preview if self._creatable(p))
        blocked = len(preview) - ok
        self.status.configure(text=f"解析 {len(preview)} 張，可開立 {ok}（{blocked} 張對不到客戶/規格需人工/商品缺漏，跳過）")

    @staticmethod
    def _creatable(p):
        return (not p["customer_missing"] and not p["ambiguous"]
                and not p.get("skipped") and bool(p["items"]))

    def _go(self):
        if not self.preview:
            mbox.showwarning("提醒", "請先解析預覽"); return
        # 排除：對不到客戶、規格歧義、商品缺漏、空單（避免開錯/殘缺單，請用 CLI）
        creatable = [p for p in self.preview if self._creatable(p)]
        if not creatable:
            mbox.showwarning("提醒", "沒有可安全開立的訂單（客戶對不到/規格需人工/商品缺漏，請用 CLI）"); return
        ot, ost, tax = self.otype.get(), self.ostat.get(), self.otax.get()
        if self.dry_switch.get():
            mbox.showinfo("預覽（未寫入）",
                          f"將開立 {len(creatable)} 張銷貨單（單別 {ot} / 狀態 {ost} / 稅率 {tax}）。\n"
                          "關閉「預覽模式」後再按開立才會實際寫入。")
            return
        if not mbox.askyesno("確認開立", f"確定開立 {len(creatable)} 張銷貨單到 Ragic？\n（已自動防重複、帶入 PO#）"):
            return

        self._active_file = getattr(self, "_preview_file", None)   # 用「預覽時」的檔，不看現在下拉

        def work():
            res = []
            for p in creatable:
                res.append(self.SLC.create_order(
                    p["customer"], p["items"], ot, ost, tax,
                    po_number=p["po"], client=p["client"], store=p["store"], commit=True))
            return res
        self.go.configure(state="disabled", text="開立中…")
        self.run_async(work, self._created)

    def _created(self, res):
        self.go.configure(state="normal", text="開立")
        ok = sum(1 for r in res if r["ok"])
        dup = sum(1 for r in res if r.get("dup"))
        fail = [r["msg"] for r in res if not r["ok"] and not r.get("dup")]
        msg = f"完成！{ok}/{len(res)} 張已建立。"
        if dup:
            msg += f"\n防重複跳過 {dup} 張（之前已開過）。"
        if fail:
            msg += "\n\n失敗：\n" + "\n".join(fail[:8])
        # 移檔到 done/ 的條件：無失敗，且整檔每張單都已處理（沒有被擋的列）。
        # 若還有對不到客戶/規格歧義/商品缺漏的單，留著檔案讓人用 CLI 處理，不可移走。
        all_done = all(self._creatable(p) for p in self.preview)
        moved = False
        f = getattr(self, "_active_file", None)
        # 只搬 client_order/ 內的檔到 done/；「選擇檔案」丟進來的外部檔不動（不在電腦亂建 done/）
        from_pending = bool(f) and f in self.files.values()
        if f and from_pending and not fail and (ok or dup) and all_done:
            try:
                import shutil
                done_dir = f.parent / "done"
                done_dir.mkdir(exist_ok=True)
                shutil.move(str(f), str(done_dir / f.name))
                moved = True
            except Exception as e:
                msg += f"\n（移檔到 done/ 失敗：{e}）"
        if moved:
            msg += "\n\n已將檔案移至 done/。"
            self.preview = []
            for w in self.scroll.winfo_children():
                w.destroy()
            self.run_async(self.SLC.list_pending, self._refresh_files)
        elif f and not all_done:
            msg += "\n\n此檔仍有未處理的單（對不到客戶/規格需人工/商品缺漏），檔案保留，請用 CLI 補完。"
        mbox.showinfo("結果", msg)

    def _refresh_files(self, files):
        self.files = {f"{f.parent.name}/{f.name}": f for f in files}
        names = list(self.files) or ["（client_order/ 內無待處理檔案）"]
        self.file_menu.configure(values=names)
        self.file_menu.set(names[0])
        self.status.configure(text=f"剩 {len(files)} 個待處理檔案")


# ════════════════════════════════════════════════════════════
#  尚未搬入的功能（佔位）
# ════════════════════════════════════════════════════════════
class PlaceholderScreen(Screen):
    def __init__(self, master, name):
        super().__init__(master)
        self.toolbar(name)
        ctk.CTkLabel(self, text=f"「{name}」尚未搬入 GUI，目前請用 CLI（./start.command）。",
                     text_color=GRAY, font=ctk.CTkFont(size=14)).pack(pady=40)


# ════════════════════════════════════════════════════════════
#  設定（Ragic API Key）
# ════════════════════════════════════════════════════════════
class SettingsScreen(Screen):
    def __init__(self, master):
        super().__init__(master)
        self.toolbar("設定")
        wrap = ctk.CTkFrame(self, fg_color="transparent")
        wrap.pack(fill="x", padx=26, pady=20)

        ctk.CTkLabel(wrap, text="Ragic API Key", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(wrap, text="可讀寫正式 ERP。從 Ragic 個人設定頁取得（已是 Base64 格式，整段貼上即可）。",
                     text_color=GRAY, font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(2, 2))
        ctk.CTkLabel(wrap, text=f"儲存位置：{R._KEY_FILE}", text_color="#B0B0B5",
                     font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(0, 10))

        self.status = ctk.CTkLabel(wrap, text="", font=ctk.CTkFont(size=13))
        self.status.pack(anchor="w", pady=(0, 8))

        self.entry = ctk.CTkEntry(wrap, show="•", width=560, placeholder_text="貼上 Ragic API Key")
        self.entry.pack(anchor="w", pady=4)

        btns = ctk.CTkFrame(wrap, fg_color="transparent")
        btns.pack(anchor="w", pady=(10, 0))
        ctk.CTkButton(btns, text="儲存", fg_color=BLUE, width=90, command=self._save).pack(side="left")
        ctk.CTkButton(btns, text="測試連線", fg_color="#8E8E93", width=100, command=self._test).pack(side="left", padx=8)
        self.show_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(btns, text="顯示", variable=self.show_var, width=60,
                        command=self._toggle_show, font=ctk.CTkFont(size=12)).pack(side="left", padx=8)

        self._refresh_status()

    def _refresh_status(self):
        if has_api_key():
            self.status.configure(text="● 目前已設定 API Key（如需更換，貼上新的後按儲存）", text_color=GREEN)
        else:
            self.status.configure(text="● 尚未設定 API Key —— 請貼上後按儲存才能使用各功能", text_color=ORANGE)

    def _toggle_show(self):
        self.entry.configure(show="" if self.show_var.get() else "•")

    def _save(self):
        key = self.entry.get().strip()
        if not key:
            mbox.showwarning("提醒", "請先貼上 API Key"); return
        try:
            R._KEY_FILE.write_text(key, encoding="utf-8")
            os.environ["RAGIC_API_KEY"] = key
        except Exception as e:
            mbox.showerror("儲存失敗", str(e)); return
        self.entry.delete(0, "end")
        self._refresh_status()
        mbox.showinfo("已儲存", "API Key 已儲存，立即生效。建議按「測試連線」確認。")

    def _test(self):
        if not has_api_key():
            mbox.showwarning("提醒", "尚未設定 API Key"); return
        self.status.configure(text="測試連線中…", text_color=GRAY)
        self.run_async(lambda: len(R.load_customers()), self._test_done, self.status)

    def _test_done(self, n):
        if n > 0:
            self.status.configure(text=f"✓ 連線成功（讀到 {n} 筆客戶），API Key 有效", text_color=GREEN)
        else:
            self.status.configure(text="⚠ 連線回傳 0 筆，請確認 API Key 是否正確/有權限", text_color=ORANGE)


SCREENS = {
    "新建銷售單": NewSalesScreen,
    "批次發樣": SampleOrderScreen,
    "在途查詢": InTransitScreen,
    "電商對帳": EcomScreen,
    "建立出貨單": DeliveryScreen,
    "建立出庫單": OutboundScreen,
    "匯出庫存報表": ExportScreen,
    "設定": SettingsScreen,
}


# ════════════════════════════════════════════════════════════
#  主視窗
# ════════════════════════════════════════════════════════════
_APP_BASES = (ctk.CTk, TkinterDnD.DnDWrapper) if HAS_DND else (ctk.CTk,)


class App(*_APP_BASES):
    def __init__(self):
        super().__init__()
        if HAS_DND:
            self.TkdndVersion = TkinterDnD._require(self)   # 啟用整個視窗的拖放能力
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
        # 首次（或 Windows 新機）尚未設定金鑰 → 先進設定，避免功能畫面去問金鑰而卡住
        if has_api_key():
            self.show("批次發樣")
        else:
            self.show("設定")
            self.after(300, lambda: mbox.showinfo(
                "歡迎", "第一次使用：請先在「設定」貼上 Ragic API Key 並儲存，才能使用各功能。"))

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
        # 沒設金鑰時，其他功能會去問金鑰(questionary)→GUI 內回 None→sys.exit 關掉程式。
        # 故無金鑰時一律導到「設定」，逼先填 key。
        if label != "設定" and not has_api_key():
            mbox.showwarning("尚未設定", "請先到「設定」填入 Ragic API Key。")
            label = "設定"
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
