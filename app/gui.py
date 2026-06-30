"""潮玩波普 ERP — 桌面 GUI（customtkinter）。

階段 2 雛形：先實作「批次發樣」一個功能，驗證桌面操作質感。
其餘功能為佔位，之後逐個搬入。所有邏輯走 sample_core（與 CLI 共用）。

執行：python3 app/gui.py
需求：Python 需含 tkinter（macOS: brew install python-tk）；pip 裝 customtkinter。
安全：右上「預覽模式」開＝只預覽不寫入；關閉並按開立才會 POST，且跳確認框。
"""
import os
import queue
import sys
import threading
import tkinter.messagebox as mbox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk   # noqa: E402
import sample_core as SC      # noqa: E402

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BLUE = "#007AFF"
GREEN = "#34C759"
GRAY = "#6C6C70"

NAV = [
    ("開單", [("📄 新建銷售單", False), ("🎁 批次發樣", True)]),
    ("拋轉", [("🚚 建立出貨單", False), ("📦 建立出庫單", False)]),
    ("查詢與對帳", [("📊 匯出庫存報表", False), ("🚢 在途查詢", False), ("🛒 電商對帳", False)]),
]


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("潮玩波普 ERP")
        self.geometry("1080x700")
        self.minsize(940, 600)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        # 內容區：先放批次發樣
        self.content = SampleOrderFrame(self)
        self.content.grid(row=0, column=1, sticky="nsew")

    def _build_sidebar(self):
        bar = ctk.CTkFrame(self, width=248, corner_radius=0, fg_color="#F4F4F6")
        bar.grid(row=0, column=0, sticky="nsew")
        bar.grid_propagate(False)
        ctk.CTkLabel(bar, text="潮玩波普 ERP", font=ctk.CTkFont(size=15, weight="bold"),
                     text_color=GRAY).pack(anchor="w", padx=20, pady=(22, 12))
        for sec, items in NAV:
            ctk.CTkLabel(bar, text=sec, font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="#B0B0B5").pack(anchor="w", padx=20, pady=(10, 2))
            for label, active in items:
                btn = ctk.CTkButton(
                    bar, text=label, anchor="w", height=34, corner_radius=7,
                    fg_color=BLUE if active else "transparent",
                    text_color="#FFFFFF" if active else "#1C1C1E",
                    hover_color="#E5E5EA" if not active else BLUE,
                    font=ctk.CTkFont(size=14),
                    command=(lambda l=label: self._nav(l)))
                btn.pack(fill="x", padx=10, pady=1)
        # 底部：預覽模式總開關（與內容區同步）
        foot = ctk.CTkFrame(bar, fg_color="transparent")
        foot.pack(side="bottom", fill="x", padx=16, pady=16)
        ctk.CTkLabel(foot, text="預覽模式（不寫入）", text_color=GRAY,
                     font=ctk.CTkFont(size=12)).pack(side="left")

    def _nav(self, label):
        if "批次發樣" not in label:
            mbox.showinfo("開發中", f"「{label.split(' ',1)[-1]}」之後會搬進 GUI。\n目前先用 CLI（./start.command）。")


class SampleOrderFrame(ctk.CTkFrame):
    """批次發樣畫面：單別 + 組合 + 客戶搜尋勾選 + dry-run + 開立。"""

    def __init__(self, master):
        super().__init__(master, fg_color="#FFFFFF", corner_radius=0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.all_customers = []     # [{code,name}]
        self.cust_vars = {}         # code -> BooleanVar
        self.chosen = set()         # 已勾的 code（跨搜尋累積）
        self.combos = {}            # 名稱 -> 品項（資料載入前先空，避免提早點選 crash）

        self._build_toolbar()
        self._build_body()
        self._build_actionbar()
        # 所有 widget 都建好後才綁 command，避免 set()/select() 提早觸發 _update_summary
        self.seg.configure(command=lambda v: self._update_summary())
        self.dry_switch.configure(command=lambda: self._update_summary())
        self._update_summary()
        self._load_data_async()

    # ── 頂部工具列 ──
    def _build_toolbar(self):
        bar = ctk.CTkFrame(self, height=52, fg_color="transparent")
        bar.grid(row=0, column=0, sticky="ew", padx=22, pady=(14, 0))
        ctk.CTkLabel(bar, text="批次發樣", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")
        self.dry_switch = ctk.CTkSwitch(bar, text="預覽模式（不寫入）", progress_color=GREEN,
                                        font=ctk.CTkFont(size=12))
        self.dry_switch.select()   # 預設開＝安全（先設值，下方再綁 command 避免提早觸發）
        self.dry_switch.pack(side="right")

    # ── 主體：單別 + 雙欄 ──
    def _build_body(self):
        self.seg = ctk.CTkSegmentedButton(self, values=SC.ORDER_TYPES,
                                          font=ctk.CTkFont(size=13))
        self.seg.set(SC.ORDER_TYPES[0])
        self.seg.grid(row=1, column=0, sticky="w", padx=24, pady=(16, 8))

        cols = ctk.CTkFrame(self, fg_color="transparent")
        cols.grid(row=2, column=0, sticky="nsew", padx=24, pady=8)
        cols.grid_columnconfigure((0, 1), weight=1, uniform="c")
        cols.grid_rowconfigure(0, weight=1)

        # 左：組合
        left = ctk.CTkFrame(cols, border_width=1, border_color="#E5E5EA", corner_radius=12)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(left, text="樣品組合", font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=GRAY).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        self.combo_menu = ctk.CTkOptionMenu(left, values=["（讀取中…）"],
                                            command=self._on_combo_change, fg_color=BLUE)
        self.combo_menu.grid(row=1, column=0, sticky="ew", padx=16, pady=4)
        self.combo_items_box = ctk.CTkTextbox(left, height=180, font=ctk.CTkFont(size=13))
        self.combo_items_box.grid(row=2, column=0, sticky="nsew", padx=16, pady=(8, 14))
        left.grid_rowconfigure(2, weight=1)

        # 右：客戶搜尋 + 勾選
        right = ctk.CTkFrame(cols, border_width=1, border_color="#E5E5EA", corner_radius=12)
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=1)
        self.cust_hd = ctk.CTkLabel(right, text="發給客戶 · 已選 0",
                                    font=ctk.CTkFont(size=13, weight="bold"), text_color=GRAY)
        self.cust_hd.grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        self.search = ctk.CTkEntry(right, placeholder_text="🔍 搜尋客戶（中文 / 代號）")
        self.search.grid(row=1, column=0, sticky="ew", padx=16, pady=4)
        self.search.bind("<KeyRelease>", lambda e: self._refresh_cust_list())
        self.cust_scroll = ctk.CTkScrollableFrame(right, fg_color="transparent")
        self.cust_scroll.grid(row=2, column=0, sticky="nsew", padx=8, pady=(6, 12))
        self.cust_scroll.grid_columnconfigure(0, weight=1)

    # ── 底部動作列 ──
    def _build_actionbar(self):
        bar = ctk.CTkFrame(self, height=64, fg_color="#FAFAFB")
        bar.grid(row=3, column=0, sticky="ew")
        bar.grid_columnconfigure(0, weight=1)
        self.summary = ctk.CTkLabel(bar, text="", text_color=GRAY, font=ctk.CTkFont(size=13))
        self.summary.grid(row=0, column=0, sticky="w", padx=26, pady=14)
        self.go_btn = ctk.CTkButton(bar, text="預覽 0 張單", width=160, height=38,
                                    corner_radius=9, fg_color=BLUE,
                                    font=ctk.CTkFont(size=14, weight="bold"),
                                    command=self._on_go)
        self.go_btn.grid(row=0, column=1, sticky="e", padx=26, pady=12)
        self._update_summary()

    # ── 資料載入（背景執行緒取資料，主執行緒輪詢 queue 更新 UI）──
    # Tkinter 不可跨執行緒呼叫；背景只放結果進 queue，由主執行緒 _poll_data 取出更新。
    def _load_data_async(self):
        self._dataq = queue.Queue()

        def work():
            try:
                custs = SC.load_customers()
                combos = SC.load_combos()
                self._dataq.put(("ok", custs, combos))
            except Exception as e:
                self._dataq.put(("err", e, None))
        threading.Thread(target=work, daemon=True).start()
        self._poll_data()   # 主執行緒啟動輪詢

    def _poll_data(self):
        try:
            kind, a, b = self._dataq.get_nowait()
        except queue.Empty:
            self.after(120, self._poll_data)   # 在主執行緒呼叫 .after()，安全
            return
        if kind == "ok":
            self._on_data_loaded(a, b)
        else:
            mbox.showerror("載入失敗", str(a))

    def _on_data_loaded(self, custs, combos):
        self.all_customers = [c for c in custs if c.get("code")]
        names = list(combos.keys()) or ["（尚無組合，請先用 CLI 建立）"]
        self.combos = combos
        self.combo_menu.configure(values=names)
        self.combo_menu.set(names[0])
        self._on_combo_change(names[0])
        self._refresh_cust_list()

    # ── 組合切換 ──
    def _on_combo_change(self, name):
        items = self.combos.get(name, [])
        self.combo_items_box.delete("1.0", "end")
        if items:
            for it in items:
                self.combo_items_box.insert("end", f"{it['code']:<12} {it['name'][:18]}  ×{it['qty']}\n")
        else:
            self.combo_items_box.insert("end", "（此組合無品項）")
        self._update_summary()

    # ── 客戶清單（依搜尋過濾，中文可）──
    def _refresh_cust_list(self):
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
            cb = ctk.CTkCheckBox(self.cust_scroll, text=f"{c['name']}｜{c['code']}",
                                 variable=var, font=ctk.CTkFont(size=13),
                                 command=lambda code=c["code"]: self._toggle(code))
            cb.pack(anchor="w", padx=6, pady=2)
        self.cust_hd.configure(text=f"發給客戶 · 已選 {len(self.chosen)}")

    def _toggle(self, code):
        var = self.cust_vars[code]
        if var.get():
            self.chosen.add(code)
        else:
            self.chosen.discard(code)
        self.cust_hd.configure(text=f"發給客戶 · 已選 {len(self.chosen)}")
        self._update_summary()

    def _selected_customers(self):
        by_code = {c["code"]: c for c in self.all_customers}
        return [by_code[cc] for cc in self.chosen if cc in by_code]

    def _current_combo_items(self):
        return self.combos.get(self.combo_menu.get(), [])

    def _update_summary(self):
        n = len(self.chosen)
        ot = self.seg.get()
        self.summary.configure(text=f"將開立 {n} 張「{ot}」單 · 單價全 0 · 狀態未出貨")
        mode = "預覽" if self.dry_switch.get() else "開立"
        self.go_btn.configure(text=f"{mode} {n} 張單")

    # ── 開立 / 預覽 ──
    def _on_go(self):
        combo = self._current_combo_items()
        custs = self._selected_customers()
        if not combo:
            mbox.showwarning("提醒", "請先選一個有品項的組合"); return
        if not custs:
            mbox.showwarning("提醒", "請至少勾選一個客戶"); return
        ot = self.seg.get()
        dry = bool(self.dry_switch.get())

        if dry:
            res = SC.create_sample_orders(ot, combo, custs, commit=False)
            lines = "\n".join(f"  {r['customer']['name']}（{r['customer']['code']}）" for r in res)
            mbox.showinfo("預覽（未寫入）",
                          f"將開立 {len(res)} 張「{ot}」單，單價全 0、狀態未出貨：\n\n{lines}\n\n"
                          f"關閉「預覽模式」後再按開立才會實際寫入。")
            return

        if not mbox.askyesno("確認開立",
                             f"確定開立 {len(custs)} 張「{ot}」單到 Ragic？\n（單價全 0、狀態未出貨）"):
            return
        res = SC.create_sample_orders(ot, combo, custs, commit=True)
        ok = sum(1 for r in res if r["ok"])
        fail = [f"{r['customer']['name']}：{r['msg']}" for r in res if not r["ok"]]
        msg = f"完成！{ok}/{len(res)} 張已建立。"
        if fail:
            msg += "\n\n失敗：\n" + "\n".join(fail)
        mbox.showinfo("結果", msg)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
