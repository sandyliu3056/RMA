# -*- coding: utf-8 -*-
"""
售後零件規劃平台（桌面版）
------------------------
六個模組對應工作架構圖：需求規劃、庫存與缺料、訂單與交期、出貨與到貨、報表與分析、跨單位協調。
資料來源：AIO 匯出檔（AllInOneData）。分析邏輯在 rma_engine.py。

執行：python 零件規劃平台.py
需要：pandas、openpyxl、matplotlib、python-pptx、pillow
"""
import os
import sys
import json
import threading
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import rma_engine as E

APP_TITLE = "售後零件規劃平台"
NAVY, ORANGE, BG, CARD, INK, MUTE = "#1F3F6E", "#E8862B", "#F3F6FA", "#FFFFFF", "#1E2430", "#5B6472"
ACTIONS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "協調待辦.json")
UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"]


# ----------------------------------------------------------------------------
# 共用小工具
# ----------------------------------------------------------------------------
def fmt_cell(v):
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        if abs(v) < 1 and v != 0:
            return f"{v:.1%}" if abs(v) <= 1 else f"{v:,.2f}"
        return f"{v:,.1f}"
    if isinstance(v, (int,)) or (hasattr(v, "dtype") and "int" in str(getattr(v, "dtype", ""))):
        return f"{int(v):,}"
    return str(v)


class Grid(ttk.Frame):
    """把 DataFrame 顯示成表格（Treeview），附捲軸與匯出。"""

    def __init__(self, master, height=12):
        super().__init__(master)
        self.tree = ttk.Treeview(self, show="headings", height=height)
        ys = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.df = None

    def show(self, df, index=True, max_rows=2000):
        self.tree.delete(*self.tree.get_children())
        if df is None or len(df) == 0:
            self.tree["columns"] = ["訊息"]
            self.tree.heading("訊息", text="訊息")
            self.tree.insert("", "end", values=["（沒有資料）"])
            self.df = None
            return
        d = df.reset_index() if index else df
        d = d.rename(columns={"index": "項目"})
        self.df = d
        cols = [str(c) for c in d.columns]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            width = max(80, min(320, int(max([len(c)] + [len(fmt_cell(v)) for v in d[c].head(50)]) * 9 + 20)))
            self.tree.column(c, width=width, anchor="w", stretch=False)
        for _, row in d.head(max_rows).iterrows():
            self.tree.insert("", "end", values=[fmt_cell(v) for v in row])

    def export(self, default_name="表格.xlsx"):
        if self.df is None:
            messagebox.showinfo(APP_TITLE, "目前沒有可匯出的表格")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=default_name,
                                            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        if path.lower().endswith(".csv"):
            self.df.to_csv(path, index=False, encoding="utf-8-sig")
        else:
            self.df.to_excel(path, index=False)
        messagebox.showinfo(APP_TITLE, f"已匯出：{path}")


def kpi_row(master, items):
    """一列 KPI 方塊。items = [(標題, 數值)…]，回傳可更新的 label 字典。"""
    frame = ttk.Frame(master)
    labels = {}
    for i, (title, value) in enumerate(items):
        card = tk.Frame(frame, bg=CARD, bd=0, highlightthickness=1, highlightbackground="#D5DDE8")
        card.grid(row=0, column=i, padx=6, pady=4, sticky="nsew", ipadx=10, ipady=6)
        frame.columnconfigure(i, weight=1)
        tk.Label(card, text=title, bg=CARD, fg=MUTE, font=("Microsoft JhengHei", 9)).pack(anchor="w", padx=8)
        lab = tk.Label(card, text=value, bg=CARD, fg=NAVY, font=("Microsoft JhengHei", 16, "bold"))
        lab.pack(anchor="w", padx=8)
        labels[title] = lab
    return frame, labels


# ----------------------------------------------------------------------------
# 主程式
# ----------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1380x860")
        self.configure(bg=BG)
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TNotebook.Tab", font=("Microsoft JhengHei", 11, "bold"), padding=(14, 8))
        style.configure("Treeview", font=("Microsoft JhengHei", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("Microsoft JhengHei", 9, "bold"))
        style.configure("TLabel", font=("Microsoft JhengHei", 10))
        style.configure("TButton", font=("Microsoft JhengHei", 10))
        style.configure("Accent.TButton", font=("Microsoft JhengHei", 10, "bold"), foreground="white", background=ORANGE)
        style.map("Accent.TButton", background=[("active", "#D0741F")])

        self.raw = None
        self.data = None
        self.report = None
        self.parts = None
        self.chart_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_charts")
        self._photo = None

        header = tk.Frame(self, bg=NAVY, height=56)
        header.pack(fill="x")
        tk.Label(header, text="售後零件規劃平台", bg=NAVY, fg="white", font=("Microsoft JhengHei", 16, "bold")).pack(side="left", padx=16, pady=10)
        self.status_var = tk.StringVar(value="尚未載入資料")
        tk.Label(header, textvariable=self.status_var, bg=NAVY, fg="#DDE6F3", font=("Microsoft JhengHei", 10)).pack(side="right", padx=16)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_source = ttk.Frame(self.nb); self.nb.add(self.tab_source, text="資料來源")
        self.tab_demand = ttk.Frame(self.nb); self.nb.add(self.tab_demand, text="需求規劃")
        self.tab_stock = ttk.Frame(self.nb); self.nb.add(self.tab_stock, text="庫存與缺料")
        self.tab_po = ttk.Frame(self.nb); self.nb.add(self.tab_po, text="訂單與交期")
        self.tab_ship = ttk.Frame(self.nb); self.nb.add(self.tab_ship, text="出貨與到貨")
        self.tab_report = ttk.Frame(self.nb); self.nb.add(self.tab_report, text="報表與分析")
        self.tab_coord = ttk.Frame(self.nb); self.nb.add(self.tab_coord, text="跨單位協調")
        self._build_source()
        self._build_demand()
        self._build_stock()
        self._build_po()
        self._build_ship()
        self._build_report()
        self._build_coord()

    # ------------------------------------------------------------------ 資料來源
    def _build_source(self):
        f = self.tab_source
        top = ttk.Frame(f); top.pack(fill="x", padx=12, pady=12)
        ttk.Label(top, text="AIO 匯出檔：").grid(row=0, column=0, sticky="w")
        self.path_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.path_var, width=80).grid(row=0, column=1, padx=6)
        ttk.Button(top, text="瀏覽…", command=self._browse).grid(row=0, column=2)
        ttk.Label(top, text="分析起日（留空＝資料最後日期的前一個月 1 日）：").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.start_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.start_var, width=16).grid(row=1, column=1, sticky="w", padx=6, pady=(8, 0))
        ttk.Button(top, text="載入並分析", style="Accent.TButton", command=self._load_clicked).grid(row=1, column=2, pady=(8, 0))
        self.kpi_frame, self.kpi_labels = kpi_row(f, [("案件數", "－"), ("在途", "－"), ("缺料（Awaiting Spares）", "－"),
                                                     ("零件未使用率", "－"), ("SLA 達成率", "－"), ("保固 30 天內到期在途", "－")])
        self.kpi_frame.pack(fill="x", padx=8)
        self.log = tk.Text(f, height=18, font=("Consolas", 10), bg="#FBFCFE")
        self.log.pack(fill="both", expand=True, padx=12, pady=8)
        self._log("1. 選擇 AIO 匯出檔（.xlsx，工作表 AllInOneData）\n2. 按「載入並分析」，約 10–30 秒\n3. 到各分頁查看，或在「報表與分析」匯出 Excel／PPT")

    def _log(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if p:
            self.path_var.set(p)

    def _load_clicked(self):
        p = self.path_var.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showwarning(APP_TITLE, "請先選擇 AIO 匯出檔")
            return
        self.status_var.set("載入中…")
        threading.Thread(target=self._load_worker, args=(p, self.start_var.get().strip() or None), daemon=True).start()

    def _load_worker(self, path, start):
        try:
            self._log(f"讀取 {os.path.basename(path)} …")
            raw = E.load_aio(path)
            self._log(f"  共 {len(raw):,} 列，欄位 {raw.shape[1]} 個")
            data = E.prepare(raw, start)
            self._log(f"  分析期間 {data.attrs['period_start']:%Y-%m-%d} 起，資料日期 {data.attrs['ref_date']:%Y-%m-%d}，{len(data):,} 件")
            rep = E.run_all(data, progress=lambda i, n, t: self._log(f"  [{i}/{n}] {t}"))
            self._log(f"  零件明細 {len(rep.parts):,} 顆")
            self.raw, self.data, self.report, self.parts = raw, data, rep, rep.parts
            self.after(0, self._after_load)
        except Exception as ex:  # noqa
            import traceback
            self._log("錯誤：" + traceback.format_exc())
            self.after(0, lambda: messagebox.showerror(APP_TITLE, f"載入失敗：{ex}"))
            self.after(0, lambda: self.status_var.set("載入失敗"))

    def _after_load(self):
        rep, d = self.report, self.data
        k12, k22, k41, k42, k52 = rep.get("1-2").kpis, rep.get("2-2").kpis, rep.get("4-1").kpis, rep.get("4-2").kpis, rep.get("5-2").kpis
        self.kpi_labels["案件數"].config(text=f"{len(d):,}")
        self.kpi_labels["在途"].config(text=f"{k12['在途件數']:,}（{k12['在途占比']:.0%}）")
        self.kpi_labels["缺料（Awaiting Spares）"].config(text=f"{k42['缺料件數']:,}（{k42['缺料占在途']:.0%}）")
        self.kpi_labels["零件未使用率"].config(text=f"{k41['未使用率']:.1%}")
        self.kpi_labels["SLA 達成率"].config(text=f"{k22['整體達成率']:.0%}")
        self.kpi_labels["保固 30 天內到期在途"].config(text=f"{k52['30天內到期件數']:,}")
        self.status_var.set(f"{os.path.basename(self.path_var.get())}｜{d.attrs['period_start']:%m/%d}–{d.attrs['ref_date']:%m/%d}｜{len(d):,} 件")
        self._log("完成。")
        self._refresh_demand()
        self._refresh_stock()
        self._refresh_ship()
        self._refresh_report_list()

    def _need_data(self):
        if self.report is None:
            messagebox.showinfo(APP_TITLE, "請先在「資料來源」載入 AIO 匯出檔")
            return False
        return True

    # ------------------------------------------------------------------ 需求規劃
    def _build_demand(self):
        f = self.tab_demand
        bar = ttk.Frame(f); bar.pack(fill="x", padx=12, pady=8)
        ttk.Label(bar, text="彙總依：").pack(side="left")
        self.demand_by = tk.StringVar(value="國家")
        ttk.Combobox(bar, textvariable=self.demand_by, values=["國家", "類型分組"], width=10, state="readonly").pack(side="left", padx=4)
        ttk.Label(bar, text="預測未來（週）：").pack(side="left", padx=(12, 0))
        self.demand_weeks = tk.IntVar(value=4)
        ttk.Spinbox(bar, from_=1, to=13, textvariable=self.demand_weeks, width=5).pack(side="left", padx=4)
        ttk.Button(bar, text="重新計算", command=self._refresh_demand).pack(side="left", padx=8)
        ttk.Button(bar, text="匯出上表", command=lambda: self.demand_grid.export("週申請零件量.xlsx")).pack(side="right")
        ttk.Button(bar, text="匯出料號表", command=lambda: self.parts_grid.export("Top料號.xlsx")).pack(side="right", padx=6)
        ttk.Label(f, text="每週申請零件顆數（完整週）與未來需求推估", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12)
        self.demand_grid = Grid(f, height=8); self.demand_grid.pack(fill="both", expand=False, padx=12, pady=4)
        self.forecast_grid = Grid(f, height=7); self.forecast_grid.pack(fill="both", expand=False, padx=12, pady=4)
        ttk.Label(f, text="申請次數最多的料號（Top 30）：申請顆數、未使用率、目前待供貨顆數", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        self.parts_grid = Grid(f, height=10); self.parts_grid.pack(fill="both", expand=True, padx=12, pady=4)

    def _refresh_demand(self):
        if self.report is None:
            return
        by = self.demand_by.get()
        self.demand_grid.show(E.demand_weekly(self.data, by))
        self.forecast_grid.show(E.demand_forecast(self.data, self.demand_weeks.get(), by))
        self.parts_grid.show(E.top_parts(self.parts, 30))

    # ------------------------------------------------------------------ 庫存與缺料
    def _build_stock(self):
        f = self.tab_stock
        self.stock_kpi_frame, self.stock_kpi = kpi_row(f, [("缺料案件", "－"), ("等待超過兩週", "－"), ("等待中位數（天）", "－"),
                                                          ("零件在途／已配案件", "－"), ("待供貨料號數", "－")])
        self.stock_kpi_frame.pack(fill="x", padx=8, pady=(8, 0))
        bar = ttk.Frame(f); bar.pack(fill="x", padx=12, pady=6)
        ttk.Label(bar, text="催料清單：等待天數 ≥").pack(side="left")
        self.stock_days = tk.IntVar(value=14)
        ttk.Spinbox(bar, from_=1, to=60, textvariable=self.stock_days, width=5, command=self._refresh_stock).pack(side="left", padx=4)
        ttk.Button(bar, text="重新整理", command=self._refresh_stock).pack(side="left", padx=8)
        ttk.Button(bar, text="匯出催料清單", command=lambda: self.short_case_grid.export("催料清單.xlsx")).pack(side="right")
        ttk.Button(bar, text="匯出料號清單", command=lambda: self.short_part_grid.export("缺料料號.xlsx")).pack(side="right", padx=6)
        ttk.Label(f, text="缺料料號（Awaiting Spares 案件中待供貨的料號）", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12)
        self.short_part_grid = Grid(f, height=9); self.short_part_grid.pack(fill="both", expand=False, padx=12, pady=4)
        ttk.Label(f, text="催料案件清單", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        self.short_case_grid = Grid(f, height=10); self.short_case_grid.pack(fill="both", expand=True, padx=12, pady=4)

    def _refresh_stock(self):
        if self.report is None:
            return
        k = self.report.get("4-2").kpis
        sp = E.shortage_parts(self.parts)
        self.stock_kpi["缺料案件"].config(text=f"{k['缺料件數']:,}（占在途 {k['缺料占在途']:.0%}）")
        self.stock_kpi["等待超過兩週"].config(text=f"{k['超過兩週件數']:,}（{k['超過兩週比例']:.0%}）")
        self.stock_kpi["等待中位數（天）"].config(text=f"{k['等待中位數']:.0f}")
        self.stock_kpi["零件在途／已配案件"].config(text=f"{k['零件在途或已配件數']:,}")
        self.stock_kpi["待供貨料號數"].config(text=f"{len(sp):,}")
        self.short_part_grid.show(sp)
        self.short_case_grid.show(E.shortage_cases(self.data, self.stock_days.get()), index=False)

    # ------------------------------------------------------------------ 訂單與交期
    def _build_po(self):
        f = self.tab_po
        ttk.Label(f, text="AIO 資料沒有採購訂單，這個模組用匯入的訂單檔計算準交率與逾期清單。", foreground=MUTE).pack(anchor="w", padx=12, pady=(10, 0))
        bar = ttk.Frame(f); bar.pack(fill="x", padx=12, pady=8)
        ttk.Button(bar, text="產生訂單範本", command=self._po_template).pack(side="left")
        ttk.Button(bar, text="匯入訂單檔…", style="Accent.TButton", command=self._po_import).pack(side="left", padx=8)
        ttk.Button(bar, text="匯出逾期清單", command=lambda: self.po_late_grid.export("逾期訂單.xlsx")).pack(side="right")
        self.po_kpi_frame, self.po_kpi = kpi_row(f, [("訂單筆數", "－"), ("已到貨", "－"), ("準交率", "－"), ("平均交期（天）", "－"), ("未到貨且已逾期", "－")])
        self.po_kpi_frame.pack(fill="x", padx=8)
        ttk.Label(f, text="供應商準交率", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12)
        self.po_sup_grid = Grid(f, height=7); self.po_sup_grid.pack(fill="both", expand=False, padx=12, pady=4)
        ttk.Label(f, text="逾期與未交訂單", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        self.po_late_grid = Grid(f, height=10); self.po_late_grid.pack(fill="both", expand=True, padx=12, pady=4)

    def _po_template(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="零件訂單範本.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        tmpl = pd.DataFrame([
            ["PO-2026-0001", "KP.04501.017", "供應商A", 50, "2026-09-10", "2026-09-12", "2026-09-11", "Thailand"],
            ["PO-2026-0002", "KT.CTE00.014", "供應商B", 20, "2026-09-12", "2026-09-20", "", "Indonesia"],
        ], columns=["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日", "目的國"])
        tmpl.to_excel(path, index=False)
        messagebox.showinfo(APP_TITLE, f"範本已存到：{path}\n欄位：訂單號、料號、供應商、數量、下單日、承諾交期、實際到貨日（未到貨留空）、目的國")

    def _po_import(self):
        path = filedialog.askopenfilename(filetypes=[("Excel／CSV", "*.xlsx *.xls *.csv")])
        if not path:
            return
        try:
            po = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
            need = ["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日"]
            miss = [c for c in need if c not in po.columns]
            if miss:
                raise ValueError("缺少欄位：" + "、".join(miss))
            for c in ["下單日", "承諾交期", "實際到貨日"]:
                po[c] = pd.to_datetime(po[c], errors="coerce")
            today = pd.Timestamp(dt.date.today())
            po["已到貨"] = po["實際到貨日"].notna()
            po["準交"] = po["已到貨"] & (po["實際到貨日"] <= po["承諾交期"])
            po["交期天數"] = (po["實際到貨日"] - po["下單日"]).dt.days
            po["逾期天數"] = ((po["實際到貨日"].fillna(today) - po["承諾交期"]).dt.days).clip(lower=0)
            arrived = po[po["已到貨"]]
            sup = po.groupby("供應商").agg(訂單筆數=("訂單號", "size"), 已到貨=("已到貨", "sum"), 準交筆數=("準交", "sum"),
                                         平均交期天數=("交期天數", "mean"), 平均逾期天數=("逾期天數", "mean"))
            sup["準交率"] = sup["準交筆數"] / sup["已到貨"].replace(0, pd.NA)
            late = po[(~po["已到貨"] & (po["承諾交期"] < today)) | (po["已到貨"] & ~po["準交"])].sort_values("逾期天數", ascending=False)
            self.po_kpi["訂單筆數"].config(text=f"{len(po):,}")
            self.po_kpi["已到貨"].config(text=f"{int(po['已到貨'].sum()):,}")
            self.po_kpi["準交率"].config(text=f"{(arrived['準交'].mean() if len(arrived) else 0):.0%}")
            self.po_kpi["平均交期（天）"].config(text=f"{(arrived['交期天數'].mean() if len(arrived) else 0):.1f}")
            self.po_kpi["未到貨且已逾期"].config(text=f"{int((~po['已到貨'] & (po['承諾交期'] < today)).sum()):,}")
            self.po_sup_grid.show(sup)
            self.po_late_grid.show(late[["訂單號", "料號", "供應商", "數量", "承諾交期", "實際到貨日", "逾期天數"] + (["目的國"] if "目的國" in po.columns else [])], index=False)
        except Exception as ex:  # noqa
            messagebox.showerror(APP_TITLE, f"匯入失敗：{ex}")

    # ------------------------------------------------------------------ 出貨與到貨
    def _build_ship(self):
        f = self.tab_ship
        self.ship_kpi_frame, self.ship_kpi = kpi_row(f, [("待出貨案件", "－"), ("待出貨超過 7 天", "－"), ("零件運送中（顆）", "－"),
                                                        ("零件已配貨（顆）", "－"), ("保固 30 天內到期在途", "－")])
        self.ship_kpi_frame.pack(fill="x", padx=8, pady=(8, 0))
        bar = ttk.Frame(f); bar.pack(fill="x", padx=12, pady=6)
        ttk.Label(bar, text="保固到期清單：天數內").pack(side="left")
        self.exp_days = tk.IntVar(value=30)
        ttk.Spinbox(bar, from_=7, to=180, increment=7, textvariable=self.exp_days, width=5, command=self._refresh_ship).pack(side="left", padx=4)
        ttk.Button(bar, text="重新整理", command=self._refresh_ship).pack(side="left", padx=8)
        ttk.Button(bar, text="匯出到期清單", command=lambda: self.exp_grid.export("保固到期在途案.xlsx")).pack(side="right")
        ttk.Button(bar, text="匯出待出貨積壓", command=lambda: self.backlog_grid.export("待出貨積壓.xlsx")).pack(side="right", padx=6)
        row = ttk.Frame(f); row.pack(fill="both", expand=False, padx=12)
        left = ttk.Frame(row); left.pack(side="left", fill="both", expand=True)
        right = ttk.Frame(row); right.pack(side="left", fill="both", expand=True, padx=(8, 0))
        ttk.Label(left, text="待出貨積壓（依據點）", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")
        self.backlog_grid = Grid(left, height=9); self.backlog_grid.pack(fill="both", expand=True, pady=4)
        ttk.Label(right, text="零件運送中／已配貨（依國家，在途案件）", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")
        self.transit_grid = Grid(right, height=9); self.transit_grid.pack(fill="both", expand=True, pady=4)
        ttk.Label(f, text="在途且保固即將到期（含建案時保固內、現已過期）", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w", padx=12, pady=(8, 0))
        self.exp_grid = Grid(f, height=10); self.exp_grid.pack(fill="both", expand=True, padx=12, pady=4)

    def _refresh_ship(self):
        if self.report is None:
            return
        bl = E.ship_backlog(self.data)
        tr = E.parts_in_transit(self.parts)
        ex = E.expiring_cases(self.data, self.exp_days.get())
        self.ship_kpi["待出貨案件"].config(text=f"{int(bl['待出貨件數'].sum()):,}")
        self.ship_kpi["待出貨超過 7 天"].config(text=f"{int(bl['超過7天'].sum()):,}")
        self.ship_kpi["零件運送中（顆）"].config(text=f"{int(tr['運送中'].sum()) if '運送中' in tr.columns else 0:,}")
        self.ship_kpi["零件已配貨（顆）"].config(text=f"{int(tr['已配貨'].sum()) if '已配貨' in tr.columns else 0:,}")
        self.ship_kpi["保固 30 天內到期在途"].config(text=f"{self.report.get('5-2').kpis['30天內到期件數']:,}")
        self.backlog_grid.show(bl)
        self.transit_grid.show(tr)
        self.exp_grid.show(ex, index=False)

    # ------------------------------------------------------------------ 報表與分析
    def _build_report(self):
        f = self.tab_report
        left = ttk.Frame(f, width=300); left.pack(side="left", fill="y", padx=(12, 6), pady=8)
        ttk.Label(left, text="分析項目", font=("Microsoft JhengHei", 11, "bold")).pack(anchor="w")
        self.an_list = tk.Listbox(left, width=34, height=14, font=("Microsoft JhengHei", 10), exportselection=False)
        self.an_list.pack(fill="y", pady=4)
        self.an_list.bind("<<ListboxSelect>>", lambda e: self._show_analysis())
        ttk.Button(left, text="匯出 Excel 表格", command=self._export_excel).pack(fill="x", pady=2)
        ttk.Button(left, text="匯出 PPT 報告", command=self._export_pptx).pack(fill="x", pady=2)
        ttk.Button(left, text="匯出全部圖 PNG", command=self._export_charts).pack(fill="x", pady=2)
        ttk.Button(left, text="全部匯出到資料夾", style="Accent.TButton", command=self._export_all).pack(fill="x", pady=(8, 2))
        right = ttk.Frame(f); right.pack(side="left", fill="both", expand=True, padx=(6, 12), pady=8)
        self.an_title = ttk.Label(right, text="請先載入資料，再從左側選一項分析", font=("Microsoft JhengHei", 13, "bold"), foreground=NAVY)
        self.an_title.pack(anchor="w")
        body = ttk.Panedwindow(right, orient="horizontal"); body.pack(fill="both", expand=True)
        textf = ttk.Frame(body); body.add(textf, weight=1)
        self.an_text = tk.Text(textf, wrap="word", font=("Microsoft JhengHei", 10), width=52, bg="#FBFCFE")
        self.an_text.pack(fill="both", expand=True)
        chartf = ttk.Frame(body); body.add(chartf, weight=2)
        self.chart_label = tk.Label(chartf, bg="white", text="（圖表）")
        self.chart_label.pack(fill="both", expand=True)
        tabf = ttk.Frame(right); tabf.pack(fill="both", expand=True, pady=(6, 0))
        bar = ttk.Frame(tabf); bar.pack(fill="x")
        ttk.Label(bar, text="表格：").pack(side="left")
        self.table_var = tk.StringVar()
        self.table_cb = ttk.Combobox(bar, textvariable=self.table_var, width=40, state="readonly")
        self.table_cb.pack(side="left", padx=4)
        self.table_cb.bind("<<ComboboxSelected>>", lambda e: self._show_table())
        ttk.Button(bar, text="匯出此表", command=lambda: self.an_grid.export("分析表格.xlsx")).pack(side="right")
        self.an_grid = Grid(tabf, height=8); self.an_grid.pack(fill="both", expand=True, pady=4)

    def _refresh_report_list(self):
        self.an_list.delete(0, "end")
        for a in self.report.analyses.values():
            self.an_list.insert("end", f"{a.key}  {a.title.split('｜')[-1]}")
        self.an_list.selection_set(0)
        self._show_analysis()

    def _current_analysis(self):
        sel = self.an_list.curselection()
        if not sel or self.report is None:
            return None
        return list(self.report.analyses.values())[sel[0]]

    def _show_analysis(self):
        a = self._current_analysis()
        if a is None:
            return
        self.an_title.config(text=f"{a.key} {a.title}")
        self.an_text.delete("1.0", "end")
        self.an_text.insert("end", a.subtitle + "\n\n")
        for b in a.bullets:
            self.an_text.insert("end", "・" + b + "\n")
        self.an_text.insert("end", "\n" + a.paragraph + "\n\n結論：" + a.conclusion + "\n")
        self.table_cb["values"] = list(a.tables.keys())
        self.table_var.set(list(a.tables.keys())[0])
        self._show_table()
        self._show_chart(a)

    def _show_table(self):
        a = self._current_analysis()
        if a is None:
            return
        self.an_grid.show(a.tables.get(self.table_var.get()))

    def _show_chart(self, a):
        try:
            os.makedirs(self.chart_dir, exist_ok=True)
            p = os.path.join(self.chart_dir, f"{a.key}.png")
            if not (a.chart and os.path.exists(a.chart)):
                E.plot_analysis(a, p)
            from PIL import Image, ImageTk
            im = Image.open(a.chart)
            w = max(self.chart_label.winfo_width(), 700)
            h = max(self.chart_label.winfo_height(), 260)
            r = min(w / im.width, h / im.height, 1.0)
            im = im.resize((int(im.width * r), int(im.height * r)))
            self._photo = ImageTk.PhotoImage(im)
            self.chart_label.config(image=self._photo, text="")
        except Exception as ex:  # noqa
            self.chart_label.config(text=f"圖表無法顯示：{ex}", image="")

    def _ask_dir(self):
        return filedialog.askdirectory()

    def _export_excel(self):
        if not self._need_data():
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="RMA分析結果.xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            E.export_excel(self.report, p); messagebox.showinfo(APP_TITLE, f"已匯出：{p}")

    def _export_pptx(self):
        if not self._need_data():
            return
        p = filedialog.asksaveasfilename(defaultextension=".pptx", initialfile="RMA分析報告.pptx", filetypes=[("PowerPoint", "*.pptx")])
        if p:
            E.export_pptx(self.report, p, os.path.join(os.path.dirname(p), "charts")); messagebox.showinfo(APP_TITLE, f"已匯出：{p}")

    def _export_charts(self):
        if not self._need_data():
            return
        d = self._ask_dir()
        if d:
            out = E.export_charts(self.report, d); messagebox.showinfo(APP_TITLE, f"已匯出 {len(out)} 張圖到 {d}")

    def _export_all(self):
        if not self._need_data():
            return
        d = self._ask_dir()
        if d:
            out = E.export_all(self.report, d)
            messagebox.showinfo(APP_TITLE, f"已匯出：\n{out['excel']}\n{out['pptx']}\n{len(out['charts'])} 張圖")

    # ------------------------------------------------------------------ 跨單位協調
    def _build_coord(self):
        f = self.tab_coord
        form = ttk.LabelFrame(f, text="新增／更新待辦"); form.pack(fill="x", padx=12, pady=8)
        self.c_unit = tk.StringVar(value=UNITS[0]); self.c_topic = tk.StringVar(); self.c_owner = tk.StringVar()
        self.c_due = tk.StringVar(value=(dt.date.today() + dt.timedelta(days=7)).isoformat()); self.c_status = tk.StringVar(value="進行中")
        ttk.Label(form, text="對象單位").grid(row=0, column=0, padx=6, pady=4, sticky="e")
        ttk.Combobox(form, textvariable=self.c_unit, values=UNITS, width=14, state="readonly").grid(row=0, column=1, sticky="w")
        ttk.Label(form, text="議題").grid(row=0, column=2, padx=6, sticky="e")
        ttk.Entry(form, textvariable=self.c_topic, width=60).grid(row=0, column=3, sticky="w")
        ttk.Label(form, text="負責人").grid(row=1, column=0, padx=6, pady=4, sticky="e")
        ttk.Entry(form, textvariable=self.c_owner, width=16).grid(row=1, column=1, sticky="w")
        ttk.Label(form, text="到期日").grid(row=1, column=2, padx=6, sticky="e")
        ttk.Entry(form, textvariable=self.c_due, width=14).grid(row=1, column=3, sticky="w")
        ttk.Label(form, text="狀態").grid(row=1, column=4, padx=6, sticky="e")
        ttk.Combobox(form, textvariable=self.c_status, values=["進行中", "待回覆", "已完成", "取消"], width=8, state="readonly").grid(row=1, column=5, sticky="w")
        btns = ttk.Frame(form); btns.grid(row=0, column=6, rowspan=2, padx=10)
        ttk.Button(btns, text="新增", style="Accent.TButton", command=self._coord_add).pack(fill="x", pady=2)
        ttk.Button(btns, text="更新所選", command=self._coord_update).pack(fill="x", pady=2)
        ttk.Button(btns, text="刪除所選", command=self._coord_delete).pack(fill="x", pady=2)
        bar = ttk.Frame(f); bar.pack(fill="x", padx=12)
        ttk.Button(bar, text="由分析結論產生待辦", command=self._coord_seed).pack(side="left")
        ttk.Button(bar, text="匯出待辦", command=self._coord_export).pack(side="right")
        self.coord_tree = ttk.Treeview(f, columns=("id", "對象單位", "議題", "負責人", "到期日", "狀態", "建立日"), show="headings", height=18)
        for c, w in zip(("id", "對象單位", "議題", "負責人", "到期日", "狀態", "建立日"), (40, 110, 640, 90, 100, 70, 100)):
            self.coord_tree.heading(c, text=c); self.coord_tree.column(c, width=w, anchor="w")
        self.coord_tree.pack(fill="both", expand=True, padx=12, pady=8)
        self.coord_tree.bind("<<TreeviewSelect>>", self._coord_pick)
        self.actions = self._coord_load()
        self._coord_render()

    def _coord_load(self):
        if os.path.exists(ACTIONS_FILE):
            try:
                with open(ACTIONS_FILE, encoding="utf-8") as fh:
                    return json.load(fh)
            except Exception:  # noqa
                return []
        return []

    def _coord_save(self):
        with open(ACTIONS_FILE, "w", encoding="utf-8") as fh:
            json.dump(self.actions, fh, ensure_ascii=False, indent=2)

    def _coord_render(self):
        self.coord_tree.delete(*self.coord_tree.get_children())
        for a in sorted(self.actions, key=lambda x: (x.get("狀態") == "已完成", x.get("到期日", ""))):
            tag = "late" if a.get("狀態") in ("進行中", "待回覆") and a.get("到期日", "9999") < dt.date.today().isoformat() else ""
            self.coord_tree.insert("", "end", values=(a["id"], a["對象單位"], a["議題"], a["負責人"], a["到期日"], a["狀態"], a["建立日"]), tags=(tag,))
        self.coord_tree.tag_configure("late", foreground="#B5544C")

    def _coord_add(self, unit=None, topic=None, owner="", due=None, status="進行中"):
        topic = topic if topic is not None else self.c_topic.get().strip()
        if not topic:
            messagebox.showwarning(APP_TITLE, "請填議題")
            return
        nid = (max([a["id"] for a in self.actions]) + 1) if self.actions else 1
        self.actions.append({"id": nid, "對象單位": unit or self.c_unit.get(), "議題": topic, "負責人": owner or self.c_owner.get(),
                             "到期日": due or self.c_due.get(), "狀態": status, "建立日": dt.date.today().isoformat()})
        self._coord_save(); self._coord_render()
        if topic == self.c_topic.get().strip():
            self.c_topic.set("")

    def _coord_pick(self, _e=None):
        sel = self.coord_tree.selection()
        if not sel:
            return
        v = self.coord_tree.item(sel[0])["values"]
        self.c_unit.set(v[1]); self.c_topic.set(v[2]); self.c_owner.set(v[3]); self.c_due.set(v[4]); self.c_status.set(v[5])

    def _coord_update(self):
        sel = self.coord_tree.selection()
        if not sel:
            return
        nid = int(self.coord_tree.item(sel[0])["values"][0])
        for a in self.actions:
            if a["id"] == nid:
                a.update({"對象單位": self.c_unit.get(), "議題": self.c_topic.get(), "負責人": self.c_owner.get(), "到期日": self.c_due.get(), "狀態": self.c_status.get()})
        self._coord_save(); self._coord_render()

    def _coord_delete(self):
        sel = self.coord_tree.selection()
        if not sel:
            return
        nid = int(self.coord_tree.item(sel[0])["values"][0])
        self.actions = [a for a in self.actions if a["id"] != nid]
        self._coord_save(); self._coord_render()

    def _coord_seed(self):
        if not self._need_data():
            return
        due = (dt.date.today() + dt.timedelta(days=14)).isoformat()
        seeds = {"1-2": "維修據點", "2-1": "維修據點", "2-2": "維修據點", "3-2": "總部服務團隊", "4-1": "各國規劃人員",
                 "4-2": "供應商", "5-1": "總部服務團隊", "5-2": "維修據點"}
        n = 0
        for key, unit in seeds.items():
            a = self.report.get(key)
            topic = f"[{a.key}] {a.conclusion}"
            if not any(x["議題"] == topic for x in self.actions):
                self._coord_add(unit=unit, topic=topic, owner="", due=due); n += 1
        messagebox.showinfo(APP_TITLE, f"已新增 {n} 筆待辦")

    def _coord_export(self):
        if not self.actions:
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="跨單位協調待辦.xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            pd.DataFrame(self.actions).to_excel(p, index=False); messagebox.showinfo(APP_TITLE, f"已匯出：{p}")


if __name__ == "__main__":
    App().mainloop()
