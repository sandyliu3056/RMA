# -*- coding: utf-8 -*-
"""
售後零件規劃平台（桌面版）
------------------------
六個模組對應工作架構圖：需求規劃、庫存與缺料、訂單與交期、出貨與到貨、報表與分析、跨單位協調。
資料來源：AIO 匯出檔（AllInOneData）。分析邏輯在 rma_engine.py，語言與名詞說明在 i18n.py。

執行：python 零件規劃平台.py
需要：pandas、openpyxl、matplotlib、python-pptx、pillow
"""
import os
import re
import sys
import json
import threading
import datetime as dt
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import i18n
from i18n import T, L, is_en, tr_df, bullet, define
import rma_engine as E

i18n.load_settings()

NAVY, ORANGE, BG, CARD, INK, MUTE = '#3B2A1A', '#E8862B', '#FBF6EC', '#FFFDF7', '#3B2A1A', '#7A6652'
GOLD, TABBAR, LINE, STRIPE, HOVER, SELECT = '#F2B134', '#F5A623', '#D9CBB6', '#F6EFE3', '#FFE2A8', '#F8D98F'
BROWN, BROWN_DK = '#6B4E31', '#4F3A24'
FONT = 'Microsoft JhengHei'
ACTIONS_FILE = os.path.join(BASE_DIR, '協調待辦.json')
UNITS = ['各國規劃人員', '供應商', '總部服務團隊', '倉庫／物流', '維修據點', '其他']
STATUSES = ['進行中', '待回覆', '已完成', '取消']
PO_COLS = ['訂單號', '料號', '供應商', '數量', '下單日', '承諾交期', '實際到貨日', '目的國']

# 欄名符合這個規則就以百分比顯示
PCT_RE = re.compile(r'(率|比例|占比|占在途|占保固外總數|累計占比)$')


def app_title():
    return T('售後零件規劃平台')


def ui_font():
    return 'Segoe UI' if is_en() else 'Microsoft JhengHei'


# ---------------------------------------------------------------- 數字格式
def col_kind(name, s, pct=False):
    """依欄位內容決定顯示格式：pct／int／num／bool／text。"""
    if s.dtype == bool:
        return 'bool'
    if not pd.api.types.is_numeric_dtype(s):
        return 'text'
    if pct or PCT_RE.search(str(name)):
        return 'pct'
    v = s.dropna()
    if len(v) and (v == v.round()).all():
        return 'int'
    return 'num'


def fmt_cell(v, kind='text'):
    try:
        if v is None or pd.isna(v):
            return ''
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return v.strftime('%Y-%m-%d')
    if kind == 'bool' or isinstance(v, bool):
        return L('是', 'Yes') if v else L('否', 'No')
    if kind == 'pct':
        return f'{float(v):.1%}'
    if kind == 'int':
        return f'{int(round(float(v))):,}'
    if kind == 'num':
        return f'{float(v):,.1f}'
    return str(v)


# ---------------------------------------------------------------- 滑鼠提示
class Tooltip:
    """滑鼠停留時顯示的小說明框。text 可為字串或回傳字串的函式。"""

    def __init__(self, widget, text, delay=450):
        self.widget, self.text, self.delay = widget, text, delay
        self.tip = None
        self._job = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self.hide, add='+')
        widget.bind('<ButtonPress>', self.hide, add='+')

    def _schedule(self, _e=None):
        self.hide()
        self._job = self.widget.after(self.delay, self.show)

    def show(self, x=None, y=None, text=None):
        msg = text if text is not None else (self.text() if callable(self.text) else self.text)
        if not msg:
            return
        self.hide()
        x = x if x is not None else self.widget.winfo_rootx() + 12
        y = y if y is not None else self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f'+{x}+{y}')
        tk.Label(tw, text=msg, justify='left', bg='#FFF8E1', fg=INK, relief='solid', bd=1,
                 font=(FONT, 9), wraplength=360, padx=8, pady=5).pack()

    def hide(self, _e=None):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ---------------------------------------------------------------- 表格
class Grid(ttk.Frame):
    """把 DataFrame 顯示成表格（Treeview），附捲軸、欄名說明與匯出。"""

    def __init__(self, master, height=12):
        super().__init__(master)
        self.max_height = height   # 最多顯示幾列；資料較少時表格自動縮短，不留空列
        self.tree = ttk.Treeview(self, show='headings', height=height)
        self.ys = ys = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        self.xs = xs = ttk.Scrollbar(self, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        ys.grid(row=0, column=1, sticky='ns')
        ys.grid_remove()   # 資料超過顯示列數時才出現
        xs.grid(row=1, column=0, sticky='ew')
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure('odd', background=STRIPE)
        self.tree.tag_configure('even', background=CARD)
        self.tree.tag_configure('warn', foreground='#B5544C', font=(FONT, 9, 'bold'))
        self.df = None
        self._defs = {}
        self._tip = Tooltip(self.tree, '')
        self.tree.unbind('<Enter>')
        self.tree.bind('<Motion>', self._hover, add='+')
        self._hover_col = None
        self._nat = {}          # 每欄依內容算出的「自然寬度」
        self._last_w = 0
        self.tree.bind('<Configure>', self._fit_columns, add='+')

    def _fit_columns(self, _e=None, force=False):
        """滿版：表格比內容寬時，把多出來的寬度按比例分給每一欄；不夠寬時保留原寬度，用橫向捲軸。"""
        if not self._nat:
            return
        w = self.tree.winfo_width()
        if w <= 1 or (w == self._last_w and not force):
            return
        self._last_w = w
        avail = w - 4
        cols = [c for c in self.tree['columns'] if c in self._nat]
        total = sum(self._nat[c] for c in cols)
        if not cols or total <= 0:
            return
        if avail <= total:
            for c in cols:
                self.tree.column(c, width=self._nat[c])
            self.xs.grid()          # 放不下才需要橫向捲軸
            return
        self.xs.grid_remove()
        used = 0
        for c in cols[:-1]:
            cw = int(self._nat[c] * avail / total)
            self.tree.column(c, width=cw)
            used += cw
        self.tree.column(cols[-1], width=avail - used)

    def _hover(self, e):
        if self.tree.identify_region(e.x, e.y) != 'heading':
            if self._hover_col:
                self._tip.hide(); self._hover_col = None
            return
        col = self.tree.identify_column(e.x)
        if col == self._hover_col:
            return
        self._hover_col = col
        self._tip.hide()
        try:
            name = self.tree['columns'][int(col[1:]) - 1]
        except (ValueError, IndexError):
            return
        msg = self._defs.get(name)
        if msg:
            self._tip.show(e.x_root + 12, e.y_root + 18, text=msg)

    def show(self, df, index=True, max_rows=2000, pct=False, highlight=None):
        """highlight：傳入一個函式，參數是內部欄名的資料列，回傳 True 的列以紅色粗體標示。"""
        self.tree.delete(*self.tree.get_children())
        if df is None or len(df) == 0:
            self.tree['columns'] = ['msg']
            self.tree.heading('msg', text=T('訊息'))
            self.tree.column('msg', width=300, anchor='w', stretch=True)
            self._nat = {'msg': 300}
            self.after_idle(lambda: self._fit_columns(force=True))
            self.tree.configure(height=2)
            self.ys.grid_remove()
            self.tree.insert('', 'end', values=[T('（沒有資料）')])
            self.df = None
            return
        d = df.reset_index() if index else df.reset_index(drop=True)
        d = d.rename(columns={'index': '項目'})
        raw_cols = [str(c) for c in d.columns]
        kinds = [col_kind(c, d.iloc[:, i], pct and i >= (d.shape[1] - df.shape[1])) for i, c in enumerate(raw_cols)]
        shown = tr_df(d)
        self.df = shown
        cols = [str(c) for c in shown.columns]
        # 欄名重複時 Treeview 會出錯，加上序號區分
        seen = {}
        ids = []
        for c in cols:
            seen[c] = seen.get(c, 0) + 1
            ids.append(c if seen[c] == 1 else f'{c} ({seen[c]})')
        self._defs = {i: define(r) for i, r in zip(ids, raw_cols) if define(r)}
        self.tree['columns'] = ids
        self._nat = {}
        for i, c in enumerate(ids):
            mark = ' ⓘ' if c in self._defs else ''
            self.tree.heading(c, text=c + mark)
            vals = [fmt_cell(v, kinds[i]) for v in shown.iloc[:50, i]]
            width = max(80, min(340, int(max([len(c) + 4] + [len(v) for v in vals]) * (7 if is_en() else 9) + 20)))
            num = kinds[i] in ('pct', 'int', 'num')
            self._nat[c] = width
            # 寬度由 _fit_columns 統一分配（滿版），這裡不讓 Tk 自行伸縮
            self.tree.column(c, width=width, minwidth=40, anchor='e' if num else 'w', stretch=False)
        self.tree.configure(height=max(3, min(len(shown), self.max_height)))
        self.after_idle(lambda: self._fit_columns(force=True))
        # 全部列都放得下就不需要表格自己的捲軸，交給整頁的大捲軸
        if len(shown) > self.max_height:
            self.ys.grid()
        else:
            self.ys.grid_remove()
        for r in range(min(len(shown), max_rows)):
            row = shown.iloc[r]
            tags = ['odd' if r % 2 else 'even']
            if highlight is not None:
                try:
                    if highlight(d.iloc[r]):
                        tags.append('warn')
                except Exception:
                    pass
            self.tree.insert('', 'end', values=[fmt_cell(row.iloc[i], kinds[i]) for i in range(len(ids))], tags=tuple(tags))

    def export(self, default_name='表格.xlsx'):
        if self.df is None:
            messagebox.showinfo(app_title(), T('目前沒有可匯出的表格'))
            return
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=T(default_name),
                                            filetypes=[('Excel', '*.xlsx'), ('CSV', '*.csv')])
        if not path:
            return
        def go():
            if path.lower().endswith('.csv'):
                self.df.to_csv(path, index=False, encoding='utf-8-sig')
            else:
                self.df.to_excel(path, index=False)
            messagebox.showinfo(app_title(), L(f'已匯出：{path}', f'Exported: {path}'))
        self.winfo_toplevel()._safe_export(go)


# ---------------------------------------------------------------- 圖表框（Tk 原生繪製，不需要 matplotlib）
class BarChart(tk.Canvas):
    """水平堆疊長條圖，最大的在最上面；視窗大小改變時自動重畫。"""

    def __init__(self, master, height=320):
        super().__init__(master, bg='white', height=height, highlightthickness=1, highlightbackground=LINE)
        self._data = None
        self.bind('<Configure>', lambda e: self._draw())

    def set_data(self, title, cats, series, hide_below=1):
        """series = [(名稱, 數值 list, 顏色)…]"""
        self._data = (title, cats, series, hide_below)
        self._draw()

    def _draw(self):
        self.delete('all')
        if not self._data:
            return
        title, cats, series, hide_below = self._data
        W, H = self.winfo_width(), self.winfo_height()
        if W < 120 or H < 80 or not cats:
            return
        from tkinter import font as tkfont
        f_lab = tkfont.Font(family=FONT, size=8)
        f_val = tkfont.Font(family=FONT, size=8)
        f_tit = tkfont.Font(family=FONT, size=10, weight='bold')
        self.create_text(12, 14, text=title, anchor='w', font=f_tit, fill=NAVY)
        # 圖例
        lx = W - 14
        for name, _, col in reversed(series):
            tw = f_lab.measure(name)
            lx -= tw
            self.create_text(lx, 34, text=name, anchor='w', font=f_lab, fill=INK)
            lx -= 16
            self.create_rectangle(lx, 29, lx + 10, 39, fill=col, outline='')
            lx -= 14
        left = min(max(f_lab.measure(str(c)) for c in cats) + 22, int(W * 0.45))
        top, bottom, right = 46, 8, f_val.measure('00,000') + 14
        n = len(cats)
        row = (H - top - bottom) / n
        bh = max(8, row * 0.64)
        totals = [sum(v[i] for _, v, _ in series) for i in range(n)]
        vmax = max(totals) or 1
        scale = (W - left - right) / vmax
        for i, cat in enumerate(cats):
            y = top + i * row + (row - bh) / 2
            label = str(cat)
            while f_lab.measure(label) > left - 16 and len(label) > 4:
                label = label[:-2]
            if label != str(cat):
                label = label.rstrip() + '…'
            self.create_text(left - 10, y + bh / 2, text=label, anchor='e', font=f_lab, fill=INK)
            x = left
            for name, vals, col in series:
                v = vals[i]
                w = v * scale
                if w > 0:
                    self.create_rectangle(x, y, x + w, y + bh, fill=col, outline='')
                    txt = f'{int(v):,}'
                    if v >= hide_below and f_val.measure(txt) + 6 < w:
                        dark = col.upper() in ('#C9B79C', '#F2B134', '#FBE7C6')
                        self.create_text(x + w / 2, y + bh / 2, text=txt, font=f_val, fill=NAVY if dark else 'white')
                x += w
            # 合計標在長條右邊
            self.create_text(x + 6, y + bh / 2, text=f'{int(totals[i]):,}', anchor='w', font=f_val, fill=MUTE)


# ---------------------------------------------------------------- 缺套件時的安裝提示
def pip_python():
    """pythonw.exe 沒有主控台，安裝套件改用同資料夾的 python.exe。"""
    exe = sys.executable
    if os.path.basename(exe).lower() == 'pythonw.exe':
        alt = os.path.join(os.path.dirname(exe), 'python.exe')
        if os.path.exists(alt):
            return alt
    return exe


def offer_install(parent_label, module='matplotlib'):
    """在圖表位置顯示「安裝圖表套件」按鈕；用目前執行程式的 Python 安裝，避免裝到別的 Python。"""
    host = parent_label.master
    for w in host.winfo_children():
        if getattr(w, '_install_btn', False):
            w.destroy()
    parent_label.config(image='', text=L(f'圖表需要 {module} 套件，目前尚未安裝。', f'Charts need the {module} package, which is not installed.'))

    def run():
        import subprocess
        btn.config(state='disabled', text=L('安裝中，約 1–3 分鐘…', 'Installing, about 1–3 minutes…'))

        def work():
            try:
                r = subprocess.run([pip_python(), '-m', 'pip', 'install', module], capture_output=True, text=True,
                                   creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
                ok = r.returncode == 0
                msg = r.stdout[-800:] + r.stderr[-800:]
            except Exception as ex:
                ok, msg = False, str(ex)
            def done():
                if ok:
                    messagebox.showinfo(app_title(), L('安裝完成。請關閉程式再重新開啟，圖表就會出現。', 'Installed. Close and reopen the program to see the charts.'))
                    btn.config(text=L('已安裝，請重新開啟程式', 'Installed – please restart'))
                else:
                    messagebox.showerror(app_title(), L('安裝失敗：\n', 'Install failed:\n') + msg)
                    btn.config(state='normal', text=L('重試安裝', 'Retry install'))
            host.after(0, done)
        threading.Thread(target=work, daemon=True).start()

    btn = ttk.Button(host, text=L(f'安裝圖表套件（{module}）', f'Install chart package ({module})'), style='Accent.TButton', command=run)
    btn._install_btn = True
    btn.place(relx=0.5, rely=0.62, anchor='center')


# ---------------------------------------------------------------- 整頁捲動
class ScrollPage(ttk.Frame):
    """分頁容器：內容放在 .body，超出視窗高度時以右側大捲軸整頁捲動。"""

    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)
        self.body = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window(0, 0, window=self.body, anchor='nw')
        self.body.bind('<Configure>', self._sync)
        self.canvas.bind('<Configure>', self._sync)

    def _sync(self, _e=None):
        # 內容寬度跟著視窗；捲動範圍跟著內容高度
        self.canvas.itemconfigure(self._win, width=self.canvas.winfo_width())
        self.canvas.configure(scrollregion=(0, 0, self.canvas.winfo_width(),
                                            max(self.body.winfo_reqheight(), self.canvas.winfo_height())))

    def scroll(self, steps):
        if self.body.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(steps, 'units')


# ---------------------------------------------------------------- 版面小工具
def kpi_row(master, items):
    """一列 KPI 方塊。items = [(標題, 數值)…]（標題用內部中文名），回傳以內部名稱為鍵的 label 字典。"""
    frame = ttk.Frame(master)
    labels = {}
    for i, (title, value) in enumerate(items):
        card = tk.Frame(frame, bg=CARD, bd=0, highlightthickness=1, highlightbackground=LINE)
        card.grid(row=0, column=i, padx=4, pady=4, sticky='nsew', ipadx=6, ipady=4)
        frame.columnconfigure(i, weight=1, uniform='kpi')
        tip = define(title)
        head = tk.Label(card, text=T(title) + ('  ⓘ' if tip else ''), bg=CARD, fg=MUTE, font=(FONT, 8),
                        wraplength=170, justify='left', anchor='w')
        head.pack(anchor='w', fill='x', padx=10)
        lab = tk.Label(card, text=value, bg=CARD, fg=NAVY, font=(FONT, 14, 'bold'))
        lab.pack(anchor='w', padx=10, pady=(0, 0))
        if tip:
            for w in (card, head, lab):
                Tooltip(w, tip)
        labels[title] = lab
    return frame, labels


def section(master, text, **pack):
    """分頁內的小節標題：橘色短槓＋粗體字。"""
    row = ttk.Frame(master)
    tk.Frame(row, bg=ORANGE, width=4, height=15).pack(side='left', padx=(0, 8))
    ttk.Label(row, text=T(text), style='Section.TLabel').pack(side='left')
    row.pack(anchor='w', padx=12, **(pack or {'pady': (6, 2)}))
    return row


def toolbar(master):
    """工具列：左邊放設定與主要動作，右邊放「匯出」群組。回傳 (left, right)。"""
    bar = ttk.Frame(master); bar.pack(fill='x', padx=12, pady=5)
    left = ttk.Frame(bar); left.pack(side='left')
    right = ttk.Frame(bar); right.pack(side='right')
    ttk.Label(right, text=L('匯出：', 'Export:'), foreground=MUTE).pack(side='left', padx=(0, 4))
    return left, right


def kv(n, share=None, digits=0):
    """KPI 數值：件數（占比）。"""
    if share is None:
        return f'{n:,}'
    return L(f'{n:,}（{share:.{digits}%}）', f'{n:,} ({share:.{digits}%})')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry('1240x760')
        self.minsize(980, 620)
        self.configure(bg=BG)
        self.raw = None
        self.data = None
        self.report = None
        self.parts = None
        self.chart_dir = os.path.join(BASE_DIR, '_charts')
        self._photo = None
        self._build_ui()
        for seq in ('<MouseWheel>', '<Button-4>', '<Button-5>'):
            self.bind_all(seq, self._on_wheel, add='+')

    def _on_wheel(self, e):
        """滑鼠滾輪：游標下的表格或文字框還能捲就交給它，否則捲動整頁。"""
        try:
            if e.widget.winfo_toplevel() is not self:
                return
            page = self.nametowidget(self.nb.select())
        except (KeyError, tk.TclError, AttributeError):
            return
        if not isinstance(page, ScrollPage):
            return
        w = e.widget
        while w is not None and w is not page:
            if isinstance(w, (ttk.Treeview, tk.Text, tk.Listbox)):
                first, last = w.yview()
                if first > 0 or last < 1:
                    return
                break
            w = w.master
        up = getattr(e, 'num', 0) == 4 or getattr(e, 'delta', 0) > 0
        page.scroll(-3 if up else 3)

    # ------------------------------------------------------------ 介面建立／語言切換
    def _apply_style(self):
        global FONT
        FONT = ui_font()
        style = ttk.Style(self)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        style.configure('.', background=BG, foreground=INK, font=(FONT, 9), bordercolor=LINE)
        style.configure('TFrame', background=BG)
        style.configure('TLabel', background=BG, foreground=INK, font=(FONT, 9))
        style.configure('Section.TLabel', font=(FONT, 10, 'bold'), foreground=INK)
        style.configure('Hint.TLabel', font=(FONT, 8), foreground=MUTE)
        style.configure('TLabelframe', background=BG, bordercolor=LINE)
        style.configure('TLabelframe.Label', background=BG, foreground=INK, font=(FONT, 9, 'bold'))

        style.configure('TNotebook', background=TABBAR, borderwidth=0, tabmargins=(6, 6, 6, 0))
        style.configure('TNotebook.Tab', font=(FONT, 9, 'bold'), padding=(12, 5), background=GOLD, foreground=INK, borderwidth=0)
        style.map('TNotebook.Tab', background=[('selected', BG), ('active', HOVER)], foreground=[('selected', INK)])

        style.configure('TButton', font=(FONT, 9), padding=(10, 3), background=CARD, foreground=INK, bordercolor=LINE, relief='flat')
        style.map('TButton', background=[('active', HOVER), ('pressed', SELECT)])
        style.configure('Accent.TButton', font=(FONT, 9, 'bold'), padding=(12, 3), foreground='white', background=BROWN, bordercolor=BROWN)
        style.map('Accent.TButton', background=[('active', BROWN_DK), ('pressed', NAVY)])
        style.configure('Export.TButton', font=(FONT, 8), padding=(8, 2), background=CARD, foreground=BROWN, bordercolor=LINE, borderwidth=1, relief='solid')
        style.map('Export.TButton', background=[('active', HOVER), ('pressed', SELECT)])

        style.configure('TEntry', fieldbackground=CARD, bordercolor=LINE, padding=2)
        style.configure('TCombobox', fieldbackground=CARD, background=CARD, bordercolor=LINE, padding=2)
        style.configure('TSpinbox', fieldbackground=CARD, bordercolor=LINE, padding=2)

        style.configure('Treeview', font=(FONT, 9), rowheight=21, background=CARD, fieldbackground=CARD, foreground=INK, bordercolor=LINE)
        style.configure('Treeview.Heading', font=(FONT, 9, 'bold'), background=BROWN, foreground='white', relief='flat', padding=3)
        style.map('Treeview.Heading', background=[('active', BROWN_DK)])
        style.map('Treeview', background=[('selected', SELECT)], foreground=[('selected', INK)])
        style.configure('TPanedwindow', background=BG)
        style.configure('Vertical.TScrollbar', background=BG, troughcolor=STRIPE, bordercolor=LINE, arrowcolor=MUTE)
        style.configure('Horizontal.TScrollbar', background=BG, troughcolor=STRIPE, bordercolor=LINE, arrowcolor=MUTE)

    def _build_ui(self):
        self._apply_style()
        self.title(app_title())

        header = tk.Frame(self, bg=NAVY)
        header.pack(fill='x')
        left = tk.Frame(header, bg=NAVY); left.pack(side='left', padx=14, pady=5)
        tk.Label(left, text=app_title(), bg=NAVY, fg=GOLD, font=(FONT, 14, 'bold')).pack(anchor='w')
        self.status_var = tk.StringVar(value=T('尚未載入資料'))
        tk.Label(left, textvariable=self.status_var, bg=NAVY, fg='#D9C7A8', font=(FONT, 8)).pack(anchor='w')
        right = tk.Frame(header, bg=NAVY); right.pack(side='right', padx=14, pady=4)
        self.clock_var = tk.StringVar(); self.date_var = tk.StringVar()
        btns = tk.Frame(right, bg=NAVY); btns.pack(side='left', padx=(0, 18))
        hb = dict(bg=BROWN, fg='white', activebackground=BROWN_DK, activeforeground='white', relief='flat',
                  bd=0, padx=10, pady=2, cursor='hand2', font=(FONT, 8, 'bold'))
        lang_btn = tk.Menubutton(btns, text=L('語言：中文 ▾', 'Language: English ▾'), **hb)
        lang_menu = tk.Menu(lang_btn, tearoff=0, font=(FONT, 9), bg=CARD, fg=INK,
                            activebackground=HOVER, activeforeground=INK)
        self._lang_choice = tk.StringVar(value=i18n.get_lang())
        for code, name in (('zh', '中文'), ('en', 'English')):
            lang_menu.add_radiobutton(label=name, value=code, variable=self._lang_choice,
                                      command=lambda c=code: self._set_lang(c))
        lang_btn.config(menu=lang_menu)
        lang_btn.pack(side='left', padx=4)
        tk.Button(btns, text=L('名詞說明', 'Glossary'), command=self._show_glossary, **hb).pack(side='left', padx=4)
        clock = tk.Frame(right, bg=NAVY); clock.pack(side='left')
        tk.Label(clock, textvariable=self.clock_var, bg=NAVY, fg=GOLD, font=('Consolas', 15, 'bold')).pack(anchor='e')
        tk.Label(clock, textvariable=self.date_var, bg=NAVY, fg='#D9C7A8', font=(FONT, 8)).pack(anchor='e')
        self._tick()

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True)
        self.tab_source = ScrollPage(self.nb); self.nb.add(self.tab_source, text=T('資料來源'))
        self.tab_demand = ScrollPage(self.nb); self.nb.add(self.tab_demand, text=T('需求規劃'))
        self.tab_stock = ScrollPage(self.nb); self.nb.add(self.tab_stock, text=T('庫存與缺料'))
        self.tab_po = ScrollPage(self.nb); self.nb.add(self.tab_po, text=T('訂單與交期'))
        self.tab_ship = ScrollPage(self.nb); self.nb.add(self.tab_ship, text=T('出貨與到貨'))
        self.tab_report = ScrollPage(self.nb); self.nb.add(self.tab_report, text=T('報表與分析'))
        self.tab_coord = ScrollPage(self.nb); self.nb.add(self.tab_coord, text=T('跨單位協調'))
        self._build_source()
        self._build_demand()
        self._build_stock()
        self._build_po()
        self._build_ship()
        self._build_report()
        self._build_coord()

    def _set_lang(self, code):
        if code != i18n.get_lang():
            self._switch_lang(code)

    def _switch_lang(self, code=None):
        keep = (self.path_var.get(), self.start_var.get(), self.nb.index('current'))
        i18n.set_lang(code or ('zh' if is_en() else 'en'))
        i18n.save_settings()
        if getattr(self, '_tick_job', None):
            self.after_cancel(self._tick_job)
        for w in self.winfo_children():
            w.destroy()
        self.config(cursor='watch'); self.update_idletasks()
        self._build_ui()
        self.path_var.set(keep[0]); self.start_var.set(keep[1])
        try:
            if self.data is not None:
                # 分析文字依語言產生，切換後重跑一次（不需重新讀檔）
                self.report = E.run_all(self.data)
                self.parts = self.report.parts
                self._after_load()
            self.nb.select(keep[2])
        finally:
            self.config(cursor='')

    def _show_glossary(self):
        win = tk.Toplevel(self)
        win.title(L('名詞說明', 'Glossary'))
        win.geometry('760x560')
        win.configure(bg=BG)
        ttk.Label(win, text=L('畫面上標有 ⓘ 的欄位或指標，滑鼠停在上面也會顯示說明。',
                              'Anything marked ⓘ on screen also shows its definition when you hover over it.'),
                  style='Hint.TLabel').pack(anchor='w', padx=12, pady=(10, 4))
        box = ttk.Frame(win); box.pack(fill='both', expand=True, padx=12, pady=(0, 12))
        txt = tk.Text(box, wrap='word', font=(FONT, 9), bg=CARD, relief='flat', highlightthickness=1,
                      highlightbackground=LINE, padx=10, pady=8)
        sb = ttk.Scrollbar(box, orient='vertical', command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        txt.pack(side='left', fill='both', expand=True)
        txt.tag_configure('term', font=(FONT, 9, 'bold'), foreground=BROWN)
        seen = set()
        for key in i18n.GLOSSARY:
            name = T(key)
            if name in seen:
                continue
            seen.add(name)
            txt.insert('end', name + '\n', 'term')
            txt.insert('end', define(key) + '\n\n')
        txt.config(state='disabled')

    def _tick(self):
        now = dt.datetime.now()
        self.clock_var.set(now.strftime('%H:%M'))
        self.date_var.set(now.strftime('%Y/%m/%d %a'))
        self._tick_job = self.after(30000, self._tick)

    # ------------------------------------------------------------ 資料來源
    def _build_source(self):
        f = self.tab_source.body
        top = ttk.Frame(f); top.pack(fill='x', padx=12, pady=12)
        ttk.Label(top, text=T('AIO 匯出檔：')).grid(row=0, column=0, sticky='w')
        self.path_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.path_var, width=80).grid(row=0, column=1, padx=6, sticky='w')
        ttk.Button(top, text=T('瀏覽…'), command=self._browse).grid(row=0, column=2, sticky='w')
        ttk.Label(top, text=L('從哪一天開始分析：', 'Analyze cases from:')).grid(row=1, column=0, sticky='w', pady=(8, 0))
        self.start_var = tk.StringVar()
        start = ttk.Frame(top); start.grid(row=1, column=1, sticky='w', padx=6, pady=(8, 0))
        ttk.Entry(start, textvariable=self.start_var, width=14).pack(side='left')
        ttk.Label(start, text=L('（選填，例：2026-08-01）', '(optional, e.g. 2026-08-01)'), style='Hint.TLabel').pack(side='left', padx=6)
        self.period_var = tk.StringVar()
        ttk.Label(start, textvariable=self.period_var, font=(FONT, 9, 'bold'), foreground=BROWN).pack(side='left', padx=10)
        ttk.Label(top, text=L('只分析這天以後建立的案件。不填的話，會自動從「上個月 1 日」開始：例如資料最後一天是 9/21，就分析 8/1 到 9/21。',
                              'Only cases created on or after this date are included. Leave it blank to start from the 1st of the previous month: '
                              'if your data runs through Sep 21, the analysis covers Aug 1 – Sep 21.'),
                  style='Hint.TLabel', wraplength=700, justify='left').grid(row=2, column=1, sticky='w', padx=6, pady=(2, 0))
        ttk.Button(top, text=T('載入並分析'), style='Accent.TButton', command=self._load_clicked).grid(row=1, column=2, pady=(8, 0), sticky='w')
        self.kpi_frame, self.kpi_labels = kpi_row(f, [('案件數', '－'), ('在途', '－'), ('缺料（Awaiting Spares）', '－'),
                                                      ('零件未使用率', '－'), ('SLA 達成率', '－'), ('保固 30 天內到期在途', '－')])
        self.kpi_frame.pack(fill='x', padx=8)
        self.log = tk.Text(f, height=18, font=('Consolas', 9), bg=CARD, relief='flat', highlightthickness=1, highlightbackground=LINE)
        self.log.pack(fill='both', expand=True, padx=12, pady=8)
        self._log(T('1. 選擇 AIO 匯出檔（.xlsx，工作表 AllInOneData）\n2. 按「載入並分析」，約 10–30 秒\n3. 到各分頁查看，或在「報表與分析」匯出 Excel／PPT'))

    def _log(self, msg):
        self.log.insert('end', msg + '\n')
        self.log.see('end')
        self.update_idletasks()

    def _browse(self):
        p = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx *.xlsm')])
        if p:
            self.path_var.set(p)

    def _load_clicked(self):
        p = self.path_var.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showwarning(app_title(), T('請先選擇 AIO 匯出檔'))
            return
        s = self.start_var.get().strip()
        if s:
            try:
                if pd.Timestamp(s).year < 2000:   # 例如只打「8/1」會被當成西元 1 年
                    raise ValueError(s)
            except (ValueError, TypeError):
                messagebox.showwarning(app_title(), L(f'「{s}」不是有效的日期。\n請輸入像 2026-08-01 這樣的日期，或留空讓程式自動決定。',
                                                      f'"{s}" isn\'t a valid date.\nEnter a date like 2026-08-01, or leave it blank to use the default.'))
                return
        self.status_var.set(T('載入中…'))
        threading.Thread(target=self._load_worker, args=(p, self.start_var.get().strip() or None), daemon=True).start()

    def _load_worker(self, path, start):
        try:
            self._log(L(f'讀取 {os.path.basename(path)} …', f'Reading {os.path.basename(path)} …'))
            raw = E.load_aio(path)
            self._log(L(f'  共 {len(raw):,} 列，欄位 {raw.shape[1]} 個', f'  {len(raw):,} rows, {raw.shape[1]} columns'))
            data = E.prepare(raw, start)
            self._log(L(f"  分析期間 {data.attrs['period_start']:%Y-%m-%d} 起，資料日期 {data.attrs['ref_date']:%Y-%m-%d}，{len(data):,} 件",
                        f"  Period from {data.attrs['period_start']:%Y-%m-%d}, data as of {data.attrs['ref_date']:%Y-%m-%d}, {len(data):,} cases"))
            rep = E.run_all(data, progress=lambda i, n, t: self._log(f'  [{i}/{n}] {t}'))
            self._log(L(f'  零件明細 {len(rep.parts):,} 顆', f'  {len(rep.parts):,} part lines'))
            self.raw, self.data, self.report, self.parts = raw, data, rep, rep.parts
            self.after(0, self._after_load)
        except Exception as ex:
            import traceback
            self._log(T('錯誤：') + traceback.format_exc())
            self.after(0, lambda: messagebox.showerror(app_title(), L(f'載入失敗：{ex}', f'Load failed: {ex}')))
            self.after(0, lambda: self.status_var.set(T('載入失敗')))

    def _after_load(self):
        rep, d = self.report, self.data
        k12, k22, k41, k42, k52 = rep.get('1-2').kpis, rep.get('2-2').kpis, rep.get('4-1').kpis, rep.get('4-2').kpis, rep.get('5-2').kpis
        self.kpi_labels['案件數'].config(text=f'{len(d):,}')
        self.kpi_labels['在途'].config(text=kv(k12['在途件數'], k12['在途占比']))
        self.kpi_labels['缺料（Awaiting Spares）'].config(text=kv(k42['缺料件數'], k42['缺料占在途']))
        self.kpi_labels['零件未使用率'].config(text=f"{k41['未使用率']:.1%}")
        self.kpi_labels['SLA 達成率'].config(text=f"{k22['整體達成率']:.0%}")
        self.kpi_labels['保固 30 天內到期在途'].config(text=f"{k52['30天內到期件數']:,}")
        self.status_var.set(L(f"{os.path.basename(self.path_var.get())}｜{d.attrs['period_start']:%m/%d}–{d.attrs['ref_date']:%m/%d}｜{len(d):,} 件",
                              f"{os.path.basename(self.path_var.get())} | {d.attrs['period_start']:%b %d}–{d.attrs['ref_date']:%b %d} | {len(d):,} cases"))
        self.period_var.set(L(f"本次分析：{d.attrs['period_start']:%Y-%m-%d} ～ {d.attrs['ref_date']:%Y-%m-%d}",
                              f"Analyzing {d.attrs['period_start']:%b %d, %Y} – {d.attrs['ref_date']:%b %d, %Y}"))
        self._log(T('完成。'))
        self._refresh_demand()
        self._refresh_stock()
        self._refresh_ship()
        self._refresh_report_list()

    def _need_data(self):
        if self.report is None:
            messagebox.showinfo(app_title(), T('請先在「資料來源」載入 AIO 匯出檔'))
            return False
        return True

    # ------------------------------------------------------------ 需求規劃
    def _build_demand(self):
        f = self.tab_demand.body
        left, right = toolbar(f)
        ttk.Label(left, text=T('彙總依：')).pack(side='left')
        self._by_map = {T(k): k for k in ('國家', '類型分組')}
        self.demand_by = tk.StringVar(value=T('國家'))
        cb = ttk.Combobox(left, textvariable=self.demand_by, values=list(self._by_map), width=12, state='readonly')
        cb.pack(side='left', padx=4)
        cb.bind('<<ComboboxSelected>>', lambda e: self._refresh_demand())
        ttk.Label(left, text=T('預測未來（週）：')).pack(side='left', padx=(12, 0))
        self.demand_weeks = tk.IntVar(value=4)
        ttk.Spinbox(left, from_=1, to=13, textvariable=self.demand_weeks, width=5, command=self._refresh_demand).pack(side='left', padx=4)
        ttk.Button(left, text=T('重新計算'), style='Accent.TButton', command=self._refresh_demand).pack(side='left', padx=8)
        ttk.Button(right, text=L('每週申請量', 'Weekly Requests'), style='Export.TButton', command=lambda: self.demand_grid.export('週申請零件量.xlsx')).pack(side='left', padx=2)
        ttk.Button(right, text=L('料號排行', 'Top Parts'), style='Export.TButton', command=lambda: self.parts_grid.export('Top料號.xlsx')).pack(side='left', padx=2)
        section(f, '每週申請零件顆數（完整週）與未來需求推估')
        self.demand_grid = Grid(f, height=30); self.demand_grid.pack(fill='both', expand=False, padx=12, pady=4)
        self.forecast_grid = Grid(f, height=30); self.forecast_grid.pack(fill='both', expand=False, padx=12, pady=4)
        section(f, '申請次數最多的料號（Top 30）：申請顆數、未使用率、目前待供貨顆數', pady=(8, 0))
        self.parts_grid = Grid(f, height=30); self.parts_grid.pack(fill='both', expand=True, padx=12, pady=4)

    def _refresh_demand(self):
        if self.report is None:
            return
        by = self._by_map.get(self.demand_by.get(), '國家')
        self.demand_grid.show(E.demand_weekly(self.data, by))
        self.forecast_grid.show(E.demand_forecast(self.data, self.demand_weeks.get(), by))
        self.parts_grid.show(E.top_parts(self.parts, 30))

    # ------------------------------------------------------------ 庫存與缺料
    def _build_stock(self):
        f = self.tab_stock.body
        self.stock_kpi_frame, self.stock_kpi = kpi_row(f, [('缺料案件', '－'), ('等待超過兩週', '－'), ('等待中位數', '－'),
                                                           ('零件在途／已配案件', '－'), ('待供貨料號數', '－')])
        self.stock_kpi_frame.pack(fill='x', padx=8, pady=(8, 0))
        left, right = toolbar(f)
        ttk.Label(left, text=T('催料清單：等待天數 ≥')).pack(side='left')
        self.stock_days = tk.IntVar(value=14)
        ttk.Spinbox(left, from_=1, to=60, textvariable=self.stock_days, width=5, command=self._refresh_stock).pack(side='left', padx=4)
        ttk.Label(left, text=L('天', 'days')).pack(side='left')
        ttk.Button(left, text=T('重新整理'), style='Accent.TButton', command=self._refresh_stock).pack(side='left', padx=8)
        ttk.Button(right, text=L('缺貨料號', 'Shortage Parts'), style='Export.TButton', command=lambda: self.short_part_grid.export('缺料料號.xlsx')).pack(side='left', padx=2)
        ttk.Button(right, text=L('追零件清單', 'Expedite List'), style='Export.TButton', command=lambda: self.short_case_grid.export('催料清單.xlsx')).pack(side='left', padx=2)
        section(f, '缺料料號（Awaiting Spares 案件中待供貨的料號）')
        self.short_part_grid = Grid(f, height=30); self.short_part_grid.pack(fill='both', expand=False, padx=12, pady=4)
        section(f, '催料案件清單', pady=(8, 0))
        self.short_case_grid = Grid(f, height=30); self.short_case_grid.pack(fill='both', expand=True, padx=12, pady=4)

    def _refresh_stock(self):
        if self.report is None:
            return
        k = self.report.get('4-2').kpis
        sp = E.shortage_parts(self.parts)
        self.stock_kpi['缺料案件'].config(text=kv(k['缺料件數'], k['缺料占在途']))
        self.stock_kpi['等待超過兩週'].config(text=kv(k['超過兩週件數'], k['超過兩週比例']))
        self.stock_kpi['等待中位數'].config(text=L(f"{k['等待中位數']:.0f} 天", f"{k['等待中位數']:.0f} days"))
        self.stock_kpi['零件在途／已配案件'].config(text=f"{k['零件在途或已配件數']:,}")
        self.stock_kpi['待供貨料號數'].config(text=f'{len(sp):,}')
        self.short_part_grid.show(sp)
        self.short_case_grid.show(E.shortage_cases(self.data, self.stock_days.get()), index=False)

    # ------------------------------------------------------------ 訂單與交期
    def _build_po(self):
        f = self.tab_po.body
        ttk.Label(f, text=T("AIO 資料沒有採購訂單，這個模組用匯入的訂單檔計算準交率與逾期清單。"), style='Hint.TLabel').pack(anchor="w", padx=12, pady=(10, 0))
        left, right = toolbar(f)
        ttk.Button(left, text=T("產生訂單範本"), command=self._po_template).pack(side="left")
        ttk.Button(left, text=T("匯入訂單檔…"), style="Accent.TButton", command=self._po_import).pack(side="left", padx=8)
        ttk.Button(right, text=L("逾期清單", "Overdue List"), style="Export.TButton", command=lambda: self.po_late_grid.export("逾期訂單.xlsx")).pack(side="left", padx=2)
        self.po_kpi_frame, self.po_kpi = kpi_row(f, [("訂單筆數", "－"), ("已到貨", "－"), ("準交率", "－"), ("平均交期（天）", "－"), ("未到貨且已逾期", "－")])
        self.po_kpi_frame.pack(fill="x", padx=8)
        section(f, "供應商準交率")
        self.po_sup_grid = Grid(f, height=30); self.po_sup_grid.pack(fill="both", expand=False, padx=12, pady=4)
        section(f, "逾期與未交訂單", pady=(8, 0))
        self.po_late_grid = Grid(f, height=30); self.po_late_grid.pack(fill="both", expand=True, padx=12, pady=4)

    def _po_template(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=T("零件訂單範本.xlsx"), filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        tmpl = pd.DataFrame([
            ["PO-2026-0001", "KP.04501.017", T("供應商A"), 50, "2026-09-10", "2026-09-12", "2026-09-11", "Thailand"],
            ["PO-2026-0002", "KT.CTE00.014", T("供應商B"), 20, "2026-09-12", "2026-09-20", "", "Indonesia"]],
            columns=[T(c) for c in PO_COLS])
        tmpl.to_excel(path, index=False)
        messagebox.showinfo(app_title(), L(f"範本已存到：{path}\n欄位：訂單號、料號、供應商、數量、下單日、承諾交期、實際到貨日（未到貨留空）、目的國",
                                           f"Template saved to: {path}\nColumns: {', '.join(T(c) for c in PO_COLS)} (leave Received Date blank if not received yet)"))

    def _po_import(self):
        path = filedialog.askopenfilename(filetypes=[(L("Excel／CSV", "Excel/CSV"), "*.xlsx *.xls *.csv")])
        if not path:
            return
        try:
            po = pd.read_csv(path) if path.lower().endswith(".csv") else pd.read_excel(path)
            # 中英文欄名都接受，內部統一成中文
            alias = {}
            for c in PO_COLS:
                i18n_lang = i18n.get_lang()
                for lang in ('zh', 'en'):
                    i18n.set_lang(lang); alias[T(c)] = c
                i18n.set_lang(i18n_lang)
            po = po.rename(columns={c: alias.get(str(c).strip(), c) for c in po.columns})
            need = PO_COLS[:7]
            miss = [c for c in need if c not in po.columns]
            if miss:
                raise ValueError(T("缺少欄位：") + i18n.list_sep().join(T(c) for c in miss))
            for c in ("下單日", "承諾交期", "實際到貨日"):
                po[c] = pd.to_datetime(po[c], errors="coerce")
            today = pd.Timestamp(dt.date.today())
            po["已到貨"] = po["實際到貨日"].notna()
            po["準交"] = po["已到貨"] & (po["實際到貨日"] <= po["承諾交期"])
            po["交期天數"] = (po["實際到貨日"] - po["下單日"]).dt.days
            po["逾期天數"] = (po["實際到貨日"].fillna(today) - po["承諾交期"]).dt.days.clip(lower=0)
            arrived = po[po["已到貨"]]
            sup = po.groupby("供應商").agg(訂單筆數=("訂單號", "size"), 已到貨=("已到貨", "sum"), 準交筆數=("準交", "sum"),
                                           平均交期天數=("交期天數", "mean"), 平均逾期天數=("逾期天數", "mean"))
            sup["準交率"] = (sup["準交筆數"] / sup["已到貨"].replace(0, float("nan"))).astype(float)
            late = po[(~po["已到貨"] & (po["承諾交期"] < today)) | (po["已到貨"] & ~po["準交"])].sort_values("逾期天數", ascending=False)
            self.po_kpi["訂單筆數"].config(text=f"{len(po):,}")
            self.po_kpi["已到貨"].config(text=f"{int(po['已到貨'].sum()):,}")
            self.po_kpi["準交率"].config(text=f"{arrived['準交'].mean() if len(arrived) else 0:.0%}")
            self.po_kpi["平均交期（天）"].config(text=f"{arrived['交期天數'].mean() if len(arrived) else 0:.1f}")
            self.po_kpi["未到貨且已逾期"].config(text=f"{int((~po['已到貨'] & (po['承諾交期'] < today)).sum()):,}")
            self.po_sup_grid.show(sup)
            self.po_late_grid.show(late[["訂單號", "料號", "供應商", "數量", "承諾交期", "實際到貨日", "逾期天數"] + (["目的國"] if "目的國" in po.columns else [])], index=False)
        except Exception as ex:
            messagebox.showerror(app_title(), L(f"匯入失敗：{ex}", f"Import failed: {ex}"))

    # ------------------------------------------------------------ 出貨與到貨
    def _build_ship(self):
        f = self.tab_ship.body
        self.ship_kpi_frame, self.ship_kpi = kpi_row(f, [("待出貨案件", "－"), ("待出貨超過 7 天", "－"), ("零件運送中（顆）", "－"),
                                                         ("零件已配貨（顆）", "－"), ("保固 30 天內到期在途", "－")])
        self.ship_kpi_frame.pack(fill="x", padx=8, pady=(8, 0))
        left, right = toolbar(f)
        ttk.Label(left, text=T("保固到期清單：天數內")).pack(side="left")
        self.exp_days = tk.IntVar(value=30)
        ttk.Spinbox(left, from_=7, to=180, increment=7, textvariable=self.exp_days, width=5, command=self._refresh_ship).pack(side="left", padx=4)
        ttk.Button(left, text=T("重新整理"), style="Accent.TButton", command=self._refresh_ship).pack(side="left", padx=8)
        ttk.Button(right, text=L("待寄回清單", "Ship Backlog"), style="Export.TButton", command=lambda: self.backlog_grid.export("待出貨積壓.xlsx")).pack(side="left", padx=2)
        ttk.Button(right, text=L("保固到期清單", "Expiry List"), style="Export.TButton", command=lambda: self.exp_grid.export("保固到期在途案.xlsx")).pack(side="left", padx=2)
        # 重點圖：一眼看出積壓在哪、零件卡在哪
        charts = ttk.Frame(f); charts.pack(fill="x", padx=12, pady=(4, 0))
        charts.columnconfigure(0, weight=1, uniform="ch"); charts.columnconfigure(1, weight=1, uniform="ch")
        self.backlog_chart = BarChart(charts, height=250); self.backlog_chart.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.transit_chart = BarChart(charts, height=250); self.transit_chart.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        section(f, "待出貨積壓（依據點）", pady=(12, 0))
        ttk.Label(f, text=L("紅字＝放超過 7 天的案件有 10 件以上的維修站", "Red = sites with 10+ cases waiting over 7 days"),
                  style="Hint.TLabel").pack(anchor="w", padx=26)
        self.backlog_grid = Grid(f, height=15); self.backlog_grid.pack(fill="both", expand=True, padx=12, pady=4)
        section(f, "零件運送中／已配貨（依國家，在途案件）", pady=(8, 0))
        self.transit_grid = Grid(f, height=15); self.transit_grid.pack(fill="both", expand=True, padx=12, pady=4)
        section(f, "在途且保固即將到期（含建案時保固內、現已過期）", pady=(8, 0))
        ttk.Label(f, text=L("紅字＝保固已經過期（建案時在保固內），結案時要按保固內處理", "Red = warranty already expired (was in warranty at creation); close as in-warranty"),
                  style="Hint.TLabel").pack(anchor="w", padx=26)
        self.exp_grid = Grid(f, height=30); self.exp_grid.pack(fill="both", expand=True, padx=12, pady=4)

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
        self.backlog_grid.show(bl, highlight=lambda r: r.get("超過7天", 0) >= 10)
        self.transit_grid.show(tr)
        self.exp_grid.show(ex, index=False, highlight=lambda r: r.get("距到期天數", 0) < 0)
        # 圖：待寄回 Top 10、各國零件運送中 vs 已配到 Top 7（Tk 原生繪製）
        top = bl.head(10)
        over = top["超過7天"].astype(int).tolist()
        within = (top["待出貨件數"] - top["超過7天"]).astype(int).tolist()
        self.backlog_chart.set_data(L("修好待寄回 Top 10 維修站", "Ready to ship: top 10 sites"), [str(x) for x in top.index],
                                    [(L("7 天內", "Within 7 days"), within, "#C9B79C"), (L("超過 7 天", "Over 7 days"), over, ORANGE)], hide_below=10)
        t7 = tr.head(7)
        zero = [0] * len(t7)
        alloc = t7["已配貨"].astype(int).tolist() if "已配貨" in t7.columns else zero
        transit = t7["運送中"].astype(int).tolist() if "運送中" in t7.columns else zero
        self.transit_chart.set_data(L("零件運送中 vs 已配到（前 7 國）", "Parts in transit vs allocated (top 7)"), [T(str(x)) for x in t7.index],
                                    [(L("已配到", "Allocated"), alloc, "#7FA650"), (L("運送中", "In transit"), transit, ORANGE)], hide_below=15)

    # ------------------------------------------------------------ 報表與分析
    def _build_report(self):
        f = self.tab_report.body
        left = ttk.Frame(f, width=300); left.pack(side="left", fill="y", padx=(12, 6), pady=8)
        section(left, "分析項目")
        self.an_list = tk.Listbox(left, width=32, height=14, font=(FONT, 9), exportselection=False, bg=CARD,
                                  relief="flat", highlightthickness=1, highlightbackground=LINE, selectbackground=SELECT,
                                  selectforeground=INK, activestyle="none")
        self.an_list.pack(fill="y", pady=4)
        self.an_list.bind("<<ListboxSelect>>", lambda e: self._show_analysis())
        ttk.Label(left, text=L("匯出", "Export"), style="Section.TLabel").pack(anchor="w", pady=(12, 2))
        ttk.Button(left, text=T("全部匯出到資料夾"), style="Accent.TButton", command=self._export_all).pack(fill="x", pady=2)
        ttk.Label(left, text=L("Excel＋PPT＋全部圖表一次產出", "Excel + PowerPoint + all charts at once"), style="Hint.TLabel").pack(anchor="w", pady=(0, 6))
        ttk.Button(left, text=T("匯出 Excel 表格"), style="Export.TButton", command=self._export_excel).pack(fill="x", pady=1)
        ttk.Button(left, text=T("匯出 PPT 報告"), style="Export.TButton", command=self._export_pptx).pack(fill="x", pady=1)
        ttk.Button(left, text=T("匯出全部圖 PNG"), style="Export.TButton", command=self._export_charts).pack(fill="x", pady=1)
        right = ttk.Frame(f); right.pack(side="left", fill="both", expand=True, padx=(6, 12), pady=8)
        self.an_title = ttk.Label(right, text=T("請先載入資料，再從左側選一項分析"), font=(FONT, 11, "bold"), foreground=BROWN)
        self.an_title.pack(anchor="w")
        tabf = ttk.Frame(right); tabf.pack(side="bottom", fill="both", expand=False, pady=(6, 0))
        # 上下排列：文字（高度依內容）→ 圖表（全寬）→ 表格；整頁用大捲軸捲動
        self.an_text = tk.Text(right, wrap="word", font=(FONT, 9), height=12, bg=CARD, relief="flat", highlightthickness=1,
                               highlightbackground=LINE, padx=10, pady=8, spacing1=2, spacing3=4)
        self.an_text.tag_configure("sub", foreground=MUTE)
        self.an_text.tag_configure("h", font=(FONT, 9, "bold"), foreground=BROWN, spacing1=8)
        self.an_text.tag_configure("concl", font=(FONT, 9, "bold"), foreground=INK, background="#FFF1D6", lmargin1=6, lmargin2=6)
        self.an_text.pack(fill="x", pady=(4, 0))
        self.an_text.bind("<Configure>", self._fit_text)
        chartf = tk.Frame(right, bg=CARD, height=380, highlightthickness=1, highlightbackground=LINE)
        chartf.pack(fill="x", pady=(8, 0))
        chartf.pack_propagate(False)
        self.chart_label = tk.Label(chartf, bg=CARD, fg=MUTE, text=T("（圖表）"))
        self.chart_label.pack(fill="both", expand=True)
        self.chart_label.bind("<Configure>", self._render_chart)
        bar = ttk.Frame(tabf); bar.pack(fill="x")
        ttk.Label(bar, text=T("表格：")).pack(side="left")
        self.table_var = tk.StringVar()
        self.table_cb = ttk.Combobox(bar, textvariable=self.table_var, width=40, state="readonly")
        self.table_cb.pack(side="left", padx=4)
        self.table_cb.bind("<<ComboboxSelected>>", lambda e: self._show_table())
        ttk.Label(bar, text=L("欄名有 ⓘ 的，滑鼠停上去看說明", "Hover over ⓘ column headers for definitions"), style="Hint.TLabel").pack(side="left", padx=10)
        ttk.Button(bar, text=T("匯出此表"), style="Export.TButton", command=lambda: self.an_grid.export("分析表格.xlsx")).pack(side="right")
        self.an_grid = Grid(tabf, height=20); self.an_grid.pack(fill="both", expand=True, pady=4)
        self._table_keys = []

    def _refresh_report_list(self):
        self.an_list.delete(0, "end")
        for a in self.report.analyses.values():
            self.an_list.insert("end", f"{a.key}  {E.title_part(a.title, -1)}")
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
        self.an_text.insert("end", a.subtitle + "\n", "sub")
        self.an_text.insert("end", L("重點", "Key points") + "\n", "h")
        for b in a.bullets:
            self.an_text.insert("end", bullet() + b + "\n")
        if a.paragraph:
            self.an_text.insert("end", L("說明", "What it shows") + "\n", "h")
            self.an_text.insert("end", a.paragraph + "\n")
        self.an_text.insert("end", "\n")
        self.an_text.insert("end", T("結論：") + a.conclusion + "\n", "concl")
        self._text_w = None
        self.after(30, self._fit_text)
        self._table_keys = list(a.tables.keys())
        self.table_cb["values"] = [T(k) for k in self._table_keys]
        self.table_cb.current(0)
        self._show_table()
        self._show_chart(a)

    def _fit_text(self, _e=None):
        """分析文字框的高度配合內容（以實際寬度換行後的行數），不留空白也不用框內捲動。"""
        w = self.an_text.winfo_width()
        if w < 200:
            self.after(50, self._fit_text)
            return
        if _e is not None and w == getattr(self, "_text_w", None):
            return
        self._text_w = w
        try:
            lines = self.an_text.count("1.0", "end", "displaylines")
            lines = lines[0] if isinstance(lines, tuple) else lines
            self.an_text.configure(height=max(6, min(int(lines or 0) + 1, 60)))
        except tk.TclError:
            pass

    def _show_table(self):
        a = self._current_analysis()
        if a is None or not self._table_keys:
            return
        key = self._table_keys[max(self.table_cb.current(), 0)]
        # 整張表都是比例的（例如「占比」「組成」）全部以百分比顯示
        whole_pct = ("占比)" in key) or key.endswith("_占比") or ("組成" in key)
        self.an_grid.show(a.tables.get(key), pct=whole_pct)

    def _show_chart(self, a):
        try:
            os.makedirs(self.chart_dir, exist_ok=True)
            p = os.path.join(self.chart_dir, f"{a.key}.png")
            if not a.chart or not os.path.exists(a.chart):
                E.plot_analysis(a, p)
            self._chart_path = a.chart
            self._render_chart()
        except ModuleNotFoundError as ex:
            self._chart_path = None
            offer_install(self.chart_label, getattr(ex, "name", None) or "matplotlib")
        except Exception as ex:
            self._chart_path = None
            self.chart_label.config(text=L(f"圖表無法顯示：{ex}", f"Can't show the chart: {ex}"), image="")

    def _render_chart(self, _e=None):
        """依圖表框目前大小縮放圖片；版面或視窗大小改變時重畫，圖不會被切掉。"""
        path = getattr(self, "_chart_path", None)
        if not path or not os.path.exists(path):
            return
        w, h = self.chart_label.winfo_width() - 8, self.chart_label.winfo_height() - 8
        if w < 100 or h < 100:
            return
        if getattr(self, "_chart_size", None) == (path, w, h) and _e is not None:
            return
        self._chart_size = (path, w, h)
        from PIL import Image, ImageTk
        im = Image.open(path)
        r = min(w / im.width, h / im.height, 1.0)
        im = im.resize((max(1, int(im.width * r)), max(1, int(im.height * r))))
        self._photo = ImageTk.PhotoImage(im)
        self.chart_label.config(image=self._photo, text="")

    def _ask_dir(self):
        return filedialog.askdirectory()

    # 匯出：缺套件時直接在程式裡安裝後重試；其他錯誤（例如檔案正開著）以對話框說明，不再默默失敗
    PIP_NAME = {"pptx": "python-pptx", "PIL": "pillow", "openpyxl": "openpyxl", "matplotlib": "matplotlib"}

    def _safe_export(self, fn, retried=False):
        self.config(cursor="watch"); self.update_idletasks()
        try:
            fn()
        except ModuleNotFoundError as ex:
            mod = (ex.name or "").split(".")[0]
            pkg = self.PIP_NAME.get(mod, mod)
            if retried or not pkg:
                messagebox.showerror(app_title(), L(f"匯出失敗：缺少 {pkg} 套件。", f"Export failed: the {pkg} package is missing."))
                return
            if messagebox.askyesno(app_title(), L(f"匯出需要 {pkg} 套件，這台電腦還沒安裝。\n\n要現在自動安裝嗎？（約 1–3 分鐘，需要網路）",
                                                  f"Exporting needs the {pkg} package, which isn't installed on this computer.\n\nInstall it now? (about 1–3 minutes, needs internet)")):
                self._install_then(pkg, lambda: self._safe_export(fn, retried=True))
        except PermissionError as ex:
            messagebox.showerror(app_title(), L(f"沒辦法寫入檔案：\n{ex.filename or ex}\n\n同名檔案可能正用 Excel 或 PowerPoint 開著，請先關閉再匯出；或換一個資料夾。",
                                               f"Couldn't write the file:\n{ex.filename or ex}\n\nA file with the same name may be open in Excel or PowerPoint. Close it and try again, or pick another folder."))
        except Exception as ex:
            messagebox.showerror(app_title(), L("匯出失敗：\n", "Export failed:\n") + f"{type(ex).__name__}: {ex}")
        finally:
            self.config(cursor="")

    def _install_then(self, pkg, after):
        """用目前執行程式的 Python 安裝套件（避免裝到別的 Python），完成後執行 after。"""
        import subprocess, importlib
        win = tk.Toplevel(self); win.title(app_title()); win.transient(self); win.resizable(False, False)
        ttk.Label(win, text=L(f"正在安裝 {pkg}，約 1–3 分鐘，請稍候…", f"Installing {pkg}, about 1–3 minutes…"), padding=20).pack()
        bar = ttk.Progressbar(win, mode="indeterminate", length=320); bar.pack(padx=20, pady=(0, 20)); bar.start(12)
        win.grab_set()

        def work():
            try:
                r = subprocess.run([pip_python(), "-m", "pip", "install", pkg], capture_output=True, text=True,
                                   creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                ok, msg = r.returncode == 0, (r.stdout[-600:] + r.stderr[-600:])
            except Exception as ex:
                ok, msg = False, str(ex)

            def done():
                win.grab_release(); win.destroy()
                if ok:
                    importlib.invalidate_caches()
                    after()
                else:
                    messagebox.showerror(app_title(), L("安裝失敗：\n", "Install failed:\n") + msg)
            self.after(0, done)
        threading.Thread(target=work, daemon=True).start()

    def _export_excel(self):
        if not self._need_data():
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=T("RMA分析結果.xlsx"), filetypes=[("Excel", "*.xlsx")])
        if p:
            def go():
                E.export_excel(self.report, p); messagebox.showinfo(app_title(), L(f"已匯出：{p}", f"Exported: {p}"))
            self._safe_export(go)

    def _export_pptx(self):
        if not self._need_data():
            return
        p = filedialog.asksaveasfilename(defaultextension=".pptx", initialfile=T("RMA分析報告.pptx"), filetypes=[("PowerPoint", "*.pptx")])
        if p:
            def go():
                E.export_pptx(self.report, p, os.path.join(os.path.dirname(p), "charts")); messagebox.showinfo(app_title(), L(f"已匯出：{p}", f"Exported: {p}"))
            self._safe_export(go)

    def _export_charts(self):
        if not self._need_data():
            return
        d = self._ask_dir()
        if d:
            def go():
                out = E.export_charts(self.report, d)
                messagebox.showinfo(app_title(), L(f"已匯出 {len(out)} 張圖到 {d}", f"Exported {len(out)} charts to {d}"))
            self._safe_export(go)

    def _export_all(self):
        if not self._need_data():
            return
        d = self._ask_dir()
        if d:
            def go():
                out = E.export_all(self.report, d)
                messagebox.showinfo(app_title(), L(f"已匯出：\n{out['excel']}\n{out['pptx']}\n{len(out['charts'])} 張圖",
                                                   f"Exported:\n{out['excel']}\n{out['pptx']}\n{len(out['charts'])} charts"))
            self._safe_export(go)

    # ------------------------------------------------------------ 跨單位協調
    # 待辦檔內一律存中文內部值（單位、狀態），兩種語言共用同一個檔案；畫面上才翻譯
    def _build_coord(self):
        f = self.tab_coord.body
        self._unit_rev = {T(u): u for u in UNITS}
        self._status_rev = {T(s): s for s in STATUSES}
        form = ttk.LabelFrame(f, text=T("新增／更新待辦")); form.pack(fill="x", padx=12, pady=8)
        self.c_unit = tk.StringVar(value=T(UNITS[0])); self.c_topic = tk.StringVar(); self.c_owner = tk.StringVar()
        self.c_due = tk.StringVar(value=(dt.date.today() + dt.timedelta(days=7)).isoformat()); self.c_status = tk.StringVar(value=T("進行中"))
        ttk.Label(form, text=T("對象單位")).grid(row=0, column=0, padx=6, pady=4, sticky="e")
        ttk.Combobox(form, textvariable=self.c_unit, values=[T(u) for u in UNITS], width=20, state="readonly").grid(row=0, column=1, sticky="w")
        ttk.Label(form, text=T("議題")).grid(row=0, column=2, padx=6, sticky="e")
        ttk.Entry(form, textvariable=self.c_topic, width=60).grid(row=0, column=3, columnspan=3, sticky="w")
        ttk.Label(form, text=T("負責人")).grid(row=1, column=0, padx=6, pady=4, sticky="e")
        ttk.Entry(form, textvariable=self.c_owner, width=22).grid(row=1, column=1, sticky="w")
        ttk.Label(form, text=T("到期日")).grid(row=1, column=2, padx=6, sticky="e")
        ttk.Entry(form, textvariable=self.c_due, width=14).grid(row=1, column=3, sticky="w")
        ttk.Label(form, text=T("狀態")).grid(row=1, column=4, padx=6, sticky="e")
        ttk.Combobox(form, textvariable=self.c_status, values=[T(s) for s in STATUSES], width=14, state="readonly").grid(row=1, column=5, sticky="w")
        btns = ttk.Frame(form); btns.grid(row=0, column=6, rowspan=2, padx=10)
        ttk.Button(btns, text=T("新增"), style="Accent.TButton", command=self._coord_add).pack(fill="x", pady=2)
        ttk.Button(btns, text=T("更新所選"), command=self._coord_update).pack(fill="x", pady=2)
        ttk.Button(btns, text=T("刪除所選"), command=self._coord_delete).pack(fill="x", pady=2)
        left, right = toolbar(f)
        ttk.Button(left, text=T("由分析結論產生待辦"), command=self._coord_seed).pack(side="left")
        ttk.Label(left, text=L("逾期未完成的待辦會以紅字顯示", "Overdue open items show in red"), style="Hint.TLabel").pack(side="left", padx=10)
        ttk.Button(right, text=L("待辦清單", "Follow-ups"), style="Export.TButton", command=self._coord_export).pack(side="left", padx=2)
        cols = ("id", "對象單位", "議題", "負責人", "到期日", "狀態", "建立日")
        self.coord_tree = ttk.Treeview(f, columns=cols, show="headings", height=18)
        for c, w in zip(cols, (40, 140, 600, 100, 100, 110, 100)):
            self.coord_tree.heading(c, text=T(c) if c != "id" else "#"); self.coord_tree.column(c, width=w, anchor="w")
        self.coord_tree.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        self.coord_tree.bind("<<TreeviewSelect>>", self._coord_pick)
        self.actions = self._coord_load()
        self._coord_render()

    def _coord_load(self):
        if os.path.exists(ACTIONS_FILE):
            try:
                with open(ACTIONS_FILE, encoding="utf-8") as fh:
                    return json.load(fh)
            except Exception:
                return []
        return []

    def _coord_save(self):
        with open(ACTIONS_FILE, "w", encoding="utf-8") as fh:
            json.dump(self.actions, fh, ensure_ascii=False, indent=2)

    def _coord_render(self):
        self.coord_tree.delete(*self.coord_tree.get_children())
        for a in sorted(self.actions, key=lambda x: (x.get("狀態") == "已完成", x.get("到期日", ""))):
            tag = "late" if a.get("狀態") in ("進行中", "待回覆") and a.get("到期日", "9999") < dt.date.today().isoformat() else ""
            self.coord_tree.insert("", "end", values=(a["id"], T(a["對象單位"]), a["議題"], a["負責人"], a["到期日"], T(a["狀態"]), a["建立日"]), tags=(tag,))
            if len(self.coord_tree.get_children()) % 2 == 0 and not tag:
                self.coord_tree.item(self.coord_tree.get_children()[-1], tags=("odd",))
        self.coord_tree.tag_configure("late", foreground="#B5544C")
        self.coord_tree.tag_configure("odd", background=STRIPE)

    def _coord_add(self, unit=None, topic=None, owner="", due=None, status="進行中"):
        topic = topic if topic is not None else self.c_topic.get().strip()
        if not topic:
            messagebox.showwarning(app_title(), T("請填議題"))
            return
        nid = max([a["id"] for a in self.actions]) + 1 if self.actions else 1
        unit = unit or self._unit_rev.get(self.c_unit.get(), self.c_unit.get())
        self.actions.append({"id": nid, "對象單位": unit, "議題": topic, "負責人": owner or self.c_owner.get(),
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
                a.update({"對象單位": self._unit_rev.get(self.c_unit.get(), self.c_unit.get()), "議題": self.c_topic.get(),
                          "負責人": self.c_owner.get(), "到期日": self.c_due.get(),
                          "狀態": self._status_rev.get(self.c_status.get(), self.c_status.get())})
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
        messagebox.showinfo(app_title(), L(f"已新增 {n} 筆待辦", f"Added {n} follow-ups"))

    def _coord_export(self):
        if not self.actions:
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=T("跨單位協調待辦.xlsx"), filetypes=[("Excel", "*.xlsx")])
        if p:
            tr_df(pd.DataFrame(self.actions)).to_excel(p, index=False)
            messagebox.showinfo(app_title(), L(f"已匯出：{p}", f"Exported: {p}"))


if __name__ == "__main__":
    App().mainloop()
