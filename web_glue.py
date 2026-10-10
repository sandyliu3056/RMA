# -*- coding: utf-8 -*-
"""
瀏覽器版（Pyodide）的 Python 端。
worker.js 把 SheetJS 讀出來的列資料丟進來，這裡呼叫 rma_engine 做分析，
回傳可直接畫在畫面上的 JSON（表格、KPI、圖的 base64）。
所有運算都在使用者自己的瀏覽器裡完成，資料不會上傳到任何伺服器。

顯示規則與桌面版（零件規劃平台.py）一致：
- 內部欄名、分類值一律中文；顯示時經 i18n.T()／tr_df() 轉成目前語言（中文模式套用白話用詞）
- 欄名符合 PCT_RE 的以百分比顯示；整張都是比例的表（占比／組成）全部以百分比顯示
- 欄名或 KPI 有名詞定義的，一併回傳 defs 讓畫面加上 ⓘ 提示
"""
import io
import os
import re
import json
import base64
import zipfile
import tempfile
import datetime as dt

import numpy as np
import pandas as pd

import i18n
from i18n import T, L, is_en, tr_df, define, bullet, list_sep
import rma_engine as E

STATE = {"rep": None, "data": None, "po": None, "charts": {}, "chart_dir": None}
UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"]
STATUSES = ["進行中", "待回覆", "已完成", "取消"]
PO_COLS = ["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日", "目的國"]
SEEDS = {"1-2": "維修據點", "2-1": "維修據點", "2-2": "維修據點", "3-2": "總部服務團隊", "4-1": "各國規劃人員",
         "4-2": "供應商", "5-1": "總部服務團隊", "5-2": "維修據點"}

# 欄名符合這個規則就以百分比顯示（與桌面版相同）
PCT_RE = re.compile(r"(率|比例|占比|占在途|占保固外總數|累計占比)$")

# 網頁版才有的字串（i18n.py 沒有的）：英文模式的對照；中文模式直接用原字串
WEB_EN = {
    "名詞說明": "Glossary",
    "畫面上標有 ⓘ 的欄位或指標，滑鼠停在上面也會顯示說明。": "Anything marked ⓘ on screen also shows its definition when you hover over it.",
    "Windows 版下載說明": "Windows download guide",
    "啟動瀏覽器版 Python…": "Starting Python in your browser…",
    "瀏覽器版 Python 已就緒，請選擇 AIO 匯出檔。": "Python is ready. Pick an AIO export file.",
    "瀏覽器版 Python 還在啟動": "Python is still starting up",
    "啟動失敗：": "Startup failed: ",
    "AIO 匯出檔（.xlsx，工作表 AllInOneData）": "AIO export file (.xlsx, sheet AllInOneData)",
    "分析起日（留空＝資料最後日期的前一個月 1 日）": "Analysis start date (blank = 1st of the month before the last data date)",
    "只分析這天以後建立的案件。不填的話，會自動從「上個月 1 日」開始：例如資料最後一天是 9/21，就分析 8/1 到 9/21。":
        "Only cases created on or after this date are included. Leave it blank to start from the 1st of the previous month: "
        "if your data runs through Sep 21, the analysis covers Aug 1 – Sep 21.",
    "分析起日格式不對，請輸入像 2026-08-01 這樣的日期，或留空。": "The start date isn't valid. Enter a date like 2026-08-01, or leave it blank.",
    "1. 選擇 AIO 匯出檔（.xlsx，工作表 AllInOneData）\n2. 按「載入並分析」，第一次約 1 分鐘（分析在你的瀏覽器裡進行，檔案不會上傳到任何伺服器）\n3. 到各分頁查看，或在「報表與分析」下載 Excel／PPT":
        "1. Pick the AIO export file (.xlsx, sheet AllInOneData)\n2. Click \"Load & Analyze\" (about a minute the first time; everything runs in your browser and the file is never uploaded)\n3. Browse the tabs, or download Excel/PowerPoint from \"Reports & Analysis\"",
    "各狀態在途件數": "Open cases by status group",
    "載入失敗：": "Load failed: ",
    "匯出失敗：": "Export failed: ",
    "匯入失敗：": "Import failed: ",
    "讀取失敗：": "Read failed: ",
    "處理中…": "Working…",
    "分析中…": "Analyzing…",
    "產生中…": "Generating…",
    "匯入中…": "Importing…",
    "切換語言中…": "Switching language…",
    "下載瀏覽器版 Python 失敗，請檢查網路後重新整理頁面。": "Couldn't download the browser Python runtime. Check your connection and reload the page.",
    "記憶體不足，請關閉其他分頁後重試。": "Out of memory. Close other tabs and try again.",
    "只顯示前 {0} 列，共 {1} 列；匯出 Excel 可取得全部。": "Showing the first {0} of {1} rows; export to Excel for the full table.",
    "每週申請零件顆數（完整週）": "Parts requested per week (full weeks)",
    "匯入訂單檔（Excel／CSV）": "Import PO file (Excel/CSV)",
    "紅字＝放超過 7 天的案件有 10 件以上的維修站": "Red = sites with 10+ cases waiting over 7 days",
    "紅字＝保固已經過期（建案時在保固內），結案時要按保固內處理": "Red = warranty already expired (was in warranty at creation); close as in-warranty",
    "匯出": "Export",
    "欄名有 ⓘ 的，滑鼠停上去看說明": "Hover over ⓘ column headers for definitions",
    "畫圖中…": "Drawing the chart…",
    "（這一項沒有圖）": "(No chart for this one)",
    "無法顯示：": "Can't show: ",
    "待辦存在這台電腦的瀏覽器裡（換電腦看不到）。要帶到別台電腦，請用「下載待辦」存成 JSON，再「還原待辦」上傳。":
        "Follow-ups are stored in this browser only. To move them to another computer, use \"Download follow-ups\" to save a JSON file, then \"Restore follow-ups\" there.",
    "逾期未完成的待辦會以紅字顯示": "Overdue open items show in red",
    "下載待辦（JSON）": "Download follow-ups (JSON)",
    "還原待辦（JSON）": "Restore follow-ups (JSON)",
    "匯出待辦 CSV": "Export follow-ups (CSV)",
    "刪除": "Delete",
    "目前沒有待辦。": "No follow-ups yet.",
    "目前沒有待辦": "There are no follow-ups",
    "已新增 {0} 筆待辦": "Added {0} follow-ups",
    "已還原 {0} 筆": "Restored {0} items",
    "格式不對": "Wrong format",
    "無法儲存待辦（瀏覽器儲存空間不可用）": "Couldn't save follow-ups (browser storage unavailable)",
    "跨單位協調待辦.csv": "Cross_Team_Follow_ups.csv",
    "協調待辦.json": "follow_ups.json",
    "訂單檔沒有資料": "The PO file has no data rows",
    "請先載入 AIO 匯出檔": "Load an AIO export file first",
    "未知指令": "Unknown command",
}


def tr(s):
    """T() 的網頁版：i18n.py 查不到時再查 WEB_EN。"""
    hit = T(s)
    if hit == s and is_en():
        return WEB_EN.get(s, s)
    return hit


def ui():
    """目前語言下所有「會變」的介面字串（鍵＝內部中文字串）；沒變的不回傳，畫面端用原字串。"""
    keys = set(i18n.EN) | set(i18n.ZH) | set(WEB_EN)
    out = {}
    for k in keys:
        v = tr(k)
        if v != k:
            out[k] = v
    return out


def set_lang(lang):
    """切換語言。分析文字依語言產生，已載入資料時重跑分析（不需重新讀檔）。"""
    i18n.set_lang(lang)
    STATE["charts"] = {}
    if STATE["data"] is not None:
        STATE["rep"] = E.run_all(STATE["data"])
    return {"lang": i18n.get_lang(), "ui": ui(), "loaded": STATE["rep"] is not None}


def glossary():
    out, seen = [], set()
    for key in i18n.GLOSSARY:
        name = T(key)
        if name in seen:
            continue
        seen.add(name)
        out.append([name, define(key)])
    return out


# ----------------------------------------------------------------------------
# 格式化（與桌面版 col_kind／fmt_cell 相同）
# ----------------------------------------------------------------------------
def col_kind(name, s, pct=False):
    """依欄位內容決定顯示格式：pct／int／num／bool／text。"""
    if s.dtype == bool:
        return "bool"
    if not pd.api.types.is_numeric_dtype(s):
        return "text"
    if pct or PCT_RE.search(str(name)):
        return "pct"
    v = s.dropna()
    if len(v) and (v == v.round()).all():
        return "int"
    return "num"


def fmt_cell(v, kind="text"):
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if kind == "bool" or isinstance(v, (bool, np.bool_)):
        return L("是", "Yes") if v else L("否", "No")
    try:
        if kind == "pct":
            return f"{float(v):.1%}"
        if kind == "int":
            return f"{int(round(float(v))):,}"
        if kind == "num":
            return f"{float(v):,.1f}"
    except (TypeError, ValueError):
        pass
    return str(v)


def _flat(df):
    """reset_index 後的表；MultiIndex 欄名壓成一層（Excel 匯出不支援多層欄名＋無索引）。"""
    if isinstance(df.columns, pd.MultiIndex):
        df = df.copy()
        df.columns = [" / ".join(str(x) for x in c if str(x)) for c in df.columns]
    return df


def df_json(df, index=True, pct=False, highlight=None, max_rows=2000):
    """DataFrame → 畫面用 JSON。
    columns：顯示用欄名（已翻譯）  defs：各欄名詞定義（沒有就空字串）
    rows：已格式化成文字的儲存格  numeric：是否靠右對齊  warn：highlight(列) 為 True 的列（紅字）
    pct：整張表的資料欄都以百分比顯示（索引欄除外）。"""
    if df is None or len(df) == 0:
        return {"columns": [], "defs": [], "rows": [], "numeric": [], "warn": [], "total": 0}
    d = df.reset_index() if index else df.reset_index(drop=True)
    d = _flat(d).rename(columns={"index": "項目"})
    raw_cols = [str(c) for c in d.columns]
    n_idx = d.shape[1] - df.shape[1]
    kinds = [col_kind(c, d.iloc[:, i], pct and i >= n_idx) for i, c in enumerate(raw_cols)]
    shown = tr_df(d)
    seen, ids = {}, []
    for c in [str(c) for c in shown.columns]:
        seen[c] = seen.get(c, 0) + 1
        ids.append(c if seen[c] == 1 else f"{c} ({seen[c]})")
    defs = [define(r) for r in raw_cols]
    rows, warn = [], []
    for r, rec in enumerate(shown.itertuples(index=False, name=None)):
        if r >= max_rows:
            break
        rows.append([fmt_cell(v, kinds[i]) for i, v in enumerate(rec)])
        w = False
        if highlight is not None:
            try:
                w = bool(highlight(d.iloc[r]))
            except Exception:
                w = False
        warn.append(w)
    numeric = [k in ("pct", "int", "num") for k in kinds]
    return {"columns": ids, "defs": defs, "rows": rows, "numeric": numeric, "warn": warn, "total": int(len(shown))}


def _xlsx_raw(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False)
    return base64.b64encode(buf.getvalue()).decode()


def xlsx_b64(df, index=True):
    """表格匯出（欄名、索引與文字已翻譯，與畫面一致）。"""
    if df is None or len(df) == 0:
        return None
    d = df.reset_index() if index else df.reset_index(drop=True)
    d = _flat(d).rename(columns={"index": "項目"})
    return _xlsx_raw(tr_df(d))


def kv(n, share=None, digits=0):
    """KPI 數值：件數（占比）。"""
    n = int(n) if n == n else 0
    if share is None or share != share:
        return f"{n:,}"
    return L(f"{n:,}（{share:.{digits}%}）", f"{n:,} ({share:.{digits}%})")


def kpi(key, value):
    """KPI 方塊：[顯示名稱, 數值, 名詞定義]。"""
    return [tr(key), value, define(key)]


def _f(v, default=0.0):
    try:
        v = float(v)
        return default if v != v else v
    except (TypeError, ValueError):
        return default


def _rep():
    if STATE["rep"] is None:
        raise RuntimeError(tr("請先載入 AIO 匯出檔"))
    return STATE["rep"]


# ----------------------------------------------------------------------------
# 載入
# ----------------------------------------------------------------------------
def load(rows, start=None, progress=print):
    """rows：SheetJS sheet_to_json(header=1, cellDates=true) 的結果（list of list）。回傳 summary() 的 JSON 字串。"""
    rows = list(rows)
    progress(L(f"整理 {len(rows):,} 列資料…", f"Preparing {len(rows):,} rows…"))
    raw = pd.DataFrame(rows)
    df = E.clean_raw(raw)
    progress(L(f"  共 {len(df):,} 列，欄位 {df.shape[1]} 個", f"  {len(df):,} rows, {df.shape[1]} columns"))
    data = E.prepare(df, start or None)
    ps, rd = data.attrs["period_start"], data.attrs["ref_date"]
    progress(L(f"  分析期間 {ps:%Y-%m-%d} 起，資料日期 {rd:%Y-%m-%d}，{len(data):,} 件",
               f"  Period from {ps:%Y-%m-%d}, data as of {rd:%Y-%m-%d}, {len(data):,} cases"))
    rep = E.run_all(data, progress=lambda i, n, t: progress(f"  [{i}/{n}] {t}"))
    progress(L(f"  零件明細 {len(rep.parts):,} 顆", f"  {len(rep.parts):,} part lines"))
    STATE["rep"], STATE["data"], STATE["charts"] = rep, data, {}
    if STATE["chart_dir"] is None:
        STATE["chart_dir"] = tempfile.mkdtemp()
    return json.dumps(summary(), ensure_ascii=False)


def summary():
    rep = _rep()
    d = rep.data
    k12, k22, k41, k42, k52 = (rep.get(k).kpis for k in ("1-2", "2-2", "4-1", "4-2", "5-2"))
    sc = d[d["在途"]]["狀態分類"].value_counts().rename("件數").to_frame()
    sc.index.name = "狀態分類"
    ps, rd = d.attrs["period_start"], d.attrs["ref_date"]
    return {
        "status": L(f"{ps:%m/%d}–{rd:%m/%d}｜{len(d):,} 件", f"{ps:%b %d}–{rd:%b %d} | {len(d):,} cases"),
        "period": L(f"本次分析：{ps:%Y-%m-%d} ～ {rd:%Y-%m-%d}", f"Analyzing {ps:%b %d, %Y} – {rd:%b %d, %Y}"),
        "caption": L(f"分析期間 {ps:%Y-%m-%d} 起，資料日期 {rd:%Y-%m-%d}；{len(d):,} 件、{d['Location'].nunique()} 個據點、"
                     f"{d['國家'].nunique()} 個國家；零件明細 {len(rep.parts):,} 顆",
                     f"Period from {ps:%Y-%m-%d}, data as of {rd:%Y-%m-%d}; {len(d):,} cases, {d['Location'].nunique()} sites, "
                     f"{d['國家'].nunique()} countries; {len(rep.parts):,} part lines"),
        "kpis": [kpi("案件數", f"{len(d):,}"),
                 kpi("在途", kv(k12["在途件數"], k12["在途占比"])),
                 kpi("缺料（Awaiting Spares）", kv(k42["缺料件數"], k42["缺料占在途"])),
                 kpi("零件未使用率", f"{_f(k41['未使用率']):.1%}"),
                 kpi("SLA 達成率", f"{_f(k22['整體達成率']):.0%}"),
                 kpi("保固 30 天內到期在途", f"{int(k52['30天內到期件數']):,}")],
        "status_title": tr("各狀態在途件數"),
        "status_table": df_json(sc),
        "status_chart": {"labels": [T(str(i)) for i in sc.index], "values": [int(v) for v in sc["件數"]]},
        "analyses": [[a.key, E.title_part(a.title, -1)] for a in rep.analyses.values()],
    }


# ----------------------------------------------------------------------------
# 各分頁
# ----------------------------------------------------------------------------
def tab_demand(by="國家", weeks=4):
    rep = _rep()
    by = by if by in ("國家", "類型分組") else "國家"
    weeks = max(1, min(int(weeks), 13))
    wk = E.demand_weekly(rep.data, by)
    fc = E.demand_forecast(rep.data, weeks, by)
    tp = E.top_parts(rep.parts, 30)
    chart = None
    if "合計" in wk.columns and len(wk):
        chart = {"labels": [str(i) for i in wk.index], "values": [_f(v) for v in wk["合計"]]}
    return {"weekly": df_json(wk), "weekly_xlsx": xlsx_b64(wk), "weekly_name": T("週申請零件量.xlsx"), "weekly_chart": chart,
            "forecast_title": L(f"未來 {weeks} 週需求推估（最近 4 個完整週平均 × 週數；高峰＝最高週 × 週數）",
                                f"Next {weeks} weeks demand forecast (average of the last 4 full weeks × weeks; peak = highest week × weeks)"),
            "forecast": df_json(fc), "top": df_json(tp), "top_xlsx": xlsx_b64(tp), "top_name": T("Top料號.xlsx")}


def tab_stock(days=14):
    rep = _rep()
    k = rep.get("4-2").kpis
    days = max(1, min(int(days), 60))
    sp = E.shortage_parts(rep.parts)
    sc = E.shortage_cases(rep.data, days)
    med = _f(k.get("等待中位數"))
    return {"kpis": [kpi("缺料案件", kv(k["缺料件數"], k["缺料占在途"])),
                     kpi("等待超過兩週", kv(k["超過兩週件數"], k["超過兩週比例"])),
                     kpi("等待中位數", L(f"{med:.0f} 天", f"{med:.0f} days")),
                     kpi("零件在途／已配案件", f"{int(k['零件在途或已配件數']):,}"),
                     kpi("待供貨料號數", f"{len(sp):,}")],
            "parts": df_json(sp), "parts_xlsx": xlsx_b64(sp), "parts_name": T("缺料料號.xlsx"),
            "cases_title": tr("催料案件清單") + L(f"（等待 ≥ {days} 天）", f" (waiting ≥ {days} days)"),
            "cases": df_json(sc, index=False), "cases_xlsx": xlsx_b64(sc, index=False), "cases_name": T("催料清單.xlsx")}


def tab_ship(days=30):
    rep = _rep()
    days = max(1, min(int(days), 365))
    bl = E.ship_backlog(rep.data)
    tr_ = E.parts_in_transit(rep.parts)
    ex = E.expiring_cases(rep.data, days)
    k52 = rep.get("5-2").kpis
    transit = int(tr_["運送中"].sum()) if "運送中" in tr_.columns else 0
    alloc = int(tr_["已配貨"].sum()) if "已配貨" in tr_.columns else 0
    top = bl.head(10)
    over = top["超過7天"].astype(int).tolist()
    within = (top["待出貨件數"] - top["超過7天"]).astype(int).tolist()
    t7 = tr_.head(7)
    zero = [0] * len(t7)
    return {"kpis": [kpi("待出貨案件", f"{int(bl['待出貨件數'].sum()):,}"),
                     kpi("待出貨超過 7 天", f"{int(bl['超過7天'].sum()):,}"),
                     kpi("零件運送中（顆）", f"{transit:,}"),
                     kpi("零件已配貨（顆）", f"{alloc:,}"),
                     kpi("保固 30 天內到期在途", f"{int(k52['30天內到期件數']):,}")],
            "backlog_chart": {"title": L("修好待寄回 Top 10 維修站", "Ready to ship: top 10 sites"), "labels": [str(x) for x in top.index],
                              "series": [{"name": L("7 天內", "Within 7 days"), "values": within, "color": "#C9B79C"},
                                         {"name": L("超過 7 天", "Over 7 days"), "values": over, "color": "#E8862B"}]},
            "transit_chart": {"title": L("零件運送中 vs 已配到（前 7 國）", "Parts in transit vs allocated (top 7)"), "labels": [T(str(x)) for x in t7.index],
                              "series": [{"name": L("已配到", "Allocated"), "values": t7["已配貨"].astype(int).tolist() if "已配貨" in t7.columns else zero, "color": "#7FA650"},
                                         {"name": L("運送中", "In transit"), "values": t7["運送中"].astype(int).tolist() if "運送中" in t7.columns else zero, "color": "#E8862B"}]},
            "backlog_hint": tr("紅字＝放超過 7 天的案件有 10 件以上的維修站"),
            "backlog": df_json(bl, highlight=lambda r: r.get("超過7天", 0) >= 10), "backlog_xlsx": xlsx_b64(bl), "backlog_name": T("待出貨積壓.xlsx"),
            "transit": df_json(tr_),
            "expiring_title": L(f"在途且保固 {days} 天內到期（含建案時保固內、現已過期）",
                                f"Open cases with warranty ending within {days} days (incl. in warranty at creation, now expired)"),
            "expiring_hint": tr("紅字＝保固已經過期（建案時在保固內），結案時要按保固內處理"),
            "expiring": df_json(ex, index=False, highlight=lambda r: r.get("距到期天數", 0) < 0),
            "expiring_xlsx": xlsx_b64(ex, index=False), "expiring_name": T("保固到期在途案.xlsx")}


def _chart_b64(key):
    rep = _rep()
    if key in STATE["charts"]:
        return STATE["charts"][key]
    a = rep.get(key)
    p = os.path.join(STATE["chart_dir"], f"{key}.png")
    b64 = None
    if E.plot_analysis(a, p) and os.path.exists(p):
        with open(p, "rb") as fh:
            b64 = base64.b64encode(fh.read()).decode()
    STATE["charts"][key] = b64
    return b64


def _get(key):
    rep = _rep()
    if key not in rep.analyses:
        raise ValueError(L(f"沒有這項分析：{key}", f"No such analysis: {key}"))
    return rep.get(key)


def analysis(key):
    a = _get(key)
    return {"key": a.key, "title": a.title, "subtitle": a.subtitle, "bullets": list(a.bullets),
            "paragraph": a.paragraph, "conclusion": a.conclusion, "bullet": bullet(),
            "labels": {"points": L("重點", "Key points"), "about": L("說明", "What it shows"), "takeaway": T("結論：")},
            "chart": _chart_b64(key), "tables": [[n, T(n)] for n in a.tables]}


def analysis_table(key, name):
    a = _get(key)
    df = a.tables.get(name)
    # 整張表都是比例的（例如「占比」「組成」）全部以百分比顯示
    whole_pct = ("占比)" in name) or name.endswith("_占比") or ("組成" in name)
    return {"table": df_json(df, pct=whole_pct), "xlsx": xlsx_b64(df), "name": f"{key}_{T(name)}.xlsx"}


def conclusions():
    rep = _rep()
    return [[k, unit, f"[{k}] {rep.get(k).conclusion}"] for k, unit in SEEDS.items() if k in rep.analyses]


# ----------------------------------------------------------------------------
# 匯出
# ----------------------------------------------------------------------------
def export(kind, progress=print):
    """kind: xlsx / pptx / charts。回傳 {"data": base64, "name": 檔名}。"""
    rep = _rep()
    td = STATE["chart_dir"]
    if kind == "xlsx":
        name = T("RMA分析結果.xlsx")
        p = os.path.join(td, "out.xlsx")
        E.export_excel(rep, p)
    elif kind == "pptx":
        progress(L("畫全部圖表…", "Drawing all charts…"))
        for k in rep.analyses:
            _chart_b64(k)
        progress(L("組裝 PPT…", "Building the PowerPoint…"))
        name = T("RMA分析報告.pptx")
        p = os.path.join(td, "out.pptx")
        E.export_pptx(rep, p, td)
    elif kind == "charts":
        progress(L("畫全部圖表…", "Drawing all charts…"))
        name = "charts.zip"
        p = os.path.join(td, "charts.zip")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
            for k, a in rep.analyses.items():
                b64 = _chart_b64(k)
                if b64:
                    fn = f"{k}_{E.title_part(a.title, -1)}.png".replace("/", "-").replace(":", "")
                    z.writestr(fn, base64.b64decode(b64))
    else:
        raise ValueError(kind)
    with open(p, "rb") as fh:
        return {"data": base64.b64encode(fh.read()).decode(), "name": name}


# ----------------------------------------------------------------------------
# 訂單與交期（匯入訂單檔）
# ----------------------------------------------------------------------------
def po_template():
    tmpl = pd.DataFrame([
        ["PO-2026-0001", "KP.04501.017", T("供應商A"), 50, "2026-09-10", "2026-09-12", "2026-09-11", "Thailand"],
        ["PO-2026-0002", "KT.CTE00.014", T("供應商B"), 20, "2026-09-12", "2026-09-20", "", "Indonesia"],
    ], columns=[T(c) for c in PO_COLS])
    return {"data": _xlsx_raw(tmpl), "name": T("零件訂單範本.xlsx")}


def _po_alias():
    """中英文欄名都接受，內部統一成中文。"""
    alias = {}
    cur = i18n.get_lang()
    try:
        for lang in ("zh", "en"):
            i18n.set_lang(lang)
            for c in PO_COLS:
                alias[T(c)] = c
    finally:
        i18n.set_lang(cur)
    return alias


def po(rows):
    """rows：訂單檔 sheet_to_json(header=1) 的結果，第一列是標題。回傳 JSON 字串。"""
    rows = list(rows)
    if len(rows) < 2:
        raise ValueError(tr("訂單檔沒有資料"))
    po = pd.DataFrame(rows[1:], columns=[str(c).strip() if c is not None else "" for c in rows[0]])
    alias = _po_alias()
    po = po.rename(columns={c: alias.get(str(c).strip(), c) for c in po.columns})
    need = PO_COLS[:7]
    miss = [c for c in need if c not in po.columns]
    if miss:
        raise ValueError(T("缺少欄位：") + list_sep().join(T(c) for c in miss))
    po = po.dropna(subset=["訂單號"])
    STATE["po"] = po
    return json.dumps(po_result(), ensure_ascii=False)


def po_result():
    """用最近一次匯入的訂單檔計算（切換語言時重算顯示文字用）。"""
    if STATE["po"] is None:
        return None
    po = STATE["po"].copy()
    for c in ("下單日", "承諾交期", "實際到貨日"):
        po[c] = pd.to_datetime(po[c], errors="coerce")
    today = pd.Timestamp(dt.date.today())
    po["已到貨"] = po["實際到貨日"].notna()
    po["準交"] = po["已到貨"] & (po["實際到貨日"] <= po["承諾交期"])
    po["交期天數"] = (po["實際到貨日"] - po["下單日"]).dt.days
    po["逾期天數"] = ((po["實際到貨日"].fillna(today) - po["承諾交期"]).dt.days).clip(lower=0)
    arrived = po[po["已到貨"]]
    sup = po.groupby("供應商").agg(訂單筆數=("訂單號", "size"), 已到貨=("已到貨", "sum"), 準交筆數=("準交", "sum"),
                                 平均交期天數=("交期天數", "mean"), 平均逾期天數=("逾期天數", "mean"))
    sup["準交率"] = (sup["準交筆數"] / sup["已到貨"].replace(0, np.nan)).astype(float)
    late = po[(~po["已到貨"] & (po["承諾交期"] < today)) | (po["已到貨"] & ~po["準交"])].sort_values("逾期天數", ascending=False)
    cols = ["訂單號", "料號", "供應商", "數量", "承諾交期", "實際到貨日", "逾期天數"] + (["目的國"] if "目的國" in po.columns else [])
    return {"kpis": [kpi("訂單筆數", f"{len(po):,}"),
                     kpi("已到貨", f"{int(po['已到貨'].sum()):,}"),
                     kpi("準交率", f"{(_f(arrived['準交'].mean()) if len(arrived) else 0):.0%}"),
                     kpi("平均交期（天）", f"{(_f(arrived['交期天數'].mean()) if len(arrived) else 0):.1f}"),
                     kpi("未到貨且已逾期", f"{int((~po['已到貨'] & (po['承諾交期'] < today)).sum()):,}")],
            "suppliers": df_json(sup), "late": df_json(late[cols], index=False),
            "late_xlsx": xlsx_b64(late[cols], index=False), "late_name": T("逾期訂單.xlsx")}


# ----------------------------------------------------------------------------
# worker.js 的單一入口
# ----------------------------------------------------------------------------
def dispatch(cmd, args_json, progress=print):
    args = json.loads(args_json) if args_json else {}
    fns = {"summary": summary, "tab_demand": tab_demand, "tab_stock": tab_stock, "tab_ship": tab_ship,
           "analysis": analysis, "analysis_table": analysis_table, "conclusions": conclusions,
           "po_template": po_template, "po_result": po_result, "set_lang": set_lang, "ui": ui, "glossary": glossary}
    if cmd == "export":
        return json.dumps(export(args["kind"], progress), ensure_ascii=False)
    if cmd not in fns:
        raise ValueError(f"{tr('未知指令')} {cmd}")
    return json.dumps(fns[cmd](**args), ensure_ascii=False)
