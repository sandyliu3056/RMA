# -*- coding: utf-8 -*-
"""
瀏覽器版（Pyodide）的 Python 端。
worker.js 把 SheetJS 讀出來的列資料丟進來，這裡呼叫 rma_engine 做分析，
回傳可直接畫在畫面上的 JSON（表格、KPI、圖的 base64）。
所有運算都在使用者自己的瀏覽器裡完成，資料不會上傳到任何伺服器。
"""
import io
import os
import json
import base64
import zipfile
import tempfile
import datetime as dt

import numpy as np
import pandas as pd

import rma_engine as E

STATE = {"rep": None, "charts": {}, "chart_dir": None}
UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"]
SEEDS = {"1-2": "維修據點", "2-1": "維修據點", "2-2": "維修據點", "3-2": "總部服務團隊", "4-1": "各國規劃人員",
         "4-2": "供應商", "5-1": "總部服務團隊", "5-2": "維修據點"}


# ----------------------------------------------------------------------------
# 格式化
# ----------------------------------------------------------------------------
def _cell(v, pct=False):
    try:
        if v is None or pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (bool, np.bool_)):
        return "是" if v else "否"
    if isinstance(v, (float, np.floating)):
        if pct:
            return f"{v:.1%}"
        return f"{v:,.1f}" if abs(v - round(v)) > 1e-9 else f"{int(round(v)):,}"
    if isinstance(v, (int, np.integer)):
        return f"{int(v):,}"
    return str(v)


def df_json(df, index=True):
    """DataFrame → {"columns": [...], "rows": [[...]], "numeric": [bool]}，數字已轉成顯示用文字。"""
    if df is None or len(df) == 0:
        return {"columns": [], "rows": [], "numeric": []}
    d = df.reset_index() if index else df.copy()
    d = d.rename(columns={"index": "項目"})
    cols = [str(c) for c in d.columns]
    pct_cols = set()
    for c in d.columns:
        s = d[c]
        if pd.api.types.is_float_dtype(s) and len(s.dropna()) and s.dropna().between(-1, 1).all() \
                and any(k in str(c) for k in ("率", "比例", "占比", "%")):
            pct_cols.add(c)
    numeric = [bool(pd.api.types.is_numeric_dtype(d[c]) and not pd.api.types.is_bool_dtype(d[c])) for c in d.columns]
    rows = [[_cell(v, c in pct_cols) for c, v in zip(d.columns, rec)] for rec in d.itertuples(index=False, name=None)]
    return {"columns": cols, "rows": rows, "numeric": numeric}


def xlsx_b64(df, index=True):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=index)
    return base64.b64encode(buf.getvalue()).decode()


def _rep():
    if STATE["rep"] is None:
        raise RuntimeError("請先載入 AIO 匯出檔")
    return STATE["rep"]


# ----------------------------------------------------------------------------
# 載入
# ----------------------------------------------------------------------------
def load(rows, start=None, progress=print):
    """rows：SheetJS sheet_to_json(header=1, cellDates=true) 的結果（list of list）。"""
    progress(f"整理 {len(rows):,} 列資料…")
    raw = pd.DataFrame(list(rows))
    df = E.clean_raw(raw)
    progress(f"  共 {len(df):,} 列，欄位 {df.shape[1]} 個")
    data = E.prepare(df, start or None)
    progress(f"  分析期間 {data.attrs['period_start']:%Y-%m-%d} 起，資料日期 {data.attrs['ref_date']:%Y-%m-%d}，{len(data):,} 件")
    rep = E.run_all(data, progress=lambda i, n, t: progress(f"  [{i}/{n}] {t}"))
    progress(f"  零件明細 {len(rep.parts):,} 顆")
    STATE["rep"] = rep
    STATE["charts"] = {}
    if STATE["chart_dir"] is None:
        STATE["chart_dir"] = tempfile.mkdtemp()
    return summary()


def summary():
    rep = _rep()
    d = rep.data
    k12, k22, k41, k42, k52 = (rep.get(k).kpis for k in ("1-2", "2-2", "4-1", "4-2", "5-2"))
    sc = d[d["在途"]]["狀態分類"].value_counts().rename("件數").to_frame()
    return {
        "status": f"{d.attrs['period_start']:%m/%d}–{d.attrs['ref_date']:%m/%d}｜{len(d):,} 件",
        "caption": (f"分析期間 {d.attrs['period_start']:%Y-%m-%d} 起，資料日期 {d.attrs['ref_date']:%Y-%m-%d}；"
                    f"{len(d):,} 件、{d['Location'].nunique()} 個據點、{d['國家'].nunique()} 個國家；零件明細 {len(rep.parts):,} 顆"),
        "kpis": [["案件數", f"{len(d):,}"],
                 ["在途", f"{k12['在途件數']:,}（{k12['在途占比']:.0%}）"],
                 ["缺料（Awaiting Spares）", f"{k42['缺料件數']:,}（{k42['缺料占在途']:.0%}）"],
                 ["零件未使用率", f"{k41['未使用率']:.1%}"],
                 ["SLA 達成率", f"{k22['整體達成率']:.0%}"],
                 ["保固 30 天內到期在途", f"{k52['30天內到期件數']:,}"]],
        "status_table": df_json(sc),
        "status_chart": {"labels": [str(i) for i in sc.index], "values": [int(v) for v in sc["件數"]]},
        "analyses": [[a.key, a.title.split("｜")[-1]] for a in rep.analyses.values()],
    }


# ----------------------------------------------------------------------------
# 各分頁
# ----------------------------------------------------------------------------
def tab_demand(by="國家", weeks=4):
    rep = _rep()
    wk = E.demand_weekly(rep.data, by)
    fc = E.demand_forecast(rep.data, int(weeks), by)
    tp = E.top_parts(rep.parts, 30)
    return {"weekly": df_json(wk), "weekly_xlsx": xlsx_b64(wk),
            "weekly_chart": {"labels": [str(i) for i in wk.index], "values": [float(v) for v in wk["合計"]]} if "合計" in wk.columns else None,
            "forecast": df_json(fc), "top": df_json(tp), "top_xlsx": xlsx_b64(tp)}


def tab_stock(days=14):
    rep = _rep()
    k = rep.get("4-2").kpis
    sp = E.shortage_parts(rep.parts)
    sc = E.shortage_cases(rep.data, int(days))
    return {"kpis": [["缺料案件", f"{k['缺料件數']:,}（占在途 {k['缺料占在途']:.0%}）"],
                     ["等待超過兩週", f"{k['超過兩週件數']:,}（{k['超過兩週比例']:.0%}）"],
                     ["等待中位數（天）", f"{k['等待中位數']:.0f}"],
                     ["零件在途／已配案件", f"{k['零件在途或已配件數']:,}"],
                     ["待供貨料號數", f"{len(sp):,}"]],
            "parts": df_json(sp), "parts_xlsx": xlsx_b64(sp),
            "cases": df_json(sc, index=False), "cases_xlsx": xlsx_b64(sc, index=False)}


def tab_ship(days=30):
    rep = _rep()
    bl = E.ship_backlog(rep.data)
    tr = E.parts_in_transit(rep.parts)
    ex = E.expiring_cases(rep.data, int(days))
    return {"kpis": [["待出貨案件", f"{int(bl['待出貨件數'].sum()):,}"],
                     ["待出貨超過 7 天", f"{int(bl['超過7天'].sum()):,}"],
                     ["零件運送中（顆）", f"{int(tr['運送中'].sum()) if '運送中' in tr.columns else 0:,}"],
                     ["零件已配貨（顆）", f"{int(tr['已配貨'].sum()) if '已配貨' in tr.columns else 0:,}"],
                     [f"保固 {int(days)} 天內到期在途", f"{len(ex):,}"]],
            "backlog": df_json(bl), "backlog_xlsx": xlsx_b64(bl),
            "transit": df_json(tr),
            "expiring": df_json(ex, index=False), "expiring_xlsx": xlsx_b64(ex, index=False)}


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


def analysis(key):
    rep = _rep()
    a = rep.get(key)
    return {"key": a.key, "title": a.title, "subtitle": a.subtitle, "bullets": list(a.bullets),
            "paragraph": a.paragraph, "conclusion": a.conclusion,
            "chart": _chart_b64(key), "tables": list(a.tables.keys())}


def analysis_table(key, name):
    a = _rep().get(key)
    df = a.tables.get(name)
    return {"table": df_json(df), "xlsx": xlsx_b64(df) if df is not None and len(df) else None}


def conclusions():
    rep = _rep()
    return [[k, unit, f"[{k}] {rep.get(k).conclusion}"] for k, unit in SEEDS.items() if k in rep.analyses]


# ----------------------------------------------------------------------------
# 匯出
# ----------------------------------------------------------------------------
def export(kind, progress=print):
    """kind: xlsx / pptx / charts。回傳 base64。"""
    rep = _rep()
    td = STATE["chart_dir"]
    if kind == "xlsx":
        p = os.path.join(td, "RMA分析結果.xlsx")
        E.export_excel(rep, p)
    elif kind == "pptx":
        progress("畫全部圖表…")
        for k in rep.analyses:
            _chart_b64(k)
        progress("組裝 PPT…")
        p = os.path.join(td, "RMA分析報告.pptx")
        E.export_pptx(rep, p, td)
    elif kind == "charts":
        progress("畫全部圖表…")
        p = os.path.join(td, "charts.zip")
        with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
            for k, a in rep.analyses.items():
                b64 = _chart_b64(k)
                if b64:
                    z.writestr(f"{k}_{a.title.split('｜')[-1]}.png", base64.b64decode(b64))
    else:
        raise ValueError(kind)
    with open(p, "rb") as fh:
        return base64.b64encode(fh.read()).decode()


# ----------------------------------------------------------------------------
# 訂單與交期（匯入訂單檔）
# ----------------------------------------------------------------------------
PO_TEMPLATE = pd.DataFrame([
    ["PO-2026-0001", "KP.04501.017", "供應商A", 50, "2026-09-10", "2026-09-12", "2026-09-11", "Thailand"],
    ["PO-2026-0002", "KT.CTE00.014", "供應商B", 20, "2026-09-12", "2026-09-20", "", "Indonesia"],
], columns=["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日", "目的國"])


def po_template():
    return xlsx_b64(PO_TEMPLATE, index=False)


def po(rows):
    """rows：訂單檔 sheet_to_json(header=1) 的結果，第一列是標題。"""
    rows = list(rows)
    if len(rows) < 2:
        raise ValueError("訂單檔沒有資料")
    po = pd.DataFrame(rows[1:], columns=[str(c).strip() for c in rows[0]])
    need = ["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日"]
    miss = [c for c in need if c not in po.columns]
    if miss:
        raise ValueError("缺少欄位：" + "、".join(miss))
    po = po.dropna(subset=["訂單號"])
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
    sup["準交率"] = sup["準交筆數"] / sup["已到貨"].replace(0, np.nan)
    late = po[(~po["已到貨"] & (po["承諾交期"] < today)) | (po["已到貨"] & ~po["準交"])].sort_values("逾期天數", ascending=False)
    cols = ["訂單號", "料號", "供應商", "數量", "承諾交期", "實際到貨日", "逾期天數"] + (["目的國"] if "目的國" in po.columns else [])
    return {"kpis": [["訂單筆數", f"{len(po):,}"], ["已到貨", f"{int(po['已到貨'].sum()):,}"],
                     ["準交率", f"{(arrived['準交'].mean() if len(arrived) else 0):.0%}"],
                     ["平均交期（天）", f"{(arrived['交期天數'].mean() if len(arrived) else 0):.1f}"],
                     ["未到貨且已逾期", f"{int((~po['已到貨'] & (po['承諾交期'] < today)).sum()):,}"]],
            "suppliers": df_json(sup), "late": df_json(late[cols], index=False), "late_xlsx": xlsx_b64(late[cols], index=False)}


# ----------------------------------------------------------------------------
# worker.js 的單一入口
# ----------------------------------------------------------------------------
def dispatch(cmd, args_json, progress=print):
    args = json.loads(args_json) if args_json else {}
    fns = {"summary": summary, "tab_demand": tab_demand, "tab_stock": tab_stock, "tab_ship": tab_ship,
           "analysis": analysis, "analysis_table": analysis_table, "conclusions": conclusions,
           "po_template": po_template}
    if cmd == "export":
        return json.dumps(export(args["kind"], progress), ensure_ascii=False)
    if cmd not in fns:
        raise ValueError(f"未知指令 {cmd}")
    return json.dumps(fns[cmd](**args), ensure_ascii=False)
