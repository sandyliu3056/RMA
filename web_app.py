# -*- coding: utf-8 -*-
"""
售後零件規劃平台（網頁版）
------------------------
與桌面版 零件規劃平台.py 相同的七個模組，改用 Streamlit 跑在瀏覽器裡。
分析邏輯全部在 rma_engine.py，這裡只負責畫面。

本機執行：streamlit run web_app.py
雲端：推到 GitHub 後在 share.streamlit.io 選這個檔案部署（見 README）。
"""
import io
import os
import json
import zipfile
import hashlib
import tempfile
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import rma_engine as E

APP_TITLE = "售後零件規劃平台"
UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"]
NAVY, GOLD, ORANGE, CARD, LINE, MUTE, INK = "#3B2A1A", "#F2B134", "#E8862B", "#FFFDF7", "#D9CBB6", "#7A6652", "#3B2A1A"

st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")

# ----------------------------------------------------------------------------
# 外觀：咖啡金主題（與桌面版一致）
# ----------------------------------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
html, body, [class*="css"], .stMarkdown, .stDataFrame {{ font-family: "Microsoft JhengHei", "Noto Sans TC", sans-serif; }}
.block-container {{ padding-top: 0.8rem; padding-bottom: 2rem; max-width: 1400px; }}
.rma-header {{ background:{NAVY}; color:{GOLD}; border-radius:10px; padding:14px 22px; display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }}
.rma-header h1 {{ margin:0; font-size:1.6rem; color:{GOLD}; font-weight:700; }}
.rma-header .sub {{ color:#D9C7A8; font-size:0.85rem; margin-top:2px; }}
.rma-header .clock {{ text-align:right; }}
.rma-header .clock .t {{ font-family:Consolas, monospace; font-size:1.6rem; font-weight:700; color:{GOLD}; line-height:1; }}
.rma-header .clock .d {{ color:#D9C7A8; font-size:0.8rem; }}
.stTabs [data-baseweb="tab-list"] {{ background:#F5A623; border-radius:8px; padding:4px 6px; gap:4px; }}
.stTabs [data-baseweb="tab"] {{ background:{GOLD}; color:{INK}; border-radius:6px; padding:6px 16px; font-weight:700; }}
.stTabs [aria-selected="true"] {{ background:#FBF6EC !important; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}
.kpi {{ background:{CARD}; border:1px solid {LINE}; border-radius:10px; padding:10px 14px; min-height:78px; }}
.kpi .k {{ color:{MUTE}; font-size:0.8rem; }}
.kpi .v {{ color:{NAVY}; font-size:1.5rem; font-weight:700; line-height:1.3; }}
.section {{ border-left:4px solid {ORANGE}; padding-left:8px; font-weight:700; font-size:1.05rem; margin:14px 0 6px; }}
.stButton > button[kind="primary"] {{ background:#6B4E31; border-color:#6B4E31; }}
.stDownloadButton > button {{ background:{CARD}; border:1px solid {LINE}; color:{INK}; }}
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# 共用小工具
# ----------------------------------------------------------------------------
def header(status="尚未載入資料"):
    now = dt.datetime.now(ZoneInfo("Asia/Taipei"))
    st.markdown(f"""
<div class="rma-header">
  <div><h1>{APP_TITLE}</h1><div class="sub">{status}</div></div>
  <div class="clock"><div class="t">{now:%H:%M}</div><div class="d">{now:%Y/%m/%d %a} Taipei</div></div>
</div>""", unsafe_allow_html=True)


def kpis(items, cols=None):
    """一列 KPI 卡片。items = [(標題, 數值)…]"""
    cs = st.columns(cols or len(items))
    for c, (k, v) in zip(cs, items):
        c.markdown(f'<div class="kpi"><div class="k">{k}</div><div class="v">{v}</div></div>', unsafe_allow_html=True)


def section(text):
    st.markdown(f'<div class="section">{text}</div>', unsafe_allow_html=True)


def fmt_df(df):
    """百分比欄顯示成 %，日期去掉時間。"""
    if df is None or len(df) == 0:
        return df
    out = df.copy()
    for c in out.columns:
        s = out[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            out[c] = s.dt.strftime("%Y-%m-%d")
        elif pd.api.types.is_float_dtype(s) and s.dropna().between(-1, 1).all() and (s.dropna().abs() < 1).any() and any(k in str(c) for k in ("率", "比例", "占比", "%")):
            out[c] = s.map(lambda v: "" if pd.isna(v) else f"{v:.1%}")
    return out


def table(df, index=True, height=None):
    if df is None or len(df) == 0:
        st.info("（沒有資料）")
        return
    d = fmt_df(df.reset_index() if index else df)
    d = d.rename(columns={"index": "項目"})
    kw = {"height": height} if height else {}
    st.dataframe(d, hide_index=True, width="stretch", **kw)


def xlsx_bytes(df, index=True):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=index)
    return buf.getvalue()


def download(df, name, label="匯出 Excel", index=True, key=None):
    if df is not None and len(df):
        st.download_button(label, xlsx_bytes(df, index), file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)


@st.cache_resource(show_spinner=False)
def build_report(raw_bytes: bytes, start: str):
    """讀檔、整理、跑 11 項分析。用 cache_resource 保留 DataFrame.attrs（期間與資料日期）。"""
    raw = E.load_aio(io.BytesIO(raw_bytes))
    data = E.prepare(raw, start or None)
    rep = E.run_all(data)
    return rep


_PARTS_FUNCS = {"top_parts", "shortage_parts", "parts_in_transit"}


@st.cache_data(show_spinner=False)
def calc(digest: str, fn: str, *args):
    """呼叫引擎的某個表格函式並快取。Streamlit 每次互動都會重跑整個腳本，
    不快取的話七個分頁的計算全部重算，每個動作要等十幾秒。digest 讓快取跟著資料檔走。"""
    rep = st.session_state["report"]
    src = rep.parts if fn in _PARTS_FUNCS else rep.data
    return getattr(E, fn)(src, *args)


@st.cache_data(show_spinner=False)
def export_bundle(digest: str):
    """Excel、PPTX、全部圖 zip、每張圖的 PNG bytes（給報表頁顯示用）。"""
    rep = st.session_state["report"]
    with tempfile.TemporaryDirectory() as td:
        out = E.export_all(rep, td)
        with open(out["excel"], "rb") as fh:
            xlsx = fh.read()
        with open(out["pptx"], "rb") as fh:
            pptx = fh.read()
        zb = io.BytesIO()
        charts = {}
        with zipfile.ZipFile(zb, "w", zipfile.ZIP_DEFLATED) as z:
            for p in out["charts"]:
                z.write(p, os.path.basename(p))
                with open(p, "rb") as fh:
                    charts[os.path.basename(p).split("_")[0]] = fh.read()
        return xlsx, pptx, zb.getvalue(), charts


# ----------------------------------------------------------------------------
# 資料來源
# ----------------------------------------------------------------------------
rep = st.session_state.get("report")
digest = st.session_state.get("digest", "")
data = rep.data if rep else None
parts = rep.parts if rep else None
header(st.session_state.get("status", "尚未載入資料"))

tabs = st.tabs(["資料來源", "需求規劃", "庫存與缺料", "訂單與交期", "出貨與到貨", "報表與分析", "跨單位協調"])

with tabs[0]:
    c1, c2 = st.columns([3, 1])
    up = c1.file_uploader("AIO 匯出檔（.xlsx，工作表 AllInOneData）", type=["xlsx", "xlsm"])
    start = c2.text_input("分析起日（留空＝資料最後日期的前一個月 1 日）", placeholder="2026-08-01")
    if st.button("載入並分析", type="primary", disabled=up is None):
        raw_bytes = up.getvalue()
        digest = hashlib.md5(raw_bytes).hexdigest()[:12] + "-" + start.strip()
        with st.spinner("分析中，約 10–30 秒…"):
            try:
                rep = build_report(raw_bytes, start.strip())
                st.session_state["report"] = rep
                st.session_state["digest"] = digest
                st.session_state["file_name"] = up.name
                d = rep.data
                st.session_state["status"] = f"{up.name}｜{d.attrs['period_start']:%m/%d}–{d.attrs['ref_date']:%m/%d}｜{len(d):,} 件"
                calc.clear(); export_bundle.clear()
            except Exception as ex:  # noqa
                st.error(f"載入失敗：{ex}")
                st.stop()
        st.rerun()

    if rep is None:
        st.markdown("""
1. 選擇 AIO 匯出檔（.xlsx，工作表 AllInOneData）
2. 按「載入並分析」，約 10–30 秒
3. 到各分頁查看，或在「報表與分析」下載 Excel／PPT
""")
    else:
        d = data
        k12, k22, k41, k42, k52 = (rep.get(k).kpis for k in ("1-2", "2-2", "4-1", "4-2", "5-2"))
        kpis([("案件數", f"{len(d):,}"),
              ("在途", f"{k12['在途件數']:,}（{k12['在途占比']:.0%}）"),
              ("缺料（Awaiting Spares）", f"{k42['缺料件數']:,}（{k42['缺料占在途']:.0%}）"),
              ("零件未使用率", f"{k41['未使用率']:.1%}"),
              ("SLA 達成率", f"{k22['整體達成率']:.0%}"),
              ("保固 30 天內到期在途", f"{k52['30天內到期件數']:,}")])
        st.caption(f"分析期間 {d.attrs['period_start']:%Y-%m-%d} 起，資料日期 {d.attrs['ref_date']:%Y-%m-%d}；"
                   f"{len(d):,} 件、{d['Location'].nunique()} 個據點、{d['國家'].nunique()} 個國家；零件明細 {len(parts):,} 顆")
        section("各狀態在途件數")
        sc = d[d["在途"]]["狀態分類"].value_counts().rename("件數").to_frame()
        c1, c2 = st.columns([1, 2])
        with c1:
            table(sc)
        with c2:
            st.bar_chart(sc, color=ORANGE)

# ----------------------------------------------------------------------------
# 需求規劃
# ----------------------------------------------------------------------------
with tabs[1]:
    if rep is None:
        st.info("請先在「資料來源」載入 AIO 匯出檔")
    else:
        c1, c2, c3 = st.columns([1, 1, 4])
        by = c1.selectbox("彙總依", ["國家", "類型分組"])
        weeks = c2.number_input("預測未來（週）", 1, 13, 4)
        section("每週申請零件顆數（完整週）")
        wk = calc(digest, "demand_weekly", by)
        c1, c2 = st.columns([3, 2])
        with c1:
            table(wk)
        with c2:
            if "合計" in wk.columns:
                st.bar_chart(wk["合計"], color=ORANGE)
        download(wk, "週申請零件量.xlsx", key="dl_wk")
        section(f"未來 {weeks} 週需求推估（最近 4 個完整週平均 × 週數；高峰＝最高週 × 週數）")
        table(calc(digest, "demand_forecast", int(weeks), by))
        section("申請次數最多的料號（Top 30）：申請顆數、未使用率、目前待供貨顆數")
        tp = calc(digest, "top_parts", 30)
        table(tp, height=420)
        download(tp, "Top料號.xlsx", key="dl_tp")

# ----------------------------------------------------------------------------
# 庫存與缺料
# ----------------------------------------------------------------------------
with tabs[2]:
    if rep is None:
        st.info("請先在「資料來源」載入 AIO 匯出檔")
    else:
        k = rep.get("4-2").kpis
        sp = calc(digest, "shortage_parts")
        kpis([("缺料案件", f"{k['缺料件數']:,}（占在途 {k['缺料占在途']:.0%}）"),
              ("等待超過兩週", f"{k['超過兩週件數']:,}（{k['超過兩週比例']:.0%}）"),
              ("等待中位數（天）", f"{k['等待中位數']:.0f}"),
              ("零件在途／已配案件", f"{k['零件在途或已配件數']:,}"),
              ("待供貨料號數", f"{len(sp):,}")])
        section("缺料料號（Awaiting Spares 案件中待供貨的料號）")
        table(sp, height=320)
        download(sp, "缺料料號.xlsx", key="dl_sp")
        days = st.slider("催料清單：等待天數 ≥", 1, 60, 14)
        section(f"催料案件清單（等待 ≥ {days} 天）")
        sc = calc(digest, "shortage_cases", days)
        table(sc, index=False, height=380)
        download(sc, "催料清單.xlsx", index=False, key="dl_sc")

# ----------------------------------------------------------------------------
# 訂單與交期
# ----------------------------------------------------------------------------
with tabs[3]:
    st.caption("AIO 資料沒有採購訂單，這個模組用匯入的訂單檔計算準交率與逾期清單。")
    tmpl = pd.DataFrame([
        ["PO-2026-0001", "KP.04501.017", "供應商A", 50, "2026-09-10", "2026-09-12", "2026-09-11", "Thailand"],
        ["PO-2026-0002", "KT.CTE00.014", "供應商B", 20, "2026-09-12", "2026-09-20", "", "Indonesia"],
    ], columns=["訂單號", "料號", "供應商", "數量", "下單日", "承諾交期", "實際到貨日", "目的國"])
    c1, c2 = st.columns([1, 3])
    c1.download_button("下載訂單範本", xlsx_bytes(tmpl, index=False), "零件訂單範本.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_tmpl")
    pof = c2.file_uploader("匯入訂單檔（Excel／CSV）", type=["xlsx", "xls", "csv"], key="po_up")
    if pof is not None:
        try:
            po = pd.read_csv(pof) if pof.name.lower().endswith(".csv") else pd.read_excel(pof)
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
            kpis([("訂單筆數", f"{len(po):,}"), ("已到貨", f"{int(po['已到貨'].sum()):,}"),
                  ("準交率", f"{(arrived['準交'].mean() if len(arrived) else 0):.0%}"),
                  ("平均交期（天）", f"{(arrived['交期天數'].mean() if len(arrived) else 0):.1f}"),
                  ("未到貨且已逾期", f"{int((~po['已到貨'] & (po['承諾交期'] < today)).sum()):,}")])
            section("供應商準交率")
            table(sup)
            section("逾期與未交訂單")
            cols = ["訂單號", "料號", "供應商", "數量", "承諾交期", "實際到貨日", "逾期天數"] + (["目的國"] if "目的國" in po.columns else [])
            table(late[cols], index=False)
            download(late[cols], "逾期訂單.xlsx", index=False, key="dl_late")
        except Exception as ex:  # noqa
            st.error(f"匯入失敗：{ex}")

# ----------------------------------------------------------------------------
# 出貨與到貨
# ----------------------------------------------------------------------------
with tabs[4]:
    if rep is None:
        st.info("請先在「資料來源」載入 AIO 匯出檔")
    else:
        bl = calc(digest, "ship_backlog")
        tr = calc(digest, "parts_in_transit")
        days = st.slider("保固到期清單：天數內", 7, 180, 30, step=7)
        ex = calc(digest, "expiring_cases", days)
        kpis([("待出貨案件", f"{int(bl['待出貨件數'].sum()):,}"),
              ("待出貨超過 7 天", f"{int(bl['超過7天'].sum()):,}"),
              ("零件運送中（顆）", f"{int(tr['運送中'].sum()) if '運送中' in tr.columns else 0:,}"),
              ("零件已配貨（顆）", f"{int(tr['已配貨'].sum()) if '已配貨' in tr.columns else 0:,}"),
              (f"保固 {days} 天內到期在途", f"{len(ex):,}")])
        c1, c2 = st.columns(2)
        with c1:
            section("待出貨積壓（依據點）")
            table(bl, height=330)
            download(bl, "待出貨積壓.xlsx", key="dl_bl")
        with c2:
            section("零件運送中／已配貨（依國家，在途案件）")
            table(tr, height=330)
        section("在途且保固即將到期（含建案時保固內、現已過期）")
        table(ex, index=False, height=380)
        download(ex, "保固到期在途案.xlsx", index=False, key="dl_ex")

# ----------------------------------------------------------------------------
# 報表與分析
# ----------------------------------------------------------------------------
with tabs[5]:
    if rep is None:
        st.info("請先在「資料來源」載入 AIO 匯出檔")
    else:
        left, right = st.columns([1, 3])
        with left:
            section("分析項目")
            keys = list(rep.analyses.keys())
            labels = {k: f"{k}  {rep.get(k).title.split('｜')[-1]}" for k in keys}
            key = st.radio("分析項目", keys, format_func=lambda k: labels[k], label_visibility="collapsed")
            section("匯出")
            with st.spinner("準備圖表與匯出檔（第一次約 20 秒）…"):
                xlsx, pptx, zipb, charts = export_bundle(digest)
            st.download_button("匯出 Excel 表格", xlsx, "RMA分析結果.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_x")
            st.download_button("匯出 PPT 報告", pptx, "RMA分析報告.pptx",
                               "application/vnd.openxmlformats-officedocument.presentationml.presentation", key="dl_p")
            st.download_button("匯出全部圖 PNG（zip）", zipb, "charts.zip", "application/zip", key="dl_z")
        with right:
            a = rep.get(key)
            st.markdown(f"### {a.key} {a.title}")
            st.caption(a.subtitle)
            t1, t2 = st.columns([1, 1.6])
            with t1:
                for b in a.bullets:
                    st.markdown(f"- {b}")
                st.markdown(a.paragraph)
                st.markdown(f"**結論：** {a.conclusion}")
            with t2:
                png = charts.get(key)
                if png:
                    st.image(png, width="stretch")
                else:
                    st.info("（這一項沒有圖）")
            if a.tables:
                tname = st.selectbox("表格", list(a.tables.keys()), key=f"tbl_{key}")
                table(a.tables[tname], height=360)
                download(a.tables[tname], f"{a.key}_{tname}.xlsx", "匯出此表", key=f"dl_t_{key}")

# ----------------------------------------------------------------------------
# 跨單位協調
# ----------------------------------------------------------------------------
with tabs[6]:
    st.caption("待辦存在這個瀏覽器分頁的工作階段裡；關掉就消失。要保留請「下載待辦」，下次用「還原待辦」上傳回來。")
    acts = st.session_state.setdefault("actions", [])

    with st.form("add_action", clear_on_submit=True):
        c = st.columns([1.2, 3, 1, 1, 1])
        unit = c[0].selectbox("對象單位", UNITS)
        topic = c[1].text_input("議題")
        owner = c[2].text_input("負責人")
        due = c[3].date_input("到期日", dt.date.today() + dt.timedelta(days=7))
        status = c[4].selectbox("狀態", ["進行中", "待回覆", "已完成", "取消"])
        if st.form_submit_button("新增", type="primary"):
            if topic.strip():
                nid = (max(a["id"] for a in acts) + 1) if acts else 1
                acts.append({"id": nid, "對象單位": unit, "議題": topic.strip(), "負責人": owner, "到期日": due.isoformat(),
                             "狀態": status, "建立日": dt.date.today().isoformat()})
                st.rerun()
            else:
                st.warning("請填議題")

    c1, c2, c3 = st.columns([1.5, 1, 1.5])
    if c1.button("由分析結論產生待辦", disabled=rep is None):
        due = (dt.date.today() + dt.timedelta(days=14)).isoformat()
        seeds = {"1-2": "維修據點", "2-1": "維修據點", "2-2": "維修據點", "3-2": "總部服務團隊", "4-1": "各國規劃人員",
                 "4-2": "供應商", "5-1": "總部服務團隊", "5-2": "維修據點"}
        n = 0
        for k, unit in seeds.items():
            a = rep.get(k)
            t = f"[{a.key}] {a.conclusion}"
            if not any(x["議題"] == t for x in acts):
                nid = (max(x["id"] for x in acts) + 1) if acts else 1
                acts.append({"id": nid, "對象單位": unit, "議題": t, "負責人": "", "到期日": due, "狀態": "進行中",
                             "建立日": dt.date.today().isoformat()}); n += 1
        st.toast(f"已新增 {n} 筆待辦")
        st.rerun()
    if acts:
        c2.download_button("下載待辦（JSON）", json.dumps(acts, ensure_ascii=False, indent=2).encode("utf-8"),
                           "協調待辦.json", "application/json", key="dl_json")
    restore = c3.file_uploader("還原待辦（JSON）", type=["json"], key="restore_json", label_visibility="collapsed")
    if restore is not None and st.session_state.get("restored_name") != restore.name:
        try:
            st.session_state["actions"] = json.load(restore)
            st.session_state["restored_name"] = restore.name
            st.rerun()
        except Exception as ex:  # noqa
            st.error(f"讀取失敗：{ex}")

    if acts:
        df = pd.DataFrame(acts).sort_values(["狀態", "到期日"], key=lambda s: s.map({"已完成": 1}).fillna(0) if s.name == "狀態" else s).reset_index(drop=True)
        today = dt.date.today().isoformat()
        edited = st.data_editor(
            df, hide_index=True, width="stretch", num_rows="dynamic", key="act_editor",
            column_config={
                "id": st.column_config.NumberColumn("id", disabled=True, width="small"),
                "對象單位": st.column_config.SelectboxColumn("對象單位", options=UNITS),
                "議題": st.column_config.TextColumn("議題", width="large"),
                "到期日": st.column_config.TextColumn("到期日（YYYY-MM-DD）"),
                "狀態": st.column_config.SelectboxColumn("狀態", options=["進行中", "待回覆", "已完成", "取消"]),
                "建立日": st.column_config.TextColumn("建立日", disabled=True),
            })
        late = edited[(edited["狀態"].isin(["進行中", "待回覆"])) & (edited["到期日"].astype(str) < today)]
        if len(late):
            st.warning(f"逾期未完成 {len(late)} 筆：" + "、".join(late["議題"].astype(str).str[:30]))
        if st.button("儲存修改"):
            st.session_state["actions"] = edited.fillna("").to_dict("records")
            st.toast("已更新")
            st.rerun()
        download(pd.DataFrame(st.session_state["actions"]), "跨單位協調待辦.xlsx", "匯出待辦 Excel", index=False, key="dl_act")
    else:
        st.info("目前沒有待辦。")
