# -*- coding: utf-8 -*-
"""
RMA 案件分析引擎
----------------
讀取 AIO 匯出檔（AllInOneData），做清理、輔助欄位、12 項分析，
並可匯出 Excel 表格、PNG 圖表與 PPTX 報告。

用法（指令列）：
    python rma_engine.py <AIO.xlsx> [輸出資料夾] [--start 2026-08-01] [--lang en]

用法（程式）：
    from rma_engine import load_aio, prepare, run_all, export_all
    raw = load_aio("0820-0921 AIO data.xlsx")
    data = prepare(raw)                 # 期間、輔助欄
    report = run_all(data)              # 12 項分析
    export_all(report, "輸出")          # Excel + PNG + PPTX
"""
import os
import re
import sys
import json
import datetime as dt
import textwrap
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from i18n import T, L, is_en, tr_df, bullet, list_sep




STATUS_MAP = {
    "Closed": "已結案", "Rejected": "已結案",
    "Ready To Ship": "待出貨", "Repair Complete": "待出貨",
    "Awaiting Defective Parts": "待收壞件",
    "Awaiting Spares": "待料", "Part In Transit": "待料", "Partially Part In Transit": "待料",
    "Allocated": "待料", "Allocated At Hub": "待料", "Partially Allocated At Hub": "待料",
    "Awaiting Customer Approval": "待客戶", "Awaiting CC Approval": "待客戶", "Customer Cancelled": "待客戶",
    "Open": "維修中", "Assigned": "維修中", "ReAssign": "維修中", "Under Repair": "維修中",
    "Under Testing": "維修中", "Received In Warehouse": "維修中", "Awaiting Onsite Visit": "維修中",
    "In Transit to Repair Center": "維修中", "Waiting": "維修中", "Suspended": "維修中",
    "Transfered to NS": "維修中",
}
BUCKET_ORDER = ["維修中", "待料", "待客戶", "待收壞件", "待出貨", "已結案"]
MAIN_TYPES = ["Hardware", "On-Site", "DOA", "Web Request", "Component Swap"]
SYMPTOM_MAP = {
    "45-CAN'T POWER ON SYSTEM": "開機", "44-SYSTEM CAN'T BOOT": "開機", "20-SYSTEM AUTO SHUTDOWN": "開機",
    "07-NO DISPLAY": "顯示", "10-ABNORMAL LINE": "顯示", "06-DISPLAY FLICKERING": "顯示",
    "01-ABNORMAL COLOR": "顯示", "08-LCD DAMAGED": "顯示", "09-DISPLAY UNIFORMITY": "顯示",
    "02-ADAPTER DEAD": "電源", "02-BATTERY CAN'T CHARGE": "電源",
}
PART_CODE = {"C": "已使用", "U": "未使用退回", "D": "零件DOA", "W": "WPB", "R": "已申請待供貨",
             "I": "運送中", "A": "已配貨", "H": "保留中", "J": "J", "O": "O", "N": "N"}
DATE_COLS = ["CloseDate", "CreateDate", "OnSiteResolvedDate", "ProposalApprovalDate", "ProposalCreateDate",
             "ReceiveDate", "RepairCompleteDate", "ShipDate", "PurchaseDate", "WarrantyExpiryDate"]
NUM_COLS = ["CurrentStatusDays", "HFPTAT", "HFPTATWorkingDays", "TRT", "TATInNetworkDays", "DaysSinceRcvd",
            "AIO_GlobalTATNetworkDays", "RepairCompleteTATInNetworkDays", "RepairCompleteTATWeekend",
            "NumberOfPartsInHStatus", "NumberOfPartsOnHold", "NumberOfRequestedParts", "NumberOfUsedParts",
            "NumberOfUnUsedParts", "NumberOfDOAParts", "NumberOfWPBParts", "MachineMonthsPurchaseDate"]
SLA_DAYS = 5
AGING_LIMIT = 14
MIN_GROUP = 300
PRESET_COUNTRIES = ["Thailand", "Indonesia", "Philippines", "Vietnam", "Malaysia", "Singapore"]





def load_aio(path, sheet=None):
    """讀 AIO 匯出檔，回傳乾淨的 DataFrame（尚未篩期間）。"""
    xls = pd.ExcelFile(path)
    if sheet is None:
        sheet = "AllInOneData" if "AllInOneData" in xls.sheet_names else xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)
    return clean_raw(raw)


def clean_raw(raw):
    """把「整張工作表、無標題」的 DataFrame 整理成乾淨資料：找標題列、轉型態、去空白。
    load_aio 與網頁版（瀏覽器端用 SheetJS 讀檔後交給這裡）共用。"""
    first = raw.iloc[:, 0].astype(str).str.strip()
    hits = raw.index[first == "CaseID"]
    if len(hits) == 0:
        raise ValueError(T("找不到標題列（第一欄應有 CaseID）"))
    h = hits[0]
    df = raw.iloc[h + 1:].copy()
    # 標題列：空白、None、NaN 的儲存格不是欄名（工作表範圍比標題寬時會有）；重複欄名只留第一個
    df.columns = ["" if (c is None or (isinstance(c, float) and np.isnan(c))) else str(c).strip() for c in raw.iloc[h]]
    df = df.loc[:, [c for c in df.columns if c and c != "nan"]]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df = df[df["CaseID"].notna()]
    df = df[df["CaseID"].astype(str).str.strip() != "案件編號"]
    for c in DATE_COLS:
        if c in df.columns:
            df[c] = to_datetime_mixed(df[c])
    for c in NUM_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in df.columns:
        if pd.api.types.is_object_dtype(df[c].dtype) or pd.api.types.is_string_dtype(df[c].dtype):
            df[c] = df[c].astype(str).str.strip().replace({"nan": np.nan, "None": np.nan, "": np.nan})
    df = df.dropna(subset=["CreateDate"]).reset_index(drop=True)
    return df


def to_datetime_mixed(s):
    """日期欄轉 datetime：先用快速路徑；Excel 日期與手打文字日期混在一起時，解析失敗的再逐格處理。"""
    out = pd.to_datetime(s, errors="coerce")
    if pd.api.types.is_datetime64_any_dtype(s):
        return out
    bad = out.isna() & s.notna() & (s.astype(str).str.strip() != "")
    if bad.any():
        out = out.copy()
        out[bad] = pd.to_datetime(s[bad], errors="coerce", format="mixed")
    return out


def default_period_start(ref_date):
    """預設分析起日：資料最後日期的前一個月 1 日（例：9/21 → 8/1）。"""
    first = ref_date.replace(day=1)
    prev = (first - pd.Timedelta(days=1)).replace(day=1)
    return prev.normalize()


def _week_monday(s):
    return (s - pd.to_timedelta(s.dt.weekday, unit="D")).dt.normalize()


def prepare(df, period_start=None, ref_date=None):
    """篩期間並加輔助欄位。回傳新的 DataFrame，屬性 attrs 記錄期間與資料日期。"""
    d = df.copy()
    ref_date = pd.Timestamp(ref_date).normalize() if ref_date is not None else d["CreateDate"].max().normalize()
    period_start = pd.Timestamp(period_start).normalize() if period_start is not None else default_period_start(ref_date)
    d["同序號案件數"] = d.groupby("SerialNumber")["CaseID"].transform("size")
    d = d[d["CreateDate"] >= period_start].copy()

    d["國家"] = d["Country"].astype(str).str.replace(r"\s*\[.*\]\s*$", "", regex=True).replace({"nan": "未填"})
    d["狀態分類"] = d["Status"].map(STATUS_MAP).fillna("其他")
    d["已結案"] = d["Status"].isin(["Closed", "Rejected"])
    d["在途"] = ~d["已結案"]
    d["建案週"] = _week_monday(d["CreateDate"])
    last_full_week_end = ref_date - pd.Timedelta(days=ref_date.weekday() + 1)
    cal_full = (d["建案週"] >= period_start) & (d["建案週"] + pd.Timedelta(days=6) <= last_full_week_end)

    days_per_week = d[cal_full].groupby("建案週")["CreateDate"].apply(lambda s: s.dt.normalize().nunique())
    vol_per_week = d[cal_full].groupby("建案週").size()
    ok_weeks = vol_per_week[(days_per_week >= 6) & (vol_per_week >= 0.7 * vol_per_week[days_per_week >= 6].median())].index
    d["完整週"] = cal_full & d["建案週"].isin(ok_weeks)
    d["類型分組"] = np.where(d["Type"].isin(MAIN_TYPES), d["Type"], "Other")
    d["重複維修"] = (d["同序號案件數"] >= 2).astype(int)
    tat = d["TATInNetworkDays"]
    d["SLA達成"] = np.where(d["Status"].eq("Closed"), (tat <= SLA_DAYS).astype(float), np.nan)
    d["TAT區間"] = pd.cut(tat, [-1, 1, 3, 5, 7, 9, 11, 1_000_000],
                        labels=["00-01", "02-03", "04-05", "06-07", "08-09", "10-11", "12+"]).astype(str).replace({"nan": None})
    d["症狀類別"] = d["RCFailureCodes"].map(SYMPTOM_MAP)
    d.loc[d["RCFailureCodes"].isna(), "症狀類別"] = "未填"
    d["症狀類別"] = d["症狀類別"].fillna("其他")
    d["維修天數"] = (d["RepairCompleteDate"] - d["ReceiveDate"]).dt.total_seconds() / 86400
    d["出貨等待"] = (d["ShipDate"] - d["RepairCompleteDate"]).dt.total_seconds() / 86400
    d["結案等待"] = (d["CloseDate"] - d["ShipDate"]).dt.total_seconds() / 86400
    d["收件至結案"] = (d["CloseDate"] - d["ReceiveDate"]).dt.total_seconds() / 86400
    exp = d["WarrantyExpiryDate"]
    d["保固到期區間"] = np.select(
        [exp < ref_date, exp <= ref_date + pd.Timedelta(days=30), exp <= ref_date + pd.Timedelta(days=90)],
        ["1 已過期", "2 30天內", "3 31-90天"], default="4 90天以上")
    d["等待天數區間"] = pd.cut(d["CurrentStatusDays"], [-1, 6, 13, 20, 27, 1_000_000],
                          labels=["0-6", "7-13", "14-20", "21-27", "28+"]).astype(str)
    d["缺料"] = d["Status"].eq("Awaiting Spares")
    d["零件在途或已配"] = d["Status"].isin(["Part In Transit", "Partially Part In Transit", "Allocated", "Allocated At Hub", "Partially Allocated At Hub"])

    d.attrs["period_start"] = period_start
    d.attrs["ref_date"] = ref_date
    return d.reset_index(drop=True)


def main_countries(d, k=6):
    """案件量前 k 名國家（依期間資料）。"""
    vc = d["國家"].value_counts()
    return [c for c in vc.index[:k]]





_PART_RE = re.compile(r"^\s*(\d+)\s*-\s*(.+?)\s*-\s*([A-Z])\s*-\s*(\d+)\s*$")


def parts_table(d):
    """把 PartsRequested 拆成一列一顆零件：CaseID、序、料號、狀態碼、狀態、國家、Type、Location、Status、天數。"""
    rows = []
    cols = ["CaseID", "國家", "Type", "Location", "Status", "狀態分類", "CurrentStatusDays", "DaysSinceRcvd",
            "建案週", "完整週", "PartsRequested", "已結案"]
    sub = d.loc[d["PartsRequested"].notna(), cols]
    for rec in sub.itertuples(index=False):
        for item in str(rec.PartsRequested).split(", "):
            m = _PART_RE.match(item)
            if not m:
                continue
            code = m.group(3)
            rows.append((rec.CaseID, int(m.group(1)), m.group(2), code, PART_CODE.get(code, code),
                         rec.國家, rec.Type, rec.Location, rec.Status, rec.狀態分類, rec.CurrentStatusDays,
                         rec.DaysSinceRcvd, rec.建案週, rec.完整週, rec.已結案))
    p = pd.DataFrame(rows, columns=["CaseID", "序", "料號", "狀態碼", "零件狀態", "國家", "Type", "Location", "Status",
                                    "狀態分類", "CurrentStatusDays", "DaysSinceRcvd", "建案週", "完整週", "已結案"])
    p["料號前綴"] = p["料號"].str.split(".").str[0].str[:2]
    return p





@dataclass
class Analysis:
    key: str
    title: str
    subtitle: str = ""
    tables: dict = field(default_factory=dict)
    kpis: dict = field(default_factory=dict)
    bullets: list = field(default_factory=list)
    paragraph: str = ""
    conclusion: str = ""
    chart: str = ""

    def main_table(self):
        return next(iter(self.tables.values())) if self.tables else pd.DataFrame()


@dataclass
class Report:
    period_start: pd.Timestamp
    ref_date: pd.Timestamp
    n_cases: int
    analyses: dict = field(default_factory=dict)
    data: pd.DataFrame = None
    parts: pd.DataFrame = None

    def get(self, key):
        return self.analyses[key]

    def kpi_frame(self):
        rows = []
        for a in self.analyses.values():
            for k, v in a.kpis.items():
                rows.append((a.key, a.title, k, v))
        return pd.DataFrame(rows, columns=["分析", "名稱", "指標", "數值"])


def title_part(title, i):
    """分析標題「主題｜項目」取其中一段（同時支援英文的 " | "）。"""
    parts = re.split(r"\s*[｜|]\s*", title)
    try:
        return parts[i]
    except IndexError:
        return title


def _pct(x, digits=1):
    return f"{x * 100:.{digits}f}%"


def _n(x):
    return f"{int(round(x)):,}"





def a11_volume(d):
    a = Analysis('1-1', T('營運量與負荷｜案件量趨勢'), T('每週案件量、Top 10 據點、國家與類型占比'))
    n = len(d)
    wk = d[d['完整週']]
    weeks = sorted(wk['建案週'].unique())
    by_type = pd.crosstab(wk['建案週'], wk['類型分組']).reindex(columns=MAIN_TYPES + ['Other'], fill_value=0)
    by_type['合計'] = by_type.sum(axis=1)
    by_type.index = [w.strftime('%m/%d') for w in by_type.index]
    top6 = main_countries(d)
    cty = wk['國家'].where(wk['國家'].isin(top6), 'Other')
    by_cty = pd.crosstab(wk['建案週'], cty).reindex(columns=top6 + ['Other'], fill_value=0)
    by_cty.index = [w.strftime('%m/%d') for w in by_cty.index]
    loc = d['Location'].value_counts()
    top10 = loc.head(10).rename('案件數').to_frame()
    top10['占比'] = top10['案件數'] / n
    top10['累計占比'] = top10['占比'].cumsum()
    types = d['Type'].value_counts().rename('案件數').to_frame()
    types['占比'] = types['案件數'] / n
    ctys = d['國家'].value_counts().rename('案件數').to_frame()
    ctys['占比'] = ctys['案件數'] / n

    a.tables = {'週案件量_依類型': by_type, '週案件量_依國家': by_cty, 'Top10據點': top10,
                '類型占比': types, '國家占比': ctys}
    k = a.kpis
    k['期間案件數'] = n
    k['據點數'] = int(d['Location'].nunique())
    k['國家數'] = int(d['國家'].nunique())
    k['完整週數'] = len(weeks)
    k['完整週合計'] = int(by_type['合計'].sum()) if len(by_type) else 0
    k['Top10據點占比'] = float(top10['占比'].sum())
    k['前六國占比'] = float(ctys.head(6)['占比'].sum())
    top1, top2 = top10.index[0], top10.index[1]
    k['前兩據點占比'] = float(top10['占比'].iloc[:2].sum())
    hw = types['占比'].get('Hardware', 0)
    a.bullets = [
        f"母體：{_n(n)} 件、{k['國家數']} 個國家、{k['據點數']} 個據點；完整週 {len(weeks)} 週合計 {_n(k['完整週合計'])} 件"
        + (f"，每週 {_n(by_type['合計'].min())}–{_n(by_type['合計'].max())} 件" if len(by_type) else ''),
        f"據點集中：Top 10 據點占 {_pct(k['Top10據點占比'], 0)}，{top1} 與 {top2} 合計 {_pct(k['前兩據點占比'], 0)}",
        f"國家集中：前六國占 {_pct(k['前六國占比'], 0)}，前三名 {'、'.join(ctys.index[:3])} 合計 {_pct(ctys.head(3)['占比'].sum(), 0)}",
        f"類型：{types.index[0]} {_pct(types['占比'].iloc[0], 0)}、{types.index[1]} {_pct(types['占比'].iloc[1], 0)}、"
        f"{types.index[2]} {_pct(types['占比'].iloc[2], 0)}，其餘 {len(types) - 3} 種合計 {_pct(types['占比'].iloc[3:].sum(), 0)}",
    ]
    a.paragraph = (
        f"期間共 {_n(n)} 件案子，分布在 {k['國家數']} 個國家、{k['據點數']} 個據點。以完整週看每週約 "
        f"{_n(by_type['合計'].mean()) if len(by_type) else '－'} 件，週與週之間起伏不大。量高度集中：Top 10 據點占 "
        f"{_pct(k['Top10據點占比'], 0)}，光 {top1} 和 {top2} 兩個據點就占 {_pct(k['前兩據點占比'], 0)}；國家看前六國占 "
        f"{_pct(k['前六國占比'], 0)}。類型以 {types.index[0]} 為主（{_pct(hw, 0)}），其次 {types.index[1]}、{types.index[2]}。"
    )
    a.conclusion = ("資源配置先看 Top 10 據點和前三國，這些地方占了絕大多數的量；週量穩定，人力與零件備料可以用每週約 "
                    f"{_n(by_type['合計'].mean()) if len(by_type) else '－'} 件當基準。")
    if is_en():
        avg = _n(by_type['合計'].mean()) if len(by_type) else '–'
        a.bullets = [
            f"Base: {_n(n)} cases across {k['國家數']} countries and {k['據點數']} sites; {len(weeks)} full weeks totaling {_n(k['完整週合計'])} cases"
            + (f", {_n(by_type['合計'].min())}–{_n(by_type['合計'].max())} per week" if len(by_type) else ''),
            f"Site concentration: the top 10 sites handle {_pct(k['Top10據點占比'], 0)}; {top1} and {top2} alone account for {_pct(k['前兩據點占比'], 0)}",
            f"Country concentration: the top 6 countries make up {_pct(k['前六國占比'], 0)}; the top 3 ({', '.join(T(c) for c in ctys.index[:3])}) total {_pct(ctys.head(3)['占比'].sum(), 0)}",
            f"Case type: {types.index[0]} {_pct(types['占比'].iloc[0], 0)}, {types.index[1]} {_pct(types['占比'].iloc[1], 0)}, "
            f"{types.index[2]} {_pct(types['占比'].iloc[2], 0)}; the other {len(types) - 3} types add up to {_pct(types['占比'].iloc[3:].sum(), 0)}",
        ]
        a.paragraph = (
            f"The period had {_n(n)} cases spread across {k['國家數']} countries and {k['據點數']} sites. Full weeks averaged about {avg} cases, "
            f"with little swing from week to week. Volume is highly concentrated: the top 10 sites handle {_pct(k['Top10據點占比'], 0)}, and "
            f"{top1} and {top2} alone account for {_pct(k['前兩據點占比'], 0)}. By country, the top 6 make up {_pct(k['前六國占比'], 0)}. "
            f"{types.index[0]} is the main case type ({_pct(hw, 0)}), followed by {types.index[1]} and {types.index[2]}."
        )
        a.conclusion = (f"Start resource planning with the top 10 sites and top 3 countries, which carry most of the volume. Weekly volume is steady, "
                        f"so about {avg} cases a week works as the baseline for staffing and parts stocking.")
    return a





def a12_in_progress(d):
    a = Analysis('1-2', T('營運量與負荷｜在途案件數與 aging'), T('在途案件的狀態分類、類型組成與停留天數'))
    o = d[d['在途']]
    n, no = len(d), len(o)
    order = [b for b in BUCKET_ORDER if b != '已結案'] + ['其他']
    by_bucket = o['狀態分類'].value_counts().reindex(order, fill_value=0).rename('案件數').to_frame()
    by_bucket = by_bucket[by_bucket['案件數'] > 0]
    by_bucket['占在途比例'] = by_bucket['案件數'] / no
    type_bucket = pd.crosstab(o['類型分組'], o['狀態分類']).reindex(index=MAIN_TYPES + ['Other'], fill_value=0)
    type_bucket = type_bucket[[c for c in order if c in type_bucket.columns]]
    type_bucket['合計'] = type_bucket.sum(axis=1)
    top10 = o['Location'].value_counts().head(10).index
    loc_bucket = pd.crosstab(o['Location'], o['狀態分類']).loc[top10]
    loc_bucket = loc_bucket[[c for c in order if c in loc_bucket.columns]]
    loc_bucket['合計'] = loc_bucket.sum(axis=1)
    ag = o.groupby('狀態分類')['DaysSinceRcvd'].agg(
        案件數='size', 中位數='median', P90=lambda s: s.quantile(0.9),
        超過14天=lambda s: int((s > AGING_LIMIT).sum()), 超過14天比例=lambda s: (s > AGING_LIMIT).mean()).reindex(order).dropna()
    aging_all = o['DaysSinceRcvd']
    a.tables = {'在途_依狀態分類': by_bucket, '在途_類型×狀態': type_bucket, '在途_Top10據點×狀態': loc_bucket, '在途_停留天數': ag}
    k = a.kpis
    k['在途件數'] = no
    k['在途占比'] = no / n
    big = by_bucket['案件數'].idxmax()
    k['最大狀態分類'] = big
    k['最大狀態分類件數'] = int(by_bucket.loc[big, '案件數'])
    k['Top10據點在途占比'] = float(loc_bucket['合計'].sum() / no)
    k['停留中位數'] = float(aging_all.median())
    k['停留P90'] = float(aging_all.quantile(0.9))
    k['超過14天件數'] = int((aging_all > AGING_LIMIT).sum())
    k['超過14天比例'] = float((aging_all > AGING_LIMIT).mean())
    worst = ag['超過14天比例'].idxmax() if len(ag) else '－'
    a.bullets = [
        f"在途 {_n(no)} 件，占全部 {_pct(k['在途占比'], 0)}；已結案 {_n(n - no)} 件",
        f"狀態：{'、'.join(f'{b} {_n(r.案件數)}（{_pct(r.占在途比例, 0)}）' for b, r in by_bucket.iterrows())}",
        f"據點：Top 10 據點的在途占 {_pct(k['Top10據點在途占比'], 0)}，{loc_bucket['合計'].idxmax()} 最多 {_n(loc_bucket['合計'].max())} 件",
        f"停留天數（收件至今）：中位數 {k['停留中位數']:.0f} 天、P90 {k['停留P90']:.0f} 天，超過 {AGING_LIMIT} 天 {_n(k['超過14天件數'])} 件（{_pct(k['超過14天比例'], 0)}）",
        f"停留最久的狀態是 {worst}：{_pct(ag.loc[worst, '超過14天比例'], 0)} 超過 {AGING_LIMIT} 天" if worst != '－' else '',
    ]
    a.bullets = [b for b in a.bullets if b]
    a.paragraph = (
        f"期間 {_n(n)} 件中還有 {_n(no)} 件沒結案，占 {_pct(k['在途占比'], 0)}。在途最大的一塊是{big} "
        f"{_n(k['最大狀態分類件數'])} 件（{_pct(by_bucket.loc[big, '占在途比例'], 0)}）。停留天數一半在 "
        f"{k['停留中位數']:.0f} 天內，但有 {_n(k['超過14天件數'])} 件超過 {AGING_LIMIT} 天；分狀態看，{worst}停留超過 "
        f"{AGING_LIMIT} 天的比例最高。"
    )
    a.conclusion = (f"在途要先處理{big}與停留超過 {AGING_LIMIT} 天的案子，這兩塊不是維修能力問題，是流程與追蹤問題；建議每週列出超過 "
                    f"{AGING_LIMIT} 天的清單給據點。")
    if is_en():
        a.bullets = [
            f"{_n(no)} cases are still open ({_pct(k['在途占比'], 0)} of all); {_n(n - no)} are closed",
            f"By status: {', '.join(f'{T(b)} {_n(r.案件數)} ({_pct(r.占在途比例, 0)})' for b, r in by_bucket.iterrows())}",
            f"By site: the top 10 sites hold {_pct(k['Top10據點在途占比'], 0)} of open cases; {loc_bucket['合計'].idxmax()} has the most at {_n(loc_bucket['合計'].max())}",
            f"Days since received: median {k['停留中位數']:.0f} days, 90% within {k['停留P90']:.0f} days; {_n(k['超過14天件數'])} cases ({_pct(k['超過14天比例'], 0)}) are over {AGING_LIMIT} days",
            f"The status group that sits longest is {T(worst)}: {_pct(ag.loc[worst, '超過14天比例'], 0)} are over {AGING_LIMIT} days" if worst != '－' else '',
        ]
        a.bullets = [b for b in a.bullets if b]
        a.paragraph = (
            f"Of the {_n(n)} cases in the period, {_n(no)} are still open ({_pct(k['在途占比'], 0)}). The biggest open group is {T(big)} with "
            f"{_n(k['最大狀態分類件數'])} cases ({_pct(by_bucket.loc[big, '占在途比例'], 0)}). Half of open cases were received {k['停留中位數']:.0f} "
            f"days ago or less, but {_n(k['超過14天件數'])} have been open more than {AGING_LIMIT} days. By status group, {T(worst)} has the highest "
            f"share over {AGING_LIMIT} days."
        )
        a.conclusion = (f"Tackle {T(big)} and anything open more than {AGING_LIMIT} days first. Neither is a repair-capacity problem; both are process "
                        f"and follow-up problems. Send each site a weekly list of its cases open more than {AGING_LIMIT} days.")
    return a





def a13_internal(d):
    a = Analysis('1-3', T('營運量與負荷｜內外部客戶占比'), T('內部與外部客戶的案件類型組成'))
    flag = d['IsInternalCustomers'].astype(str).str.upper().map({'YES': '內部客戶', 'NO': '外部客戶'}).fillna('未填')
    ct = pd.crosstab(flag, d['類型分組']).reindex(columns=MAIN_TYPES + ['Other'], fill_value=0)
    ct['合計'] = ct.sum(axis=1)
    pct = ct.drop(columns='合計').div(ct['合計'], axis=0)
    a.tables = {'內外部×類型_件數': ct, '內外部×類型_占比': pct}
    k = a.kpis
    ni = int(ct['合計'].get('內部客戶', 0))
    k['內部客戶件數'] = ni
    k['內部客戶占比'] = ni / len(d)
    ext = pct.loc['外部客戶'] if '外部客戶' in pct.index else pct.iloc[0]
    itn = pct.loc['內部客戶'] if '內部客戶' in pct.index else None
    ext_s = '、'.join(f"{t} {_pct(v, 0)}" for t, v in ext.sort_values(ascending=False).head(4).items())
    a.bullets = [f"量：內部客戶 {_n(ni)} 件，占 {_pct(k['內部客戶占比'], 0)}，服務量上無需區分",
                 f"外部組成：{ext_s}，以送修為主的一般維修結構"]
    if itn is not None:
        itn_s = '、'.join(f"{t} {_pct(v, 0)}" for t, v in itn.sort_values(ascending=False).head(3).items() if v > 0)
        ratio = itn.get('DOA', 0) / ext.get('DOA', 1) if ext.get('DOA', 0) else 0
        a.bullets.append(f"內部組成：{itn_s}，DOA 比例為外部的 {ratio:.0f} 倍" if ratio else f"內部組成：{itn_s}")
        a.bullets.append('意義：內部通報集中在新機開箱不良，是品質面的早期訊號')
        a.paragraph = (f"內部客戶只有 {_n(ni)} 件、{_pct(k['內部客戶占比'], 0)}，對服務量沒有影響，但組成完全不同：外部客戶是"
                       f"{ext_s}；內部客戶 {itn_s}，沒有到府案件。內部案子多半是新機開箱檢查發現的不良，比外部客戶更早反映品質問題。")

    a.conclusion = T('內部客戶的案件量可以併入整體看，但 DOA 內容值得單獨追蹤，當作新機品質的早期警訊。')
    if is_en():
        ext_s = ', '.join(f"{t} {_pct(v, 0)}" for t, v in ext.sort_values(ascending=False).head(4).items())
        a.bullets = [f"Volume: {_n(ni)} internal-customer cases ({_pct(k['內部客戶占比'], 0)}), not enough to plan for separately",
                     f"External mix: {ext_s}, a typical carry-in repair profile"]
        a.paragraph = ""
        if itn is not None:
            itn_s = ', '.join(f"{t} {_pct(v, 0)}" for t, v in itn.sort_values(ascending=False).head(3).items() if v > 0)
            a.bullets.append(f"Internal mix: {itn_s}; the DOA share is {ratio:.0f}x the external rate" if ratio else f"Internal mix: {itn_s}")
            a.bullets.append(T('意義：內部通報集中在新機開箱不良，是品質面的早期訊號'))
            a.paragraph = (f"Internal customers account for only {_n(ni)} cases ({_pct(k['內部客戶占比'], 0)}), so they don't move the overall "
                           f"workload, but their mix is completely different. External customers are {ext_s}; internal customers are {itn_s}, with "
                           f"no on-site cases. Most internal cases are defects found during out-of-box checks on new units, so they flag quality "
                           f"issues earlier than external customers do.")
    return a





def a21_tat(d):
    a = Analysis('2-1', T('時效與 SLA｜TAT 分布'), T('已結案 TATInNetworkDays（收件至維修完成，工作天）與各段耗時'))
    c = d[d['Status'].eq('Closed')]
    top6 = main_countries(d)
    order = ['00-01', '02-03', '04-05', '06-07', '08-09', '10-11', '12+']
    dist = pd.crosstab(c['TAT區間'], c['國家']).reindex(index=order, columns=top6, fill_value=0)
    dist_pct = dist.div(dist.sum(axis=0), axis=1)
    stat = c.groupby('國家')['TATInNetworkDays'].agg(
        案件數='size', 平均='mean', 中位數='median', P90=lambda s: s.quantile(0.9),
        一天內比例=lambda s: (s <= 1).mean(), 超過5天比例=lambda s: (s > SLA_DAYS).mean()).loc[top6]
    stat_t = c[c['Type'].isin(MAIN_TYPES)].groupby('Type')['TATInNetworkDays'].agg(
        案件數='size', 平均='mean', 中位數='median', P90=lambda s: s.quantile(0.9),
        一天內比例=lambda s: (s <= 1).mean(), 超過5天比例=lambda s: (s > SLA_DAYS).mean()).reindex(MAIN_TYPES).dropna()
    seg = c[c['Type'].isin(MAIN_TYPES)].groupby('Type')[['維修天數', '出貨等待', '結案等待', '收件至結案']].mean().reindex(MAIN_TYPES)
    seg.loc['全部'] = c[['維修天數', '出貨等待', '結案等待', '收件至結案']].mean()
    seg.columns = ['收件→維修完成', '維修完成→出貨', '出貨→結案', '收件→結案']
    a.tables = {'TAT分布_依國家(占比)': dist_pct, 'TAT分布_依國家(件數)': dist, 'TAT統計_依國家': stat,
                'TAT統計_依類型': stat_t, '各段耗時_依類型(日曆天)': seg}
    t = c['TATInNetworkDays']
    k = a.kpis
    k['已結案件數'] = len(c)
    k['TAT平均'] = float(t.mean()); k['TAT中位數'] = float(t.median()); k['TAT_P90'] = float(t.quantile(0.9))
    k['一天內比例'] = float((t <= 1).mean()); k['超過7天比例'] = float((t > 7).mean())
    k['收件至結案平均'] = float(seg.loc['全部', '收件→結案'])
    k['出貨至結案平均'] = float(seg.loc['全部', '出貨→結案'])
    fast = stat['一天內比例'].idxmax(); slow = stat['P90'].idxmax()
    peak = {cty: dist_pct[cty].idxmax() for cty in top6}
    delayed = [cty for cty, p in peak.items() if p != '00-01']
    a.bullets = [
        f"母體與定義：已結案 {_n(len(c))} 件，TAT 採 TATInNetworkDays（收件至維修完成，工作天）；X 軸每兩天一格，Y 軸為各國占比",
        f"整體：平均 {k['TAT平均']:.1f} 個工作天、中位數 {k['TAT中位數']:.0f} 天、P90 {k['TAT_P90']:.0f} 天；"
        f"{_pct(k['一天內比例'], 0)} 在 1 天內完成，超過 7 天僅 {_pct(k['超過7天比例'], 0)}",
        f"當日完成型：{fast} 0–1 天占 {_pct(stat.loc[fast, '一天內比例'], 0)}；長尾型：{slow} P90 {stat.loc[slow, 'P90']:.0f} 天、超過 5 天 "
        f"{_pct(stat.loc[slow, '超過5天比例'], 0)}",
        f"固定延遲型：{'、'.join(delayed)} 高峰不在第一格而在 {peak[delayed[0]]} 天，多數案件並非拖延，較像流程固定多出一至兩天"
        if delayed else '各國高峰都在 0–1 天',
        f"各段耗時：收件→維修完成 {seg.loc['全部', '收件→維修完成']:.1f} 天、維修完成→出貨 {seg.loc['全部', '維修完成→出貨']:.1f} 天、出貨→結案 "
        f"{seg.loc['全部', '出貨→結案']:.1f} 天，客戶感受的收件→結案平均 {k['收件至結案平均']:.1f} 個日曆天",
    ]
    slow_seg = seg.drop(index='全部')['出貨→結案'].idxmax()
    a.paragraph = (
        f"維修速度本身不慢：{_n(len(c))} 件已結案平均 {k['TAT平均']:.1f} 個工作天，一半在 {k['TAT中位數']:.0f} 天內修好，90% 在 "
        f"{k['TAT_P90']:.0f} 天內。各國曲線型態不同：{fast} 大多當天或隔天完成；{slow} 尾巴最長，P90 到 "
        f"{stat.loc[slow, 'P90']:.0f} 天；{'、'.join(delayed) + ' 的高峰落在 ' + peak[delayed[0]] + ' 天，是流程固定多出一兩天。' if delayed else ''}但客戶感受的是收件到結案，平均 "
        f"{k['收件至結案平均']:.1f} 個日曆天，其中維修只占 {seg.loc['全部', '收件→維修完成']:.1f} 天，出貨後到結案還要 "
        f"{k['出貨至結案平均']:.1f} 天；{slow_seg} 這段最長（{seg.loc[slow_seg, '出貨→結案']:.1f} 天）。"
    )
    a.conclusion = ("TAT 整體穩定，差異來自各國流程結構而非效率。要縮短客戶感受的時間，先看出貨後到結案這 "
                    f"{k['出貨至結案平均']:.1f} 天，特別是 {slow_seg}；{slow} 把超過 {SLA_DAYS} 個工作天的案子挑出來找共同特徵。")
    if is_en():
        a.bullets = [
            f"Base and definition: {_n(len(c))} closed cases; TAT is TATInNetworkDays (received to repair complete, business days). "
            f"Each X-axis bucket is two days wide; the Y-axis is each country's share",
            f"Overall: mean {k['TAT平均']:.1f} business days, median {k['TAT中位數']:.0f}, 90% within {k['TAT_P90']:.0f}; "
            f"{_pct(k['一天內比例'], 0)} finish within 1 day and only {_pct(k['超過7天比例'], 0)} take more than 7",
            f"Same-day finishers: {T(fast)} completes {_pct(stat.loc[fast, '一天內比例'], 0)} in 0–1 days. Long tail: {T(slow)} has 90% within "
            f"{stat.loc[slow, 'P90']:.0f} days, with {_pct(stat.loc[slow, '超過5天比例'], 0)} over 5 days",
            f"Built-in delay: {', '.join(T(x) for x in delayed)} peak at {peak[delayed[0]]} days rather than the first bucket. Most cases aren't "
            f"stalling; the process just seems to add a day or two" if delayed else T('各國高峰都在 0–1 天'),
            f"Time by stage: received → repaired {seg.loc['全部', '收件→維修完成']:.1f} days, repaired → shipped {seg.loc['全部', '維修完成→出貨']:.1f}, "
            f"shipped → closed {seg.loc['全部', '出貨→結案']:.1f}. What customers actually feel, received → closed, averages "
            f"{k['收件至結案平均']:.1f} calendar days",
        ]
        late = (f"{', '.join(T(x) for x in delayed)} peak at {peak[delayed[0]]} days, which points to a fixed extra day or two in the process. "
                if delayed else "")
        a.paragraph = (
            f"Repair speed itself isn't the problem: {_n(len(c))} closed cases averaged {k['TAT平均']:.1f} business days, half were fixed within "
            f"{k['TAT中位數']:.0f} days, and 90% within {k['TAT_P90']:.0f}. Countries have different curve shapes: {T(fast)} finishes most cases "
            f"the same or next day; {T(slow)} has the longest tail, with 90% taking up to {stat.loc[slow, 'P90']:.0f} days; {late}"
            f"But customers experience received-to-closed, which averages {k['收件至結案平均']:.1f} calendar days. Repair is only "
            f"{seg.loc['全部', '收件→維修完成']:.1f} of those days, and shipped-to-closed takes another {k['出貨至結案平均']:.1f}; "
            f"{slow_seg} has the longest stretch ({seg.loc[slow_seg, '出貨→結案']:.1f} days)."
        )
        a.conclusion = (f"TAT is stable overall; the differences come from how each country's process is set up, not from efficiency. To shorten "
                        f"what customers feel, start with the {k['出貨至結案平均']:.1f} days between shipping and closing, especially for {slow_seg}. "
                        f"In {T(slow)}, pull the cases over {SLA_DAYS} business days and look for what they have in common.")
    return a





def a22_sla(d):
    a = Analysis('2-2', T('時效與 SLA｜SLA 達成率'), L(f'已結案（Closed）TAT ≤ {SLA_DAYS} 個工作天的比例', f'Share of closed cases with TAT ≤ {SLA_DAYS} business days'))
    c = d[d['Status'].eq('Closed')]
    top6 = main_countries(d)
    by_c = c.groupby('國家')['TATInNetworkDays'].agg(
        案件數='size', 達成率=lambda s: (s <= SLA_DAYS).mean(), 未達成件數=lambda s: int((s > SLA_DAYS).sum()),
        三天內達成率=lambda s: (s <= 3).mean()).loc[top6].sort_values('達成率', ascending=False)
    by_t = c[c['Type'].isin(MAIN_TYPES)].groupby('Type')['TATInNetworkDays'].agg(
        案件數='size', 達成率=lambda s: (s <= SLA_DAYS).mean(), 未達成件數=lambda s: int((s > SLA_DAYS).sum())).reindex(MAIN_TYPES).dropna()
    top10 = c['Location'].value_counts().head(10).index
    by_l = c[c['Location'].isin(top10)].groupby('Location')['TATInNetworkDays'].agg(
        案件數='size', 達成率=lambda s: (s <= SLA_DAYS).mean(), 未達成件數=lambda s: int((s > SLA_DAYS).sum())).sort_values('達成率')
    a.tables = {'SLA_依國家': by_c, 'SLA_依類型': by_t, 'SLA_Top10據點': by_l}
    k = a.kpis
    k['整體達成率'] = float((c['TATInNetworkDays'] <= SLA_DAYS).mean())
    k['未達成件數'] = int((c['TATInNetworkDays'] > SLA_DAYS).sum())
    k['三天內達成率'] = float((c['TATInNetworkDays'] <= 3).mean())
    best, worst = by_c.index[0], by_c.index[-1]
    a.bullets = [
        f"定義：SLA 門檻 {SLA_DAYS} 個工作天（TATInNetworkDays），母體為已結案 {_n(len(c))} 件",
        f"整體達成 {_pct(k['整體達成率'], 0)}，未達成 {_n(k['未達成件數'])} 件；若門檻改為 3 天則為 {_pct(k['三天內達成率'], 0)}",
        f"國家：{'、'.join(f'{i} {_pct(r.達成率, 0)}（{_n(r.未達成件數)} 件未達）' for i, r in by_c.iterrows())}",
        f"類型：{'、'.join(f'{i} {_pct(r.達成率, 0)}' for i, r in by_t.iterrows())}",
    ]
    a.paragraph = (
        f"以 {SLA_DAYS} 個工作天當門檻，{_n(len(c))} 件已結案有 {_pct(k['整體達成率'], 0)} 達成，未達成 {_n(k['未達成件數'])} 件。"
        f"{best} 最高 {_pct(by_c.loc[best, '達成率'], 0)}，{worst} 最低 {_pct(by_c.loc[worst, '達成率'], 0)}、未達成 "
        f"{_n(by_c.loc[worst, '未達成件數'])} 件。未達成件數最多的是 {by_c['未達成件數'].idxmax()}（{_n(by_c['未達成件數'].max())} 件），因為量大，比例不算最差但件數最多。門檻若收緊到 3 天，整體會掉到 "
        f"{_pct(k['三天內達成率'], 0)}，各國差距會拉大。"
    )
    a.conclusion = (f"SLA 整體達標，改善優先順序看未達成件數而不是比例：{by_c['未達成件數'].idxmax()} 件數最多、"
                    f"{worst} 比例最低，先處理這兩國超過 {SLA_DAYS} 天的案子。")
    if is_en():
        most = by_c['未達成件數'].idxmax()
        a.bullets = [
            f"Definition: the SLA target is {SLA_DAYS} business days (TATInNetworkDays); base is {_n(len(c))} closed cases",
            f"Overall, {_pct(k['整體達成率'], 0)} met SLA and {_n(k['未達成件數'])} missed; at a 3-day target it would be {_pct(k['三天內達成率'], 0)}",
            f"By country: {', '.join(f'{T(i)} {_pct(r.達成率, 0)} ({_n(r.未達成件數)} missed)' for i, r in by_c.iterrows())}",
            f"By type: {', '.join(f'{i} {_pct(r.達成率, 0)}' for i, r in by_t.iterrows())}",
        ]
        a.paragraph = (
            f"With a {SLA_DAYS}-business-day target, {_pct(k['整體達成率'], 0)} of {_n(len(c))} closed cases met SLA and {_n(k['未達成件數'])} missed. "
            f"{T(best)} is highest at {_pct(by_c.loc[best, '達成率'], 0)}; {T(worst)} is lowest at {_pct(by_c.loc[worst, '達成率'], 0)}, with "
            f"{_n(by_c.loc[worst, '未達成件數'])} missed. {T(most)} has the most misses ({_n(by_c['未達成件數'].max())}): its rate isn't the worst, "
            f"but its volume is high. Tightening the target to 3 days would drop overall attainment to {_pct(k['三天內達成率'], 0)} and widen the "
            f"gaps between countries."
        )
        a.conclusion = (f"SLA is on target overall. Prioritize by number of misses, not by rate: {T(most)} has the most misses and {T(worst)} the "
                        f"lowest rate, so start with their cases over {SLA_DAYS} days.")
    return a





def a31_pareto(d):
    a = Analysis('3-1', T('品質與故障｜故障代碼 Pareto'), T('RCFailureCodes 前 20 名與累計占比、症狀歸類'))
    coded = d[d["RCFailureCodes"].notna()]
    vc = coded["RCFailureCodes"].value_counts()
    n = len(coded)
    top20 = vc.head(20).rename("案件數").to_frame()
    top20["占比"] = top20["案件數"] / n
    top20["累計占比"] = top20["占比"].cumsum()
    cum = (vc / n).cumsum()
    codes80 = int((cum < 0.8).sum() + 1)
    sym = coded["症狀類別"].value_counts().rename("案件數").to_frame()
    sym["占比"] = sym["案件數"] / n
    sym_detail = coded[coded["症狀類別"].isin(["開機", "顯示", "電源"])].groupby(["症狀類別", "RCFailureCodes"]).size().rename("案件數").to_frame()
    by_type = pd.crosstab(coded["Type"], coded["RCFailureCodes"]).reindex(index=MAIN_TYPES).fillna(0)
    top_by_type = pd.DataFrame({t: by_type.loc[t].sort_values(ascending=False).head(5).index.tolist() for t in MAIN_TYPES if t in by_type.index}).T
    top_by_type.columns = [f"第{i + 1}名" for i in range(top_by_type.shape[1])]
    a.tables = {"Top20故障代碼": top20, "症狀歸類": sym, "症狀歸類_明細": sym_detail, "各類型前五代碼": top_by_type}
    k = a.kpis
    k["有代碼件數"] = n; k["未填件數"] = int(len(d) - n); k["代碼種類"] = int(vc.size)
    k["前3名占比"] = float(top20["占比"].head(3).sum()); k["前10名占比"] = float(top20["占比"].head(10).sum())
    k["前20名占比"] = float(top20["占比"].sum()); k["達80%所需代碼數"] = codes80
    g = {s: int(sym.loc[s, "案件數"]) if s in sym.index else 0 for s in ("開機", "顯示", "電源")}
    k["開機類件數"], k["顯示類件數"], k["電源類件數"] = g["開機"], g["顯示"], g["電源"]
    three = sum(g.values()) / n
    a.bullets = [
        f"母體：有故障代碼 {_n(n)} 件（{_pct(n / len(d), 0)}），未填 {_n(k['未填件數'])} 件；共 {_n(vc.size)} 種代碼",
        f"集中度：前 3 名占 {_pct(k['前3名占比'], 0)}、前 10 名 {_pct(k['前10名占比'], 0)}、前 20 名 {_pct(k['前20名占比'], 0)}；{codes80} 個代碼就到 80%",
        f"前三名：{'、'.join(f'{c} {_n(v)}（{_pct(v / n, 0)}）' for c, v in vc.head(3).items())}",
        f"依症狀歸類（Top 20 代碼內）：開機類 {_n(g['開機'])} 件、顯示類 {_n(g['顯示'])} 件、電源類 {_n(g['電源'])} 件，三類合計 {_pct(three, 0)}",
    ]
    a.paragraph = (
        f"{_n(n)} 件有故障代碼的案子分散在 {_n(vc.size)} 種代碼上，但非常集中：前 20 個代碼就占 {_pct(k['前20名占比'], 0)}，"
        f"{codes80} 個代碼覆蓋 80%。第一名 {vc.index[0]} 一個代碼就占 {_pct(vc.iloc[0] / n, 0)}。把 Top 20 依症狀歸類，開機（無法開機、無法啟動、自動關機）"
        f"{_n(g['開機'])} 件、顯示（無畫面、異常線條、閃爍等）{_n(g['顯示'])} 件、電源（變壓器、電池充電）"
        f"{_n(g['電源'])} 件，三類合計 {_pct(three, 0)}。"
    )
    a.conclusion = T("品質改善與零件備料都可以聚焦：盯住前 20 個代碼就覆蓋六成案件，開機、顯示、電源三類是主軸；備料清單以這三類對應的零件為優先。")

    if is_en():
        a.bullets = [
            f"Base: {_n(n)} cases with a failure code ({_pct(n / len(d), 0)}), {_n(k['未填件數'])} without one; {_n(vc.size)} distinct codes",
            f"Concentration: the top 3 codes make up {_pct(k['前3名占比'], 0)}, top 10 {_pct(k['前10名占比'], 0)}, top 20 {_pct(k['前20名占比'], 0)}; "
            f"just {codes80} codes reach 80%",
            f"Top 3: {', '.join(f'{c} {_n(v)} ({_pct(v / n, 0)})' for c, v in vc.head(3).items())}",
            f"By symptom group (within the top 20 codes): boot {_n(g['開機'])} cases, display {_n(g['顯示'])}, power {_n(g['電源'])}; "
            f"together {_pct(three, 0)}",
        ]
        a.paragraph = (
            f"The {_n(n)} cases with a failure code are spread across {_n(vc.size)} codes, but they're highly concentrated: the top 20 codes account "
            f"for {_pct(k['前20名占比'], 0)}, and {codes80} codes cover 80%. The #1 code, {vc.index[0]}, makes up {_pct(vc.iloc[0] / n, 0)} on its "
            f"own. Grouping the top 20 by symptom: boot (won't power on, won't boot, shuts down on its own) {_n(g['開機'])} cases, display (no "
            f"picture, abnormal lines, flicker, etc.) {_n(g['顯示'])}, and power (adapter, battery not charging) {_n(g['電源'])}, together "
            f"{_pct(three, 0)}."
        )
    return a





def a32_repeat(d):
    a = Analysis('3-2', T('品質與故障｜重複維修'), T('同一序號在期間內出現 2 件以上'))
    n = len(d)
    sn = d["同序號案件數"]
    dist = pd.DataFrame({"案件數": d.groupby("同序號案件數").size(),
                         "機器數": d.groupby("同序號案件數")["SerialNumber"].nunique()})
    rep_cases = int((sn >= 2).sum())
    rep_serials = int(d.loc[sn >= 2, "SerialNumber"].nunique())
    top6 = main_countries(d)
    by_c = d.groupby("國家").agg(案件數=("重複維修", "size"), 重複案件=("重複維修", "sum"))
    by_c["重複率"] = by_c["重複案件"] / by_c["案件數"]
    by_c = by_c.loc[top6].sort_values("重複率", ascending=False)
    by_t = d[d["Type"].isin(MAIN_TYPES)].groupby("Type").agg(案件數=("重複維修", "size"), 重複案件=("重複維修", "sum"))
    by_t["重複率"] = by_t["重複案件"] / by_t["案件數"]
    by_t = by_t.sort_values("重複率", ascending=False)

    r = d[sn >= 2].sort_values(["SerialNumber", "CreateDate"]).copy()
    r["前案狀態"] = r.groupby("SerialNumber")["Status"].shift(1)
    r["前案結案日"] = r.groupby("SerialNumber")["CloseDate"].shift(1)
    r["前案代碼"] = r.groupby("SerialNumber")["RCFailureCodes"].shift(1)
    f = r[r["前案狀態"].notna()].copy()
    f["間隔天數"] = (f["CreateDate"] - f["前案結案日"]).dt.days
    gap = pd.cut(f["間隔天數"], [-1000000, 0, 7, 14, 30, 1000000], labels=["同日重開", "1-7天", "8-14天", "15-30天", ">30天"]).value_counts().sort_index().rename("後續案件數").to_frame()
    reopen = float(f["前案狀態"].isin(["Rejected", "Customer Cancelled"]).mean()) if len(f) else 0
    same_code = float((f["RCFailureCodes"] == f["前案代碼"]).mean()) if len(f) else 0
    a.tables = {"重複次數分布": dist, "重複率_依國家": by_c, "重複率_依類型": by_t, "後續案件與前案間隔": gap}
    k = a.kpis
    k["機器數"] = int(d["SerialNumber"].nunique()); k["重複機器數"] = rep_serials; k["重複案件數"] = rep_cases
    k["重複率"] = rep_cases / n; k["前案被拒或取消比例"] = reopen; k["同日重開比例"] = float((f["間隔天數"] <= 0).mean()) if len(f) else 0
    k["同故障代碼比例"] = same_code
    two = int(dist.loc[2, "機器數"]) if 2 in dist.index else 0
    three = int(dist.loc[3:, "機器數"].sum()) if len(dist) else 0
    hi, lo = by_c.index[0], by_c.index[-1]
    a.bullets = [
        f"母體：{_n(n)} 件、{_n(k['機器數'])} 台機器；其中 {_n(rep_serials)} 台重複進案，共 {_n(rep_cases)} 件，占 {_pct(k['重複率'])}",
        f"次數：{_n(two)} 台 2 次、{_n(three)} 台 3 次以上",
        f"國家：{hi} {_pct(by_c.loc[hi, '重複率'])} 最高，{lo} {_pct(by_c.loc[lo, '重複率'])} 最低，相差 {by_c.loc[hi, '重複率'] / max(by_c.loc[lo, '重複率'], 1e-9):.1f} 倍",
        f"類型：{by_t.index[0]} {_pct(by_t['重複率'].iloc[0])} 最高，{by_t.index[-1]} {_pct(by_t['重複率'].iloc[-1])} 最低",
        f"注意：後續案件有 {_pct(k['同日重開比例'], 0)} 是前案結案當天重開、{_pct(reopen, 0)} 的前案被 Rejected 或客戶取消，真正「修不好再回來」的比例低於 "
        f"{_pct(k['重複率'])}",
    ]
    a.paragraph = (
        f"{_n(n)} 件案子對應 {_n(k['機器數'])} 台機器，其中 {_n(rep_serials)} 台在期間內進來兩次以上，合計 {_n(rep_cases)} 件、占 "
        f"{_pct(k['重複率'])}；絕大多數是兩次（{_n(two)} 台）。依國家看差距很大，{hi} {_pct(by_c.loc[hi, '重複率'])}、"
        f"{lo} {_pct(by_c.loc[lo, '重複率'])}；依類型看 {by_t.index[0]} 最高。要留意的是這裡的「重複」是同一序號在一個多月內出現兩件，其中 "
        f"{_pct(k['同日重開比例'], 0)} 是前案結案當天重開，並不是修好又壞，所以真正的返修率會比 {_pct(k['重複率'])} 低。"
    )
    a.conclusion = f"重複進案整體 {_pct(k['重複率'])}，{hi} 與 {by_t.index[0]} 明顯偏高；建議先把 {_n(rep_serials)} 台的案由抓出來，區分「重開案」和「真返修」，並把觀察期拉長到 30／90 天再定正式返修率指標。"

    if is_en():
        a.bullets = [
            f"Base: {_n(n)} cases on {_n(k['機器數'])} machines; {_n(rep_serials)} machines came in more than once, totaling {_n(rep_cases)} cases "
            f"({_pct(k['重複率'])})",
            f"How often: {_n(two)} machines twice, {_n(three)} three or more times",
            f"By country: {T(hi)} is highest at {_pct(by_c.loc[hi, '重複率'])}, {T(lo)} lowest at {_pct(by_c.loc[lo, '重複率'])}, "
            f"a {by_c.loc[hi, '重複率'] / max(by_c.loc[lo, '重複率'], 1e-9):.1f}x gap",
            f"By type: {by_t.index[0]} is highest at {_pct(by_t['重複率'].iloc[0])}, {by_t.index[-1]} lowest at {_pct(by_t['重複率'].iloc[-1])}",
            f"Note: {_pct(k['同日重開比例'], 0)} of follow-up cases were opened the same day the prior case closed, and {_pct(reopen, 0)} followed a "
            f"prior case that was Rejected or cancelled by the customer, so the true \"fixed, then came back\" rate is below {_pct(k['重複率'])}",
        ]
        a.paragraph = (
            f"{_n(n)} cases map to {_n(k['機器數'])} machines; {_n(rep_serials)} of them came in two or more times during the period, totaling "
            f"{_n(rep_cases)} cases ({_pct(k['重複率'])}). Nearly all came in twice ({_n(two)} machines). The gap by country is wide: {T(hi)} "
            f"{_pct(by_c.loc[hi, '重複率'])} vs. {T(lo)} {_pct(by_c.loc[lo, '重複率'])}. By type, {by_t.index[0]} is highest. Keep in mind that "
            f"\"repeat\" here means the same serial number showing up twice within about a month, and {_pct(k['同日重開比例'], 0)} of those were "
            f"reopened the same day the prior case closed, not repaired and then broken again. The real return rate is lower than "
            f"{_pct(k['重複率'])}."
        )
        a.conclusion = (f"Repeat cases run {_pct(k['重複率'])} overall and are noticeably higher for {T(hi)} and {by_t.index[0]}. Pull the case "
                        f"history for the {_n(rep_serials)} machines, separate reopened cases from true repeat repairs, and extend the window to "
                        f"30/90 days before setting a formal repeat-repair KPI.")
    return a





def a41_parts(d):
    a = Analysis('4-1', T('零件與成本｜零件使用結果'), T('已結案（Closed）案件的零件去向：使用／未使用／DOA／WPB'))
    c = d[d["Status"].eq("Closed")]
    cols = ["NumberOfUsedParts", "NumberOfUnUsedParts", "NumberOfDOAParts", "NumberOfWPBParts"]
    names = ["已使用", "未使用", "零件DOA", "WPB"]
    t = c.groupby("類型分組")[cols].sum().reindex(MAIN_TYPES + ["Other"]).fillna(0)  # 全部類型都算，主要五類以外併為 Other
    t.columns = names
    t.loc["全部合計"] = t.sum()
    t["合計"] = t[["已使用", "未使用", "零件DOA"]].sum(axis=1)
    for nm in names[:3]:
        t[nm + "率"] = t[nm] / t["合計"]
    top6 = main_countries(d)
    cc = c.groupby("國家")[cols].sum().reindex(top6)
    cc.columns = names
    cc["合計"] = cc[["已使用", "未使用", "零件DOA"]].sum(axis=1)
    cc["未使用率"] = cc["未使用"] / cc["合計"]
    cc["零件DOA率"] = cc["零件DOA"] / cc["合計"]
    onsite = c[c["Type"].eq("On-Site")]
    oc = onsite.groupby("國家")[cols].sum()
    oc.columns = names
    oc["合計"] = oc[["已使用", "未使用", "零件DOA"]].sum(axis=1)
    oc["未使用率"] = oc["未使用"] / oc["合計"]
    oc = oc[oc["合計"] >= 100].sort_values("未使用率", ascending=False)
    ppc = c[c["NumberOfRequestedParts"] > 0].groupby("類型分組")["NumberOfRequestedParts"].mean().reindex(MAIN_TYPES + ["Other"]).rename("每案平均申請顆數").to_frame()
    a.tables = {"零件去向_依類型": t, "零件去向_依國家": cc, "On-Site未使用率_依國家": oc, "每案平均申請顆數": ppc}
    tot = t.loc["全部合計"]
    k = a.kpis
    k["已結案件數"] = len(c); k["有申請零件件數"] = int((c["NumberOfRequestedParts"] > 0).sum())
    k["零件總數"] = int(tot["合計"]); k["使用率"] = float(tot["已使用率"]); k["未使用率"] = float(tot["未使用率"]); k["零件DOA率"] = float(tot["零件DOA率"])
    worst_t = t.drop(index="全部合計")["未使用率"].idxmax()
    k["未使用率最高類型"] = worst_t; k["未使用率最高類型數值"] = float(t.loc[worst_t, "未使用率"])
    worst_c = cc["未使用率"].idxmax()
    onsite_share = float(onsite["國家"].eq(worst_c).mean()) if len(onsite) else 0
    a.bullets = [
        f"母體：已結案 {_n(len(c))} 件中 {_n(k['有申請零件件數'])} 件有申請零件，共 {_n(k['零件總數'])} 顆",
        f"整體：已使用 {_pct(k['使用率'])}（{_n(tot['已使用'])}）、未使用 {_pct(k['未使用率'])}（{_n(tot['未使用'])}）、零件 DOA "
        f"{_pct(k['零件DOA率'])}（{_n(tot['零件DOA'])}）、WPB {_n(tot['WPB'])} 顆",
        f"類型：{worst_t} 未使用 {_pct(k['未使用率最高類型數值'])}（{_n(t.loc[worst_t, '未使用'])}／{_n(t.loc[worst_t, '合計'])}），每案平均申請 "
        f"{ppc.loc[worst_t, '每案平均申請顆數']:.1f} 顆；其他類型都在 {_pct(t.drop(index=['全部合計', worst_t])['未使用率'].max(), 0)} 以下",
        f"國家：{worst_c} 未使用 {_pct(cc.loc[worst_c, '未使用率'])}（{_n(cc.loc[worst_c, '未使用'])} 顆）；{worst_c} 占 On-Site 結案的 {_pct(onsite_share, 0)}"
        + (f"，{worst_c} On-Site 的未使用率 {_pct(oc.loc[worst_c, '未使用率'])}" if worst_c in oc.index else ""),
        f"零件 DOA 各類型都在 {_pct(t.drop(index='全部合計')['零件DOA率'].max())} 以下，不是主要問題",
    ]
    a.paragraph = (
        f"已結案的案子共申請 {_n(k['零件總數'])} 顆零件，真正用掉 {_pct(k['使用率'])}，退回未使用 {_pct(k['未使用率'])}，零件本身是壞的只有 "
        f"{_pct(k['零件DOA率'])}。每領出去約 {1 / max(k['未使用率'], 1e-9):.0f} 顆就有一顆原封退回，倉庫出貨、運送、再收回、重新上架都是成本。未使用幾乎全部集中在 "
        f"{worst_t}：申請 {_n(t.loc[worst_t, '合計'])} 顆只用掉 "
        f"{_n(t.loc[worst_t, '已使用'])} 顆；到府維修出發前無法確定故障零件，一次帶多顆備選，修完再退，每案平均申請 "
        f"{ppc.loc[worst_t, '每案平均申請顆數']:.1f} 顆，Hardware 只有 {ppc.loc['Hardware', '每案平均申請顆數']:.2f} 顆。依國家看集中在 {worst_c}。"
    )
    a.conclusion = (f"零件退回率 {_pct(k['未使用率'], 0)} 的來源是 {worst_t} 的備選領料方式，尤其是 {worst_c}；建議檢討派工前的遠端診斷與領料規則，先從 "
                    f"{worst_c} 做起；零件 DOA 率 {_pct(k['零件DOA率'])} 屬正常水準不需處理。")
    if is_en():
        a.bullets = [
            f"Base: of {_n(len(c))} closed cases, {_n(k['有申請零件件數'])} requested parts, {_n(k['零件總數'])} units in total",
            f"Overall: used {_pct(k['使用率'])} ({_n(tot['已使用'])}), unused {_pct(k['未使用率'])} ({_n(tot['未使用'])}), part DOA "
            f"{_pct(k['零件DOA率'])} ({_n(tot['零件DOA'])}), WPB {_n(tot['WPB'])} units",
            f"By type: {worst_t} unused {_pct(k['未使用率最高類型數值'])} ({_n(t.loc[worst_t, '未使用'])}/{_n(t.loc[worst_t, '合計'])}), averaging "
            f"{ppc.loc[worst_t, '每案平均申請顆數']:.1f} parts requested per case; every other type is under "
            f"{_pct(t.drop(index=['全部合計', worst_t])['未使用率'].max(), 0)}",
            f"By country: {T(worst_c)} unused {_pct(cc.loc[worst_c, '未使用率'])} ({_n(cc.loc[worst_c, '未使用'])} units); {T(worst_c)} accounts for "
            f"{_pct(onsite_share, 0)} of closed On-Site cases"
            + (f", with an On-Site unused rate of {_pct(oc.loc[worst_c, '未使用率'])}" if worst_c in oc.index else ""),
            f"Part DOA stays under {_pct(t.drop(index='全部合計')['零件DOA率'].max())} for every type, so it isn't a major issue",
        ]
        a.paragraph = (
            f"Closed cases requested {_n(k['零件總數'])} parts. {_pct(k['使用率'])} were actually used, {_pct(k['未使用率'])} "
            f"came back unused, and only {_pct(k['零件DOA率'])} were defective themselves. Roughly 1 in every {1 / max(k['未使用率'], 1e-9):.0f} parts "
            f"sent out comes back untouched, and picking, shipping, returning, and restocking all cost money. Nearly all unused parts come from "
            f"{worst_t}: {_n(t.loc[worst_t, '合計'])} requested, only {_n(t.loc[worst_t, '已使用'])} used. Technicians can't confirm the failed part "
            f"before an on-site visit, so they bring several candidates and return the rest, averaging {ppc.loc[worst_t, '每案平均申請顆數']:.1f} "
            f"parts per case vs. just {ppc.loc['Hardware', '每案平均申請顆數']:.2f} for Hardware. By country, it's concentrated in {T(worst_c)}."
        )
        a.conclusion = (f"The {_pct(k['未使用率'], 0)} parts return rate comes from {worst_t}'s bring-several-candidates practice, especially in "
                        f"{T(worst_c)}. Review remote diagnosis before dispatch and the parts-request rules, starting with {T(worst_c)}. Part DOA at "
                        f"{_pct(k['零件DOA率'])} is normal and needs no action.")
    return a





def a42_shortage(d):
    a = Analysis('4-2', T('零件與成本｜缺料與延誤'), T('在途案件中 Awaiting Spares 的等待天數（CurrentStatusDays）'))
    o = d[d["在途"]]
    s = o[o["缺料"]]
    top6 = main_countries(d)
    order = ["0-6", "7-13", "14-20", "21-27", "28+"]
    wait = pd.crosstab(s["國家"], s["等待天數區間"]).reindex(index=top6, columns=order, fill_value=0)
    wait = wait.loc[:, wait.sum() > 0]
    wait["合計"] = wait.sum(axis=1)
    wait["超過兩週"] = s[s["CurrentStatusDays"] >= 14].groupby("國家").size().reindex(top6, fill_value=0)
    wait["超過兩週比例"] = wait["超過兩週"] / wait["合計"].replace(0, np.nan)
    wait["等待中位數"] = s.groupby("國家")["CurrentStatusDays"].median().reindex(top6)
    share = o.groupby("國家").agg(在途=("缺料", "size"), 缺料=("缺料", "sum"), 零件在途或已配=("零件在途或已配", "sum"))
    share["缺料占在途比例"] = share["缺料"] / share["在途"]
    share = share.loc[top6].sort_values("缺料占在途比例", ascending=False)
    mix = pd.crosstab(s["國家"], s["類型分組"]).reindex(index=top6, fill_value=0)
    mix = mix.div(mix.sum(axis=1).replace(0, np.nan), axis=0)
    a.tables = {"缺料等待天數_依國家": wait, "缺料占在途比例_依國家": share, "缺料案類型組成_依國家": mix}
    k = a.kpis
    k["在途件數"] = len(o); k["缺料件數"] = len(s); k["缺料占在途"] = len(s) / max(len(o), 1)
    k["零件在途或已配件數"] = int(o["零件在途或已配"].sum())
    k["六天內比例"] = float((s["CurrentStatusDays"] <= 6).mean()) if len(s) else 0
    k["超過兩週件數"] = int((s["CurrentStatusDays"] >= 14).sum()); k["超過兩週比例"] = k["超過兩週件數"] / max(len(s), 1)
    k["等待中位數"] = float(s["CurrentStatusDays"].median()) if len(s) else 0
    slowest = wait["超過兩週比例"].idxmax() if wait["超過兩週比例"].notna().any() else top6[0]
    top_type = mix.loc[slowest].idxmax() if slowest in mix.index else "－"
    a.bullets = [
        f"在途 {_n(len(o))} 件中缺料等貨（Awaiting Spares）{_n(len(s))} 件，占 {_pct(k['缺料占在途'], 0)}；另有 "
        f"{_n(k['零件在途或已配件數'])} 件零件已配貨或在途中，不算缺料",
        f"等待天數：{_pct(k['六天內比例'], 0)} 在 6 天內，{_pct(k['超過兩週比例'], 0)} 已超過兩週（{_n(k['超過兩週件數'])} 件），中位數 {k['等待中位數']:.0f} 天",
        f"缺料占各國在途案的比例：{'、'.join(f'{i} {_pct(r.缺料占在途比例, 0)}' for i, r in share.iterrows())}",
        f"延誤集中：{slowest} 缺料案 {_pct(wait.loc[slowest, '超過兩週比例'], 0)} 已等超過兩週（{_n(wait.loc[slowest, '超過兩週'])} 件，中位數 "
        f"{wait.loc[slowest, '等待中位數']:.0f} 天）；{slowest} 的缺料案 {_pct(mix.loc[slowest, top_type], 0)} 是 {top_type}",
    ]
    others = wait.drop(index=slowest)
    quick = others["超過兩週比例"].idxmin() if len(others) else "－"
    a.paragraph = (
        f"在途的 {_n(len(o))} 件裡有 {_n(len(s))} 件狀態是 Awaiting Spares，就是零件缺貨在等供應商或總倉出貨，占 "
        f"{_pct(k['缺料占在途'], 0)}；另外 {_n(k['零件在途或已配件數'])} 件零件已經配到或在路上，不算缺料。等待天數一半以上不長，"
        f"{_pct(k['六天內比例'], 0)} 在 6 天內，但有 {_n(k['超過兩週件數'])} 件超過兩週。分國家看，{share.index[0]}、{share.index[1]} 的在途案有 "
        f"{_pct(share['缺料占在途比例'].iloc[1], 0)}–{_pct(share['缺料占在途比例'].iloc[0], 0)} 卡在缺料。真正慢的是 {slowest}："
        f"{_n(wait.loc[slowest, '合計'])} 件缺料案有 {_n(wait.loc[slowest, '超過兩週'])} 件超過兩週，中位數已等 "
        f"{wait.loc[slowest, '等待中位數']:.0f} 天，而且 {_pct(mix.loc[slowest, top_type], 0)} 是 {top_type}"
        f"{'，問題出在新機備品的庫存，不是維修零件' if top_type == 'DOA' else ''}；{quick} 幾乎沒有超過兩週的。"
    )
    a.conclusion = (f"缺料延誤不是全區問題，集中在 {slowest}（{top_type}）；建議先補 {slowest} 的"
                    f"{'DOA 備機庫存' if top_type == 'DOA' else '對應零件安全庫存'}，其他國家維持現狀，並每週追蹤等待超過兩週的缺料清單。")

    if is_en():
        a.bullets = [
            f"Of {_n(len(o))} open cases, {_n(len(s))} are waiting on parts (Awaiting Spares), {_pct(k['缺料占在途'], 0)}; another "
            f"{_n(k['零件在途或已配件數'])} cases have parts allocated or on the way and don't count as shortages",
            f"Days waiting: {_pct(k['六天內比例'], 0)} within 6 days, {_pct(k['超過兩週比例'], 0)} over two weeks ({_n(k['超過兩週件數'])} cases), "
            f"median {k['等待中位數']:.0f} days",
            f"Share of each country's open cases waiting on parts: {', '.join(f'{T(i)} {_pct(r.缺料占在途比例, 0)}' for i, r in share.iterrows())}",
            f"Delays are concentrated: {_pct(wait.loc[slowest, '超過兩週比例'], 0)} of {T(slowest)}'s shortage cases have waited over two weeks "
            f"({_n(wait.loc[slowest, '超過兩週'])} cases, median {wait.loc[slowest, '等待中位數']:.0f} days); "
            f"{_pct(mix.loc[slowest, top_type], 0)} of them are {T(top_type)}",
        ]
        a.paragraph = (
            f"{_n(len(s))} of the {_n(len(o))} open cases are in Awaiting Spares, meaning the part is out of stock and waiting on the supplier or "
            f"central warehouse ({_pct(k['缺料占在途'], 0)}). Another {_n(k['零件在途或已配件數'])} cases already have parts allocated or on the way "
            f"and don't count as shortages. Most waits aren't long: {_pct(k['六天內比例'], 0)} are within 6 days, but {_n(k['超過兩週件數'])} cases "
            f"are past two weeks. By country, {_pct(share['缺料占在途比例'].iloc[1], 0)}–{_pct(share['缺料占在途比例'].iloc[0], 0)} of open cases in "
            f"{T(share.index[0])} and {T(share.index[1])} are stuck waiting on parts. The real slowdown is {T(slowest)}: "
            f"{_n(wait.loc[slowest, '超過兩週'])} of {_n(wait.loc[slowest, '合計'])} shortage cases are past two weeks, with a median wait of "
            f"{wait.loc[slowest, '等待中位數']:.0f} days, and {_pct(mix.loc[slowest, top_type], 0)} are {T(top_type)}"
            f"{', so the gap is in replacement-unit stock for new machines, not repair parts' if top_type == 'DOA' else ''}. "
            f"{T(quick)} has almost none past two weeks."
        )
        a.conclusion = (f"Parts delays aren't region-wide; they're concentrated in {T(slowest)} ({T(top_type)}). Start by building up {T(slowest)}'s "
                        f"{'DOA replacement-unit stock' if top_type == 'DOA' else 'safety stock for the affected parts'}, leave the other countries "
                        f"as they are, and review the list of shortages waiting over two weeks every week.")
    return a





def a51_warranty(d):
    a = Analysis('5-1', T('保固與財務｜保固內外比例'), T('WarrantyStatus（實際處理狀態）依國家、類型、產品線'))
    w = d[d["WarrantyStatus"].notna()]
    out = w["WarrantyStatus"].eq("Out Of Warranty")
    top6 = main_countries(d)
    by_c = w.groupby("國家").agg(案件數=("CaseID", "size"))
    by_c["保固外"] = out.groupby(w["國家"]).sum()
    by_c["保固外比率"] = by_c["保固外"] / by_c["案件數"]
    by_c["占保固外總數"] = by_c["保固外"] / out.sum()
    by_c6 = by_c.loc[top6].sort_values("保固外比率", ascending=False)
    by_t = w[w["Type"].isin(MAIN_TYPES)].groupby("Type").agg(案件數=("CaseID", "size"))
    by_t["保固外"] = out.groupby(w["Type"]).sum().reindex(by_t.index)
    by_t["保固外比率"] = by_t["保固外"] / by_t["案件數"]
    by_t["占保固外總數"] = by_t["保固外"] / out.sum()
    by_t = by_t.sort_values("保固外比率", ascending=False)
    by_p = w.groupby("ProductLine").agg(案件數=("CaseID", "size"))
    by_p["保固外"] = out.groupby(w["ProductLine"]).sum()
    by_p["保固外比率"] = by_p["保固外"] / by_p["案件數"]
    by_p = by_p[by_p["案件數"] >= MIN_GROUP].sort_values("保固外比率", ascending=False)
    cmp = pd.crosstab(w["WarrantyStatus"], w["WarrantyClassification"], margins=True)
    o = d[d["在途"]]
    o_out = o[o["WarrantyStatus"].eq("Out Of Warranty")]
    a.tables = {"保固內外_依國家": by_c6, "保固內外_依類型": by_t, "保固內外_依產品線": by_p,
                "保固狀態vs保固分類": cmp, "保固內外_全部國家": by_c.sort_values("案件數", ascending=False)}
    k = a.kpis
    k["保固外件數"] = int(out.sum()); k["保固外比率"] = float(out.mean())
    k["保固外最高國家"] = by_c6.index[0]; k["保固外最高國家比率"] = float(by_c6["保固外比率"].iloc[0])
    k["Hardware占保固外"] = float(by_t.loc["Hardware", "占保固外總數"]) if "Hardware" in by_t.index else 0
    inp_out = int(((w["WarrantyStatus"] == "Out Of Warranty") & (w["WarrantyClassification"] == "In Warranty")).sum())
    k["保固期內以保固外處理"] = inp_out
    k["在途保固外件數"] = len(o_out)
    k["在途保固外待客戶"] = int(o_out["狀態分類"].eq("待客戶").sum())
    a.bullets = [
        f"母體 {_n(len(w))} 件：保固內 {_n(len(w) - out.sum())} 件（{_pct(1 - k['保固外比率'])}）、保固外 {_n(out.sum())} 件（{_pct(k['保固外比率'])}）",
        f"國家：{'、'.join(f'{i} {_pct(r.保固外比率)}（{_n(r.保固外)} 件）' for i, r in by_c6.iterrows())}",
        f"類型：保固外幾乎都是 {by_t['占保固外總數'].idxmax()}（{_n(by_t['保固外'].max())} 件，占全部保固外的 {_pct(by_t['占保固外總數'].max(), 0)}）",
        f"產品線：{'、'.join(f'{i} {_pct(r.保固外比率, 0)}' for i, r in by_p.head(3).iterrows())} 保固外比例最高",
        f"保固期內但以保固外處理的有 {_n(inp_out)} 件，多為 Hardware 送修，推測為人為損壞等不在保固範圍的收費維修",
        f"在途保固外 {_n(len(o_out))} 件中 {_n(k['在途保固外待客戶'])} 件停在待客戶核准報價",
    ]
    hi, lo = by_c6.index[0], by_c6.index[-1]
    a.paragraph = (
        f"{_n(len(w))} 件裡 {_pct(1 - k['保固外比率'])} 在保固內，保固外 {_pct(k['保固外比率'])}、{_n(out.sum())} 件。但分國家差距非常大："
        f"{hi} {_pct(by_c6.loc[hi, '保固外比率'], 0)}（{_n(by_c6.loc[hi, '保固外'])} 件），{lo} 只有 {_pct(by_c6.loc[lo, '保固外比率'])}。保固外集中在 "
        f"{by_t['占保固外總數'].idxmax()}，占全部保固外的 {_pct(by_t['占保固外總數'].max(), 0)}；產品線上 "
        f"{'、'.join(by_p.head(3).index)} 這幾條較舊機種的保固外比例最高。保固外案子要先報價給客戶，所以在途的 {_n(len(o_out))} 件保固外案裡有 "
        f"{_n(k['在途保固外待客戶'])} 件停在待客戶核准。"
    )
    a.conclusion = f"保固外維修是收費業務，{hi} 等國是主要來源；{lo} 幾乎沒有保固外案件，是市場本來就沒有、還是保固外維修沒有進這個系統，值得確認，因為這會影響收費維修營收的統計。"

    if is_en():
        a.bullets = [
            f"Base: {_n(len(w))} cases: {_n(len(w) - out.sum())} in warranty ({_pct(1 - k['保固外比率'])}), {_n(out.sum())} out of warranty "
            f"({_pct(k['保固外比率'])})",
            f"By country: {', '.join(f'{T(i)} {_pct(r.保固外比率)} ({_n(r.保固外)} cases)' for i, r in by_c6.iterrows())}",
            f"By type: almost all out-of-warranty cases are {by_t['占保固外總數'].idxmax()} ({_n(by_t['保固外'].max())} cases, "
            f"{_pct(by_t['占保固外總數'].max(), 0)} of all out-of-warranty)",
            f"Product lines with the highest out-of-warranty rates: {', '.join(f'{i} {_pct(r.保固外比率, 0)}' for i, r in by_p.head(3).iterrows())}",
            f"{_n(inp_out)} cases were inside the warranty period but handled as out of warranty, mostly Hardware carry-ins; likely paid repairs "
            f"for damage the warranty doesn't cover, such as accidental damage",
            f"Of {_n(len(o_out))} open out-of-warranty cases, {_n(k['在途保固外待客戶'])} are waiting for the customer to approve a quote",
        ]
        a.paragraph = (
            f"Of {_n(len(w))} cases, {_pct(1 - k['保固外比率'])} were in warranty and {_pct(k['保固外比率'])} ({_n(out.sum())}) out of warranty. "
            f"The gap between countries is huge, though: {T(hi)} {_pct(by_c6.loc[hi, '保固外比率'], 0)} ({_n(by_c6.loc[hi, '保固外'])} cases), while "
            f"{T(lo)} is just {_pct(by_c6.loc[lo, '保固外比率'])}. Out-of-warranty cases are concentrated in {by_t['占保固外總數'].idxmax()}, "
            f"which accounts for {_pct(by_t['占保固外總數'].max(), 0)} of them. By product line, older models ({', '.join(by_p.head(3).index)}) "
            f"have the highest out-of-warranty rates. These cases need a quote approved by the customer first, so {_n(k['在途保固外待客戶'])} of "
            f"the {_n(len(o_out))} open out-of-warranty cases are waiting on customer approval."
        )
        a.conclusion = (f"Out-of-warranty repair is paid work, and {T(hi)} and similar countries are the main source. {T(lo)} has almost no "
                        f"out-of-warranty cases. It's worth checking whether that market truly has none or whether those repairs aren't going "
                        f"through this system, since that affects how paid-repair revenue is reported.")
    return a





def a52_expiry(d):
    ref = d.attrs["ref_date"]
    a = Analysis("5-2", T("保固與財務｜保固將到期的在途案"), L(f"在途案件的保固到期狀況（以 {ref:%Y/%m/%d} 為基準）", f"Warranty expiry status of open cases (as of {ref:%b %d, %Y})"))
    o = d[d["在途"]]
    top6 = main_countries(d)
    order = ["1 已過期", "2 30天內", "3 31-90天", "4 90天以上"]
    by_c = pd.crosstab(o["國家"], o["保固到期區間"]).reindex(index=top6, columns=order, fill_value=0)
    by_c["合計"] = by_c.sum(axis=1)
    by_c["已過期比例"] = by_c["1 已過期"] / by_c["合計"]
    tot = o["保固到期區間"].value_counts().reindex(order, fill_value=0).rename("件數").to_frame()
    tot["占在途比例"] = tot["件數"] / len(o)
    exp = o[o["保固到期區間"].eq("1 已過期")]
    exp_split = exp["WarrantyStatus"].value_counts().rename("件數").to_frame()
    exp_in = exp[exp["WarrantyStatus"].eq("In Warranty")]
    exp_in_c = exp_in["國家"].value_counts().head(6).rename("件數").to_frame()
    soon = o[o["保固到期區間"].eq("2 30天內")]
    soon_c = soon["國家"].value_counts().head(6).rename("件數").to_frame()
    a.tables = {"保固到期區間_依國家": by_c, "保固到期區間_合計": tot, "已過期在途_保固狀態": exp_split,
                "建案時保固內現已過期_依國家": exp_in_c, "30天內到期_依國家": soon_c}
    k = a.kpis
    k["在途件數"] = len(o)
    k["已過期件數"] = int(tot.loc["1 已過期", "件數"]); k["30天內到期件數"] = int(tot.loc["2 30天內", "件數"]); k["31-90天到期件數"] = int(tot.loc["3 31-90天", "件數"])
    k["90天內合計"] = k["已過期件數"] + k["30天內到期件數"] + k["31-90天到期件數"]
    k["建案時保固內現已過期"] = len(exp_in)
    k["建案時已保固外"] = int(len(exp) - len(exp_in))
    hi = by_c["已過期比例"].idxmax()
    a.bullets = [
        f"在途 {_n(len(o))} 件中，保固已過期 {_n(k['已過期件數'])} 件（{_pct(tot.loc['1 已過期', '占在途比例'])}）、30 天內到期 "
        f"{_n(k['30天內到期件數'])} 件（{_pct(tot.loc['2 30天內', '占在途比例'])}）、31–90 天到期 {_n(k['31-90天到期件數'])} 件，合計 "
        f"{_n(k['90天內合計'])} 件（{_pct(k['90天內合計'] / len(o))}）需要優先處理",
        f"已過期的 {_n(len(exp))} 件裡，{_n(k['建案時已保固外'])} 件建案時就是保固外的收費維修；{_n(len(exp_in))} 件建案時還在保固內、拖到現在過期（"
        f"{'、'.join(f'{i} {_n(v.件數)}' for i, v in exp_in_c.head(3).iterrows())}）",
        f"國家：{hi} 在途案 {_pct(by_c.loc[hi, '已過期比例'], 0)}（{_n(by_c.loc[hi, '1 已過期'])} 件）保固已過期，其他國家 "
        f"{_pct(by_c.drop(index=hi)['已過期比例'].min(), 0)}–{_pct(by_c.drop(index=hi)['已過期比例'].max(), 0)}",
        f"30 天內到期件數：{'、'.join(f'{i} {_n(v.件數)}' for i, v in soon_c.iterrows())}",
    ]
    a.paragraph = (
        f"在途的 {_n(len(o))} 件裡，保固已經過期 {_n(k['已過期件數'])} 件，30 天內會到期 {_n(k['30天內到期件數'])} 件，31 到 90 天內到期 "
        f"{_n(k['31-90天到期件數'])} 件，加起來 {_n(k['90天內合計'])} 件。已過期的要分兩種看：{_n(k['建案時已保固外'])} 件建案時就是保固外，本來就是收費維修，過期不是問題；另外 "
        f"{_n(len(exp_in))} 件建案時還在保固內，是案子拖著保固到期了，這些客戶的權益要照建案當時的保固內處理，結案時不能被當成保固外收費。"
        f"{hi} 在途案有 {_pct(by_c.loc[hi, '已過期比例'], 0)} 保固已過期，跟該國保固外比例高是同一件事。真正要盯的是 30 天內到期的 "
        f"{_n(k['30天內到期件數'])} 件，如果在到期前修不完，之後零件和工資的歸屬會有爭議。"
    )
    a.conclusion = (f"建議每週跑一次「在途且保固 30 天內到期」清單給各據點優先處理（本期 {_n(k['30天內到期件數'])} 件），並把 "
                    f"{_n(len(exp_in))} 件「建案時保固內、現已過期」的案子標記起來，確保結案時按保固內處理。")
    if is_en():
        a.bullets = [
            f"Of {_n(len(o))} open cases, {_n(k['已過期件數'])} have expired warranties ({_pct(tot.loc['1 已過期', '占在途比例'])}), "
            f"{_n(k['30天內到期件數'])} expire within 30 days ({_pct(tot.loc['2 30天內', '占在途比例'])}), and {_n(k['31-90天到期件數'])} within "
            f"31–90 days: {_n(k['90天內合計'])} cases ({_pct(k['90天內合計'] / len(o))}) to prioritize",
            f"Of the {_n(len(exp))} expired, {_n(k['建案時已保固外'])} were already out-of-warranty paid repairs when created; {_n(len(exp_in))} were "
            f"in warranty at creation and lapsed while the case was open ({', '.join(f'{T(i)} {_n(v.件數)}' for i, v in exp_in_c.head(3).iterrows())})",
            f"By country: {_pct(by_c.loc[hi, '已過期比例'], 0)} of {T(hi)}'s open cases ({_n(by_c.loc[hi, '1 已過期'])}) have expired warranties, "
            f"vs. {_pct(by_c.drop(index=hi)['已過期比例'].min(), 0)}–{_pct(by_c.drop(index=hi)['已過期比例'].max(), 0)} elsewhere",
            f"Expiring within 30 days: {', '.join(f'{T(i)} {_n(v.件數)}' for i, v in soon_c.iterrows())}",
        ]
        a.paragraph = (
            f"Of the {_n(len(o))} open cases, {_n(k['已過期件數'])} already have expired warranties, {_n(k['30天內到期件數'])} expire within 30 "
            f"days, and {_n(k['31-90天到期件數'])} within 31 to 90 days: {_n(k['90天內合計'])} in total. Expired cases fall into two groups. "
            f"{_n(k['建案時已保固外'])} were already out of warranty when created, so they're paid repairs and expiry isn't an issue. The other "
            f"{_n(len(exp_in))} were in warranty at creation, and the warranty ran out while the case dragged on. Those customers are entitled to "
            f"in-warranty treatment based on their status at creation and can't be charged as out of warranty at closing. "
            f"{_pct(by_c.loc[hi, '已過期比例'], 0)} of {T(hi)}'s open cases have expired warranties, which ties back to that country's high "
            f"out-of-warranty rate. The ones to watch are the {_n(k['30天內到期件數'])} cases expiring within 30 days: if they aren't fixed before "
            f"expiry, who pays for parts and labor afterward becomes a dispute."
        )
        a.conclusion = (f"Run an \"open, warranty ends within 30 days\" list every week for sites to prioritize ({_n(k['30天內到期件數'])} cases "
                        f"this period), and flag the {_n(len(exp_in))} cases that were in warranty at creation but have since expired so they're "
                        f"closed as in-warranty.")
    return a


ANALYSES = [a11_volume, a12_in_progress, a13_internal, a21_tat, a22_sla, a31_pareto, a32_repeat,
            a41_parts, a42_shortage, a51_warranty, a52_expiry]


def run_all(d, progress=None):
    """跑完全部分析，回傳 Report。progress(i, total, title) 可選。"""
    rep = Report(period_start=d.attrs["period_start"], ref_date=d.attrs["ref_date"], n_cases=len(d), data=d)
    for i, fn in enumerate(ANALYSES, 1):
        a = fn(d)
        rep.analyses[a.key] = a
        if progress:
            progress(i, len(ANALYSES), a.title)
    rep.parts = parts_table(d)
    return rep


# ---------------------------------------------------------------- 零件規劃


def demand_weekly(d, by="國家", top6=None):
    """完整週的申請零件顆數，依國家或類型。每個國家各一欄（依申請量由大到小）；指定 top6 時其餘併為 Other。"""
    wk = d[d["完整週"]]
    key = wk[by] if by != "國家" or not top6 else wk["國家"].where(wk["國家"].isin(top6), "Other")
    t = wk.groupby([wk["建案週"], key])["NumberOfRequestedParts"].sum().unstack(fill_value=0)
    t = t[t.sum().sort_values(ascending=False).index]
    t.index = pd.Index([w.strftime("%m/%d") for w in t.index], name="週")
    t["合計"] = t.sum(axis=1)
    return t


def demand_forecast(d, weeks_ahead=4, by="國家", recent_weeks=4):
    """以最近 N 個完整週的平均申請顆數推估未來需求。"""
    t = demand_weekly(d, by)
    recent = t.tail(recent_weeks)
    f = pd.DataFrame({"最近每週平均": recent.mean(), "最近每週最高": recent.max(),
                      f"未來{weeks_ahead}週需求(平均)": recent.mean() * weeks_ahead,
                      f"未來{weeks_ahead}週需求(高峰)": recent.max() * weeks_ahead}).round(0)
    return f


def top_parts(parts, n=30):
    """申請次數最多的料號：申請顆數、案件數、未使用率、缺料中顆數、主要國家。"""
    # 向量化計算（逐群 lambda 在瀏覽器版會慢十幾秒）
    code = parts["狀態碼"]
    flags = pd.DataFrame({"已使用": code.eq("C"), "未使用退回": code.eq("U"), "零件DOA": code.eq("D"),
                          "待供貨": code.eq("R"), "運送中": code.eq("I")}, index=parts.index)
    g = parts.groupby("料號").agg(申請顆數=("CaseID", "size"), 案件數=("CaseID", "nunique"))
    g = g.join(flags.groupby(parts["料號"]).sum().astype(int))
    g["未使用率"] = g["未使用退回"] / (g["已使用"] + g["未使用退回"]).replace(0, np.nan)
    # 主要國家：該料號出現最多的國家（同數取先出現的）
    cnt = parts.groupby(["料號", "國家"], sort=False).size().reset_index(name="n")
    cnt = cnt.sort_values("n", ascending=False, kind="stable").drop_duplicates("料號")
    g["主要國家"] = cnt.set_index("料號")["國家"]
    return g.sort_values("申請顆數", ascending=False).head(n)


def shortage_parts(parts):
    """目前在 Awaiting Spares 案件裡、狀態碼為 R（已申請待供貨）的料號清單。"""
    s = parts[(parts["Status"] == "Awaiting Spares") & (parts["狀態碼"] == "R")]
    g = s.groupby("料號").agg(等待案件數=("CaseID", "nunique"), 待供貨顆數=("CaseID", "size"),
                            最長等待天數=("CurrentStatusDays", "max"), 平均等待天數=("CurrentStatusDays", "mean"),
                            國家=("國家", lambda x: list_sep().join(L(f"{k}{v}", f"{T(k)} ({v})") for k, v in x.value_counts().items())))
    return g.sort_values(["等待案件數", "最長等待天數"], ascending=False)


def shortage_cases(d, min_days=14):
    """等待超過 min_days 天的缺料案件清單（催料用）。"""
    s = d[d["缺料"] & (d["CurrentStatusDays"] >= min_days)]
    cols = ["CaseID", "國家", "Location", "Type", "ProductModel", "SerialNumber",
            "CurrentStatusDays", "DaysSinceRcvd", "WarrantyStatus", "PartsRequested"]
    return s[cols].sort_values("CurrentStatusDays", ascending=False)


def ship_backlog(d):
    """待出貨積壓：依據點的件數與停留天數。"""
    o = d[d["狀態分類"].eq("待出貨")]
    g = o.groupby("Location").agg(待出貨件數=("CaseID", "size"), 停留中位數=("CurrentStatusDays", "median"),
                                  超過7天=("CurrentStatusDays", lambda s: int((s > 7).sum())),
                                  國家=("國家", "first"))
    return g.sort_values("待出貨件數", ascending=False)


def parts_in_transit(parts):
    """零件運送中（I）與已配貨（A）的顆數，依國家。"""
    s = parts[parts["狀態碼"].isin(["I", "A"]) & ~parts["已結案"]]
    g = pd.crosstab(s["國家"], s["零件狀態"])
    g["合計"] = g.sum(axis=1)
    return g.sort_values("合計", ascending=False)


def expiring_cases(d, days=30):
    """在途且保固 days 天內到期（含已過期但建案時保固內）的案件清單。"""
    ref = d.attrs["ref_date"]
    o = d[d["在途"]]
    soon = o[(o["WarrantyExpiryDate"] <= ref + pd.Timedelta(days=days))
             & ((o["WarrantyExpiryDate"] >= ref) | o["WarrantyStatus"].eq("In Warranty"))]
    cols = ["CaseID", "國家", "Location", "Type", "ProductModel", "Status", "狀態分類",
            "WarrantyStatus", "WarrantyExpiryDate", "DaysSinceRcvd"]
    out = soon[cols].copy()
    out["距到期天數"] = (out["WarrantyExpiryDate"] - ref).dt.days
    out["到期狀態"] = np.where(out["距到期天數"] < 0, "已過期（建案時保固內）", "即將到期")
    out["_ord"] = np.where(out["距到期天數"] < 0, 1, 0)
    out["_key"] = out["距到期天數"].abs()
    return out.sort_values(["_ord", "_key"]).drop(columns=["_ord", "_key"])


# ---------------------------------------------------------------- 圖表
# matplotlib 延遲載入；中文字型只設定一次

_FONT_DONE = {}


def _mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    lang = "en" if is_en() else "zh"
    if _FONT_DONE.get("lang") != lang:
        from matplotlib import font_manager
        have = {f.name for f in font_manager.fontManager.ttflist}
        cjk = [c for c in ("Microsoft JhengHei", "Microsoft YaHei", "PingFang TC", "Noto Sans CJK TC",
                           "Noto Sans TC", "Arial Unicode MS", "SimHei") if c in have][:1]
        latin = [c for c in ("Segoe UI", "Arial", "Helvetica", "Liberation Sans") if c in have][:1]
        # 英文模式以西文字型為主、中文字型備援；中文模式反之
        fams = (latin + cjk) if lang == "en" else (cjk + latin)
        plt.rcParams["font.family"] = fams + ["DejaVu Sans"]
        plt.rcParams["axes.unicode_minus"] = False
        _FONT_DONE["lang"] = lang
    return plt


PALETTE = ["#2E6FBA", "#D9822B", "#2E9E6B", "#8E4FB5", "#C9A227", "#6B7280", "#4FB0C6", "#B5544C"]


def _bar_labels(ax, fmt="{:.0f}", pct=False, min_val=0.03):
    for cont in ax.containers:
        labels = []
        for v in cont.datavalues:
            if np.isnan(v) or (pct and v < min_val) or (not pct and v == 0):
                labels.append("")
            else:
                labels.append(f"{v:.0%}" if pct else fmt.format(v))
        ax.bar_label(cont, labels=labels, label_type="center" if len(ax.containers) > 1 else "edge", fontsize=8, color="white" if len(ax.containers) > 1 else "black")


def plot_analysis(a, path):
    plt = _mpl()
    key = a.key
    if key == "1-1":
        fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), gridspec_kw={"width_ratios": [1.3, 1]})
        t = a.tables["週案件量_依類型"].drop(columns="合計").rename(columns=T)
        t.plot(kind="bar", stacked=True, ax=axes[0], color=PALETTE, width=0.7)
        axes[0].set_title(T("每週案件量（完整週，依類型）")); axes[0].set_xlabel(""); axes[0].tick_params(axis="x", rotation=0)
        for i, v in enumerate(a.tables["週案件量_依類型"]["合計"]):
            axes[0].text(i, v, f"{int(v):,}", ha="center", va="bottom", fontsize=9)
        axes[0].legend(fontsize=8, ncol=3, frameon=False)
        top = a.tables["Top10據點"].iloc[::-1]
        axes[1].barh(top.index, top["案件數"], color=PALETTE[0])
        for i, (v, p) in enumerate(zip(top["案件數"], top["占比"])):
            axes[1].text(v, i, f" {int(v):,} ({p:.0%})", va="center", fontsize=8)
        axes[1].set_title(T("Top 10 據點")); axes[1].tick_params(axis="y", labelsize=8)
    elif key == "1-2":
        fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), gridspec_kw={"width_ratios": [1.2, 1]})
        t = a.tables["在途_類型×狀態"].drop(columns="合計").rename(columns=T)
        t.plot(kind="barh", stacked=True, ax=axes[0], color=PALETTE, width=0.7)
        axes[0].invert_yaxis(); axes[0].set_title(T("在途案件：類型 × 狀態分類")); axes[0].set_ylabel("")
        _bar_labels(axes[0]); axes[0].legend(fontsize=8, ncol=3, frameon=False)
        ag = a.tables["在途_停留天數"]
        axes[1].bar([textwrap.fill(str(T(x)), 12) for x in ag.index], ag["超過14天比例"], color=PALETTE[1])
        for i, v in enumerate(ag["超過14天比例"]):
            axes[1].text(i, v, f"{v:.0%}", ha="center", va="bottom", fontsize=9)
        axes[1].set_title(L(f"停留超過 {AGING_LIMIT} 天的比例（依狀態分類）", f"Share Open Over {AGING_LIMIT} Days (by Status Group)")); axes[1].set_ylim(0, max(ag["超過14天比例"].max() * 1.3, 0.1))
        axes[1].yaxis.set_major_formatter(lambda v, _: f"{v:.0%}")
    elif key == "1-3":
        fig, ax = plt.subplots(figsize=(11, 3.8))
        t = a.tables["內外部×類型_占比"].iloc[::-1].rename(index=T)
        t.plot(kind="barh", stacked=True, ax=ax, color=PALETTE, width=0.6)
        _bar_labels(ax, pct=True); ax.set_title(T("內部 vs 外部客戶：案件類型組成")); ax.set_ylabel("")
        ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); ax.legend(fontsize=8, ncol=6, frameon=False, loc="lower center", bbox_to_anchor=(0.5, -0.35))
    elif key == "2-1":
        fig, axes = plt.subplots(1, 2, figsize=(13, 4.8), gridspec_kw={"width_ratios": [1.2, 1]})
        t = a.tables["TAT分布_依國家(占比)"]
        for i, c in enumerate(t.columns):
            axes[0].plot(t.index, t[c], marker="o", label=T(c), color=PALETTE[i % len(PALETTE)])
        axes[0].set_title(T("已結案 TAT 分布（依國家，收件至維修完成，工作天）")); axes[0].yaxis.set_major_formatter(lambda v, _: f"{v:.0%}")
        axes[0].legend(fontsize=8, frameon=False)
        seg = a.tables["各段耗時_依類型(日曆天)"].drop(columns="收件→結案").rename(columns=T, index=T)
        seg.plot(kind="barh", stacked=True, ax=axes[1], color=[PALETTE[0], PALETTE[2], PALETTE[1]], width=0.6)
        axes[1].invert_yaxis(); _bar_labels(axes[1], fmt="{:.1f}"); axes[1].set_title(T("各階段平均耗時（日曆天）")); axes[1].set_ylabel("")
        axes[1].legend(fontsize=8, frameon=False, ncol=3, loc="upper center", bbox_to_anchor=(0.5, -0.08))
    elif key == "2-2":
        fig, ax = plt.subplots(figsize=(11, 4.2))
        t = a.tables["SLA_依國家"]
        bars = ax.bar([T(x) for x in t.index], t["達成率"], color=PALETTE[0])
        ax.axhline(a.kpis["整體達成率"], color=PALETTE[1], ls="--", lw=1.5, label=L(f"整體 {a.kpis['整體達成率']:.0%}", f"Overall {a.kpis['整體達成率']:.0%}"))
        for b, v, m in zip(bars, t["達成率"], t["未達成件數"]):
            ax.text(b.get_x() + b.get_width() / 2, v, L(f"{v:.0%}\n({m} 件未達)", f"{v:.0%}\n({m} missed)"), ha="center", va="bottom", fontsize=8)
        ax.set_ylim(0, 1.15); ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); ax.set_title(L(f"SLA 達成率（TAT ≤ {SLA_DAYS} 工作天，依國家）", f"SLA Attainment (TAT ≤ {SLA_DAYS} Business Days, by Country)")); ax.legend(frameon=False)
    elif key == "3-1":
        fig, ax = plt.subplots(figsize=(13, 4.8))
        t = a.tables["Top20故障代碼"]
        ax.bar(range(len(t)), t["案件數"], color=PALETTE[0])
        ax.set_xticks(range(len(t))); ax.set_xticklabels(t.index, rotation=45, ha="right", fontsize=8)
        ax2 = ax.twinx(); ax2.plot(range(len(t)), t["累計占比"], color=PALETTE[1], marker="o")
        ax2.set_ylim(0, 1.05); ax2.yaxis.set_major_formatter(lambda v, _: f"{v:.0%}")
        for i, v in enumerate(t["累計占比"]):
            if i % 2 == 0 or i == len(t) - 1:
                ax2.text(i, v + 0.03, f"{v:.0%}", ha="center", fontsize=8, color=PALETTE[1])
        ax.set_title(T("故障代碼 Pareto（前 20 名，柱＝案件數，線＝累計占比）"))
    elif key == "3-2":
        fig, axes = plt.subplots(1, 2, figsize=(12, 4.2))
        for ax, name in zip(axes, ["重複率_依國家", "重複率_依類型"]):
            t = a.tables[name].sort_values("重複率")
            ax.barh([T(x) for x in t.index], t["重複率"], color=PALETTE[0])
            for i, (v, c) in enumerate(zip(t["重複率"], t["重複案件"])):
                ax.text(v, i, L(f" {v:.1%} ({int(c)} 件)", f" {v:.1%} ({int(c)} cases)"), va="center", fontsize=8)
            ax.xaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); ax.set_title(L(name.replace("_", "（") + "）", T(name))); ax.set_xlim(0, t["重複率"].max() * 1.35)
    elif key == "4-1":
        fig, ax = plt.subplots(figsize=(11, 4.2))
        t = a.tables["零件去向_依類型"].drop(index="全部合計")[["已使用", "未使用", "零件DOA"]].iloc[::-1].rename(columns=T)
        t.plot(kind="barh", stacked=True, ax=ax, color=[PALETTE[0], PALETTE[1], PALETTE[2]], width=0.6)
        _bar_labels(ax); ax.set_title(T("已結案案件的零件去向（依類型，顆數）")); ax.set_ylabel(""); ax.legend(fontsize=8, frameon=False)
    elif key == "4-2":
        fig, axes = plt.subplots(1, 2, figsize=(13, 4.6), gridspec_kw={"width_ratios": [1.2, 1]})
        t = a.tables["缺料等待天數_依國家"]
        cols = [c for c in ("0-6", "7-13", "14-20", "21-27", "28+") if c in t.columns]
        t[cols].rename(index=T).plot(kind="barh", stacked=True, ax=axes[0], color=PALETTE, width=0.6)
        axes[0].invert_yaxis(); _bar_labels(axes[0]); axes[0].set_title(T("缺料案件已等待天數（依國家）")); axes[0].set_ylabel(""); axes[0].legend(fontsize=8, frameon=False, title=T("天"))
        m = a.tables["缺料案類型組成_依國家"].rename(index=T, columns=T)
        m.plot(kind="barh", stacked=True, ax=axes[1], color=PALETTE, width=0.6)
        axes[1].invert_yaxis(); _bar_labels(axes[1], pct=True, min_val=0.08); axes[1].set_title(T("缺料案件的類型組成（依國家）")); axes[1].set_ylabel("")
        axes[1].xaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); axes[1].legend(fontsize=7, frameon=False, ncol=6, loc="upper center", bbox_to_anchor=(0.5, -0.12))
    elif key == "5-1":
        fig, axes = plt.subplots(1, 2, figsize=(13, 4.4))
        t = a.tables["保固內外_依國家"].iloc[::-1].rename(index=T)
        both = pd.DataFrame({"In Warranty": 1 - t["保固外比率"], "Out Of Warranty": t["保固外比率"]}, index=t.index)
        both.plot(kind="barh", stacked=True, ax=axes[0], color=[PALETTE[0], PALETTE[7]], width=0.6)
        _bar_labels(axes[0], pct=True, min_val=0.04); axes[0].set_title(T("保固內外比例（依國家）")); axes[0].set_ylabel("")
        axes[0].xaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); axes[0].legend(fontsize=8, frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(0.5, -0.12))
        p = a.tables["保固內外_依產品線"].iloc[::-1]
        axes[1].barh(p.index, p["保固外比率"], color=PALETTE[7])
        for i, v in enumerate(p["保固外比率"]):
            axes[1].text(v, i, f" {v:.0%}", va="center", fontsize=8)
        axes[1].set_title(L(f"保固外比率（依產品線，案件數 ≥ {MIN_GROUP}）", f"Out-of-Warranty Rate (by Product Line, ≥ {MIN_GROUP} Cases)")); axes[1].xaxis.set_major_formatter(lambda v, _: f"{v:.0%}"); axes[1].tick_params(axis="y", labelsize=8)
    elif key == "5-2":
        fig, ax = plt.subplots(figsize=(11, 4.4))
        t = a.tables["保固到期區間_依國家"][["1 已過期", "2 30天內", "3 31-90天", "4 90天以上"]].iloc[::-1].rename(index=T, columns=T)
        t.plot(kind="barh", stacked=True, ax=ax, color=[PALETTE[7], PALETTE[1], PALETTE[4], PALETTE[5]], width=0.6)
        _bar_labels(ax); ax.set_title(a.subtitle); ax.set_ylabel(""); ax.legend(fontsize=8, frameon=False, ncol=4)
    else:
        return None
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)
    a.chart = path
    return path


# ---------------------------------------------------------------- 匯出


def export_excel(rep, path):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        info = pd.DataFrame({T("項目"): [T(x) for x in ["分析期間起日", "資料日期", "案件數", "產出時間"]],
                             T("值"): [rep.period_start.strftime("%Y-%m-%d"), rep.ref_date.strftime("%Y-%m-%d"), rep.n_cases,
                                   dt.datetime.now().strftime("%Y-%m-%d %H:%M")]})
        info.to_excel(xw, sheet_name=T("說明"), index=False)
        tr_df(rep.kpi_frame()).to_excel(xw, sheet_name=T("KPI總表"), index=False)
        text = pd.DataFrame([(a.key, a.title, "\n".join(bullet() + b for b in a.bullets), a.paragraph, a.conclusion) for a in rep.analyses.values()],
                            columns=[T(c) for c in ["分析", "名稱", "重點", "段落", "結論"]])
        text.to_excel(xw, sheet_name=T("投影片文字"), index=False)
        used = set()
        for a in rep.analyses.values():
            for name, t in a.tables.items():
                sn = f"{a.key} {T(name)}"[:31].replace("/", "／").replace("*", "x").replace("?", "").replace("[", "(").replace("]", ")").replace(":", "")
                while sn in used:
                    sn = sn[:29] + "_2"
                used.add(sn)
                tr_df(t).to_excel(xw, sheet_name=sn)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width * 1.1 + 2), 60)
    return path


def export_charts(rep, folder):
    os.makedirs(folder, exist_ok=True)
    out = []
    for a in rep.analyses.values():
        p = os.path.join(folder, f"{a.key}_{title_part(a.title, -1)}.png".replace("/", "-").replace(":", ""))
        if plot_analysis(a, p):
            out.append(p)
    return out


def export_pptx(rep, path, chart_folder=None):
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    if chart_folder is None:
        chart_folder = os.path.join(os.path.dirname(path) or ".", "charts")
    if not all(a.chart and os.path.exists(a.chart) for a in rep.analyses.values()):
        export_charts(rep, chart_folder)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]
    NAVY, ORANGE, INK, MUTE = RGBColor(31, 63, 110), RGBColor(232, 134, 43), RGBColor(30, 36, 48), RGBColor(91, 100, 114)

    def tb(slide, x, y, w, h, text, size=14, bold=False, color=INK, align=None):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = box.text_frame; tf.word_wrap = True
        lines = text if isinstance(text, list) else [text]
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line; p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = color
            if align: p.alignment = align
            p.space_after = Pt(4)
        return box

    # 封面
    s = prs.slides.add_slide(blank)
    tb(s, 0.8, 2.4, 11.5, 1.2, T("RMA 案件資料模型與分析報告"), 40, True, NAVY)
    tb(s, 0.8, 3.6, 11.5, 0.8, L(f"分析期間 {rep.period_start:%Y/%m/%d}–{rep.ref_date:%Y/%m/%d}｜{rep.n_cases:,} 件案件｜五個主題、{len(rep.analyses)} 項分析",
                                 f"Period {rep.period_start:%b %d, %Y} – {rep.ref_date:%b %d, %Y} | {rep.n_cases:,} cases | 5 themes, {len(rep.analyses)} analyses"), 18, False, MUTE)

    s = prs.slides.add_slide(blank)
    tb(s, 0.6, 0.4, 12, 0.8, T("分析框架"), 28, True, NAVY)
    themes = {}
    for a in rep.analyses.values():
        themes.setdefault(title_part(a.title, 0), []).append(f"{a.key} {title_part(a.title, 1)}")
    x = 0.6
    for th, items in themes.items():
        tb(s, x, 1.5, 2.4, 0.6, th, 16, True, ORANGE)
        tb(s, x, 2.1, 2.4, 3, items, 13, False, INK)
        x += 2.5

    for a in rep.analyses.values():
        s = prs.slides.add_slide(blank)
        tb(s, 0.5, 0.3, 12.3, 0.8, f"{a.key} {a.title}", 26, True, NAVY)
        tb(s, 0.5, 1.0, 12.3, 0.5, a.subtitle, 13, False, MUTE)
        tb(s, 0.5, 1.6, 5.3, 4.6, [bullet() + b for b in a.bullets], 12, False, INK)
        if a.chart and os.path.exists(a.chart):
            s.shapes.add_picture(a.chart, Inches(6.0), Inches(1.6), width=Inches(7.0))
        tb(s, 0.5, 6.5, 12.3, 0.8, T("結論：") + a.conclusion, 12, True, ORANGE)
        s.notes_slide.notes_text_frame.text = a.paragraph + "\n\n" + T("結論：") + a.conclusion
    prs.save(path)
    return path


def export_all(rep, folder):
    os.makedirs(folder, exist_ok=True)
    charts = export_charts(rep, os.path.join(folder, "charts"))
    xlsx = export_excel(rep, os.path.join(folder, T("RMA分析結果.xlsx")))
    pptx = export_pptx(rep, os.path.join(folder, T("RMA分析報告.pptx")), os.path.join(folder, "charts"))
    return {"excel": xlsx, "pptx": pptx, "charts": charts}


# ---------------------------------------------------------------- 指令列


def _cli():
    import argparse
    import i18n
    i18n.load_settings()
    if "--lang" in sys.argv:
        i = sys.argv.index("--lang")
        if i + 1 < len(sys.argv):
            i18n.set_lang(sys.argv[i + 1])
    ap = argparse.ArgumentParser(description=T("RMA 案件分析引擎"))
    ap.add_argument("file", help=T("AIO 匯出檔 (.xlsx)"))
    ap.add_argument("out", nargs="?", default=T("RMA_輸出"), help=T("輸出資料夾"))
    ap.add_argument("--start", help=T("分析起日 YYYY-MM-DD（預設：資料最後日期的前一個月 1 日）"))
    ap.add_argument("--ref", help=T("資料基準日 YYYY-MM-DD（預設：最後建案日）"))
    ap.add_argument("--lang", choices=["zh", "en"], help="zh / en")
    args = ap.parse_args()
    print(T("讀檔…")); raw = load_aio(args.file)
    d = prepare(raw, args.start, args.ref)
    print(L(f"期間 {d.attrs['period_start']:%Y-%m-%d} 起，{len(d):,} 件；資料日期 {d.attrs['ref_date']:%Y-%m-%d}",
            f"Period from {d.attrs['period_start']:%Y-%m-%d}, {len(d):,} cases; data as of {d.attrs['ref_date']:%Y-%m-%d}"))
    rep = run_all(d, progress=lambda i, n, t: print(f"  [{i}/{n}] {t}"))
    out = export_all(rep, args.out)
    print(T("完成："), out["excel"], out["pptx"], L(f"{len(out['charts'])} 張圖", f"{len(out['charts'])} charts"))


if __name__ == "__main__":
    _cli()
