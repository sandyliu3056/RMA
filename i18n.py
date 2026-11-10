# -*- coding: utf-8 -*-
"""
語言切換（中文／English）
------------------------
程式內部的欄位名稱、分類值、表格名稱一律維持中文；顯示與匯出時才經過 T() 或 tr_df() 轉成目前語言。
分析文字（重點、段落、結論）含數字，在 rma_engine.py 內以 is_en() 分別產生。

語言設定存在程式旁的 settings.json（{"lang": "en"} 或 {"lang": "zh"}）。
"""
import os
import re
import sys
import json

LANG = "zh"
LANGS = ("zh", "en")

BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")


def is_en():
    return LANG == "en"


def get_lang():
    return LANG


def set_lang(lang):
    global LANG
    LANG = lang if lang in LANGS else "zh"
    return LANG


def load_settings():
    """讀 settings.json 並套用語言；檔案不存在或壞掉就維持中文。"""
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as fh:
            s = json.load(fh)
        set_lang(s.get("lang", "zh"))
    except (OSError, ValueError):
        pass
    return LANG


def save_settings():
    try:
        s = {}
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as fh:
                s = json.load(fh)
        s["lang"] = LANG
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump(s, fh, ensure_ascii=False, indent=2)
    except (OSError, ValueError):
        pass


def L(zh, en):
    """依目前語言二選一（用於含變數的句子）。"""
    return en if LANG == "en" else zh


def T(s):
    """固定字串轉成目前語言的顯示文字（中文模式套用白話用詞）；非字串或查不到時原樣回傳。"""
    if not isinstance(s, str):
        return s
    if LANG != "en":
        return ZH.get(s, s)
    hit = EN.get(s)
    if hit is not None:
        return hit
    for pat, tmpl in PATTERNS:
        m = pat.match(s)
        if m:
            return tmpl.format(*m.groups())
    return s


def define(key):
    """名詞定義（內部名稱 → 目前語言的說明）；沒有定義回傳空字串。"""
    d = GLOSSARY.get(key)
    if not d:
        return ""
    return d[1] if LANG == "en" else d[0]


def _tr_labels(idx):
    """翻譯 Index／MultiIndex 的標籤與名稱。"""
    import pandas as pd
    if isinstance(idx, pd.MultiIndex):
        out = pd.MultiIndex.from_tuples([tuple(T(x) for x in t) for t in idx], names=[T(n) for n in idx.names])
        return out
    vals = [T(x) for x in idx]
    try:
        out = pd.Index(vals, name=T(idx.name))
    except Exception:
        out = idx
    return out


def tr_df(df):
    """回傳翻譯過欄名、索引與文字儲存格的副本（中文模式直接回傳原物件）。"""
    if df is None:
        return df
    import pandas as pd
    if isinstance(df, pd.Series):
        s = df.copy()
        s.index = _tr_labels(s.index)
        s.name = T(s.name)
        return s
    cols = []
    for i in range(df.shape[1]):
        col = df.iloc[:, i]
        if (pd.api.types.is_object_dtype(col.dtype) or pd.api.types.is_string_dtype(col.dtype)
                or isinstance(col.dtype, pd.CategoricalDtype)):
            col = col.astype(object).map(lambda v: T(v) if isinstance(v, str) else v)
        cols.append(col.reset_index(drop=True))
    d = pd.concat(cols, axis=1) if cols else df.copy().reset_index(drop=True)
    d.columns = _tr_labels(df.columns)
    d.index = _tr_labels(df.index)
    return d


def bullet():
    return "• " if LANG == "en" else "・"


def list_sep():
    return ", " if LANG == "en" else "、"


# ---------------------------------------------------------------- 樣式化字串
PATTERNS = [
    (re.compile(r"^未來(\d+)週需求\(平均\)$"), "Next {0} Wks Demand (Avg)"),
    (re.compile(r"^未來(\d+)週需求\(高峰\)$"), "Next {0} Wks Demand (Peak)"),
    (re.compile(r"^第(\d+)名$"), "#{0}"),
]

# ---------------------------------------------------------------- 中文白話顯示名稱（只影響畫面與匯出，不影響內部欄位）
ZH = {
    "週": "週（週一）",
    "在途": "未結案",
    "在途件數": "未結案件數",
    "在途占比": "未結案占比",
    "占在途比例": "占未結案比例",
    "缺料（Awaiting Spares）": "等零件中",
    "缺料": "等零件",
    "缺料案件": "等零件的案件",
    "缺料件數": "等零件件數",
    "缺料占在途": "等零件占未結案",
    "缺料占在途比例": "等零件占未結案比例",
    "待料": "等零件（含運送中）",
    "待收壞件": "等收回壞零件",
    "待出貨": "修好待寄回",
    "待出貨案件": "修好待寄回",
    "待出貨超過 7 天": "修好超過 7 天未寄回",
    "待出貨件數": "待寄回件數",
    "待出貨積壓（依據點）": "修好待寄回的案件（依維修站）",
    "匯出待出貨積壓": "匯出待寄回清單",
    "保固 30 天內到期在途": "保固 30 天內到期（未結案）",
    "零件未使用率": "零件領了沒用的比例",
    "零件在途／已配案件": "零件已在路上／已配到的案件",
    "零件在途或已配": "零件已在路上或已配到",
    "零件在途或已配件數": "零件已在路上或已配到件數",
    "已配貨": "已配到零件",
    "零件已配貨（顆）": "已配到的零件（顆）",
    "準交率": "準時到貨率",
    "準交": "準時到貨",
    "準交筆數": "準時到貨筆數",
    "供應商準交率": "各供應商準時到貨率",
    "催料清單：等待天數 ≥": "追零件清單：已等待天數 ≥",
    "匯出催料清單": "匯出追零件清單",
    "催料案件清單": "需要追零件的案件",
    "缺料料號（Awaiting Spares 案件中待供貨的料號）": "缺貨料號（等零件的案件中還沒到貨的料號）",
    "類型分組": "案件類型",
    "狀態分類": "狀態分組",
    "據點數": "維修站數",
    "Top10據點占比": "前 10 大維修站占比",
    "前兩據點占比": "前 2 大維修站占比",
    "Top10據點在途占比": "前 10 大維修站占未結案比例",
    "停留中位數": "停留天數中位數",
    "停留P90": "停留天數（90% 案件在內）",
    "P90": "90% 案件在內（天）",
    "TAT_P90": "TAT（90% 案件在內）",
    "超過14天": "超過 14 天",
    "超過14天比例": "超過 14 天比例",
    "超過14天件數": "超過 14 天件數",
    "零件DOA": "零件到貨即壞（DOA）",
    "零件DOA率": "零件到貨即壞率",
    "保固外比率": "保固外比例",
    "在途_依狀態分類": "未結案_依狀態分組",
    "在途_類型×狀態": "未結案_類型×狀態",
    "在途_Top10據點×狀態": "未結案_前10維修站×狀態",
    "在途_停留天數": "未結案_停留天數",
    "Top10據點": "前10大維修站",
    "SLA_Top10據點": "SLA_前10大維修站",
    # AIO 原始欄位
    "CaseID": "案件編號",
    "Location": "維修站",
    "Type": "案件類型",
    "Status": "狀態",
    "ProductModel": "機型",
    "SerialNumber": "序號",
    "CurrentStatusDays": "目前狀態已停留天數",
    "DaysSinceRcvd": "收件至今天數",
    "WarrantyStatus": "保固狀態",
    "WarrantyExpiryDate": "保固到期日",
    "PartsRequested": "申請零件明細",
    "TATInNetworkDays": "TAT（工作天）",
}

# ---------------------------------------------------------------- 名詞定義（滑鼠停在 KPI 或欄名上顯示，也列在「名詞說明」）
GLOSSARY = {
    "在途": ("狀態不是 Closed 或 Rejected 的案件，也就是還沒結案的案子。",
             "Cases whose status isn't Closed or Rejected, meaning they're still being worked on."),
    "在途件數": ("目前還沒結案（狀態不是 Closed／Rejected）的案件數。", "Number of cases not yet closed (status isn't Closed or Rejected)."),
    "案件數": ("分析期間內建立的案件數。", "Number of cases created during the analysis period."),
    "已結案": ("狀態為 Closed 或 Rejected 的案件。", "Cases with a status of Closed or Rejected."),
    "缺料（Awaiting Spares）": ("狀態為 Awaiting Spares 的案件：零件缺貨，在等供應商或總倉出貨。括號內是占未結案的比例。",
                              "Cases in Awaiting Spares status: the part is out of stock and waiting on the supplier or central warehouse. The percentage is the share of open cases."),
    "缺料": ("狀態為 Awaiting Spares 的案件。", "Cases in Awaiting Spares status."),
    "缺料案件": ("狀態為 Awaiting Spares 的案件數；括號內是占未結案的比例。", "Number of Awaiting Spares cases; the percentage is the share of open cases."),
    "零件未使用率": ("已結案（Closed）案件中，申請的零件最後沒用到、原封退回的比例＝未使用 ÷（已使用＋未使用＋零件DOA），包含所有案件類型。",
                   "On closed cases, the share of requested parts that came back unused = unused ÷ (used + unused + part DOA), all case types."),
    "未使用率": ("沒用到、原封退回的零件 ÷（已使用＋未使用＋零件DOA）。", "Parts returned unused ÷ (used + unused + part DOA)."),
    "SLA 達成率": ("已結案（Closed）案件中，TAT 在 5 個工作天內的比例。", "Share of closed cases with a TAT of 5 business days or less."),
    "達成率": ("TAT 在 5 個工作天內的已結案案件比例。", "Share of closed cases with a TAT of 5 business days or less."),
    "整體達成率": ("所有已結案（Closed）案件中，TAT 在 5 個工作天內的比例。", "Share of all closed cases with a TAT of 5 business days or less."),
    "保固 30 天內到期在途": ("還沒結案、且保固到期日落在資料日期之後 30 天內的案件。", "Open cases whose warranty ends within 30 days after the data as-of date."),
    "30天內到期件數": ("還沒結案、且保固到期日落在資料日期之後 30 天內的案件數。", "Open cases whose warranty ends within 30 days after the data as-of date."),
    "TAT": ("Turnaround Time，用 TATInNetworkDays：收件到維修完成的工作天數（不含假日）。",
            "Turnaround time, from TATInNetworkDays: business days from receipt to repair complete (excludes weekends/holidays)."),
    "TATInNetworkDays": ("收件到維修完成的工作天數（不含假日）。", "Business days from receipt to repair complete."),
    "TAT區間": ("TAT 每兩個工作天分一格。", "TAT grouped into two-business-day buckets."),
    "中位數": ("把數字由小到大排好，正中間那一個；代表一半的案件在這個數字以內。", "The middle value: half the cases are at or below this number."),
    "P90": ("90% 的案件在這個數字以內，只有 10% 比它久；用來看最慢的那一群。", "90% of cases are at or below this number; only 10% take longer. Shows the slow tail."),
    "完整週": ("週一到週日完整落在分析期間內、至少 6 天有建案、且週量不低於一般週 70% 的週；用來避免半週資料拉低平均。",
             "A Monday–Sunday week fully inside the period, with cases created on at least 6 days and volume at least 70% of a typical week. Keeps partial weeks from skewing averages."),
    "類型分組": ("案件類型；Hardware、On-Site、DOA、Web Request、Component Swap 以外歸為 Other。",
             "Case type; anything other than Hardware, On-Site, DOA, Web Request, or Component Swap is grouped as Other."),
    "狀態分類": ("把 AIO 的細項狀態歸成六組：維修中、等零件、等客戶、等收回壞零件、修好待寄回、已結案。",
             "AIO statuses grouped into six: in repair, awaiting parts, awaiting customer, awaiting defective return, ready to ship, closed."),
    "待料": ("狀態為 Awaiting Spares、Part In Transit、Allocated 等，跟零件有關的等待。", "Statuses like Awaiting Spares, Part In Transit, and Allocated: waiting on parts."),
    "待客戶": ("等客戶核准報價或回覆，例如 Awaiting Customer Approval。", "Waiting on the customer, e.g. Awaiting Customer Approval."),
    "待收壞件": ("等客戶或據點寄回故障零件（Awaiting Defective Parts）。", "Waiting for the defective part to be returned (Awaiting Defective Parts)."),
    "待出貨": ("已修好、還沒寄回客戶（Ready To Ship、Repair Complete）。", "Repaired but not yet shipped back (Ready To Ship, Repair Complete)."),
    "維修中": ("已收件、正在檢修或測試等。", "Received and being diagnosed, repaired, or tested."),
    "DaysSinceRcvd": ("收件到資料日期為止的天數（日曆天）。", "Calendar days from receipt to the data as-of date."),
    "CurrentStatusDays": ("停在目前這個狀態的天數。", "Days the case has been in its current status."),
    "停留中位數": ("一半的案件在這個天數以內。", "Half the cases are at or below this many days."),
    "等待中位數": ("等零件的案件中，一半在這個天數以內。", "Half of the cases waiting on parts are at or below this many days."),
    "等待超過兩週": ("等零件已超過 14 天的案件數；括號內是占所有等零件案件的比例。", "Cases waiting on parts for 14+ days; the percentage is their share of all such cases."),
    "零件在途／已配案件": ("零件已經配到或在運送中（Part In Transit、Allocated 等）的未結案案件，不算缺料。",
                     "Open cases whose parts are allocated or on the way (Part In Transit, Allocated, etc.). Not counted as shortages."),
    "待供貨料號數": ("等零件的案件裡，零件狀態為「已申請待供貨」（R）的不同料號數。", "Distinct part numbers with status R (requested, awaiting supply) on Awaiting Spares cases."),
    "零件狀態": ("零件明細的狀態碼：C 已使用、U 未使用退回、D 零件DOA、W WPB、R 已申請待供貨、I 運送中、A 已配貨、H 保留中。",
             "Part status code: C used, U returned unused, D part DOA, W WPB, R requested/awaiting supply, I in transit, A allocated, H on hold."),
    "零件DOA": ("零件本身到貨就是壞的（Dead On Arrival）。", "The replacement part itself was dead on arrival."),
    "DOA": ("Dead On Arrival：新機開箱或零件到貨就故障。", "Dead on arrival: a new unit or part that fails out of the box."),
    "WPB": ("AIO 欄位 NumberOfWPBParts 的零件數。", "Part count from the AIO field NumberOfWPBParts."),
    "保固外": ("WarrantyStatus 為 Out Of Warranty 的案件（實際以保固外處理）。", "Cases with WarrantyStatus = Out Of Warranty (handled as paid repair)."),
    "保固外比率": ("保固外案件 ÷ 案件數。", "Out-of-warranty cases ÷ all cases."),
    "重複率": ("同一序號在期間內有 2 件以上案件，這些案件占全部的比例。", "Share of cases whose serial number has 2+ cases in the period."),
    "重複維修": ("同一序號在期間內有 2 件以上案件。", "A serial number with 2+ cases in the period."),
    "同序號案件數": ("同一序號在期間內的案件數。", "Number of cases for the same serial number in the period."),
    "準交率": ("已到貨的訂單中，實際到貨日不晚於承諾交期的比例。", "Of received POs, the share that arrived on or before the promised date."),
    "逾期天數": ("實際到貨日（未到貨以今天計）晚於承諾交期的天數。", "Days past the promised date (today is used if not yet received)."),
    "距到期天數": ("保固到期日減資料日期；負數表示已過期。", "Warranty expiry minus the data as-of date; negative means already expired."),
    "週": ("該週的週一日期；只列完整週（7 天中至少 6 天有建案、且量不低於一般週 70%）。",
           "The Monday of each week. Only full weeks are listed (cases on 6+ days and volume at least 70% of a typical week)."),
    "完整週合計": ("所有完整週的案件數加總。", "Total cases across all full weeks."),
    "累計占比": ("由大到小排序後，累加到這一列的占比。", "Running share after sorting from largest to smallest."),
    "收件→維修完成": ("收件到維修完成的平均日曆天。", "Average calendar days from receipt to repair complete."),
    "維修完成→出貨": ("維修完成到寄出的平均日曆天。", "Average calendar days from repair complete to shipped."),
    "出貨→結案": ("寄出到結案的平均日曆天。", "Average calendar days from shipped to closed."),
    "收件→結案": ("收件到結案的平均日曆天，最接近客戶感受的時間。", "Average calendar days from receipt to closing; closest to what the customer feels."),
}

# ---------------------------------------------------------------- 字典（中文 → English）
EN = {
    "週": "Week of",
    # ── 程式名稱、分頁
    "售後零件規劃平台": "After-Sales Parts Planner",
    "RMA 案件分析引擎": "RMA Case Analysis Engine",
    "RMA 案件資料模型與分析報告": "RMA Case Data Model & Analysis Report",
    "資料來源": "Data Source",
    "需求規劃": "Demand Planning",
    "庫存與缺料": "Inventory & Shortages",
    "訂單與交期": "Orders & Lead Times",
    "出貨與到貨": "Shipping & Receiving",
    "報表與分析": "Reports & Analysis",
    "跨單位協調": "Cross-Team Follow-ups",

    # ── 資料來源分頁
    "AIO 匯出檔：": "AIO export file:",
    "瀏覽…": "Browse…",
    "分析起日（留空＝資料最後日期的前一個月 1 日）：": "Analysis start date (blank = 1st of the month before the last data date):",
    "載入並分析": "Load & Analyze",
    "案件數": "Cases",
    "在途": "Open",
    "缺料（Awaiting Spares）": "Waiting on Parts",
    "零件未使用率": "Unused Parts Rate",
    "SLA 達成率": "SLA Attainment",
    "保固 30 天內到期在途": "Open, Warranty ≤30 Days",
    "1. 選擇 AIO 匯出檔（.xlsx，工作表 AllInOneData）\n2. 按「載入並分析」，約 10–30 秒\n3. 到各分頁查看，或在「報表與分析」匯出 Excel／PPT":
        "1. Pick the AIO export file (.xlsx, sheet AllInOneData)\n2. Click \"Load & Analyze\" (takes about 10–30 seconds)\n3. Browse the tabs, or export Excel/PowerPoint from \"Reports & Analysis\"",
    "尚未載入資料": "No data loaded",
    "載入中…": "Loading…",
    "載入失敗": "Load failed",
    "完成。": "Done.",
    "錯誤：": "Error: ",
    "請先選擇 AIO 匯出檔": "Pick an AIO export file first.",
    "請先在「資料來源」載入 AIO 匯出檔": "Load an AIO export file on the Data Source tab first.",

    # ── 表格元件
    "訊息": "Message",
    "（沒有資料）": "(No data)",
    "項目": "Item",
    "目前沒有可匯出的表格": "There's no table to export right now.",
    "表格.xlsx": "Table.xlsx",

    # ── 需求規劃
    "彙總依：": "Group by:",
    "預測未來（週）：": "Forecast ahead (weeks):",
    "重新計算": "Recalculate",
    "匯出上表": "Export Weekly Table",
    "匯出料號表": "Export Part List",
    "週申請零件量.xlsx": "Weekly_Parts_Requested.xlsx",
    "Top料號.xlsx": "Top_Parts.xlsx",
    "每週申請零件顆數（完整週）與未來需求推估": "Parts requested per week (full weeks) and demand forecast",
    "申請次數最多的料號（Top 30）：申請顆數、未使用率、目前待供貨顆數": "Most-requested part numbers (top 30): parts requested, unused rate, parts awaiting supply",

    # ── 庫存與缺料
    "缺料案件": "Shortage Cases",
    "等待超過兩週": "Waiting Over 2 Weeks",
    "等待中位數（天）": "Median Wait (Days)",
    "零件在途／已配案件": "Cases with Parts In Transit/Allocated",
    "待供貨料號數": "Part Numbers Awaiting Supply",
    "催料清單：等待天數 ≥": "Expedite list: days waiting ≥",
    "重新整理": "Refresh",
    "匯出催料清單": "Export Expedite List",
    "匯出料號清單": "Export Part Numbers",
    "催料清單.xlsx": "Expedite_List.xlsx",
    "缺料料號.xlsx": "Shortage_Parts.xlsx",
    "缺料料號（Awaiting Spares 案件中待供貨的料號）": "Shortage part numbers (parts awaiting supply on Awaiting Spares cases)",
    "催料案件清單": "Cases to expedite",

    # ── 訂單與交期
    "AIO 資料沒有採購訂單，這個模組用匯入的訂單檔計算準交率與逾期清單。":
        "The AIO data has no purchase orders, so this module uses an imported PO file to calculate on-time delivery and the overdue list.",
    "產生訂單範本": "Create PO Template",
    "匯入訂單檔…": "Import PO File…",
    "匯出逾期清單": "Export Overdue List",
    "逾期訂單.xlsx": "Overdue_POs.xlsx",
    "零件訂單範本.xlsx": "Parts_PO_Template.xlsx",
    "訂單筆數": "POs",
    "已到貨": "Received",
    "準交率": "On-Time Rate",
    "平均交期（天）": "Avg Lead Time (Days)",
    "未到貨且已逾期": "Not Received & Overdue",
    "供應商準交率": "Supplier on-time delivery",
    "逾期與未交訂單": "Late and outstanding POs",
    "訂單號": "PO Number",
    "料號": "Part No.",
    "供應商": "Supplier",
    "數量": "Qty",
    "下單日": "Order Date",
    "承諾交期": "Promised Date",
    "實際到貨日": "Received Date",
    "目的國": "Destination Country",
    "供應商A": "Supplier A",
    "供應商B": "Supplier B",
    "缺少欄位：": "Missing columns: ",
    "準交": "On Time",
    "交期天數": "Lead Time (Days)",
    "逾期天數": "Days Late",
    "準交筆數": "On-Time POs",
    "平均交期天數": "Avg Lead Time (Days)",
    "平均逾期天數": "Avg Days Late",

    # ── 出貨與到貨
    "待出貨案件": "Ready-to-Ship Cases",
    "待出貨超過 7 天": "Ready to Ship Over 7 Days",
    "零件運送中（顆）": "Parts In Transit (Units)",
    "零件已配貨（顆）": "Parts Allocated (Units)",
    "保固到期清單：天數內": "Warranty expiry list: within days",
    "匯出到期清單": "Export Expiry List",
    "匯出待出貨積壓": "Export Ship Backlog",
    "保固到期在途案.xlsx": "Open_Cases_Warranty_Expiring.xlsx",
    "待出貨積壓.xlsx": "Ship_Backlog.xlsx",
    "待出貨積壓（依據點）": "Ready-to-ship backlog (by site)",
    "零件運送中／已配貨（依國家，在途案件）": "Parts in transit / allocated (by country, open cases)",
    "在途且保固即將到期（含建案時保固內、現已過期）": "Open cases with warranty ending soon (incl. in warranty at creation, now expired)",

    # ── 報表與分析
    "分析項目": "Analyses",
    "匯出 Excel 表格": "Export Excel Tables",
    "匯出 PPT 報告": "Export PowerPoint Report",
    "匯出全部圖 PNG": "Export All Charts (PNG)",
    "全部匯出到資料夾": "Export Everything to Folder",
    "請先載入資料，再從左側選一項分析": "Load data first, then pick an analysis on the left",
    "（圖表）": "(Chart)",
    "表格：": "Table:",
    "匯出此表": "Export This Table",
    "分析表格.xlsx": "Analysis_Table.xlsx",
    "結論：": "Takeaway: ",
    "RMA分析結果.xlsx": "RMA_Analysis_Results.xlsx",
    "RMA分析報告.pptx": "RMA_Analysis_Report.pptx",
    "RMA_輸出": "RMA_Output",

    # ── 跨單位協調
    "新增／更新待辦": "Add / Update Follow-up",
    "對象單位": "Team",
    "議題": "Topic",
    "負責人": "Owner",
    "到期日": "Due Date",
    "狀態": "Status",
    "建立日": "Created",
    "進行中": "In Progress",
    "待回覆": "Awaiting Reply",
    "已完成": "Done",
    "取消": "Cancelled",
    "新增": "Add",
    "更新所選": "Update Selected",
    "刪除所選": "Delete Selected",
    "由分析結論產生待辦": "Create Follow-ups from Takeaways",
    "匯出待辦": "Export Follow-ups",
    "跨單位協調待辦.xlsx": "Cross_Team_Follow_ups.xlsx",
    "請填議題": "Enter a topic.",
    "各國規劃人員": "Country Planners",
    "總部服務團隊": "HQ Service Team",
    "倉庫／物流": "Warehouse / Logistics",
    "維修據點": "Repair Sites",

    # ── 匯出檔內容
    "說明": "About",
    "KPI總表": "KPI Summary",
    "投影片文字": "Slide Text",
    "分析期間起日": "Analysis Start Date",
    "資料日期": "Data As-of Date",
    "產出時間": "Generated At",
    "值": "Value",
    "分析": "Analysis",
    "名稱": "Name",
    "指標": "Metric",
    "數值": "Value",
    "重點": "Key Points",
    "段落": "Narrative",
    "結論": "Takeaway",
    "分析框架": "Analysis Framework",
    "找不到標題列（第一欄應有 CaseID）": "Header row not found (the first column should contain CaseID)",
    "讀檔…": "Reading file…",
    "完成：": "Done:",
    "AIO 匯出檔 (.xlsx)": "AIO export file (.xlsx)",
    "輸出資料夾": "Output folder",
    "分析起日 YYYY-MM-DD（預設：資料最後日期的前一個月 1 日）": "Analysis start date YYYY-MM-DD (default: 1st of the month before the last data date)",
    "資料基準日 YYYY-MM-DD（預設：最後建案日）": "As-of date YYYY-MM-DD (default: last case creation date)",

    # ── 分析標題與副標
    "營運量與負荷｜案件量趨勢": "Volume & Workload | Case Volume Trend",
    "每週案件量、Top 10 據點、國家與類型占比": "Weekly case volume, top 10 sites, and mix by country and case type",
    "營運量與負荷｜在途案件數與 aging": "Volume & Workload | Open Cases & Aging",
    "在途案件的狀態分類、類型組成與停留天數": "Open cases by status group, case type, and days open",
    "營運量與負荷｜內外部客戶占比": "Volume & Workload | Internal vs. External Customers",
    "內部與外部客戶的案件類型組成": "Case-type mix for internal vs. external customers",
    "時效與 SLA｜TAT 分布": "Turnaround & SLA | TAT Distribution",
    "已結案 TATInNetworkDays（收件至維修完成，工作天）與各段耗時":
        "TATInNetworkDays on closed cases (received to repair complete, business days) and time by stage",
    "時效與 SLA｜SLA 達成率": "Turnaround & SLA | SLA Attainment",
    "品質與故障｜故障代碼 Pareto": "Quality & Failures | Failure Code Pareto",
    "RCFailureCodes 前 20 名與累計占比、症狀歸類": "Top 20 RCFailureCodes, cumulative share, and symptom groups",
    "品質與故障｜重複維修": "Quality & Failures | Repeat Repairs",
    "同一序號在期間內出現 2 件以上": "Serial numbers with 2+ cases in the period",
    "零件與成本｜零件使用結果": "Parts & Cost | Parts Usage Outcomes",
    "已結案（Closed）案件的零件去向：使用／未使用／DOA／WPB": "Where parts ended up on closed cases: used, unused, DOA, WPB",
    "零件與成本｜缺料與延誤": "Parts & Cost | Parts Shortages & Delays",
    "在途案件中 Awaiting Spares 的等待天數（CurrentStatusDays）": "Days waiting on open Awaiting Spares cases (CurrentStatusDays)",
    "保固與財務｜保固內外比例": "Warranty & Finance | In vs. Out of Warranty",
    "WarrantyStatus（實際處理狀態）依國家、類型、產品線": "WarrantyStatus (how the case was handled) by country, case type, and product line",
    "保固與財務｜保固將到期的在途案": "Warranty & Finance | Open Cases Nearing Warranty Expiry",
    "內部客戶的案件量可以併入整體看，但 DOA 內容值得單獨追蹤，當作新機品質的早期警訊。":
        "Internal-customer volume can be rolled into the overall numbers, but their DOA cases deserve their own tracking as an early warning on new-unit quality.",
    "品質改善與零件備料都可以聚焦：盯住前 20 個代碼就覆蓋六成案件，開機、顯示、電源三類是主軸；備料清單以這三類對應的零件為優先。":
        "Both quality work and parts stocking can stay focused: the top 20 codes cover about 60% of cases, and boot, display, and power are the core groups. Prioritize the parts tied to those three groups on the stocking list.",
    "意義：內部通報集中在新機開箱不良，是品質面的早期訊號": "What it means: internal reports center on out-of-box failures on new units, an early quality signal",
    "各國高峰都在 0–1 天": "Every country peaks at 0–1 days",

    # ── 表格名稱
    "週案件量_依類型": "Weekly Volume by Type",
    "週案件量_依國家": "Weekly Volume by Country",
    "Top10據點": "Top 10 Sites",
    "類型占比": "Case Type Mix",
    "國家占比": "Country Mix",
    "在途_依狀態分類": "Open by Status Group",
    "在途_類型×狀態": "Open Type x Status",
    "在途_Top10據點×狀態": "Open Top 10 Sites x Status",
    "在途_停留天數": "Open Case Aging",
    "內外部×類型_件數": "Int vs Ext by Type (Count)",
    "內外部×類型_占比": "Int vs Ext by Type (Share)",
    "TAT分布_依國家(占比)": "TAT Dist by Country (Share)",
    "TAT分布_依國家(件數)": "TAT Dist by Country (Count)",
    "TAT統計_依國家": "TAT Stats by Country",
    "TAT統計_依類型": "TAT Stats by Type",
    "各段耗時_依類型(日曆天)": "Stage Days by Type (Cal.)",
    "SLA_依國家": "SLA by Country",
    "SLA_依類型": "SLA by Type",
    "SLA_Top10據點": "SLA Top 10 Sites",
    "Top20故障代碼": "Top 20 Failure Codes",
    "症狀歸類": "Symptom Groups",
    "症狀歸類_明細": "Symptom Groups Detail",
    "各類型前五代碼": "Top 5 Codes by Type",
    "重複次數分布": "Repeat Count Distribution",
    "重複率_依國家": "Repeat Rate by Country",
    "重複率_依類型": "Repeat Rate by Type",
    "後續案件與前案間隔": "Gap to Previous Case",
    "零件去向_依類型": "Parts Outcome by Type",
    "零件去向_依國家": "Parts Outcome by Country",
    "On-Site未使用率_依國家": "On-Site Unused by Country",
    "每案平均申請顆數": "Avg Parts per Case",
    "缺料等待天數_依國家": "Shortage Wait by Country",
    "缺料占在途比例_依國家": "Shortage Share by Country",
    "缺料案類型組成_依國家": "Shortage Mix by Country",
    "保固內外_依國家": "Warranty by Country",
    "保固內外_依類型": "Warranty by Type",
    "保固內外_依產品線": "Warranty by Product Line",
    "保固狀態vs保固分類": "Status vs Classification",
    "保固內外_全部國家": "Warranty All Countries",
    "保固到期區間_依國家": "Expiry Window by Country",
    "保固到期區間_合計": "Expiry Window Total",
    "已過期在途_保固狀態": "Expired Open by Status",
    "建案時保固內現已過期_依國家": "Lapsed In-Warr by Country",
    "30天內到期_依國家": "Expiring in 30d by Country",

    # ── 欄名、索引、分類值
    "國家": "Country",
    "狀態分類": "Status Group",
    "已結案": "Closed",
    "建案週": "Week Created",
    "完整週": "Full Week",
    "類型分組": "Case Type",
    "同序號案件數": "Cases per Serial",
    "重複維修": "Repeat Repair",
    "SLA達成": "SLA Met",
    "TAT區間": "TAT Band",
    "症狀類別": "Symptom Group",
    "維修天數": "Repair Days",
    "出貨等待": "Ship Wait",
    "結案等待": "Close Wait",
    "收件至結案": "Received to Closed",
    "保固到期區間": "Warranty Expiry Window",
    "等待天數區間": "Wait Band",
    "缺料": "Shortage",
    "零件在途或已配": "Parts In Transit/Allocated",
    "維修中": "In Repair",
    "待料": "Awaiting Parts",
    "待客戶": "Awaiting Customer",
    "待收壞件": "Awaiting Defective Return",
    "待出貨": "Ready to Ship",
    "其他": "Other",
    "未填": "Not Specified",
    "開機": "Boot",
    "顯示": "Display",
    "電源": "Power",
    "已使用": "Used",
    "未使用退回": "Returned Unused",
    "零件DOA": "Part DOA",
    "已申請待供貨": "Requested, Awaiting Supply",
    "運送中": "In Transit",
    "已配貨": "Allocated",
    "保留中": "On Hold",
    "序": "Seq",
    "狀態碼": "Status Code",
    "零件狀態": "Part Status",
    "料號前綴": "Part Prefix",
    "合計": "Total",
    "占比": "Share",
    "累計占比": "Cumulative Share",
    "占在途比例": "Share of Open",
    "中位數": "Median",
    "平均": "Mean",
    "超過14天": "Over 14 Days",
    "超過14天比例": "Over 14 Days %",
    "內部客戶": "Internal",
    "外部客戶": "External",
    "一天內比例": "Within 1 Day %",
    "超過5天比例": "Over 5 Days %",
    "收件→維修完成": "Received → Repaired",
    "維修完成→出貨": "Repaired → Shipped",
    "出貨→結案": "Shipped → Closed",
    "收件→結案": "Received → Closed",
    "全部": "All",
    "達成率": "Attainment",
    "未達成件數": "Missed",
    "三天內達成率": "Within 3 Days %",
    "機器數": "Machines",
    "重複案件": "Repeat Cases",
    "重複率": "Repeat Rate",
    "前案狀態": "Prior Case Status",
    "前案結案日": "Prior Case Closed",
    "前案代碼": "Prior Case Code",
    "間隔天數": "Gap (Days)",
    "後續案件數": "Follow-up Cases",
    "同日重開": "Same-Day Reopen",
    "1-7天": "1–7 Days",
    "8-14天": "8–14 Days",
    "15-30天": "15–30 Days",
    ">30天": ">30 Days",
    "未使用": "Unused",
    "已使用率": "Used %",
    "未使用率": "Unused Rate",
    "零件DOA率": "Part DOA Rate",
    "全部合計": "All Types Total",
    "超過兩週": "Over 2 Weeks",
    "超過兩週比例": "Over 2 Weeks %",
    "等待中位數": "Median Wait",
    "缺料占在途比例": "Shortage % of Open",
    "保固外": "Out of Warranty",
    "保固外比率": "OOW Rate",
    "占保固外總數": "Share of All OOW",
    "1 已過期": "1 Expired",
    "2 30天內": "2 Within 30 Days",
    "3 31-90天": "3 31–90 Days",
    "4 90天以上": "4 Over 90 Days",
    "已過期比例": "Expired %",
    "件數": "Cases",
    "天": "Days",

    # ── 零件規劃用表
    "申請顆數": "Parts Requested",
    "待供貨": "Awaiting Supply",
    "主要國家": "Top Country",
    "等待案件數": "Cases Waiting",
    "待供貨顆數": "Parts Awaiting Supply",
    "最長等待天數": "Max Days Waiting",
    "平均等待天數": "Avg Days Waiting",
    "待出貨件數": "Ready-to-Ship Cases",
    "停留中位數": "Median Days",
    "超過7天": "Over 7 Days",
    "距到期天數": "Days to Expiry",
    "到期狀態": "Expiry Status",
    "已過期（建案時保固內）": "Expired (In Warranty at Creation)",
    "即將到期": "Expiring Soon",
    "最近每週平均": "Recent Weekly Avg",
    "最近每週最高": "Recent Weekly Peak",

    # ── KPI 名稱
    "期間案件數": "Cases in Period",
    "據點數": "Sites",
    "國家數": "Countries",
    "完整週數": "Full Weeks",
    "完整週合計": "Full-Week Total",
    "Top10據點占比": "Top 10 Sites Share",
    "前六國占比": "Top 6 Countries Share",
    "前兩據點占比": "Top 2 Sites Share",
    "在途件數": "Open Cases",
    "在途占比": "Open Share",
    "最大狀態分類": "Largest Status Group",
    "最大狀態分類件數": "Largest Status Group Cases",
    "Top10據點在途占比": "Top 10 Sites Share of Open",
    "停留P90": "P90 Days Open",
    "超過14天件數": "Over 14 Days (Cases)",
    "內部客戶件數": "Internal Cases",
    "內部客戶占比": "Internal Share",
    "已結案件數": "Closed Cases",
    "TAT平均": "TAT Mean",
    "TAT中位數": "TAT Median",
    "超過7天比例": "Over 7 Days %",
    "收件至結案平均": "Avg Received to Closed",
    "出貨至結案平均": "Avg Shipped to Closed",
    "整體達成率": "Overall Attainment",
    "有代碼件數": "Cases with Code",
    "未填件數": "Cases Without Code",
    "代碼種類": "Distinct Codes",
    "前3名占比": "Top 3 Share",
    "前10名占比": "Top 10 Share",
    "前20名占比": "Top 20 Share",
    "達80%所需代碼數": "Codes to Reach 80%",
    "開機類件數": "Boot Cases",
    "顯示類件數": "Display Cases",
    "電源類件數": "Power Cases",
    "重複機器數": "Repeat Machines",
    "重複案件數": "Repeat Cases",
    "前案被拒或取消比例": "Prior Case Rejected/Cancelled %",
    "同日重開比例": "Same-Day Reopen %",
    "同故障代碼比例": "Same Failure Code %",
    "有申請零件件數": "Cases with Parts Requested",
    "零件總數": "Total Parts",
    "使用率": "Usage Rate",
    "未使用率最高類型": "Highest-Unused Type",
    "未使用率最高類型數值": "Highest Unused Rate",
    "缺料件數": "Shortage Cases",
    "缺料占在途": "Shortage % of Open",
    "零件在途或已配件數": "Cases with Parts In Transit/Allocated",
    "六天內比例": "Within 6 Days %",
    "超過兩週件數": "Over 2 Weeks (Cases)",
    "保固外件數": "OOW Cases",
    "保固外最高國家": "Highest-OOW Country",
    "保固外最高國家比率": "Highest-OOW Country Rate",
    "Hardware占保固外": "Hardware Share of OOW",
    "保固期內以保固外處理": "In Warranty, Handled as OOW",
    "在途保固外件數": "Open OOW Cases",
    "在途保固外待客戶": "Open OOW Awaiting Customer",
    "已過期件數": "Expired Cases",
    "30天內到期件數": "Expiring ≤30 Days",
    "31-90天到期件數": "Expiring 31–90 Days",
    "90天內合計": "Total ≤90 Days",
    "建案時保固內現已過期": "In Warranty at Creation, Now Expired",
    "建案時已保固外": "OOW at Creation",

    # ── AIO 原始欄位
    "CaseID": "Case ID",
    "Location": "Site",
    "Type": "Case Type",
    "ProductModel": "Model",
    "SerialNumber": "Serial No.",
    "CurrentStatusDays": "Days in Current Status",
    "DaysSinceRcvd": "Days Since Received",
    "WarrantyStatus": "Warranty Status",
    "WarrantyExpiryDate": "Warranty Expiry",
    "PartsRequested": "Requested Parts (Detail)",
    "TATInNetworkDays": "TAT (Business Days)",

    # ── 圖表標題
    "每週案件量（完整週，依類型）": "Weekly Case Volume (Full Weeks, by Type)",
    "Top 10 據點": "Top 10 Sites",
    "在途案件：類型 × 狀態分類": "Open Cases: Type × Status Group",
    "內部 vs 外部客戶：案件類型組成": "Internal vs. External Customers: Case-Type Mix",
    "已結案 TAT 分布（依國家，收件至維修完成，工作天）": "Closed-Case TAT Distribution (by Country, Received to Repaired, Business Days)",
    "各階段平均耗時（日曆天）": "Average Time by Stage (Calendar Days)",
    "故障代碼 Pareto（前 20 名，柱＝案件數，線＝累計占比）": "Failure Code Pareto (Top 20; Bars = Cases, Line = Cumulative Share)",
    "已結案案件的零件去向（依類型，顆數）": "Parts Outcome on Closed Cases (by Type, Units)",
    "缺料案件已等待天數（依國家）": "Shortage Cases: Days Waiting (by Country)",
    "缺料案件的類型組成（依國家）": "Shortage Cases: Case-Type Mix (by Country)",
    "保固內外比例（依國家）": "In vs. Out of Warranty (by Country)",
}
