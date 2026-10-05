# 售後零件規劃平台

兩個檔案，一套架構：

| 檔案 | 用途 |
|---|---|
| `rma_engine.py` | 分析引擎。讀 AIO 匯出檔、清理、輔助欄位、11 項分析、零件明細解析，匯出 Excel／PNG／PPTX。可獨立以指令執行。 |
| `零件規劃平台.py` | 桌面程式。六個分頁對應工作架構圖的六個模組，所有數字來自引擎。 |

## 安裝

Python 3.9 以上（Windows 內建 tkinter）。

```
pip install -r requirements.txt
```

## 執行

桌面程式：`python 零件規劃平台.py`（Windows 可直接雙擊 `執行平台.bat`）

指令列一次產出整份報告：

```
python rma_engine.py "0820-0921 AIO data.xlsx" 輸出資料夾
python rma_engine.py "0820-0921 AIO data.xlsx" 輸出資料夾 --start 2026-08-01
```

輸出資料夾會有 `RMA分析結果.xlsx`（每項分析的表格、KPI 總表、投影片文字）、`RMA分析報告.pptx`（每項分析一頁：重點、圖、結論，段落在備忘稿）、`charts/`（11 張 PNG）。

## 不用裝 Python：單一 exe

每次推送到 GitHub，`.github/workflows/build-exe.yml` 會在 GitHub 的 Windows 機器上自動打包成 `RMA_PartsPlanner.exe`：

- 到 repo 的 **Actions** 頁 → 點最新一次「Build Windows exe」→ 最下方 **Artifacts** 下載 `RMA_PartsPlanner-windows`。
- 推送 `v1.0` 這類標籤時，會另外建立 **Release**，exe 直接掛在 Releases 頁面，不用登入也能下載。

把 exe 複製到任何 Windows 電腦雙擊即可，第一次啟動約 10–20 秒。`協調待辦.json` 和 `_charts/` 會存在 exe 旁邊。
啟動失敗時，exe 旁邊會出現 `啟動錯誤.log`。

也可以在自己電腦打包：雙擊 `build_exe.bat`，完成後在 `dist/` 資料夾。

## 網頁版：任何裝置開瀏覽器就能用（Vercel／GitHub Pages）

`index.html` + `app.js` + `worker.js` + `web_glue.py` 是純靜態的網頁版：用 Pyodide 把 Python 搬進瀏覽器，
`rma_engine.py` 直接在使用者的瀏覽器裡執行，**AIO 檔案不會上傳到任何伺服器**。所以只要能放靜態檔案的地方都能架：Vercel、GitHub Pages、Netlify 都可以。

**部署到 Vercel（一次就好）**

1. 到 https://vercel.com 用 GitHub 帳號登入 → **Add New → Project** → 選 `sandyliu3056/RMA` → **Import**。
2. Framework Preset 選 **Other**，Build Command 留空，Output Directory 留空（根目錄就是網站），按 **Deploy**。
3. 完成後會給你 `https://rma-xxxx.vercel.app`，之後每次推到 GitHub 會自動更新。
4. 想限制只有自己能看：Project → Settings → **Deployment Protection** 開啟（需要 Vercel 帳號登入才能開）。

GitHub Pages 也一樣能用：同一份檔案已經由 `.github/workflows/static.yml` 發布到 https://sandyliu3056.github.io/RMA/ 。

**第一次開網頁**會下載瀏覽器版 Python 和 pandas 等套件（約 30 MB，之後瀏覽器會快取），啟動約 20–60 秒；載入 AIO 檔分析約 1 分鐘。手機也能開，但建議用電腦。

**本機試跑**：在這個資料夾執行 `python -m http.server 8000`，開 http://localhost:8000 。

**和桌面版的差別**

- 每次開網頁都要重新選 AIO 檔（檔案只在瀏覽器裡，關掉就沒了）。
- 「跨單位協調」的待辦存在那台電腦的瀏覽器裡；要帶到別台電腦，用「下載待辦」存成 JSON 再「還原待辦」。
- Excel／PPT／圖都是按鈕下載。

另外還有一個 Streamlit 版 `web_app.py`（需要能跑 Python 的主機，例如 Streamlit Community Cloud：`streamlit run web_app.py`），功能相同，給有伺服器的情況用。

## 資料規則

- 讀取工作表 `AllInOneData`（沒有時取第一張），自動找 `CaseID` 所在列當標題，略過中文說明列。
- 分析期間預設為資料最後建案日的前一個月 1 日起（9/21 的檔案 → 8/1 起）；可在畫面或 `--start` 指定。
- 「資料日期」＝最後建案日，停留天數與保固到期都以此為基準。
- 完整週：週一到週日都在期間內、至少 6 天有案件、週量不低於中位數的七成；週趨勢與需求預估只用完整週。
- 狀態分類：25 個系統狀態歸為 維修中／待料／待客戶／待收壞件／待出貨／已結案（對照在 `STATUS_MAP`）。
- TAT 採 `TATInNetworkDays`（收件至維修完成，工作天）；SLA 門檻 5 個工作天，只算 Status = Closed。
- 缺料＝Status = Awaiting Spares；Part In Transit、Allocated 等視為零件已配貨或在途，不算缺料。
- 零件去向只算 Status = Closed 的案件；WPB 極少，另列不計入合計。
- 零件明細由 `PartsRequested` 解析，狀態碼：C 已使用、U 未使用退回、D 零件 DOA、W WPB、R 已申請待供貨、I 運送中、A 已配貨、H 保留中。
- 國家名稱去掉方括號代碼；國家比較只列案件量前六名。

## 六個模組

| 分頁 | 內容 | 資料來源 |
|---|---|---|
| 資料來源 | 選檔、指定起日、載入並分析、六個總覽 KPI | AIO |
| 需求規劃 | 完整週申請零件顆數（依國家或類型）、未來 N 週需求推估、Top 30 料號（申請量、未使用率、待供貨） | AIO、零件明細 |
| 庫存與缺料 | 缺料 KPI、待供貨料號清單、可調天數門檻的催料案件清單 | AIO、零件明細 |
| 訂單與交期 | 匯入訂單檔算準交率、平均交期、逾期與未交清單；提供範本 | 自行匯入（AIO 無採購資料） |
| 出貨與到貨 | 待出貨積壓（依據點）、零件運送中／已配貨（依國家）、保固即將到期的在途案 | AIO、零件明細 |
| 報表與分析 | 11 項分析的重點、段落、結論、圖與表格；匯出 Excel／PPT／PNG | 引擎 |
| 跨單位協調 | 待辦追蹤（對象單位、議題、負責人、到期日、狀態），可由分析結論一鍵產生；存於 `協調待辦.json` | 本機 |

## 訂單檔格式

欄位：訂單號、料號、供應商、數量、下單日、承諾交期、實際到貨日（未到貨留空）、目的國（選填）。以「產生訂單範本」取得樣板。

## 常見問題

- 找不到標題列：確認匯出檔第一欄有 `CaseID`。
- 圖上中文變方塊：系統需有中文字型（Windows 的微軟正黑體即可）。
- 數字與人工樞紐略有差異：多半是期間起日或 Closed／Rejected 是否併入的差別，依上面的資料規則對齊即可。
