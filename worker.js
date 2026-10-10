/* Web Worker：在背景載入 Pyodide（瀏覽器版 Python），讀 Excel、跑分析，主畫面不會卡住。
 * 啟動分兩段：
 *   第一段（必要）：Pyodide + pandas + 分析引擎 → 送出 ready，就能「載入並分析」
 *   第二段（背景）：matplotlib、openpyxl、python-pptx、字型 → 圖表與匯出需要時才等它
 * 訊息格式：主畫面 → {id, cmd, args}；回覆 {id, ok, result | error}；進度 {type:"log", msg} / {type:"progress", pct, msg} / {type:"ready"} */
let pyodide = null;
let glue = null;
let config = null;
let M = null;          // 啟動訊息（依語言）
let extras = null;     // 第二段的 Promise

const MSG = {
  zh: {
    download: "下載瀏覽器版 Python（第一次約 15 MB，之後會快取）…",
    packages: "載入 pandas…",
    engine: "載入分析引擎…",
    ready: "就緒",
    extras: "背景載入圖表與匯出套件（matplotlib、openpyxl、python-pptx）…",
    extrasDone: "圖表與匯出套件已就緒。",
    extrasFail: "圖表與匯出套件載入失敗（需要時會再試一次）：",
    fontFail: "字型載入失敗：",
    fetchFail: "讀取 {0} 失敗（{1}）",
    reading: "讀取 Excel…",
    rows: "工作表共 {0} 列",
  },
  en: {
    download: "Downloading the browser Python runtime (about 15 MB the first time; cached afterwards)…",
    packages: "Loading pandas…",
    engine: "Loading the analysis engine…",
    ready: "Ready",
    extras: "Loading chart and export packages in the background (matplotlib, openpyxl, python-pptx)…",
    extrasDone: "Chart and export packages are ready.",
    extrasFail: "Chart and export packages failed to load (will retry when needed): ",
    fontFail: "Font failed to load: ",
    fetchFail: "Failed to read {0} ({1})",
    reading: "Reading the Excel file…",
    rows: "{0} rows in the worksheet",
  },
};
const fmt = (s, ...a) => s.replace(/\{(\d+)\}/g, (_, i) => a[+i]);

function log(msg) { self.postMessage({ type: "log", msg }); }
function progress(pct, msg) { self.postMessage({ type: "progress", pct, msg }); if (msg) log(msg); }

async function fetchText(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(fmt(M.fetchFail, url, r.status));
  return await r.text();
}

/* 相對路徑的 .whl 轉成完整網址（micropip 需要）；套件名稱原樣保留 */
function wheelUrl(w) {
  if (/\.whl$/i.test(w)) return new URL(w, self.location.href).href;
  return w;
}

async function boot(cfg) {
  config = cfg;
  M = MSG[cfg.lang === "en" ? "en" : "zh"];
  progress(3, M.download);
  importScripts(cfg.pyodideBase + "pyodide.js");
  importScripts(cfg.xlsxUrl);
  pyodide = await loadPyodide({ indexURL: cfg.pyodideBase });
  pyodide.setStdout({ batched: (s) => log(s) });
  pyodide.setStderr({ batched: (s) => { if (!/Glyph|UserWarning|warnings\.warn|font cache/.test(s)) log(s); } });
  progress(40, M.packages);
  let n = 0;
  await pyodide.loadPackage(["pandas", "numpy"], { messageCallback: (m) => { if (/^Loaded/.test(m)) progress(Math.min(78, 45 + 8 * ++n)); log(m); } });
  progress(80, M.engine);
  pyodide.FS.mkdirTree("/app/fonts");
  pyodide.FS.writeFile("/app/i18n.py", await fetchText(cfg.i18nUrl));
  pyodide.FS.writeFile("/app/rma_engine.py", await fetchText(cfg.engineUrl));
  pyodide.FS.writeFile("/app/web_glue.py", await fetchText(cfg.glueUrl));
  progress(90);
  await pyodide.runPythonAsync(`
import sys, warnings
warnings.filterwarnings("ignore")
sys.path.insert(0, "/app")
import web_glue
`);
  glue = pyodide.pyimport("web_glue");
  progress(100, M.ready);
  self.postMessage({ type: "ready" });
  ensureExtras().catch(() => {});
}

/* 第二段：圖表與匯出需要的套件。失敗時下次需要會再試。 */
function ensureExtras() {
  if (!extras) {
    extras = loadExtras().catch((e) => { extras = null; log(M.extrasFail + ((e && e.message) || e)); throw e; });
  }
  return extras;
}
async function loadExtras() {
  log(M.extras);
  await pyodide.loadPackage(["matplotlib", "lxml", "pillow", "micropip"], { messageCallback: (m) => { if (/^Loaded/.test(m)) log(m); } });
  const micropip = pyodide.pyimport("micropip");
  // openpyxl 與 python-pptx 不在 Pyodide 內建套件裡；預設用網站自帶的 wheels/（含相依套件，不連 PyPI）
  const wheels = (config.wheels || []).map(wheelUrl);
  const onlyFiles = wheels.length && wheels.every((w) => /\.whl$/i.test(w));
  const pyWheels = pyodide.toPy(wheels);
  try {
    if (onlyFiles) await micropip.install.callKwargs(pyWheels, { deps: false });
    else await micropip.install(pyWheels);
  } finally { pyWheels.destroy(); }
  for (const f of config.fonts || []) {
    try {
      const r = await fetch(f);
      if (r.ok) {
        const buf = new Uint8Array(await r.arrayBuffer());
        pyodide.FS.writeFile("/app/fonts/" + f.split("/").pop().split("?")[0], buf);
      }
    } catch (e) { log(M.fontFail + f); }
  }
  await pyodide.runPythonAsync(`
import os
import matplotlib
matplotlib.use("Agg")
from matplotlib import font_manager
for fn in os.listdir("/app/fonts"):
    try:
        font_manager.fontManager.addfont("/app/fonts/" + fn)
    except Exception as ex:
        print("font register failed", fn, ex)
`);
  log(M.extrasDone);
}

/* 用 SheetJS 把 Excel 讀成「列的陣列」。cellDates 讓日期變成 Date 物件，到 Python 端會變成 datetime。 */
function readSheet(buf, preferName) {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, dense: true });
  const name = preferName && wb.SheetNames.includes(preferName) ? preferName : wb.SheetNames[0];
  const ws = wb.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
  // Date 物件轉成「本地時間」字串，交給 pandas 解析（toISOString 會轉成 UTC，日期會差一天）
  const p = (n) => String(n).padStart(2, "0");
  const fmtD = (d) => `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  for (const r of rows) for (let i = 0; i < r.length; i++) if (r[i] instanceof Date) r[i] = isNaN(r[i]) ? null : fmtD(r[i]);
  return rows;
}

// 這些指令只需要第一段；其他（圖表、匯出、表格的 Excel）要等第二段
const LIGHT = new Set(["boot", "ui_lang", "load", "set_lang", "summary", "conclusions", "ui", "glossary"]);

self.onmessage = async (ev) => {
  const { id, cmd, args } = ev.data;
  try {
    let result;
    if (cmd === "boot") {
      await boot(args);
      result = true;
    } else if (cmd === "ui_lang") {
      // 啟動中切換語言：後續的啟動訊息改用新語言
      M = MSG[args && args.lang === "en" ? "en" : "zh"];
      result = true;
    } else if (cmd === "load") {
      log(M.reading);
      const rows = readSheet(args.buffer, "AllInOneData");
      log(fmt(M.rows, rows.length.toLocaleString()));
      const pyRows = pyodide.toPy(rows);
      try { result = JSON.parse(glue.load(pyRows, args.start || null, (m) => log(m))); }
      finally { pyRows.destroy(); }
    } else if (cmd === "po") {
      await ensureExtras();
      const rows = readSheet(args.buffer, null);
      const pyRows = pyodide.toPy(rows);
      try { result = JSON.parse(glue.po(pyRows)); }
      finally { pyRows.destroy(); }
    } else {
      if (!LIGHT.has(cmd)) await ensureExtras();
      result = JSON.parse(glue.dispatch(cmd, JSON.stringify(args || {}), (m) => log(m)));
    }
    self.postMessage({ id, ok: true, result });
  } catch (e) {
    const msg = (e && e.message) ? e.message : String(e);
    self.postMessage({ id, ok: false, error: msg.split("\n").slice(-6).join("\n") });
  }
};
