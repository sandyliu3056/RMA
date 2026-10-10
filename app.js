/* 售後零件規劃平台（瀏覽器版）主畫面。
 * 所有分析交給 worker.js 裡的 Pyodide（瀏覽器版 Python）執行，這裡只負責畫面與互動。
 * 介面字串：HTML 元素用 data-t="內部中文" 標示，程式碼用 t("內部中文")；切換語言時向 Python（i18n.py）
 * 取得對照表（中文模式也會套用白話用詞），並存在 localStorage，下次開啟不必等 Python 啟動。 */
(function () {
  "use strict";
  // 版本字串：改了 app.js／worker.js／*.py 就一併改這裡與 index.html 的 app.js?v=，避免瀏覽器用舊快取
  const V = "2026-10-10e";
  const DEFAULTS = {
    pyodideBase: "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/",
    xlsxUrl: "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
    engineUrl: "rma_engine.py?v=" + V,
    glueUrl: "web_glue.py?v=" + V,
    i18nUrl: "i18n.py?v=" + V,
    fonts: ["fonts/NotoSansTC-Regular.otf?v=" + V],
    // openpyxl、python-pptx 與相依套件：網站自帶的 wheel 檔（不連 PyPI）；本機測試可在 config.local.js 改來源
    wheels: ["wheels/et_xmlfile-2.0.0-py3-none-any.whl", "wheels/openpyxl-3.1.5-py2.py3-none-any.whl",
             "wheels/python_pptx-1.0.2-py3-none-any.whl", "wheels/typing_extensions-4.16.0-py3-none-any.whl",
             "wheels/xlsxwriter-3.2.9-py3-none-any.whl"],
  };
  const CFG = Object.assign({}, DEFAULTS, window.APP_CONFIG || {});
  const UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"];
  const STATUSES = ["進行中", "待回覆", "已完成", "取消"];
  const $ = (id) => document.getElementById(id);
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

  // ------------------------------------------------------------ 語言
  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } },
  };
  let LANG = store.get("rma_lang") === "en" ? "en" : "zh";
  let UI = {};
  // ui_strings.js（由 make_ui_strings.py 從 i18n.py 產生）：不必等 Python 啟動就有完整字典與名詞說明
  const STATIC = (window.RMA_UI && window.RMA_UI.ui) || {};
  const STATIC_GLOSS = (window.RMA_UI && window.RMA_UI.glossary) || {};
  function loadCachedUI() {
    // 靜態字典與這次部署的 Python 同一份；舊部署留在 localStorage 的快取不再使用
    UI = Object.assign({}, STATIC[LANG] || {});
  }
  loadCachedUI();
  const t = (k) => (typeof UI[k] === "string" ? UI[k] : k);
  const tf = (k, ...a) => t(k).replace(/\{(\d+)\}/g, (_, i) => a[+i]);
  const sep = () => (LANG === "en" ? " | " : "｜");
  function applyUI() {
    document.documentElement.lang = LANG === "en" ? "en" : "zh-Hant";
    document.title = t("售後零件規劃平台");
    document.querySelectorAll("[data-t]").forEach((el) => { el.textContent = t(el.dataset.t); });
    document.querySelectorAll("select[data-t-opts]").forEach((s) => { for (const o of s.options) o.textContent = t(o.value); });
    $("lang-sel").value = LANG;
    for (const o of $("zoom-sel").options) o.textContent = o.value + "%" + (o.value === "100" ? " (" + t("一般") + ")" : "");
    const tzo = $("tz-sel").selectedOptions[0];
    $("tzname").textContent = tzo ? tzo.textContent : "";
    if (!ready && bootPct > 0 && !$("boot-pill").classList.contains("fail")) showProgress(bootPct);
  }
  function setT(el, key) { el.dataset.t = key; el.textContent = t(key); }

  // ------------------------------------------------------------ worker
  const worker = new Worker("worker.js?v=" + V);
  const pending = new Map();
  let seq = 0, ready = false, loaded = false, switching = false, curFile = "", bootPct = 0;
  let analyses = [];
  const loadedTabs = new Set();

  function call(cmd, args, transfer) {
    return new Promise((resolve, reject) => {
      const id = ++seq;
      pending.set(id, { resolve, reject });
      worker.postMessage({ id, cmd, args }, transfer || []);
    });
  }
  worker.onmessage = (ev) => {
    const m = ev.data;
    if (m.type === "log") { log(m.msg); return; }
    if (m.type === "progress") { showProgress(m.pct); return; }
    if (m.type === "ready") { onReady(); return; }
    const p = pending.get(m.id);
    if (!p) return;
    pending.delete(m.id);
    m.ok ? p.resolve(m.result) : p.reject(new Error(m.error));
  };
  worker.onerror = (e) => { bootFailed(e); };
  function showProgress(pct) {
    // 標題列膠囊與「載入並分析」按鈕顯示啟動進度；就緒後恢復
    bootPct = Math.max(bootPct, Math.min(100, pct | 0));
    $("boot-bar").style.width = bootPct + "%";
    if (!ready) {
      $("boot-text").textContent = t("準備中") + " " + bootPct + "%";
      $("load").textContent = t("準備中") + " " + bootPct + "%";
    }
  }
  function bootFailed(e) {
    const msg = t("啟動失敗：") + friendly(e);
    $("boot-pill").classList.add("fail"); $("boot-pill").title = msg;
    setT($("boot-text"), "啟動失敗");
    setT($("load"), "啟動失敗，請重新整理頁面");
    log(msg); toast(msg, 10000);
  }
  async function onReady() {
    ready = true;
    try { await syncLang(); } catch (e) { toast(friendly(e), 8000); }
    $("boot-pill").classList.add("ready"); $("boot-pill").title = "";
    setT($("boot-text"), "就緒");
    setT($("load"), "載入並分析");
    $("aio").disabled = false; $("load").disabled = false;
  }
  async function syncLang() {
    // 告訴 Python 目前語言並取得介面字串（啟動後、切換語言時）
    const r = await call("set_lang", { lang: LANG });
    UI = Object.assign({}, STATIC[LANG] || {}, r.ui || {});
    applyUI();
    return r;
  }
  async function switchLang(lang) {
    lang = lang === "en" ? "en" : "zh";
    if (lang === LANG) return;
    if (switching) { $("lang-sel").value = LANG; toast(t("切換語言中…")); return; }
    LANG = lang; store.set("rma_lang", lang);
    loadCachedUI(); applyUI(); renderActions();
    worker.postMessage({ id: 0, cmd: "ui_lang", args: { lang } });   // worker 自己的進度訊息也換語言
    if (!ready) return;   // 啟動中：先換介面字串；啟動完成時 onReady 會再與 Python 同步
    switching = true;
    $("lang-sel").disabled = true;
    if (loaded) toast(t("切換語言中…"), 60000);
    try {
      await syncLang();
      renderActions();
      if (loaded) { renderSummary(await call("summary", {})); resetTabs(); refreshActive(); }
      if (poData) { poData = await call("po_result", {}); renderPo(); }
    } catch (e) { toast(friendly(e), 8000); }
    finally {
      switching = false;
      $("lang-sel").disabled = false;
      if (loaded) $("toast").style.display = "none";
    }
  }
  $("lang-sel").addEventListener("change", (e) => switchLang(e.target.value));

  // ------------------------------------------------------------ 小工具
  let logTouched = false;
  function log(msg) {
    const el = $("log");
    if (!logTouched) { logTouched = true; el.textContent = ""; el.removeAttribute("data-t"); }
    el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
  }
  let toastTimer = null;
  function toast(msg, ms) {
    const el = $("toast"); el.textContent = msg; el.style.display = "block";
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (el.style.display = "none"), ms || 4000);
  }
  function busy(btn, on, label) {
    if (!btn) return;
    if (on) { btn.dataset.label = btn.textContent; btn.textContent = label || t("處理中…"); btn.disabled = true; }
    else { btn.textContent = btn.dataset.t ? t(btn.dataset.t) : (btn.dataset.label || btn.textContent); btn.disabled = false; }
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
  function friendly(e) {
    // 把 Pyodide／網路錯誤換成看得懂的一句話
    const m = (e && e.message) ? e.message : String(e);
    if (/Failed to fetch|NetworkError when|importScripts|net::ERR_|Load failed/.test(m)) return t("下載瀏覽器版 Python 失敗，請檢查網路後重新整理頁面。");
    if (/out of memory|allocation failed|RangeError|Aborted\(/i.test(m)) return t("記憶體不足，請關閉其他分頁後重試。");
    const lines = m.trim().split("\n");
    return lines[lines.length - 1].replace(/^[A-Za-z_.]*(Error|Exception):\s*/, "");
  }
  function clampInput(el, def) {
    let v = parseInt(el.value, 10);
    if (isNaN(v)) v = def;
    const mn = parseInt(el.min, 10), mx = parseInt(el.max, 10);
    if (!isNaN(mn)) v = Math.max(mn, v);
    if (!isNaN(mx)) v = Math.min(mx, v);
    el.value = v;
    return v;
  }
  const p2 = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
  const today = () => ymd(new Date());
  const plusDays = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return ymd(d); };
  function normDate(s) {
    // 回傳 YYYY-MM-DD；空字串＝不填；null＝格式不對
    if (!s) return "";
    const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
    if (!m) return null;
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    if (d.getFullYear() !== +m[1] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) return null;
    return `${m[1]}-${p2(m[2])}-${p2(m[3])}`;
  }

  function renderTable(el, tj, maxRows) {
    if (!tj || !tj.columns || !tj.columns.length) { el.innerHTML = `<div class="caption pad">${esc(t("（沒有資料）"))}</div>`; return; }
    const n = Math.min(tj.rows.length, maxRows || 2000);
    let h = "<table><thead><tr>" + tj.columns.map((c, i) => {
      const d = tj.defs && tj.defs[i];
      return `<th class="${tj.numeric[i] ? "num" : ""}"${d ? ` title="${esc(d)}"` : ""}>${esc(c)}${d ? ' <span class="info">ⓘ</span>' : ""}</th>`;
    }).join("") + "</tr></thead><tbody>";
    for (let r = 0; r < n; r++) {
      const row = tj.rows[r];
      h += `<tr${tj.warn && tj.warn[r] ? ' class="warn"' : ""}>` + row.map((v, i) => `<td class="${tj.numeric[i] ? "num" : ""}">${esc(v)}</td>`).join("") + "</tr>";
    }
    h += "</tbody></table>";
    const total = tj.total || tj.rows.length;
    if (total > n) h += `<div class="caption pad">${esc(tf("只顯示前 {0} 列，共 {1} 列；匯出 Excel 可取得全部。", n.toLocaleString(), total.toLocaleString()))}</div>`;
    el.innerHTML = h;
  }
  function renderKpis(el, items) {
    el.innerHTML = items.map(([k, v, d]) =>
      `<div class="kpi"${d ? ` title="${esc(d)}"` : ""}><div class="k">${esc(k)}${d ? ' <span class="info">ⓘ</span>' : ""}</div><div class="v">${esc(v)}</div></div>`).join("");
  }
  function renderBars(el, labels, values) {
    const max = Math.max(1, ...values);
    el.innerHTML = labels.map((l, i) => `<div class="b"><b>${values[i].toLocaleString()}</b><i style="height:${Math.round(values[i] / max * 150)}px"></i><span title="${esc(l)}">${esc(l)}</span></div>`).join("");
  }
  const LIGHT = new Set(["#C9B79C", "#F2B134", "#FBE7C6"]);
  function renderHBars(el, c) {
    // 水平堆疊長條圖（與桌面版出貨分頁相同）：c = {title, labels, series:[{name, values, color}]}
    if (!c || !c.labels || !c.labels.length) { el.innerHTML = `<div class="caption">${esc(t("（沒有資料）"))}</div>`; return; }
    const totals = c.labels.map((_, i) => c.series.reduce((s, x) => s + (x.values[i] || 0), 0));
    const max = Math.max(1, ...totals);
    let h = `<div class="hb-title">${esc(c.title)}</div><div class="hb-legend">${c.series.map((s) => `<span><i style="background:${esc(s.color)}"></i>${esc(s.name)}</span>`).join("")}</div>`;
    c.labels.forEach((l, i) => {
      h += `<div class="hb-row"><span class="hb-lab" title="${esc(l)}">${esc(l)}</span><span class="hb-bar">` +
        c.series.map((s) => {
          const v = s.values[i] || 0; if (!v) return "";
          const w = v / max * 100;
          return `<i class="${LIGHT.has(s.color.toUpperCase()) ? "dark" : ""}" style="width:${w}%;background:${esc(s.color)}" title="${esc(s.name)} ${v.toLocaleString()}">${w > 9 ? v.toLocaleString() : ""}</i>`;
        }).join("") + `</span><b>${totals[i].toLocaleString()}</b></div>`;
    });
    el.innerHTML = h;
    el.querySelectorAll(".hb-bar i").forEach((seg) => { if (seg.scrollWidth > seg.clientWidth) seg.textContent = ""; });
  }
  function downloadB64(b64, name, mime) {
    const bin = atob(b64), buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    downloadBlob(new Blob([buf], { type: mime }), name);
  }
  function downloadBlob(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function needData() { if (!loaded) { toast(t("請先在「資料來源」載入 AIO 匯出檔")); return false; } return true; }
  function dl(b64, name) { if (b64) downloadB64(b64, name, XLSX_MIME); else toast(t("（沒有資料）")); }

  // ------------------------------------------------------------ 標題列設定（縮放、字型、主題、時區）與時鐘
  const PREFS = { zoom: "100", font: "hand", theme: "brown", tz: "Asia/Taipei" };
  for (const k of Object.keys(PREFS)) { const v = store.get("rma_" + k); if (v) PREFS[k] = v; }
  function applyPrefs() {
    for (const k of Object.keys(PREFS)) {
      const el = $(k + "-sel");
      el.value = PREFS[k];
      if (el.value !== PREFS[k]) { el.selectedIndex = 0; PREFS[k] = el.value; }   // 儲存的值已不存在就用第一個
    }
    document.documentElement.dataset.theme = PREFS.theme;
    document.documentElement.dataset.font = PREFS.font;
    document.body.style.zoom = String((+PREFS.zoom || 100) / 100);
    const tzo = $("tz-sel").selectedOptions[0];
    $("tzname").textContent = tzo ? tzo.textContent : "";
    tick();
  }
  for (const k of Object.keys(PREFS)) $(k + "-sel").addEventListener("change", (e) => { PREFS[k] = e.target.value; store.set("rma_" + k, PREFS[k]); applyPrefs(); });
  function tick() {
    const now = new Date();
    let tw;
    try { tw = new Date(now.toLocaleString("en-US", { timeZone: PREFS.tz })); } catch (e) { tw = now; }
    $("clock").textContent = `${p2(tw.getHours())}:${p2(tw.getMinutes())}`;
    $("date").textContent = `${tw.getFullYear()}/${p2(tw.getMonth() + 1)}/${p2(tw.getDate())} ${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][tw.getDay()]}`;
  }
  setInterval(tick, 30000);

  // ------------------------------------------------------------ 分頁
  const refreshers = { demand: refreshDemand, stock: refreshStock, ship: refreshShip, report: refreshReportList };
  function activeTab() { const b = document.querySelector("#tabs button.active"); return b ? b.dataset.tab : "source"; }
  function refreshActive() {
    const name = activeTab();
    if (loaded && refreshers[name] && !loadedTabs.has(name)) { loadedTabs.add(name); refreshers[name](); }
  }
  $("tabs").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-tab]"); if (!b) return;
    document.querySelectorAll("#tabs button").forEach((x) => x.classList.toggle("active", x === b));
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x.id === "tab-" + b.dataset.tab));
    refreshActive();
  });
  function resetTabs() {
    // 載入新檔或切換語言後，各分頁改成下次打開時重算；先清掉舊畫面
    loadedTabs.clear();
    demandData = stockData = shipData = null; currentKey = null; currentTable = null; showSeq++;
    for (const id of ["d-weekly", "d-chart", "d-forecast", "d-top", "s-kpis", "s-parts", "s-cases", "h-kpis", "h-chart-backlog",
                      "h-chart-transit", "h-backlog", "h-transit", "h-exp", "r-list"]) $(id).innerHTML = "";
    $("d-fc-title").textContent = ""; $("h-backlog-hint").textContent = ""; $("h-exp-hint").textContent = "";
    $("r-body").innerHTML = `<div class="caption">${esc(t("請先載入資料，再從左側選一項分析"))}</div>`;
  }

  // ------------------------------------------------------------ 資料來源
  function renderSummary(s) {
    $("status").textContent = `${curFile}${sep()}${s.status}`;
    $("status").removeAttribute("data-t");
    renderKpis($("src-kpis"), s.kpis);
    $("src-caption").textContent = s.caption;
    $("src-period").textContent = s.period;
    $("src-status-title").textContent = s.status_title;
    renderTable($("src-status-tbl"), s.status_table);
    renderBars($("src-status-chart"), s.status_chart.labels, s.status_chart.values);
    $("src-status").style.display = "block";
    analyses = s.analyses || [];
  }
  $("load").addEventListener("click", async () => {
    const f = $("aio").files[0];
    if (!f) { toast(t("請先選擇 AIO 匯出檔")); return; }
    const start = normDate($("start").value.trim());
    if (start === null) { toast(t("分析起日格式不對，請輸入像 2026-08-01 這樣的日期，或留空。"), 6000); return; }
    busy($("load"), true, t("分析中…"));
    setT($("status"), "載入中…");
    try {
      const buffer = await f.arrayBuffer();
      const s = await call("load", { buffer, start }, [buffer]);
      loaded = true; curFile = f.name;
      resetTabs(); renderSummary(s); refreshActive();
      log(t("完成。"));
    } catch (e) {
      log(t("錯誤：") + e.message); toast(t("載入失敗：") + friendly(e), 8000);
      setT($("status"), "載入失敗");
    } finally { busy($("load"), false); }
  });

  // ------------------------------------------------------------ 需求規劃
  let demandData = null;
  async function refreshDemand() {
    if (!needData()) return;
    busy($("d-run"), true);
    try {
      const weeks = clampInput($("d-weeks"), 4);
      demandData = await call("tab_demand", { by: $("d-by").value, weeks });
      renderTable($("d-weekly"), demandData.weekly);
      if (demandData.weekly_chart) renderBars($("d-chart"), demandData.weekly_chart.labels, demandData.weekly_chart.values.map(Math.round));
      else $("d-chart").innerHTML = "";
      $("d-fc-title").textContent = demandData.forecast_title;
      renderTable($("d-forecast"), demandData.forecast);
      renderTable($("d-top"), demandData.top);
    } catch (e) { toast(friendly(e), 8000); } finally { busy($("d-run"), false); }
  }
  $("d-run").addEventListener("click", refreshDemand);
  $("d-by").addEventListener("change", refreshDemand);
  $("d-weekly-dl").addEventListener("click", () => demandData && dl(demandData.weekly_xlsx, demandData.weekly_name));
  $("d-top-dl").addEventListener("click", () => demandData && dl(demandData.top_xlsx, demandData.top_name));

  // ------------------------------------------------------------ 庫存與缺料
  let stockData = null;
  async function refreshStock() {
    if (!needData()) return;
    busy($("s-run"), true);
    try {
      const days = clampInput($("s-days"), 14);
      stockData = await call("tab_stock", { days });
      renderKpis($("s-kpis"), stockData.kpis);
      renderTable($("s-parts"), stockData.parts);
      $("s-cases-title").textContent = stockData.cases_title; $("s-cases-title").removeAttribute("data-t");
      renderTable($("s-cases"), stockData.cases);
    } catch (e) { toast(friendly(e), 8000); } finally { busy($("s-run"), false); }
  }
  $("s-run").addEventListener("click", refreshStock);
  $("s-parts-dl").addEventListener("click", () => stockData && dl(stockData.parts_xlsx, stockData.parts_name));
  $("s-cases-dl").addEventListener("click", () => stockData && dl(stockData.cases_xlsx, stockData.cases_name));

  // ------------------------------------------------------------ 訂單與交期
  let poData = null;
  function renderPo() {
    if (!poData) return;
    renderKpis($("po-kpis"), poData.kpis);
    renderTable($("po-sup"), poData.suppliers);
    renderTable($("po-late"), poData.late);
    $("po-body").style.display = "block";
  }
  $("po-tmpl").addEventListener("click", async () => {
    if (!ready) { toast(t("瀏覽器版 Python 還在啟動")); return; }
    try { const r = await call("po_template", {}); downloadB64(r.data, r.name, XLSX_MIME); } catch (e) { toast(friendly(e)); }
  });
  $("po-file").addEventListener("change", async () => {
    const f = $("po-file").files[0]; if (!f) return;
    if (!ready) { toast(t("瀏覽器版 Python 還在啟動")); $("po-file").value = ""; return; }
    toast(t("匯入中…"), 30000);
    try {
      const buffer = await f.arrayBuffer();
      poData = await call("po", { buffer }, [buffer]);
      renderPo(); $("toast").style.display = "none";
    } catch (e) { toast(t("匯入失敗：") + friendly(e), 8000); }
    finally { $("po-file").value = ""; }
  });
  $("po-late-dl").addEventListener("click", () => poData && dl(poData.late_xlsx, poData.late_name));

  // ------------------------------------------------------------ 出貨與到貨
  let shipData = null;
  async function refreshShip() {
    if (!needData()) return;
    busy($("h-run"), true);
    try {
      shipData = await call("tab_ship", { days: clampInput($("h-days"), 30) });
      renderKpis($("h-kpis"), shipData.kpis);
      renderHBars($("h-chart-backlog"), shipData.backlog_chart);
      renderHBars($("h-chart-transit"), shipData.transit_chart);
      $("h-backlog-hint").textContent = shipData.backlog_hint;
      renderTable($("h-backlog"), shipData.backlog);
      renderTable($("h-transit"), shipData.transit);
      $("h-exp-title").textContent = shipData.expiring_title; $("h-exp-title").removeAttribute("data-t");
      $("h-exp-hint").textContent = shipData.expiring_hint;
      renderTable($("h-exp"), shipData.expiring);
    } catch (e) { toast(friendly(e), 8000); } finally { busy($("h-run"), false); }
  }
  $("h-run").addEventListener("click", refreshShip);
  $("h-backlog-dl").addEventListener("click", () => shipData && dl(shipData.backlog_xlsx, shipData.backlog_name));
  $("h-exp-dl").addEventListener("click", () => shipData && dl(shipData.expiring_xlsx, shipData.expiring_name));

  // ------------------------------------------------------------ 報表與分析
  let currentKey = null, currentTable = null, showSeq = 0;
  function refreshReportList() {
    if (!needData()) return;
    const list = $("r-list");
    list.innerHTML = analyses.map(([k, title]) => `<button data-key="${esc(k)}">${esc(k)}  ${esc(title)}</button>`).join("");
    list.onclick = (e) => { const b = e.target.closest("button[data-key]"); if (b) showAnalysis(b.dataset.key); };
    if (analyses.length) showAnalysis(currentKey && analyses.some((a) => a[0] === currentKey) ? currentKey : analyses[0][0]);
  }
  async function showAnalysis(key) {
    const my = ++showSeq;   // 快速連點時只顯示最後一個
    currentKey = key; currentTable = null;
    document.querySelectorAll("#r-list button").forEach((b) => b.classList.toggle("active", b.dataset.key === key));
    $("r-body").innerHTML = `<div class="boot"><span class="spin"></span>${esc(t("畫圖中…"))}</div>`;
    try {
      const a = await call("analysis", { key });
      if (my !== showSeq) return;
      const opts = a.tables.map(([n, label]) => `<option value="${esc(n)}">${esc(label)}</option>`).join("");
      $("r-body").innerHTML = `
        <h2>${esc(a.key)} ${esc(a.title)}</h2>
        <div class="caption">${esc(a.subtitle)}</div>
        <div class="rbody">
          <div>
            <div class="h">${esc(a.labels.points)}</div>
            <ul>${a.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul>
            ${a.paragraph ? `<div class="h">${esc(a.labels.about)}</div><p>${esc(a.paragraph)}</p>` : ""}
            <p class="concl"><b>${esc(a.labels.takeaway)}</b>${esc(a.conclusion)}</p>
          </div>
          <div>${a.chart ? `<img src="data:image/png;base64,${a.chart}" alt="${esc(a.title)}">` : `<div class="caption">${esc(t("（這一項沒有圖）"))}</div>`}</div>
        </div>
        <div class="row">
          <label class="inline"><span>${esc(t("表格："))}</span><select id="r-table">${opts}</select></label>
          <button class="btn" id="r-table-dl">${esc(t("匯出此表"))}</button>
          <span class="small">${esc(t("欄名有 ⓘ 的，滑鼠停上去看說明"))}</span>
        </div>
        <div class="tbl" id="r-table-body"></div>`;
      $("r-table").onchange = () => showTable(key, $("r-table").value);
      $("r-table-dl").onclick = () => { if (currentTable) dl(currentTable.xlsx, currentTable.name); };
      if (a.tables.length) showTable(key, a.tables[0][0]);
    } catch (e) { if (my === showSeq) $("r-body").innerHTML = `<div class="notice">${esc(t("無法顯示："))}${esc(friendly(e))}</div>`; }
  }
  async function showTable(key, name) {
    const my = showSeq;
    try {
      const r = await call("analysis_table", { key, name });
      if (my !== showSeq || !$("r-table-body")) return;
      currentTable = r; renderTable($("r-table-body"), r.table);
    } catch (e) { toast(friendly(e)); }
  }
  async function exportKind(btn, kind, mime) {
    if (!needData()) return;
    busy(btn, true, t("產生中…"));
    try { const r = await call("export", { kind }); downloadB64(r.data, r.name, mime); }
    catch (e) { toast(t("匯出失敗：") + friendly(e), 8000); } finally { busy(btn, false); }
  }
  $("r-xlsx").addEventListener("click", (e) => exportKind(e.currentTarget, "xlsx", XLSX_MIME));
  $("r-pptx").addEventListener("click", (e) => exportKind(e.currentTarget, "pptx", PPTX_MIME));
  $("r-charts").addEventListener("click", (e) => exportKind(e.currentTarget, "charts", "application/zip"));

  // ------------------------------------------------------------ 名詞說明
  $("gloss-btn").addEventListener("click", async () => {
    try {
      let g = STATIC_GLOSS[LANG];
      if (!Array.isArray(g) || !g.length) {
        if (!ready) { toast(t("瀏覽器版 Python 還在啟動")); return; }
        g = await call("glossary", {});
      }
      $("gloss-body").innerHTML = g.map(([n, d]) => `<dt>${esc(n)}</dt><dd>${esc(d)}</dd>`).join("");
      $("gloss").style.display = "flex";
      $("gloss-body").scrollTop = 0; $("gloss-body").focus();
    } catch (e) { toast(friendly(e)); }
  });
  function closeGloss() { if ($("gloss").style.display === "none") return; $("gloss").style.display = "none"; $("gloss-btn").focus(); }
  $("gloss-close").addEventListener("click", closeGloss);
  $("gloss").addEventListener("click", (e) => { if (e.target === $("gloss")) closeGloss(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeGloss(); });

  // ------------------------------------------------------------ 跨單位協調（存在 localStorage，欄位名稱與桌面版的 協調待辦.json 相同）
  const KEY = "rma_actions";
  const unitFromLabel = (v) => UNITS.find((u) => u === v || t(u) === v);
  const statusFromLabel = (v) => STATUSES.find((s) => s === v || t(s) === v);
  function sanitizeActions(arr) {
    // 只接受長得像待辦的資料；壞掉的欄位補預設值，重複的 id 重新編號
    if (!Array.isArray(arr)) return [];
    const out = [], seen = new Set();
    const date = (v, d) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v || "")) ? String(v) : d);
    for (const a of arr) {
      if (!a || typeof a !== "object") continue;
      const topic = String(a.議題 ?? a.Topic ?? "").trim();
      if (!topic) continue;
      let id = Number(a.id);
      if (!Number.isInteger(id) || id < 1 || seen.has(id)) id = null; else seen.add(id);
      out.push({ id, 對象單位: unitFromLabel(a.對象單位 ?? a.Team) || "其他", 議題: topic.slice(0, 500),
                 負責人: String(a.負責人 ?? a.Owner ?? "").slice(0, 100), 到期日: date(a.到期日 ?? a["Due Date"], plusDays(7)),
                 狀態: statusFromLabel(a.狀態 ?? a.Status) || "進行中", 建立日: date(a.建立日 ?? a.Created, today()) });
    }
    let next = 1;
    for (const a of out) if (!a.id) { while (seen.has(next)) next++; a.id = next; seen.add(next); }
    return out;
  }
  let actions = [];
  try { actions = sanitizeActions(JSON.parse(store.get(KEY) || "[]")); } catch (e) { actions = []; }
  function saveActions() {
    if (!store.set(KEY, JSON.stringify(actions))) toast(t("無法儲存待辦（瀏覽器儲存空間不可用）"));
    renderActions();
  }
  $("c-unit").innerHTML = UNITS.map((u) => `<option value="${esc(u)}">${esc(u)}</option>`).join("");
  $("c-status").innerHTML = STATUSES.map((s) => `<option value="${esc(s)}">${esc(s)}</option>`).join("");
  $("c-due").value = plusDays(7);

  function addAction(unit, topic, owner, due, status) {
    const id = actions.length ? Math.max(...actions.map((a) => a.id)) + 1 : 1;
    actions.push({ id, 對象單位: unit, 議題: topic, 負責人: owner || "", 到期日: due, 狀態: status || "進行中", 建立日: today() });
  }
  function renderActions() {
    const el = $("c-table");
    if (!actions.length) { el.innerHTML = `<div class="caption pad">${esc(t("目前沒有待辦。"))}</div>`; return; }
    const sorted = [...actions].sort((a, b) => (a.狀態 === "已完成") - (b.狀態 === "已完成") || a.到期日.localeCompare(b.到期日));
    const td = today();
    const heads = ["#", ...["對象單位", "議題", "負責人", "到期日", "狀態", "建立日"].map(t), ""];
    let h = "<table><thead><tr>" + heads.map((x) => `<th>${esc(x)}</th>`).join("") + "</tr></thead><tbody>";
    for (const a of sorted) {
      const late = ["進行中", "待回覆"].includes(a.狀態) && a.到期日 < td;
      h += `<tr data-id="${a.id}"${late ? ' class="warn"' : ""}><td>${a.id}</td>
        <td><select data-f="對象單位">${UNITS.map((u) => `<option value="${esc(u)}"${u === a.對象單位 ? " selected" : ""}>${esc(t(u))}</option>`).join("")}</select></td>
        <td class="topic"><input data-f="議題" value="${esc(a.議題)}"></td>
        <td><input data-f="負責人" value="${esc(a.負責人)}" size="8"></td>
        <td><input data-f="到期日" type="date" value="${esc(a.到期日)}"></td>
        <td><select data-f="狀態">${STATUSES.map((s) => `<option value="${esc(s)}"${s === a.狀態 ? " selected" : ""}>${esc(t(s))}</option>`).join("")}</select></td>
        <td>${esc(a.建立日)}</td><td><button class="btn" data-del="${a.id}">${esc(t("刪除"))}</button></td></tr>`;
    }
    el.innerHTML = h + "</tbody></table>";
  }
  $("c-table").addEventListener("change", (e) => {
    const f = e.target.dataset.f; if (!f) return;
    const id = +e.target.closest("tr").dataset.id;
    const a = actions.find((x) => x.id === id); if (!a) return;
    let v = e.target.value;
    if (f === "議題") { v = v.trim(); if (!v) { toast(t("請填議題")); renderActions(); return; } }
    if (f === "到期日" && !/^\d{4}-\d{2}-\d{2}$/.test(v)) { renderActions(); return; }
    a[f] = v; saveActions();
  });
  $("c-table").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-del]"); if (!b) return;
    actions = actions.filter((x) => x.id !== +b.dataset.del); saveActions();
  });
  $("c-add").addEventListener("click", () => {
    const topic = $("c-topic").value.trim();
    if (!topic) { toast(t("請填議題")); return; }
    addAction($("c-unit").value, topic, $("c-owner").value.trim(), $("c-due").value || plusDays(7), $("c-status").value);
    $("c-topic").value = ""; saveActions();
  });
  $("c-seed").addEventListener("click", async () => {
    if (!needData()) return;
    try {
      const seeds = await call("conclusions", {});
      let n = 0;
      for (const [, unit, topic] of seeds) if (!actions.some((a) => a.議題 === topic)) { addAction(unit, topic, "", plusDays(14), "進行中"); n++; }
      saveActions(); toast(tf("已新增 {0} 筆待辦", n));
    } catch (e) { toast(friendly(e)); }
  });
  $("c-dl").addEventListener("click", () => downloadBlob(new Blob([JSON.stringify(actions, null, 2)], { type: "application/json" }), t("協調待辦.json")));
  $("c-restore").addEventListener("change", async () => {
    const f = $("c-restore").files[0]; if (!f) return;
    try {
      const arr = JSON.parse(await f.text());
      if (!Array.isArray(arr)) throw new Error(t("格式不對"));
      actions = sanitizeActions(arr); saveActions(); toast(tf("已還原 {0} 筆", actions.length));
    } catch (e) { toast(t("讀取失敗：") + friendly(e)); }
    $("c-restore").value = "";
  });
  $("c-csv").addEventListener("click", () => {
    if (!actions.length) { toast(t("目前沒有待辦")); return; }
    const cols = ["id", "對象單位", "議題", "負責人", "到期日", "狀態", "建立日"];
    // 以 = + - @ 開頭的字會被 Excel 當公式，前面加上 ' 避免
    const q = (v) => { let s = String(v ?? ""); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return `"${s.replace(/"/g, '""')}"`; };
    const head = ["#", ...cols.slice(1).map(t)];
    const csv = "﻿" + head.map(q).join(",") + "\n" +
      actions.map((a) => cols.map((c) => q(c === "對象單位" || c === "狀態" ? t(a[c]) : a[c])).join(",")).join("\n");
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), t("跨單位協調待辦.csv"));
  });

  // ------------------------------------------------------------ 啟動
  applyUI();
  applyPrefs();
  renderActions();
  call("boot", Object.assign({ lang: LANG }, CFG)).catch(bootFailed);
})();
