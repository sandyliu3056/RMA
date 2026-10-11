/* 售後零件規劃平台（瀏覽器版）主畫面。
 * 所有分析交給 worker.js 裡的 Pyodide（瀏覽器版 Python）執行，這裡只負責畫面與互動。
 * 介面字串：HTML 元素用 data-t="內部中文" 標示，程式碼用 t("內部中文")；切換語言時向 Python（i18n.py）
 * 取得對照表（中文模式也會套用白話用詞），並存在 localStorage，下次開啟不必等 Python 啟動。 */
(function () {
  "use strict";
  // 版本字串：改了 app.js／worker.js／*.py 就一併改這裡與 index.html 的 app.js?v=，避免瀏覽器用舊快取
  const V = "2026-10-11a";
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
    if (ready) showExtras();
    if (phase) renderSteps(phase === "boot" ? bootPct : loadPct());
  }
  function setT(el, key) { el.dataset.t = key; el.textContent = t(key); }

  // ------------------------------------------------------------ worker
  const worker = new Worker("worker.js?v=" + V);
  const pending = new Map();
  let seq = 0, ready = false, loaded = false, switching = false, curFile = "", bootPct = 0;
  let analyses = [], statusDetail = null;
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
    if (m.type === "progress") { showProgress(m.pct); if (phase === "boot") renderSteps(m.pct); return; }
    if (m.type === "step") { onStep(m); return; }
    if (m.type === "extras") { extrasState = m.state; showExtras(); return; }
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
  let extrasState = "";
  function showExtras() {
    // 第二段（圖表與匯出套件）的狀態：就緒前在膠囊顯示提示
    if (!ready) return;
    if (extrasState === "loading") { $("boot-text").textContent = t("就緒") + " · " + t("圖表套件載入中…"); $("boot-pill").title = ""; }
    else if (extrasState === "failed") { setT($("boot-text"), "就緒"); $("boot-pill").title = t("圖表套件載入失敗，需要時會再試"); }
    else { setT($("boot-text"), "就緒"); $("boot-pill").title = ""; }
  }
  function bootFailed(e) {
    const msg = t("啟動失敗：") + friendly(e);
    $("boot-pill").classList.add("fail"); $("boot-pill").title = msg;
    setT($("boot-text"), "啟動失敗");
    setT($("load"), "啟動失敗，請重新整理頁面");
    phaseDone = "fail"; const cur = Object.keys(stepState).find((x) => stepState[x] === "run"); if (cur) stepState[cur] = "fail";
    log(msg); toast(msg, 10000); renderSteps(bootPct);
  }
  async function onReady() {
    ready = true;
    try { await syncLang(); } catch (e) { toast(friendly(e), 8000); }
    $("boot-pill").classList.add("ready"); $("boot-pill").title = "";
    setT($("boot-text"), "就緒");
    setT($("load"), "載入並分析");
    $("aio").disabled = false; $("load").disabled = false;
    showExtras();
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
    loadCachedUI(); applyUI(); renderActions(); renderImports(); dbUI();
    if (window.RMAScene) RMAScene.setLang(LANG);
    worker.postMessage({ id: 0, cmd: "ui_lang", args: { lang } });   // worker 自己的進度訊息也換語言
    if (!ready) return;   // 啟動中：先換介面字串；啟動完成時 onReady 會再與 Python 同步
    switching = true;
    $("lang-sel").disabled = true;
    if (loaded) toast(t("切換語言中…"), 60000);
    try {
      await syncLang();
      renderActions(); renderImports(); dbUI();
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
  function log(msg) {
    const el = $("log");
    el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
    const sm = $("steps-msg");
    sm.textContent = String(msg).trim(); sm.style.animation = "none"; void sm.offsetWidth; sm.style.animation = "";
  }
  // ------------------------------------------------------------ 步驟進度面板（啟動／載入的動畫）
  const PHASES = {
    boot: { title: "啟動", done: "就緒", steps: [["runtime", "下載瀏覽器版 Python"], ["pandas", "載入 pandas"], ["engine", "載入分析引擎"], ["extras", "圖表與匯出套件（背景）"]] },
    load: { title: "載入並分析", done: "完成", steps: [["read", "讀取 Excel"], ["clean", "整理資料"], ["analyze", "執行分析"], ["parts", "零件明細"]] },
    save: { title: "存入資料庫", done: "已存入", steps: [["meta", "建立匯入紀錄"], ["rows", "上傳案件"]] },
  };
  const PCT = { load: { read: [0, 10], clean: [10, 25], analyze: [25, 92], parts: [92, 100] }, save: { meta: [0, 8], rows: [8, 100] } };
  let phase = null, stepState = {}, stepInfo = {}, stepLabel = {}, phaseDone = null;
  function startPhase(name) {
    phase = name; phaseDone = null; stepState = {}; stepInfo = {}; stepLabel = {};
    for (const [id] of PHASES[name].steps) stepState[id] = "wait";
    $("log").textContent = ""; $("steps-msg").textContent = "";
    renderSteps(0);
  }
  function onStep(m) {
    if (m.phase !== phase) { if (m.phase === "boot") return; startPhase(m.phase); }   // 背景套件的啟動步驟不要蓋掉載入／存入的面板
    const ids = PHASES[phase].steps.map((s) => s[0]);
    let id = m.id;
    if (id === "current") id = ids.find((x) => stepState[x] === "run") || ids[ids.length - 1];
    const k = ids.indexOf(id); if (k < 0) return;
    if (m.state === "run") ids.slice(0, k).forEach((x) => { if (stepState[x] !== "fail" && x !== "extras") stepState[x] = "done"; });
    stepState[id] = m.state;
    if (m.i && m.n) stepInfo[id] = `${m.i}/${m.n}`;
    if (m.state === "fail") phaseDone = "fail";
    else if (phase !== "boot" && id === ids[ids.length - 1] && m.state === "done") { ids.forEach((x) => (stepState[x] = "done")); phaseDone = "done"; }
    renderSteps(phase === "boot" ? bootPct : loadPct());
  }
  function loadPct() {
    let pct = 0;
    for (const [id, [a, b]] of Object.entries(PCT[phase] || PCT.load)) {
      if (stepState[id] === "done") pct = b;
      else if (stepState[id] === "run") { const f = stepInfo[id] ? (+stepInfo[id].split("/")[0]) / (+stepInfo[id].split("/")[1]) : 0.3; pct = a + (b - a) * f; break; }
    }
    return phaseDone === "done" ? 100 : pct;
  }
  function renderSteps(pct) {
    if (!phase) return;
    const P = PHASES[phase];
    const running = Object.values(stepState).some((s) => s === "run") && phaseDone !== "done";
    $("steps-title").textContent = t(phaseDone === "done" ? P.done : phaseDone === "fail" ? "失敗" : P.title);
    const p = phaseDone === "done" ? 100 : Math.max(0, Math.min(100, pct | 0));
    $("steps-pct").textContent = p + "%";
    const bar = $("steps-bar"); bar.style.width = p + "%"; bar.className = "pbar-fill" + (phaseDone === "done" ? " done" : phaseDone === "fail" ? " fail" : running ? " running" : "");
    if (window.RMAScene) RMAScene.setProgress(p / 100);
    $("steps-list").innerHTML = P.steps.map(([id, label]) => {
      const st = stepState[id] || "wait";
      const ico = st === "done" ? "✓" : st === "fail" ? "✕" : "";
      const sub = stepInfo[id] && st !== "done" ? ` <span class="sub">${esc(stepInfo[id])}</span>` : "";
      return `<li class="${st}"><span class="ico">${ico}</span><span class="lab">${esc(t(stepLabel[id] || label))}${sub}</span></li>`;
    }).join("");
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
  // 國旗：countries.js 的對照表（AIO 的國家寫法 → ISO 代碼、座標）；查不到的名字就不加旗子
  const COUNTRIES = window.RMA_COUNTRIES || {};
  function flagImg(name) { const c = COUNTRIES[String(name).trim()]; return c ? `<img class="flag" src="flags/${c.iso2}.svg" alt="" loading="lazy">` : ""; }
  // 文字裡的國家名稱（可能是「Indonesia、Thailand」這種清單）：每個名字前面加國旗
  function flagify(s) { return String(s).split(/(、|,\s*|;\s*)/).map((p, i) => (i % 2 ? esc(p) : flagImg(p) + esc(p))).join(""); }
  function friendly(e) {
    // 把 Pyodide／網路錯誤換成看得懂的一句話
    const m = (e && e.message) ? e.message : String(e);
    if (/Failed to fetch|NetworkError when|importScripts|net::ERR_|Load failed/.test(m)) return t("下載瀏覽器版 Python 失敗，請檢查網路後重新整理頁面。");
    if (/out of memory|allocation failed|RangeError|Aborted\(/i.test(m)) return t("記憶體不足，請關閉其他分頁後重試。");
    if (/loading-packages|Can't fetch|Couldn't load package|micropip/i.test(m)) return t("圖表與匯出套件下載失敗，請檢查網路後再試一次。");
    const lines = m.trim().split("\n");
    return lines[lines.length - 1].replace(/^[A-Za-z_.]*(Error|Exception):\s*/, "");
  }
  function dbFriendly(e) {
    // 資料庫（Supabase）相關的錯誤
    const m = (e && e.message) ? e.message : String(e);
    if (/NOT_LOGGED_IN/.test(m)) return t("請先登入資料庫");
    if (/SESSION_EXPIRED|JWT expired|invalid JWT|HTTP 401/i.test(m)) return t("登入已過期，請重新登入");
    if (/Invalid login credentials|invalid_credentials|invalid_grant/i.test(m)) return t("帳號或密碼不對");
    if (/Failed to fetch|NetworkError|Load failed|net::ERR_/i.test(m)) return t("連不上資料庫（網址不對或沒有網路）");
    return friendly(e);
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
      return `<th class="${tj.numeric[i] ? "num" : ""}"${d ? ` title="${esc(d)}"` : ""}>${flagImg(c)}${esc(c)}${d ? ' <span class="info">ⓘ</span>' : ""}</th>`;
    }).join("") + "</tr></thead><tbody>";
    for (let r = 0; r < n; r++) {
      const row = tj.rows[r];
      const cls = ((tj.warn && tj.warn[r]) ? "warn " : "") + ((tj.cls && tj.cls[r]) || "");
      h += `<tr${cls.trim() ? ` class="${esc(cls.trim())}"` : ""}>` + row.map((v, i) => `<td class="${tj.numeric[i] ? "num" : ""}">${tj.flags && tj.flags[i] ? flagify(v) : esc(v)}</td>`).join("") + "</tr>";
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
  function renderBars(el, labels, values, meta) {
    // 單一數列的直條圖；meta = {title, ylabel, note}
    const max = Math.max(1, ...values);
    const m = meta || {};
    el.innerHTML = (m.title ? `<div class="ch-title">${esc(m.title)}${m.ylabel ? ` <span class="ch-unit">(${esc(m.ylabel)})</span>` : ""}</div>` : "") +
      `<div class="bars">` + labels.map((l, i) => `<div class="b"><b>${values[i].toLocaleString()}</b><i style="height:${Math.round(values[i] / max * 150)}px" title="${esc(l)}：${values[i].toLocaleString()}"></i><span title="${esc(l)}">${flagImg(l)}${esc(l)}</span></div>`).join("") + `</div>` +
      (m.note ? `<div class="ch-note">${esc(m.note)}</div>` : "");
  }
  const SEG_COLORS = ["#2E6FBA", "#D9822B", "#2E9E6B", "#8E4FB5", "#C9A227", "#4FB0C6", "#B5544C", "#6B7280", "#7FA650", "#C77DB5", "#5B8FD0", "#E8862B"];
  function renderHStack(el, c) {
    // 橫向堆疊：c = {title, unit, note, rows:[{label, total, segs:[{name, value}]}]}；色塊夠寬就標名稱與數量
    if (!c || !c.rows || !c.rows.length) { el.innerHTML = `<div class="caption pad">${esc(t("（沒有資料）"))}</div>`; return; }
    const max = Math.max(1, ...c.rows.map((r) => r.total));
    let h = `<div class="ch-title">${esc(c.title || "")}${c.unit ? ` <span class="ch-unit">(${esc(c.unit)})</span>` : ""}</div><div class="hstack">`;
    for (const r of c.rows) {
      // 一組一塊：標題列（分組名＋合計）、長條（色塊）、下方逐一列出每個原始狀態與件數（顏色對應色塊）
      h += `<div class="hs-group"><div class="hs-head"><span class="hs-lab">${esc(r.label)}</span><b>${r.total.toLocaleString()}</b></div>` +
        `<span class="hb-bar">` + r.segs.map((sg, i) => { const w = sg.value / max * 100; return `<i style="width:${w}%;background:${SEG_COLORS[i % SEG_COLORS.length]}" title="${esc(sg.name)}：${sg.value.toLocaleString()}">${esc(sg.name)} ${sg.value.toLocaleString()}</i>`; }).join("") + `</span>` +
        `<div class="hs-keys">` + r.segs.map((sg, i) => `<span><i style="background:${SEG_COLORS[i % SEG_COLORS.length]}"></i>${esc(sg.name)} <b>${sg.value.toLocaleString()}</b></span>`).join("") + `</div></div>`;
    }
    h += `</div>` + (c.note ? `<div class="ch-note">${esc(c.note)}</div>` : "");
    el.innerHTML = h;
    el.querySelectorAll(".hb-bar i").forEach((seg) => { if (seg.scrollWidth > seg.clientWidth) seg.textContent = ""; });
  }
  function renderStackedBars(el, c) {
    // 堆疊直條圖：c = {title, xlabel, ylabel, note, labels, series:[{name, values, color}], totals}
    if (!c || !c.labels || !c.labels.length) { el.innerHTML = `<div class="caption pad">${esc(t("（沒有資料）"))}</div>`; return; }
    const totals = c.totals || c.labels.map((_, i) => c.series.reduce((s, x) => s + (x.values[i] || 0), 0));
    const max = Math.max(1, ...totals);
    let h = `<div class="ch-title">${esc(c.title || "")}${c.ylabel ? ` <span class="ch-unit">(${esc(c.ylabel)})</span>` : ""}</div>`;
    h += `<div class="hb-legend">${c.series.map((s) => `<span><i style="background:${esc(s.color)}"></i>${flagImg(s.name)}${esc(s.name)}</span>`).join("")}</div>`;
    h += `<div class="bars">` + c.labels.map((l, i) => {
      // 第一個數列在最下面：由上往下排，所以反轉順序
      const segs = c.series.slice().reverse().map((s) => { const v = s.values[i] || 0; return v ? `<i style="height:${v / max * 150}px;background:${esc(s.color)}" title="${esc(s.name)}：${Math.round(v).toLocaleString()}"></i>` : ""; }).join("");
      return `<div class="b"><b>${Math.round(totals[i]).toLocaleString()}</b><div class="segs">${segs}</div><span title="${esc(l)}">${esc(l)}</span></div>`;
    }).join("") + `</div>`;
    if (c.xlabel) h += `<div class="ch-x">${esc(c.xlabel)} →</div>`;
    if (c.note) h += `<div class="ch-note">${esc(c.note)}</div>`;
    el.innerHTML = h;
  }
  const LIGHT = new Set(["#C9B79C", "#F2B134", "#FBE7C6"]);
  function renderHBars(el, c) {
    // 水平堆疊長條圖（與桌面版出貨分頁相同）：c = {title, labels, series:[{name, values, color}]}
    if (!c || !c.labels || !c.labels.length) { el.innerHTML = `<div class="caption">${esc(t("（沒有資料）"))}</div>`; return; }
    const totals = c.labels.map((_, i) => c.series.reduce((s, x) => s + (x.values[i] || 0), 0));
    const max = Math.max(1, ...totals);
    let h = `<div class="hb-title">${esc(c.title)}</div><div class="hb-legend">${c.series.map((s) => `<span><i style="background:${esc(s.color)}"></i>${flagImg(s.name)}${esc(s.name)}</span>`).join("")}</div>`;
    c.labels.forEach((l, i) => {
      h += `<div class="hb-row"><span class="hb-lab" title="${esc(l)}">${flagImg(l)}${esc(l)}</span><span class="hb-bar">` +
        c.series.map((s) => {
          const v = s.values[i] || 0; if (!v) return "";
          const w = v / max * 100;
          return `<i class="${LIGHT.has(s.color.toUpperCase()) ? "dark" : ""}" style="width:${w}%;background:${esc(s.color)}" title="${esc(s.name)} ${v.toLocaleString()}">${w > 9 ? v.toLocaleString() : ""}</i>`;
        }).join("") + `</span><b>${totals[i].toLocaleString()}</b></div>`;
    });
    el.innerHTML = h;
    el.querySelectorAll(".hb-bar i").forEach((seg) => { if (seg.scrollWidth > seg.clientWidth) seg.textContent = ""; });
  }
  // ------------------------------------------------------------ 世界地圖（各國案件分布）
  // world.js：各國輪廓（等距圓柱投影到 1000×500）；countries.js：名稱 → ISO 代碼與中心座標
  const MAP_FILLS = ["#FBE7C6", "#F6CF86", "#F2B134", "#E8862B", "#B5544C"];
  let mapData = null, mapView = "fit";
  const project = (c) => [(c.lon + 180) / 360 * 1000, (90 - c.lat) / 180 * 500];
  function renderMap(el, m) {
    const W = window.RMA_WORLD;
    const rows = (m && m.rows || []).filter((r) => COUNTRIES[r.name]);
    if (!W || !rows.length) { el.innerHTML = `<div class="caption pad">${esc(t("（沒有資料）"))}</div>`; return; }
    // 視窗：fit = 剛好框住有資料的國家（留邊，至少 250×125，維持 2:1）；world = 全世界
    let vb = { x: 0, y: 0, w: 1000, h: 500 };
    if (mapView === "fit") {
      const pts = rows.map((r) => project(COUNTRIES[r.name]));
      const x0 = Math.min(...pts.map((p) => p[0])) - 45, x1 = Math.max(...pts.map((p) => p[0])) + 45;
      const y0 = Math.min(...pts.map((p) => p[1])) - 35, y1 = Math.max(...pts.map((p) => p[1])) + 35;
      let w = Math.max(250, x1 - x0), h = Math.max(125, y1 - y0);
      if (w / h > 2) h = w / 2; else w = h * 2;
      vb = { x: Math.max(0, Math.min(1000 - w, (x0 + x1) / 2 - w / 2)), y: Math.max(0, Math.min(500 - h, (y0 + y1) / 2 - h / 2)), w, h };
    }
    const u = vb.w / 900;   // 一個畫面像素約等於幾個地圖單位（旗子、字的大小用）
    const max = Math.max(1, ...rows.map((r) => r.cases));
    const lab = m.labels || {};
    const tip = (r) => `${r.label}\n${lab.cases || ""} ${r.cases.toLocaleString()}\n${lab.open || ""} ${r.open.toLocaleString()}\n${lab.shortage || ""} ${r.shortage.toLocaleString()}`;
    const byIso = {};
    for (const r of rows) byIso[COUNTRIES[r.name].iso3] = r;
    let h = `<svg viewBox="${vb.x} ${vb.y} ${vb.w} ${vb.h}" preserveAspectRatio="xMidYMid meet" role="img"><rect x="0" y="0" width="1000" height="500" class="sea"/><g stroke-width="${0.6 * u}">`;
    for (const [iso, d] of Object.entries(W.paths)) {
      const r = byIso[iso];
      if (!r) { h += `<path d="${d}"><title>${esc(W.names[iso] || iso)}</title></path>`; continue; }
      const lvl = Math.min(4, Math.round(Math.sqrt(r.cases / max) * 4));
      h += `<path d="${d}" class="has" style="fill:${MAP_FILLS[lvl]}"><title>${esc(tip(r))}</title></path>`;
    }
    h += `</g>`;
    // 旗子＋件數：先放在國家中心；擠在一起的互相推開（算上旗框與數字的高度），推開的用細線指回原來的位置
    const fw = 26 * u, fh = fw * 0.75, pad = 2 * u;
    const bw = fw + 2 * pad + 4 * u, bh = fh + 2 * pad + 15 * u;
    const pins = rows.map((r) => { const [x, y] = project(COUNTRIES[r.name]); return { r, x0: x, y0: y, x, y }; });
    for (let it = 0; it < 80; it++) {
      let moved = false;
      for (let i = 0; i < pins.length; i++) for (let j = i + 1; j < pins.length; j++) {
        const a = pins[i], b = pins[j];
        const dx = b.x - a.x, dy = b.y - a.y, ox = bw - Math.abs(dx), oy = bh - Math.abs(dy);
        if (ox <= 0 || oy <= 0) continue;
        moved = true;
        if (ox < oy) { const s = (ox / 2 + u) * (dx < 0 ? -1 : 1); a.x -= s; b.x += s; }
        else { const s = (oy / 2 + u) * (dy < 0 ? -1 : 1); a.y -= s; b.y += s; }
      }
      if (!moved) break;
    }
    for (const p of pins) {
      if (Math.hypot(p.x - p.x0, p.y - p.y0) > 4 * u) h += `<line class="lead" x1="${p.x0.toFixed(1)}" y1="${p.y0.toFixed(1)}" x2="${p.x.toFixed(1)}" y2="${p.y.toFixed(1)}" stroke-width="${1.2 * u}"/><circle class="lead" cx="${p.x0.toFixed(1)}" cy="${p.y0.toFixed(1)}" r="${2 * u}"/>`;
    }
    // 案件少的最後畫（在上層），才不會被大國的旗子蓋住
    for (const p of pins.slice().reverse()) {
      const r = p.r;
      h += `<g class="pin" transform="translate(${p.x.toFixed(1)},${p.y.toFixed(1)})"><title>${esc(tip(r))}</title>` +
        `<rect x="${-fw / 2 - pad}" y="${-fh / 2 - pad}" width="${fw + 2 * pad}" height="${fh + 2 * pad}" rx="${2 * u}" stroke-width="${u}"/>` +
        `<image href="flags/${COUNTRIES[r.name].iso2}.svg" x="${-fw / 2}" y="${-fh / 2}" width="${fw}" height="${fh}"/>` +
        `<text y="${fh / 2 + 12 * u}" font-size="${11 * u}" text-anchor="middle" stroke-width="${3 * u}">${r.cases.toLocaleString()}</text></g>`;
    }
    el.innerHTML = h + `</svg>`;
  }
  function renderMapChips(el, m) {
    // 地圖下方：每個國家一個小標籤（國旗、名稱、件數），沒有座標的國家（例如「未填」）也列在這裡
    el.innerHTML = (m && m.rows || []).map((r) => `<span title="${esc(r.label)}">${flagImg(r.name)}${esc(r.label)} <b>${r.cases.toLocaleString()}</b></span>`).join("");
  }
  document.querySelectorAll("#src-map-view button").forEach((b) => b.addEventListener("click", () => {
    mapView = b.dataset.view;
    document.querySelectorAll("#src-map-view button").forEach((x) => x.classList.toggle("active", x === b));
    if (mapData) renderMap($("src-map-svg"), mapData);
  }));
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
  // 表格匯出：按下時才在 Python 端轉成 Excel（第一次可能要等圖表與匯出套件載完）
  async function exportTable(btn, key) {
    if (!key) { toast(t("（沒有資料）")); return; }
    busy(btn, true, t("產生中…"));
    try { const r = await call("table_xlsx", { key }); if (r.data) downloadB64(r.data, r.name, XLSX_MIME); else toast(t("（沒有資料）")); }
    catch (e) { toast(t("匯出失敗：") + friendly(e), 8000); } finally { busy(btn, false); }
  }

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
    if (window.RMAScene) { RMAScene.setTz(PREFS.tz); RMAScene.resize(); }
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
    $("src-status-hint").textContent = "";   // 說明放在明細視窗裡
    $("src-status").style.display = "block";   // 先顯示再畫圖：色塊的文字要量得到寬度才知道放不放得下
    // 頁面上只放精簡版（分組＋合計）；完整的兩層表放在「顯示原始狀態明細」視窗
    statusDetail = s.status_table;
    const full = s.status_table;
    const gi = full.cls.map((c, r) => c === "grp" ? r : -1).filter((r) => r >= 0);
    renderTable($("src-status-tbl"), { columns: [full.columns[0], full.columns[2]], defs: [full.defs[0], full.defs[2]], numeric: [full.numeric[0], full.numeric[2]],
                                        rows: gi.map((r) => [full.rows[r][0], full.rows[r][2]]), warn: gi.map(() => false), cls: gi.map(() => ""), total: gi.length });
    renderHStack($("src-status-chart"), s.status_chart);
    mapData = s.map || null;
    $("src-map").style.display = mapData && mapData.rows.length ? "block" : "none";
    if (mapData) {
      $("src-map-title").textContent = mapData.title; $("src-map-title").removeAttribute("data-t");
      $("src-map-note").textContent = mapData.note || "";
      renderMap($("src-map-svg"), mapData); renderMapChips($("src-map-chips"), mapData);
    }
    analyses = s.analyses || [];
  }
  // 分析期間：下拉選單（值是內部中文字串，Python 端 period_start_for 對應）；「自訂起日」才顯示日期欄
  function periodArgs() {
    const period = $("period").value;
    let start = null;
    if (period === "自訂起日") {
      const v = $("start").value.trim();
      start = v ? normDate(v) : null;
      if (v && start === null) { toast(t("分析起日格式不對，請輸入像 2026-08-01 這樣的日期，或留空。"), 6000); return null; }
    }
    store.set("rma_period", period);
    return { period, start };
  }
  $("period").addEventListener("change", () => { $("start-wrap").style.display = $("period").value === "自訂起日" ? "" : "none"; store.set("rma_period", $("period").value); });
  $("load").addEventListener("click", async () => {
    const f = $("aio").files[0];
    if (!f) { toast(t("請先選擇 AIO 匯出檔")); return; }
    const pa = periodArgs();
    if (!pa) return;
    busy($("load"), true, t("分析中…"));
    setT($("status"), "載入中…");
    startPhase("load");
    try {
      const buffer = await f.arrayBuffer();
      const s = await call("load", { buffer, period: pa.period, start: pa.start }, [buffer]);
      loaded = true; curFile = f.name; source = "file";
      resetTabs(); renderSummary(s); refreshActive(); dbUI();
      log(t("完成。"));
    } catch (e) {
      log(t("錯誤：") + e.message); toast(t("載入失敗：") + friendly(e), 8000);
      setT($("status"), "載入失敗");
      phaseDone = "fail"; renderSteps(loadPct());
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
      if (demandData.weekly_chart) renderStackedBars($("d-chart"), demandData.weekly_chart);
      else $("d-chart").innerHTML = "";
      $("d-fc-title").textContent = demandData.forecast_title;
      renderTable($("d-forecast"), demandData.forecast);
      renderTable($("d-top"), demandData.top);
    } catch (e) { toast(friendly(e), 8000); } finally { busy($("d-run"), false); }
  }
  $("d-run").addEventListener("click", refreshDemand);
  $("d-by").addEventListener("change", refreshDemand);
  $("d-weekly-dl").addEventListener("click", (e) => demandData && exportTable(e.currentTarget, demandData.weekly_x));
  $("d-top-dl").addEventListener("click", (e) => demandData && exportTable(e.currentTarget, demandData.top_x));

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
  $("s-parts-dl").addEventListener("click", (e) => stockData && exportTable(e.currentTarget, stockData.parts_x));
  $("s-cases-dl").addEventListener("click", (e) => stockData && exportTable(e.currentTarget, stockData.cases_x));

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
  $("po-late-dl").addEventListener("click", (e) => poData && exportTable(e.currentTarget, poData.late_x));

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
  $("h-backlog-dl").addEventListener("click", (e) => shipData && exportTable(e.currentTarget, shipData.backlog_x));
  $("h-exp-dl").addEventListener("click", (e) => shipData && exportTable(e.currentTarget, shipData.expiring_x));

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
          <div id="r-chart"><div class="boot"><span class="spin"></span>${esc(extrasState === "ready" ? t("畫圖中…") : t("等待套件載入…"))}</div></div>
        </div>
        <div class="row">
          <label class="inline"><span>${esc(t("表格："))}</span><select id="r-table">${opts}</select></label>
          <button class="btn" id="r-table-dl">${esc(t("匯出此表"))}</button>
          <span class="small">${esc(t("欄名有 ⓘ 的，滑鼠停上去看說明"))}</span>
        </div>
        <div class="tbl" id="r-table-body"></div>`;
      $("r-table").onchange = () => showTable(key, $("r-table").value);
      $("r-table-dl").onclick = (e) => { if (currentTable) exportTable(e.currentTarget, currentTable.x); };
      if (a.tables.length) showTable(key, a.tables[0][0]);
      showChart(key, a.title);
    } catch (e) { if (my === showSeq) $("r-body").innerHTML = `<div class="notice">${esc(t("無法顯示："))}${esc(friendly(e))}</div>`; }
  }
  async function showChart(key, title) {
    // 圖另外取：文字與表格先出現，圖畫好再放上（第一次要等圖表套件載完）
    const my = showSeq;
    try {
      const r = await call("chart", { key });
      const el = $("r-chart");
      if (my !== showSeq || !el) return;
      el.innerHTML = r.chart ? `<img src="data:image/png;base64,${r.chart}" alt="${esc(title)}">` : `<div class="caption">${esc(t("（這一項沒有圖）"))}</div>`;
    } catch (e) { const el = $("r-chart"); if (my === showSeq && el) el.innerHTML = `<div class="caption">${esc(t("無法顯示："))}${esc(friendly(e))}</div>`; }
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
    const waiting = extrasState !== "ready";
    busy(btn, true, waiting ? t("等待套件載入…") : t("產生中…"));
    if (waiting) toast(t("圖表與匯出套件還在背景載入，載完會自動開始。"), 6000);
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
  function openDetail(title, hint, tj) {
    $("detail-title").textContent = title; $("detail-hint").textContent = hint || "";
    $("detail-body").innerHTML = '<div class="tbl"></div>'; renderTable($("detail-body").firstChild, tj, 5000);
    $("detail").style.display = "flex"; $("detail-body").scrollTop = 0; $("detail-body").focus();
  }
  function closeDetail() { if ($("detail").style.display !== "flex") return; $("detail").style.display = "none"; $("src-status-more").focus(); }
  $("src-status-more").addEventListener("click", () => { if (!statusDetail) { needData(); return; } openDetail(t("各組包含的 AIO 原始狀態"), t("粗體列是狀態分組，底下是歸入該組的 AIO 原始狀態。"), statusDetail); });
  $("detail-close").addEventListener("click", closeDetail);
  $("detail").addEventListener("click", (e) => { if (e.target === $("detail")) closeDetail(); });
  function closeGloss() { if ($("gloss").style.display === "none") return; $("gloss").style.display = "none"; $("gloss-btn").focus(); }
  $("gloss-close").addEventListener("click", closeGloss);
  $("gloss").addEventListener("click", (e) => { if (e.target === $("gloss")) closeGloss(); });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") { closeGloss(); closeDetail(); } });

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

  // ------------------------------------------------------------ 資料庫（Supabase）：留存每次匯入，才看得到多期趨勢
  // 程式在 db.js（登入、匯入紀錄、案件快照）；資料表在 supabase/schema.sql；設定在 config.js
  const DB = window.RMADB || null;
  let imports = [];
  let source = "file";   // 目前的資料來自 file（匯出檔）或 db（資料庫）
  function dbConfigured() { return !!(DB && DB.configured()); }
  function dbUI() {
    const user = dbConfigured() ? DB.user() : null;
    $("db-status").textContent = !dbConfigured() ? "" : user ? t("已登入：") + user : t("請登入資料庫（帳號由管理員建立）");
    $("db-note").textContent = dbConfigured() ? "" : t("尚未設定資料庫：照 README 的「資料庫（Supabase）」一節建立專案，再把網址與金鑰填進 config.js。");
    $("db-login-form").style.display = dbConfigured() && !user ? "" : "none";
    $("db-actions").style.display = user ? "" : "none";
    $("db-history").style.display = user && imports.length ? "block" : "none";
    $("db-save").disabled = !(user && loaded && source === "file");
  }
  function fmtTs(iso) {
    if (!iso) return "";
    try { return new Intl.DateTimeFormat(LANG === "en" ? "en-GB" : "zh-TW", { timeZone: PREFS.tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).format(new Date(iso)); }
    catch (e) { return String(iso).slice(0, 16).replace("T", " "); }
  }
  function renderImports() {
    const el = $("db-imports");
    if (!imports.length) { el.innerHTML = ""; $("db-chart").innerHTML = ""; return; }
    const k = (i, key) => Number((i.kpis && i.kpis[key]) || 0);
    const heads = ["資料日期", "檔名", "案件數", "在途", "缺料", "匯入者", "匯入時間", ""];
    let h = "<table><thead><tr>" + heads.map((c, i) => `<th class="${i >= 2 && i <= 4 ? "num" : ""}">${esc(c ? t(c) : "")}</th>`).join("") + "</tr></thead><tbody>";
    for (const i of imports) {
      h += `<tr><td>${esc(i.data_date)}</td><td title="${esc(i.file_name || "")}">${esc(i.file_name || "")}</td><td class="num">${(k(i, "cases") || Number(i.row_count || 0)).toLocaleString()}</td>` +
        `<td class="num">${k(i, "open").toLocaleString()}</td><td class="num">${k(i, "shortage").toLocaleString()}</td><td>${esc(i.imported_by || "")}</td><td>${esc(fmtTs(i.imported_at))}</td>` +
        `<td><button class="btn sm db-del" data-id="${esc(i.id)}" data-date="${esc(i.data_date)}">${esc(t("刪除"))}</button></td></tr>`;
    }
    el.innerHTML = h + "</tbody></table>";
    const asc = imports.slice().sort((a, b) => String(a.data_date).localeCompare(String(b.data_date)));
    renderBars($("db-chart"), asc.map((i) => i.data_date), asc.map((i) => k(i, "open")), { title: t("歷次匯入的未結案件數"), ylabel: t("件數"), note: t("每根長條是那次匯入時的未結案件數（整份檔案，不受分析期間影響）。") });
  }
  async function refreshImports() {
    if (!dbConfigured() || !DB.user()) { imports = []; renderImports(); dbUI(); return; }
    try { imports = (await DB.listImports()) || []; }
    catch (e) { imports = []; toast(t("讀取匯入紀錄失敗：") + dbFriendly(e), 6000); }
    renderImports(); dbUI();
  }
  async function dbLogin() {
    const email = $("db-email").value.trim(), pass = $("db-pass").value;
    if (!email || !pass) { toast(t("請輸入 email 與密碼")); return; }
    busy($("db-login"), true, t("登入中…"));
    try { await DB.login(email, pass); $("db-pass").value = ""; dbUI(); await refreshImports(); }
    catch (e) { toast(dbFriendly(e), 6000); }
    finally { busy($("db-login"), false); }
  }
  async function dbSave() {
    if (!loaded || source !== "file") { toast(t("從資料庫載入的資料不需要再存；請載入 AIO 匯出檔後再存。"), 6000); return; }
    busy($("db-save"), true, t("存入中…"));
    try {
      const p = await call("db_payload", {});
      const existing = imports.find((i) => i.data_date === p.data_date);
      if (existing && !confirm(tf("資料庫已經有 {0} 的匯入（{1}），要覆蓋嗎？", p.data_date, existing.file_name || ""))) return;
      startPhase("save"); onStep({ phase: "save", id: "meta", state: "run" });
      if (existing) await DB.deleteImport(existing.id);
      const imp = await DB.createImport({ data_date: p.data_date, file_name: curFile, row_count: 0, columns: p.columns, kpis: p.kpis, imported_by: DB.user() });
      onStep({ phase: "save", id: "meta", state: "done" }); onStep({ phase: "save", id: "rows", state: "run", i: 0, n: p.rows.length });
      await DB.upsertCases(imp.id, p.rows, (i, n) => { onStep({ phase: "save", id: "rows", state: "run", i, n }); log(tf("已存入資料庫：{0} 件", i)); });
      onStep({ phase: "save", id: "rows", state: "done" });
      toast(tf("已存入資料庫：{0} 件", p.rows.length));
      await refreshImports();
    } catch (e) {
      log(t("錯誤：") + e.message); toast(t("存入資料庫失敗：") + dbFriendly(e), 8000);
      phaseDone = "fail"; renderSteps(loadPct());
    } finally { busy($("db-save"), false); dbUI(); }
  }
  async function dbLoad() {
    if (!ready) { toast(t("瀏覽器版 Python 還在啟動")); return; }
    const pa = periodArgs();
    if (!pa) return;
    busy($("db-load"), true, t("載入中…"));
    setT($("status"), "載入中…");
    startPhase("load"); stepLabel.read = "從資料庫讀取"; onStep({ phase: "load", id: "read", state: "run" });
    try {
      imports = (await DB.listImports()) || [];
      renderImports();
      if (!imports.length) throw new Error(t("資料庫裡還沒有匯入任何檔案"));
      const refDate = imports.reduce((m, i) => (String(i.data_date) > m ? String(i.data_date) : m), "");
      const r = await call("period_from", { period: pa.period, custom: pa.start, ref_date: refDate });
      log(`${t("從資料庫讀取")} ${r.from || "…"} → ${refDate}`);
      const objs = await DB.fetchCases(r.from, null, (n) => { log(tf("已下載 {0} 件…", n.toLocaleString())); });
      // 欄位順序：各次匯入的欄位聯集（新的在前，CaseID 永遠第一欄），最後加 _data_date
      const cols = ["CaseID"];
      for (const i of imports) for (const c of (Array.isArray(i.columns) ? i.columns : [])) if (c && !cols.includes(c)) cols.push(c);
      for (const o of objs) for (const c of Object.keys(o)) if (c !== "_data_date" && !cols.includes(c)) cols.push(c);
      const rows = [cols.concat(["_data_date"])].concat(objs.map((o) => cols.map((c) => (o[c] === undefined ? null : o[c])).concat([o._data_date || null])));
      onStep({ phase: "load", id: "read", state: "done" });
      const s = await call("load", { rows, period: pa.period, start: pa.start, ref_date: refDate });
      loaded = true; source = "db"; curFile = tf("資料庫（{0} 次匯入）", imports.length);
      resetTabs(); renderSummary(s); refreshActive();
      log(t("完成。"));
    } catch (e) {
      log(t("錯誤：") + e.message); toast(t("載入失敗：") + dbFriendly(e), 8000);
      setT($("status"), "載入失敗");
      phaseDone = "fail"; renderSteps(loadPct());
    } finally { busy($("db-load"), false); dbUI(); }
  }
  async function dbInit() {
    if (!dbConfigured()) { dbUI(); return; }
    try { await DB.restore(); } catch (e) { /* 沒有登入就算了 */ }
    dbUI();
    if (DB.user()) refreshImports();
  }
  if (DB) {
    $("db-login").addEventListener("click", dbLogin);
    $("db-pass").addEventListener("keydown", (e) => { if (e.key === "Enter") dbLogin(); });
    $("db-email").addEventListener("keydown", (e) => { if (e.key === "Enter") $("db-pass").focus(); });
    $("db-logout").addEventListener("click", async () => { try { await DB.logout(); } catch (e) { /* 忽略 */ } imports = []; renderImports(); dbUI(); });
    $("db-save").addEventListener("click", dbSave);
    $("db-load").addEventListener("click", dbLoad);
    $("db-imports").addEventListener("click", async (e) => {
      const b = e.target.closest(".db-del"); if (!b) return;
      if (!confirm(tf("要刪除 {0} 的匯入嗎？那次匯入的案件版本會一併刪除（更早匯入的版本不受影響）。", b.dataset.date))) return;
      try { await DB.deleteImport(b.dataset.id); toast(t("已刪除")); } catch (err) { toast(dbFriendly(err), 6000); }
      refreshImports();
    });
  }

  // ------------------------------------------------------------ 啟動
  if (DB) DB.configure(CFG);
  const savedPeriod = store.get("rma_period");
  if (savedPeriod && Array.from($("period").options).some((o) => o.value === savedPeriod)) $("period").value = savedPeriod;
  $("start-wrap").style.display = $("period").value === "自訂起日" ? "" : "none";
  applyUI();
  applyPrefs();
  renderActions();
  startPhase("boot");
  if (window.RMAScene) RMAScene.init({ lang: LANG, tz: PREFS.tz, sign: "RMA" });
  call("boot", Object.assign({ lang: LANG }, CFG)).catch(bootFailed);
  dbInit();
})();
