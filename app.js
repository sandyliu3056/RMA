/* 售後零件規劃平台（瀏覽器版）主畫面。
 * 所有分析交給 worker.js 裡的 Pyodide（瀏覽器版 Python）執行，這裡只負責畫面與互動。 */
(function () {
  "use strict";
  const DEFAULTS = {
    pyodideBase: "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/",
    xlsxUrl: "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
    engineUrl: "rma_engine.py",
    glueUrl: "web_glue.py",
    fonts: ["fonts/NotoSansTC-Regular.otf", "fonts/NotoSansTC-Bold.otf"],
    wheels: ["openpyxl", "python-pptx"],   // micropip 從 PyPI 安裝；離線測試可改成本機 .whl 網址
  };
  const CFG = Object.assign({}, DEFAULTS, window.APP_CONFIG || {});
  const UNITS = ["各國規劃人員", "供應商", "總部服務團隊", "倉庫／物流", "維修據點", "其他"];
  const $ = (id) => document.getElementById(id);
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  // ------------------------------------------------------------ worker
  const worker = new Worker("worker.js");
  const pending = new Map();
  let seq = 0, ready = false, loaded = false;
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
    if (m.type === "ready") {
      ready = true;
      $("boot").innerHTML = "瀏覽器版 Python 已就緒，請選擇 AIO 匯出檔。";
      $("aio").disabled = false; $("load").disabled = false;
      return;
    }
    const p = pending.get(m.id);
    if (!p) return;
    pending.delete(m.id);
    m.ok ? p.resolve(m.result) : p.reject(new Error(m.error));
  };
  worker.onerror = (e) => { $("boot-msg").textContent = "啟動失敗：" + e.message; toast("啟動失敗：" + e.message); };

  // ------------------------------------------------------------ 小工具
  function log(msg) {
    const el = $("log");
    el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
    if (!ready) $("boot-msg").textContent = msg;
  }
  let toastTimer = null;
  function toast(msg, ms) {
    const t = $("toast"); t.textContent = msg; t.style.display = "block";
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.style.display = "none"), ms || 4000);
  }
  function busy(btn, on, label) {
    if (!btn) return;
    if (on) { btn.dataset.label = btn.textContent; btn.textContent = label || "處理中…"; btn.disabled = true; }
    else { btn.textContent = btn.dataset.label || btn.textContent; btn.disabled = false; }
  }
  function esc(s) { return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

  function renderTable(el, tj, maxRows) {
    if (!tj || !tj.columns.length) { el.innerHTML = '<div class="caption" style="padding:10px">（沒有資料）</div>'; return; }
    const rows = tj.rows.slice(0, maxRows || 2000);
    let h = "<table><thead><tr>" + tj.columns.map((c, i) => `<th class="${tj.numeric[i] ? "num" : ""}">${esc(c)}</th>`).join("") + "</tr></thead><tbody>";
    for (const r of rows) h += "<tr>" + r.map((v, i) => `<td class="${tj.numeric[i] ? "num" : ""}">${esc(v)}</td>`).join("") + "</tr>";
    h += "</tbody></table>";
    if (tj.rows.length > rows.length) h += `<div class="caption" style="padding:6px 10px">只顯示前 ${rows.length.toLocaleString()} 列，共 ${tj.rows.length.toLocaleString()} 列；匯出 Excel 可取得全部。</div>`;
    el.innerHTML = h;
  }
  function renderKpis(el, items) {
    el.innerHTML = items.map(([k, v]) => `<div class="kpi"><div class="k">${esc(k)}</div><div class="v">${esc(v)}</div></div>`).join("");
  }
  function renderBars(el, labels, values) {
    const max = Math.max(1, ...values);
    el.innerHTML = labels.map((l, i) => `<div class="b"><b>${values[i].toLocaleString()}</b><i style="height:${Math.round(values[i] / max * 150)}px"></i><span>${esc(l)}</span></div>`).join("");
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
  function needData() { if (!loaded) { toast("請先在「資料來源」載入 AIO 匯出檔"); return false; } return true; }

  // ------------------------------------------------------------ 時鐘
  function tick() {
    const now = new Date();
    const tw = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
    const p = (n) => String(n).padStart(2, "0");
    $("clock").textContent = `${p(tw.getHours())}:${p(tw.getMinutes())}`;
    $("date").textContent = `${tw.getFullYear()}/${p(tw.getMonth() + 1)}/${p(tw.getDate())} ${["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][tw.getDay()]}`;
  }
  tick(); setInterval(tick, 30000);

  // ------------------------------------------------------------ 分頁
  const refreshers = { demand: refreshDemand, stock: refreshStock, ship: refreshShip, report: refreshReportList };
  $("tabs").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-tab]"); if (!b) return;
    document.querySelectorAll("#tabs button").forEach((x) => x.classList.toggle("active", x === b));
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x.id === "tab-" + b.dataset.tab));
    const name = b.dataset.tab;
    if (loaded && refreshers[name] && !loadedTabs.has(name)) { loadedTabs.add(name); refreshers[name](); }
  });

  // ------------------------------------------------------------ 資料來源
  $("load").addEventListener("click", async () => {
    const f = $("aio").files[0];
    if (!f) { toast("請先選擇 AIO 匯出檔"); return; }
    busy($("load"), true, "分析中…");
    $("status").textContent = "載入中…";
    try {
      const buffer = await f.arrayBuffer();
      const s = await call("load", { buffer, start: $("start").value.trim() }, [buffer]);
      loaded = true; loadedTabs.clear();
      $("status").textContent = `${f.name}｜${s.status}`;
      renderKpis($("src-kpis"), s.kpis);
      $("src-caption").textContent = s.caption;
      renderTable($("src-status-tbl"), s.status_table);
      renderBars($("src-status-chart"), s.status_chart.labels, s.status_chart.values);
      $("src-status").style.display = "block";
      window.__analyses = s.analyses;
      log("完成。");
    } catch (e) {
      log("錯誤：" + e.message); toast("載入失敗：" + e.message, 8000);
      $("status").textContent = "載入失敗";
    } finally { busy($("load"), false); }
  });

  // ------------------------------------------------------------ 需求規劃
  let demandData = null;
  async function refreshDemand() {
    if (!needData()) return;
    busy($("d-run"), true);
    try {
      const weeks = +$("d-weeks").value || 4;
      demandData = await call("tab_demand", { by: $("d-by").value, weeks });
      renderTable($("d-weekly"), demandData.weekly);
      if (demandData.weekly_chart) renderBars($("d-chart"), demandData.weekly_chart.labels, demandData.weekly_chart.values.map(Math.round));
      $("d-fc-title").textContent = `未來 ${weeks} 週需求推估（最近 4 個完整週平均 × 週數；高峰＝最高週 × 週數）`;
      renderTable($("d-forecast"), demandData.forecast);
      renderTable($("d-top"), demandData.top);
    } catch (e) { toast(e.message, 8000); } finally { busy($("d-run"), false); }
  }
  $("d-run").addEventListener("click", refreshDemand);
  $("d-weekly-dl").addEventListener("click", () => demandData && downloadB64(demandData.weekly_xlsx, "週申請零件量.xlsx", XLSX_MIME));
  $("d-top-dl").addEventListener("click", () => demandData && downloadB64(demandData.top_xlsx, "Top料號.xlsx", XLSX_MIME));

  // ------------------------------------------------------------ 庫存與缺料
  let stockData = null;
  async function refreshStock() {
    if (!needData()) return;
    busy($("s-run"), true);
    try {
      const days = +$("s-days").value || 14;
      stockData = await call("tab_stock", { days });
      renderKpis($("s-kpis"), stockData.kpis);
      renderTable($("s-parts"), stockData.parts);
      $("s-cases-title").textContent = `催料案件清單（等待 ≥ ${days} 天）`;
      renderTable($("s-cases"), stockData.cases);
    } catch (e) { toast(e.message, 8000); } finally { busy($("s-run"), false); }
  }
  $("s-run").addEventListener("click", refreshStock);
  $("s-parts-dl").addEventListener("click", () => stockData && downloadB64(stockData.parts_xlsx, "缺料料號.xlsx", XLSX_MIME));
  $("s-cases-dl").addEventListener("click", () => stockData && downloadB64(stockData.cases_xlsx, "催料清單.xlsx", XLSX_MIME));

  // ------------------------------------------------------------ 訂單與交期
  let poData = null;
  $("po-tmpl").addEventListener("click", async () => {
    if (!ready) { toast("瀏覽器版 Python 還在啟動"); return; }
    try { downloadB64(await call("po_template", {}), "零件訂單範本.xlsx", XLSX_MIME); } catch (e) { toast(e.message); }
  });
  $("po-file").addEventListener("change", async () => {
    const f = $("po-file").files[0]; if (!f || !ready) return;
    try {
      const buffer = await f.arrayBuffer();
      poData = await call("po", { buffer }, [buffer]);
      renderKpis($("po-kpis"), poData.kpis);
      renderTable($("po-sup"), poData.suppliers);
      renderTable($("po-late"), poData.late);
      $("po-body").style.display = "block";
    } catch (e) { toast("匯入失敗：" + e.message, 8000); }
  });
  $("po-late-dl").addEventListener("click", () => poData && downloadB64(poData.late_xlsx, "逾期訂單.xlsx", XLSX_MIME));

  // ------------------------------------------------------------ 出貨與到貨
  let shipData = null;
  async function refreshShip() {
    if (!needData()) return;
    busy($("h-run"), true);
    try {
      shipData = await call("tab_ship", { days: +$("h-days").value || 30 });
      renderKpis($("h-kpis"), shipData.kpis);
      renderTable($("h-backlog"), shipData.backlog);
      renderTable($("h-transit"), shipData.transit);
      renderTable($("h-exp"), shipData.expiring);
    } catch (e) { toast(e.message, 8000); } finally { busy($("h-run"), false); }
  }
  $("h-run").addEventListener("click", refreshShip);
  $("h-backlog-dl").addEventListener("click", () => shipData && downloadB64(shipData.backlog_xlsx, "待出貨積壓.xlsx", XLSX_MIME));
  $("h-exp-dl").addEventListener("click", () => shipData && downloadB64(shipData.expiring_xlsx, "保固到期在途案.xlsx", XLSX_MIME));

  // ------------------------------------------------------------ 報表與分析
  let currentKey = null, currentTable = null;
  function refreshReportList() {
    if (!needData()) return;
    const list = $("r-list");
    list.innerHTML = (window.__analyses || []).map(([k, t]) => `<button data-key="${k}">${k}  ${esc(t)}</button>`).join("");
    list.onclick = (e) => { const b = e.target.closest("button[data-key]"); if (b) showAnalysis(b.dataset.key); };
    if (window.__analyses && window.__analyses.length) showAnalysis(window.__analyses[0][0]);
  }
  async function showAnalysis(key) {
    currentKey = key;
    document.querySelectorAll("#r-list button").forEach((b) => b.classList.toggle("active", b.dataset.key === key));
    $("r-body").innerHTML = '<div class="boot"><span class="spin"></span>畫圖中…</div>';
    try {
      const a = await call("analysis", { key });
      const tables = a.tables.map((t) => `<option>${esc(t)}</option>`).join("");
      $("r-body").innerHTML = `
        <h2 style="margin:0 0 2px">${esc(a.key)} ${esc(a.title)}</h2>
        <div class="caption">${esc(a.subtitle)}</div>
        <div class="rbody">
          <div><ul>${a.bullets.map((b) => `<li>${esc(b)}</li>`).join("")}</ul><p>${esc(a.paragraph)}</p><p><b>結論：</b>${esc(a.conclusion)}</p></div>
          <div>${a.chart ? `<img src="data:image/png;base64,${a.chart}" alt="${esc(a.title)}">` : '<div class="caption">（這一項沒有圖）</div>'}</div>
        </div>
        <div class="row"><label>表格<select id="r-table">${tables}</select></label><button class="btn" id="r-table-dl">匯出此表</button></div>
        <div class="tbl" id="r-table-body"></div>`;
      $("r-table").onchange = () => showTable(key, $("r-table").value);
      $("r-table-dl").onclick = () => currentTable && currentTable.xlsx && downloadB64(currentTable.xlsx, `${key}_${$("r-table").value}.xlsx`, XLSX_MIME);
      if (a.tables.length) showTable(key, a.tables[0]);
    } catch (e) { $("r-body").innerHTML = `<div class="notice">無法顯示：${esc(e.message)}</div>`; }
  }
  async function showTable(key, name) {
    try { currentTable = await call("analysis_table", { key, name }); renderTable($("r-table-body"), currentTable.table); }
    catch (e) { toast(e.message); }
  }
  async function exportKind(btn, kind, name, mime) {
    if (!needData()) return;
    busy(btn, true, "產生中…");
    try { downloadB64(await call("export", { kind }), name, mime); }
    catch (e) { toast("匯出失敗：" + e.message, 8000); } finally { busy(btn, false); }
  }
  $("r-xlsx").addEventListener("click", (e) => exportKind(e.target, "xlsx", "RMA分析結果.xlsx", XLSX_MIME));
  $("r-pptx").addEventListener("click", (e) => exportKind(e.target, "pptx", "RMA分析報告.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation"));
  $("r-charts").addEventListener("click", (e) => exportKind(e.target, "charts", "charts.zip", "application/zip"));

  // ------------------------------------------------------------ 跨單位協調（存在 localStorage）
  const KEY = "rma_actions";
  let actions = [];
  try { actions = JSON.parse(localStorage.getItem(KEY) || "[]"); } catch (e) { actions = []; }
  function saveActions() { try { localStorage.setItem(KEY, JSON.stringify(actions)); } catch (e) { toast("無法儲存待辦（瀏覽器儲存空間不可用）"); } renderActions(); }
  const today = () => new Date().toISOString().slice(0, 10);
  const plusDays = (n) => { const d = new Date(); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
  $("c-unit").innerHTML = UNITS.map((u) => `<option>${u}</option>`).join("");
  $("c-due").value = plusDays(7);

  function addAction(unit, topic, owner, due, status) {
    const id = actions.length ? Math.max(...actions.map((a) => a.id)) + 1 : 1;
    actions.push({ id, 對象單位: unit, 議題: topic, 負責人: owner || "", 到期日: due, 狀態: status || "進行中", 建立日: today() });
  }
  function renderActions() {
    const el = $("c-table");
    if (!actions.length) { el.innerHTML = '<div class="caption" style="padding:10px">目前沒有待辦。</div>'; return; }
    const sorted = [...actions].sort((a, b) => (a.狀態 === "已完成") - (b.狀態 === "已完成") || a.到期日.localeCompare(b.到期日));
    const td = today();
    let h = "<table><thead><tr><th>id</th><th>對象單位</th><th>議題</th><th>負責人</th><th>到期日</th><th>狀態</th><th>建立日</th><th></th></tr></thead><tbody>";
    for (const a of sorted) {
      const late = ["進行中", "待回覆"].includes(a.狀態) && a.到期日 < td;
      h += `<tr data-id="${a.id}"><td>${a.id}</td><td>${esc(a.對象單位)}</td><td style="white-space:normal;min-width:320px">${esc(a.議題)}</td>
        <td><input data-f="負責人" value="${esc(a.負責人)}" size="8"></td>
        <td><input data-f="到期日" type="date" value="${esc(a.到期日)}" class="${late ? "late" : ""}"></td>
        <td><select data-f="狀態">${["進行中", "待回覆", "已完成", "取消"].map((s) => `<option ${s === a.狀態 ? "selected" : ""}>${s}</option>`).join("")}</select></td>
        <td>${esc(a.建立日)}</td><td><button class="btn" data-del="${a.id}">刪除</button></td></tr>`;
    }
    el.innerHTML = h + "</tbody></table>";
  }
  $("c-table").addEventListener("change", (e) => {
    const f = e.target.dataset.f; if (!f) return;
    const id = +e.target.closest("tr").dataset.id;
    const a = actions.find((x) => x.id === id); if (a) { a[f] = e.target.value; saveActions(); }
  });
  $("c-table").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-del]"); if (!b) return;
    actions = actions.filter((x) => x.id !== +b.dataset.del); saveActions();
  });
  $("c-add").addEventListener("click", () => {
    const topic = $("c-topic").value.trim();
    if (!topic) { toast("請填議題"); return; }
    addAction($("c-unit").value, topic, $("c-owner").value.trim(), $("c-due").value || plusDays(7), $("c-status").value);
    $("c-topic").value = ""; saveActions();
  });
  $("c-seed").addEventListener("click", async () => {
    if (!needData()) return;
    try {
      const seeds = await call("conclusions", {});
      let n = 0;
      for (const [, unit, topic] of seeds) if (!actions.some((a) => a.議題 === topic)) { addAction(unit, topic, "", plusDays(14), "進行中"); n++; }
      saveActions(); toast(`已新增 ${n} 筆待辦`);
    } catch (e) { toast(e.message); }
  });
  $("c-dl").addEventListener("click", () => downloadBlob(new Blob([JSON.stringify(actions, null, 2)], { type: "application/json" }), "協調待辦.json"));
  $("c-restore").addEventListener("change", async () => {
    const f = $("c-restore").files[0]; if (!f) return;
    try { const arr = JSON.parse(await f.text()); if (!Array.isArray(arr)) throw new Error("格式不對"); actions = arr; saveActions(); toast(`已還原 ${arr.length} 筆`); }
    catch (e) { toast("讀取失敗：" + e.message); }
    $("c-restore").value = "";
  });
  $("c-csv").addEventListener("click", () => {
    if (!actions.length) { toast("目前沒有待辦"); return; }
    const cols = ["id", "對象單位", "議題", "負責人", "到期日", "狀態", "建立日"];
    const q = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
    const csv = "﻿" + cols.join(",") + "\n" + actions.map((a) => cols.map((c) => q(a[c])).join(",")).join("\n");
    downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "跨單位協調待辦.csv");
  });
  renderActions();

  // ------------------------------------------------------------ 啟動
  call("boot", CFG).catch((e) => { $("boot-msg").textContent = "啟動失敗：" + e.message; toast("啟動失敗：" + e.message, 10000); });
})();
