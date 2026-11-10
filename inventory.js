/* 「庫存報告」分頁：月底庫存報告（Stock Holding by Ageing）的自動計算。
 * 讀 auto mail 的 CSS Inventory、PLS Picklist、PLS Stock，與（選填）主檔 Stock Holding by Ageing，
 * 在 inv_worker.js 裡用 SheetJS 計算，不需要等瀏覽器版 Python 啟動；檔案不會上傳。
 * 介面字串用 data-i="中文" 標示，英文在下方 EN 對照表（獨立於 i18n.py，不影響 RMA 分析的字典）。 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const V = "2026-11-10a";
  const XLSX_URL = (window.APP_CONFIG && window.APP_CONFIG.xlsxUrl) || "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

  // ------------------------------------------------------------ 語言
  const EN = {
    "庫存報告": "Inventory report",
    "月底庫存報告（Stock Holding by Ageing）": "Month-end inventory report (Stock Holding by Ageing)",
    "CSS Inventory（必填）": "CSS Inventory (required)",
    "PLS Picklist（選填）": "PLS Picklist (optional)",
    "PLS Stock（選填）": "PLS Stock (optional)",
    "主檔 Stock Holding by Ageing（選填）": "Master Stock Holding by Ageing (optional)",
    "資料月份": "Data month",
    "計算": "Calculate",
    "匯出 Excel": "Export Excel",
    "選擇 auto mail 的三個檔案（AAP_Inventory_CSS_Inventory、PLS_Picklist、PLS_Stock）後按「計算」。主檔可放上個月的（用來算增減）或同一個月的（用來核對數字）。計算在你的瀏覽器裡進行，不會上傳。":
      "Pick the three auto-mail files (AAP_Inventory_CSS_Inventory, PLS_Picklist, PLS_Stock) and press Calculate. Add last month's master to see month-over-month changes, or this month's master to cross-check it. Everything runs in your browser; nothing is uploaded.",
    "規則：AAP 不含 AKC、ABD；只算 Good、FG、Defective（Scrap 不算）；open picklist 另外排除 RO、Others，FOC 不計金額；數字全部從明細（View）加總，不使用檔案裡的樞紐快取。":
      "Rules: AAP excludes AKC and ABD; only Good, FG and Defective count (Scrap does not); the open picklist also excludes RO and Others, and FOC carries no value. All figures are summed from the detail rows (View), not from the files' cached pivots.",
    "讀取中…": "Reading…", "計算中…": "Calculating…", "完成": "Done",
    "列": "rows", "月份": "months", "開單日": "Picklist dates", "主檔月份": "Master month",
    "AAP 庫存": "AAP inventory", "比上月": "vs last month", "舊庫存": "Aged stock", "舊庫存占比": "Aged share",
    "不良品占比": "Defective share", "超出年度目標": "Over 2026 target", "Feili 總庫存": "Feili inventory", "AAP 含 Feili": "AAP incl. Feili",
    "各 NS 庫存": "Inventory by NS", "各 NS 庫齡分布": "Ageing mix by NS", "本月比上月增減": "Change vs last month",
    "依倉別": "By warehouse", "依倉性質": "By nature", "舊庫存前 20 料號": "Top 20 aged parts",
    "Feili（HUB）": "Feili (HUB)", "Open picklist 依 NS": "Open picklist by NS", "資料檢查": "Data checks",
    "本月": "This month", "上月": "Last month", "增減": "Change", "增減%": "Change %", "季底": "Quarter end",
    "比去年底": "vs last Dec", "年度目標": "2026 target", "比目標": "vs target", "主檔數字": "Master",
    "差異": "Difference", "比季底": "vs quarter end", "舊庫存目標": "Aged target", "還要減": "Still to cut", "匯率": "Ex. rate", "當地幣": "Local currency",
    "倉別": "Warehouse", "性質": "Nature", "金額": "Amount", "占比": "Share", "合計": "Total",
    "料號": "Part number", "數量": "Qty", "項目": "Item", "說明": "Note",
    "以前": "& older", "（沒有資料）": "(no data)",
    "放在 HUB、還沒被揀走（HUB LTB、ACTDRPSHIP、HUB Returns WH）": "In the hub, not yet picked (HUB LTB, ACTDRPSHIP, HUB Returns WH)",
    "總部統一規劃的備料（HUB Active、HUB DR）": "Planned centrally for all NS (HUB Active, HUB DR)",
    "Take and Go 料件（HUB CTE）": "Take-and-go parts (HUB CTE)",
    "已開揀貨單、還沒出貨給 NS": "Picked for an NS but not yet shipped",
    "沒有選 PLS Stock 檔": "No PLS Stock file", "沒有選 PLS Picklist 檔": "No PLS Picklist file",
    "前 10 料號合計 {0}，占舊庫存 {1}；處理掉的話，舊庫存占比會從 {2} 降到 {3}。":
      "The top 10 parts hold {0}, {1} of aged stock; clearing them would cut the aged share from {2} to {3}.",
    "{0} 年（含）以前": "{0} & older",
    "請先選 CSS Inventory 檔": "Pick the CSS Inventory file first",
    "這個檔案看起來不是 {0}，確定要用嗎？": "This file doesn't look like {0}. Use it anyway?",
    chk_master_ok: "Master ({0}) matches the CSS totals for every NS.",
    chk_master_bad: "Master differs from CSS for: {0}",
    chk_master_prev: "Master is for {0}; it is used as last month for the comparisons.",
    chk_master_later: "Master month {0} is later than the data month {1}; comparisons are skipped.",
    chk_pick_unmapped: "Picklist sub-types not in the mapping (counted as After MP): {0}",
    chk_pick_dates: "Picklist dates run from {0} to {1}. Make sure this is the month-end picklist.",
    chk_stock_unmapped: "PLS Stock warehouses not in the mapping: {0} (not counted)",
    chk_months: "The CSS file holds months {0}; only {1} is used.",
    chk_excl_ns: "Excluded entities: {0}",
    chk_excl_nat: "Excluded natures (Scrap, Equipment): {0}",
    chk_no_date: "{0} rows have no Date_Created; they are counted as aged stock.",
  };
  const ZH_CHK = {
    chk_master_ok: "主檔（{0}）各 NS 金額與 CSS 加總一致。",
    chk_master_bad: "主檔與 CSS 加總不一致：{0}",
    chk_master_prev: "主檔是 {0} 的，當作上月來比較。",
    chk_master_later: "主檔月份 {0} 比資料月份 {1} 晚，不做比較。",
    chk_pick_unmapped: "Picklist 有對照表沒有的 SubType（先算在 After MP）：{0}",
    chk_pick_dates: "Picklist 開單日從 {0} 到 {1}，請確認是月底那一份。",
    chk_stock_unmapped: "PLS Stock 有對照表沒有的倉別：{0}（未計入）",
    chk_months: "CSS 檔裡有 {0} 這幾個月份，只用 {1}。",
    chk_excl_ns: "排除的 NS：{0}",
    chk_excl_nat: "排除的倉性質（Scrap、Equipment）：{0}",
    chk_no_date: "有 {0} 列沒有 Date_Created，算進舊庫存。",
  };
  const lang = () => { try { return localStorage.getItem("rma_lang") === "en" ? "en" : "zh"; } catch (e) { return "zh"; } };
  const t = (k) => (lang() === "en" ? (EN[k] || k) : (ZH_CHK[k] || k));
  const tf = (k, ...a) => t(k).replace(/\{(\d+)\}/g, (_, i) => a[+i]);
  function applyText() { document.querySelectorAll("[data-i]").forEach((el) => { el.textContent = t(el.dataset.i); }); }

  // ------------------------------------------------------------ 格式
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const n0 = (v) => (v == null || isNaN(v) ? "" : Math.round(v).toLocaleString("en-US"));
  const sgn = (v) => (v == null || isNaN(v) ? "" : (v > 0 ? "+" : v < 0 ? "−" : "") + Math.abs(Math.round(v)).toLocaleString("en-US"));
  const pct = (v, d) => (v == null || !isFinite(v) ? "" : (v * 100).toFixed(d == null ? 1 : d) + "%");
  const spct = (v) => (v == null || !isFinite(v) ? "" : (v > 0 ? "+" : "") + (v * 100).toFixed(1) + "%");
  const m2 = (v) => (v == null ? "" : "$" + (v / 1e6).toFixed(2) + "M");
  const cls = (v) => (v > 0.5 ? "up" : v < -0.5 ? "down" : "");

  // ------------------------------------------------------------ worker
  let worker = null, seq = 0;
  const pending = new Map();
  function getWorker() {
    if (worker) return worker;
    worker = new Worker("inv_worker.js?v=" + V);
    worker.onmessage = (ev) => {
      const m = ev.data;
      if (m.type === "log") { status(m.msg); return; }
      const p = pending.get(m.id); if (!p) return;
      pending.delete(m.id); m.ok ? p.resolve(m.result) : p.reject(new Error(m.error));
    };
    worker.onerror = (e) => { status("✕ " + (e.message || e)); };
    return worker;
  }
  let inited = null;
  function call(cmd, args, transfer) {
    const w = getWorker();
    if (!inited) inited = new Promise((res, rej) => { const id = ++seq; pending.set(id, { resolve: res, reject: rej }); w.postMessage({ id, cmd: "init", args: { xlsxUrl: XLSX_URL } }); });
    return inited.then(() => new Promise((resolve, reject) => {
      const id = ++seq; pending.set(id, { resolve, reject }); w.postMessage({ id, cmd, args }, transfer || []);
    }));
  }
  function status(msg) { $("inv-status").textContent = msg; }

  // ------------------------------------------------------------ 檔案
  const SLOTS = { css: "inv-css", pick: "inv-pick", stock: "inv-stock", master: "inv-master" };
  const LOOKS = { css: /CSS|Inventory/i, pick: /Picklist/i, stock: /PLS.*Stock|Stock/i, master: /Stock Holding|Ageing|Aging/i };
  const info = {};
  let result = null;
  for (const [kind, id] of Object.entries(SLOTS)) {
    $(id).addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0];
      delete info[kind];
      if (!f) { call("clear", { kind }).catch(() => {}); renderInfo(); return; }
      if (!LOOKS[kind].test(f.name) && !confirm(tf("這個檔案看起來不是 {0}，確定要用嗎？", kind === "master" ? "Stock Holding by Ageing" : kind.toUpperCase()))) { e.target.value = ""; e.target.dispatchEvent(new Event("change")); return; }
      status(t("讀取中…") + " " + f.name);
      setBusy(true);
      try {
        const buf = await f.arrayBuffer();
        info[kind] = await call("load", { kind, buf, name: f.name }, [buf]);
        if (kind === "css") fillMonths(info.css.months);
        status(t("完成"));
      } catch (err) { status("✕ " + err.message); e.target.value = ""; }
      finally { setBusy(false); renderInfo(); }
    });
  }
  function fillMonths(months) {
    const sel = $("inv-month");
    sel.innerHTML = months.map((m) => `<option value="${esc(m.key)}">${esc(m.key)}（${m.rows.toLocaleString()} ${esc(t("列"))}）</option>`).join("");
    // 檔案同時有兩個月時，預設較早那個（月初寄出的檔案通常帶到下個月的前幾天）
    if (months.length) sel.value = months[0].key;
  }
  function renderInfo() {
    const parts = [];
    if (info.css) parts.push(`CSS：${info.css.rows.toLocaleString()} ${t("列")}`);
    if (info.pick) parts.push(`Picklist：${info.pick.rows.toLocaleString()} ${t("列")}，${t("開單日")} ${info.pick.dmin}～${info.pick.dmax}`);
    if (info.stock) parts.push(`PLS Stock：${info.stock.rows.toLocaleString()} ${t("列")}`);
    if (info.master) parts.push(`${t("主檔月份")}：${info.master.month || "?"}`);
    $("inv-info").textContent = parts.join("　｜　");
    $("inv-run").disabled = !info.css;
  }
  function setBusy(b) { $("inv-run").disabled = b || !info.css; $("inv-xlsx").disabled = b || !result; }

  $("inv-run").addEventListener("click", async () => {
    if (!info.css) { status(t("請先選 CSS Inventory 檔")); return; }
    setBusy(true); status(t("計算中…"));
    try { result = await call("compute", { month: $("inv-month").value }); render(); status(t("完成")); }
    catch (err) { status("✕ " + err.message); }
    finally { setBusy(false); }
  });

  // ------------------------------------------------------------ 畫面
  function kpis(items) {
    return items.filter(Boolean).map(([k, v, sub, c]) => `<div class="kpi"><div class="k">${esc(k)}</div><div class="v ${c || ""}">${esc(v)}</div>${sub ? `<div class="caption">${esc(sub)}</div>` : ""}</div>`).join("");
  }
  function table(head, rows, opt) {
    const o = opt || {};
    return `<div class="tbl${o.tall ? " tall" : ""}"><table><thead><tr>${head.map((h, i) => `<th class="${o.num && o.num.includes(i) ? "num" : ""}">${esc(h)}</th>`).join("")}</tr></thead><tbody>` +
      rows.map((r) => `<tr class="${r.cls || ""}">${r.cells.map((c, i) => `<td class="${o.num && o.num.includes(i) ? "num" : ""} ${c && c.c ? c.c : ""}">${c && c.h != null ? c.h : esc(c && c.v != null ? c.v : c)}</td>`).join("")}</tr>`).join("") +
      `</tbody></table></div>`;
  }
  const C = (v, c) => ({ v, c });
  const AGE_COLORS = ["#B5544C", "#E8A87C", "#C9B79C", "#9BC48A", "#5FA06A", "#2E7D4F"];

  function render() {
    const R = result; if (!R) return;
    const prev = R.mode === "prev", chk = R.mode === "check";
    const F = R.feili && R.feili.summary;
    // KPI
    $("inv-kpis").innerHTML = kpis([
      [t("AAP 庫存") + " · " + R.month, m2(R.total)],
      prev && R.prevTotal ? [t("比上月"), sgn(R.total - R.prevTotal), spct(R.total / R.prevTotal - 1), cls(R.total - R.prevTotal)] : null,
      [t("舊庫存") + " · " + tf("{0} 年（含）以前", R.oldYear), m2(R.old), t("舊庫存占比") + " " + pct(R.oldPct)],
      [t("不良品占比"), pct(R.defPct)],
      R.targetTotal ? [t("超出年度目標"), spct(R.total / R.targetTotal - 1), m2(R.targetTotal), cls(R.total - R.targetTotal)] : null,
      F ? [t("Feili 總庫存"), m2(F.total), F.hasPick ? "" : t("沒有選 PLS Picklist 檔")] : null,
      F ? [t("AAP 含 Feili"), m2(R.total + F.total)] : null,
    ]);

    // 各 NS
    const head = ["NS", t("本月")];
    if (prev) head.push(t("上月"), t("增減"), t("增減%"));
    if (chk) head.push(t("主檔數字"), t("差異"));
    const hasQ = R.nsTable.some((r) => r.q);
    if (hasQ) head.push(t("季底"), t("比季底"));
    head.push(t("比去年底"), t("年度目標"), t("比目標"), t("舊庫存"), t("舊庫存占比"), t("舊庫存目標"), t("還要減"), t("匯率"));
    const numIdx = head.map((_, i) => i).filter((i) => i > 0);
    const rows = R.nsTable.map((r) => {
      const cells = [r.ns, n0(r.total)];
      if (prev) cells.push(n0(r.prev), C(sgn(r.prev != null ? r.total - r.prev : null), cls(r.total - (r.prev || r.total))), r.prev ? spct(r.total / r.prev - 1) : "");
      if (chk) { const d = r.master != null ? r.total - r.master : null; cells.push(n0(r.master), C(d == null ? "" : Math.abs(d) < 1 ? "✔" : sgn(d), d != null && Math.abs(d) >= 1 ? "up" : "")); }
      if (hasQ) cells.push(n0(r.q), C(r.q ? sgn(r.total - r.q) : "", r.q ? cls(r.total - r.q) : ""));
      const reduce = r.longTarget != null ? r.longTarget - r.old : null;
      cells.push(r.dec ? spct(r.total / r.dec - 1) : "", n0(r.target), C(r.target ? spct(r.total / r.target - 1) : "", r.target && r.total > r.target ? "up" : ""),
        n0(r.old), C(pct(r.oldPct), r.oldPct > 0.2 ? "up" : ""), n0(r.longTarget), C(reduce == null ? "" : sgn(reduce), reduce < 0 ? "up" : "down"), r.rate ? r.rate.toFixed(4) : "");
      return { cells };
    });
    const sum = (k) => R.nsTable.reduce((s, r) => s + (r[k] || 0), 0);
    const tr = ["AAP", n0(R.total)];
    if (prev) tr.push(n0(R.prevTotal), sgn(R.total - R.prevTotal), spct(R.total / R.prevTotal - 1));
    if (chk) { const d = R.total - sum("master"); tr.push(n0(sum("master")), Math.abs(d) < 1 ? "✔" : sgn(d)); }
    if (hasQ) tr.push(n0(sum("q")), sgn(R.total - sum("q")));
    const decT = sum("dec"), tgT = sum("target"), ltT = sum("longTarget");
    tr.push(decT ? spct(R.total / decT - 1) : "", n0(tgT || null), tgT ? spct(R.total / tgT - 1) : "", n0(R.old), pct(R.oldPct), n0(ltT || null), ltT ? sgn(ltT - R.old) : "", "");
    rows.push({ cls: "grp", cells: tr });
    $("inv-ns").innerHTML = table(head, rows, { num: numIdx, tall: true });

    // 本月比上月（有上月主檔時）
    if (prev) {
      const d = R.nsTable.filter((r) => r.prev != null).map((r) => ({ ns: r.ns, v: r.total - r.prev })).sort((a, b) => b.v - a.v);
      const max = Math.max(1, ...d.map((x) => Math.abs(x.v)));
      $("inv-mom").innerHTML = `<div class="ch-title">${esc(t("本月比上月增減"))} <span class="ch-unit">(USD)</span></div>` + d.map((x) => {
        const w = Math.abs(x.v) / max * 50;
        return `<div class="inv-div"><span>${esc(x.ns)}</span><span class="inv-axis"><i style="${x.v < 0 ? `right:50%;width:${w}%;background:#4E9A5B` : `left:50%;width:${w}%;background:#B5544C`}"></i></span><b class="${cls(x.v)}">${esc(sgn(x.v))}</b></div>`;
      }).join("");
      $("inv-mom").style.display = ""; $("inv-mom").parentElement.style.gridTemplateColumns = "";
    } else { $("inv-mom").innerHTML = ""; $("inv-mom").style.display = "none"; $("inv-mom").parentElement.style.gridTemplateColumns = "1fr"; }

    // 庫齡分布
    const segNames = [tf("{0} 年（含）以前", R.oldYear), ...R.yrs.map(String)];
    $("inv-age").innerHTML = `<div class="ch-title">${esc(t("各 NS 庫齡分布"))}</div><div class="hb-legend">${segNames.map((s, i) => `<span><i style="background:${AGE_COLORS[i % AGE_COLORS.length]}"></i>${esc(s)}</span>`).join("")}</div>` +
      [...R.nsTable].sort((a, b) => b.oldPct - a.oldPct).map((r) => {
        const vals = [r.old, ...R.yrs.map((y) => r.years[y] || 0)];
        return `<div class="hb-row"><span class="hb-lab">${esc(r.ns)}</span><span class="hb-bar">${vals.map((v, i) => { const w = r.total ? v / r.total * 100 : 0; return `<i style="width:${w}%;background:${AGE_COLORS[i % AGE_COLORS.length]}" title="${esc(segNames[i])}: ${n0(v)} (${pct(w / 100)})">${w >= 7 ? Math.round(w) + "%" : ""}</i>`; }).join("")}</span><b>${pct(r.oldPct, 0)}</b></div>`;
      }).join("");

    // 依倉別、依性質
    $("inv-wh").innerHTML = table([t("倉別"), t("性質"), t("金額"), t("占比")],
      R.whTable.map((w) => ({ cells: [w.wh, w.nat, n0(w.total), pct(w.total / R.total)] })).concat([{ cls: "grp", cells: [t("合計"), "", n0(R.total), "100.0%"] }]), { num: [2, 3], tall: true });
    $("inv-nat").innerHTML = table([t("性質"), t("金額"), t("占比")],
      R.natTable.map((n) => ({ cls: n.nat === "Defective" ? "warn" : "", cells: [n.nat, n0(n.total), pct(n.pct)] })), { num: [1, 2] }) +
      `<div class="inv-stack">${R.natTable.map((n, i) => `<i style="width:${n.pct * 100}%;background:${["#5FA06A", "#2E6FBA", "#B5544C"][i]}">${n.pct > 0.06 ? esc(n.nat) + " " + pct(n.pct, 0) : ""}</i>`).join("")}</div>`;

    // 舊庫存料號
    const after = R.total ? (R.old - R.top10) / R.total : 0;
    $("inv-top-note").textContent = tf("前 10 料號合計 {0}，占舊庫存 {1}；處理掉的話，舊庫存占比會從 {2} 降到 {3}。", m2(R.top10), pct(R.old ? R.top10 / R.old : 0), pct(R.oldPct), pct(after));
    $("inv-top").innerHTML = table(["NS", t("料號"), t("倉別"), t("數量"), t("金額")],
      R.topOld.map((o) => ({ cells: [o.ns, o.part, o.wh, n0(o.qty), n0(o.usd)] })), { num: [3, 4], tall: true });

    // Feili
    if (R.feili) {
      const f = R.feili, s = f.summary;
      const pickRows = Object.entries(f.pick).filter(([k]) => !k.includes("FOC")).sort((a, b) => b[1] - a[1]).map(([k, v]) => ({ cls: "sub", cells: ["", k.replace("Open Picklist ", ""), n0(v)] }));
      $("inv-feili").innerHTML = table([t("項目"), t("說明"), t("金額")], [
        { cells: ["Hold in Feili", t("放在 HUB、還沒被揀走（HUB LTB、ACTDRPSHIP、HUB Returns WH）"), s.hasStock ? n0(s.hold) : t("沒有選 PLS Stock 檔")] },
        { cells: ["Central", t("總部統一規劃的備料（HUB Active、HUB DR）"), s.hasStock ? n0(s.central) : ""] },
        { cells: ["CTE", t("Take and Go 料件（HUB CTE）"), s.hasStock ? n0(s.cte) : ""] },
        { cells: ["Open Picklist", t("已開揀貨單、還沒出貨給 NS"), s.hasPick ? n0(s.pick) : t("沒有選 PLS Picklist 檔")] },
        ...pickRows,
        { cls: "grp", cells: [t("合計"), "", n0(s.total)] },
      ], { num: [2] });
      const pn = Object.entries(f.pickNs).sort((a, b) => b[1] - a[1]);
      $("inv-pickns").innerHTML = pn.length ? table(["NS", t("金額"), t("占比")], pn.map(([k, v]) => ({ cells: [k, n0(v), pct(s.pick ? v / s.pick : 0)] })), { num: [1, 2] }) : "";
      $("inv-feili-wrap").style.display = "";
    } else $("inv-feili-wrap").style.display = "none";

    // 檢查
    const icon = { ok: "✔", warn: "⚠", info: "ℹ" };
    $("inv-checks").innerHTML = R.checks.map((c) => `<li class="inv-${c.level}"><span>${icon[c.level]}</span> ${esc(tf(c.key, ...c.args))}</li>`).join("");
    $("inv-result").style.display = "";
    $("inv-xlsx").disabled = false;
  }

  // ------------------------------------------------------------ 匯出
  $("inv-xlsx").addEventListener("click", async () => {
    const R = result; if (!R) return;
    const r0 = (v) => (v == null ? null : Math.round(v * 100) / 100);
    const nsHead = ["NS", t("本月"), t("上月"), t("主檔數字"), t("比去年底"), t("年度目標"), t("舊庫存"), t("舊庫存占比"), t("舊庫存目標"), t("當地幣"), t("匯率"), tf("{0} 年（含）以前", R.oldYear), ...R.yrs.map(String)];
    const nsRows = R.nsTable.map((r) => [r.ns, r0(r.total), r0(r.prev), r0(r.master), r.dec ? r0(r.total / r.dec - 1) : null, r0(r.target), r0(r.old), r0(r.oldPct), r0(r.longTarget), r0(r.loc), r.rate ? Math.round(r.rate * 1e4) / 1e4 : null, r0(r.old), ...R.yrs.map((y) => r0(r.years[y] || 0))]);
    const sheets = [
      { name: t("各 NS 庫存"), rows: [nsHead, ...nsRows] },
      { name: t("依倉別"), rows: [[t("倉別"), t("性質"), t("金額"), ...R.nsTable.map((r) => r.ns)], ...R.whTable.map((w) => [w.wh, w.nat, r0(w.total), ...R.nsTable.map((r) => r0(w.byNs[r.ns] || 0))])] },
      { name: t("依倉性質"), rows: [[t("性質"), t("金額"), t("占比")], ...R.natTable.map((n) => [n.nat, r0(n.total), r0(n.pct)])] },
      { name: t("舊庫存前 20 料號"), rows: [["NS", t("料號"), t("倉別"), t("數量"), t("金額")], ...R.topOld.map((o) => [o.ns, o.part, o.wh, o.qty, r0(o.usd)])] },
    ];
    if (R.feili) {
      const s = R.feili.summary;
      sheets.push({ name: "Feili", rows: [[t("項目"), t("金額")], ["Hold in Feili", r0(s.hold)], ["Central", r0(s.central)], ["CTE", r0(s.cte)],
        ...Object.entries(R.feili.pick).map(([k, v]) => [k, r0(v)]), [t("合計"), r0(s.total)], [], ["Open picklist by NS"], ...Object.entries(R.feili.pickNs).map(([k, v]) => [k, r0(v)])] });
    }
    sheets.push({ name: t("資料檢查"), rows: [[t("說明")], ...R.checks.map((c) => [tf(c.key, ...c.args)])] });
    try {
      const out = await call("export", { sheets, name: `Inventory_${R.month}.xlsx` });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([out.buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = out.name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    } catch (err) { status("✕ " + err.message); }
  });

  // 語言切換：app.js 改完 localStorage 後再套用
  const ls = $("lang-sel");
  if (ls) ls.addEventListener("change", () => setTimeout(() => { applyText(); renderInfo(); status(""); if (result) { render(); fillMonthsKeep(); } }, 50));
  function fillMonthsKeep() { if (info.css) { const v = $("inv-month").value; fillMonths(info.css.months); $("inv-month").value = v; } }
  applyText();
  renderInfo();
})();
