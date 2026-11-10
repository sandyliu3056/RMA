/* 庫存報告分頁的 Web Worker：用 SheetJS 讀 CSS Inventory／PLS Picklist／PLS Stock／主檔，計算月底庫存報告的數字。
 * 不需要 Pyodide，選好檔案就能算。檔案只在瀏覽器裡處理，不會上傳。
 * 訊息：主畫面 → {id, cmd, args}；回覆 {id, ok, result | error}；進度 {type:"log", msg}
 * cmd：
 *   load   {kind, buf, name}  讀一個檔案並暫存（kind = css / pick / stock / master），回傳摘要
 *   compute {month}           依暫存的檔案計算報告，回傳所有表格
 *   export  {report}          把 compute 的結果寫成 Excel，回傳 ArrayBuffer */
let XLSX_READY = false;
const S = { css: null, pick: null, stock: null, master: null };

// 依 SOP：AAP 不含 AKC、ABD；Feili 的 open picklist 另外排除 RO、Others
const EXCL_NS = new Set(["AKC", "ABD"]);
const EXCL_PICK = new Set(["AKC", "ABD", "RO", "OTHERS"]);
const KEEP_NATURE = new Set(["Good", "FG", "Defective"]);   // Scrap、Equipment 不算庫存
// PLS Stock 倉別 → 分類（檔案裡有 Mapping (3) 時以檔案為準；HUB DR 依主檔歸入 Central）
const STOCK_MAP = {
  "HUB FOC": "FOC", "HUB QUALITY": "FOC", "FOCDRPSHIP": "FOC",
  "HUB Active": "Central", "HUB DR": "Central",
  "ACTDRPSHIP": "Hold In Feili", "HUB LTB": "Hold In Feili", "HUB Returns WH": "Hold In Feili",
  "HUB CTE": "CTE",
  "Scrap Warehouse": "B2B", "CustDefective Warehouse": "B2B", "CustActive Warehouse": "B2B", "Defective Warehouse": "B2B",
  "Missing Stock": "Missing Stock",
};
// PLS Picklist 的 SubType → 階段（Mapping (3)）
const SUB_CTE = new Set(["TnG", "TnG-Last"]);
const SUB_FOC = new Set(["FOC", "Quality"]);
const SUB_MP = new Set(["BS-NPI", "BS-F", "BS-N", "BS-NR", "BS-OBS", "FCST-R", "Non-FCST", "Npro-BML", "Npro-NPI", "Pro-monthly", "Pro-NPI"]);
const SUB_AFTER = new Set(["", "(Blank)", "LTB", "Npro-AML", "Npro-PPB1", "Npro-PPB2", "PORETURN", "Pro-last", "Special support", "Special Support"]);
const NS_ALIAS = { HSN: "ASSB", TWN: "TBG", TW: "TBG" };

function log(msg) { self.postMessage({ type: "log", msg }); }

self.onmessage = async (ev) => {
  const { id, cmd, args } = ev.data;
  try {
    if (cmd === "init") { if (!XLSX_READY) { importScripts(args.xlsxUrl); XLSX_READY = true; } self.postMessage({ id, ok: true, result: true }); return; }
    let result;
    if (cmd === "load") result = load(args);
    else if (cmd === "compute") result = compute(args);
    else if (cmd === "export") result = exportXlsx(args);
    else if (cmd === "clear") { S[args.kind] = null; result = true; }
    else throw new Error("unknown command " + cmd);
    self.postMessage({ id, ok: true, result }, result && result.buf ? [result.buf] : []);
  } catch (e) {
    self.postMessage({ id, ok: false, error: (e && e.message) || String(e) });
  }
};

// ------------------------------------------------------------ 讀檔
function sheetRows(wb, name) {
  const ws = wb.Sheets[name];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
}
function headerIndex(row) {
  const m = {};
  (row || []).forEach((h, i) => { if (h != null) m[String(h).trim()] = i; });
  return m;
}
function need(h, cols, file) {
  const miss = cols.filter((c) => !(c in h));
  if (miss.length) throw new Error(`${file}：找不到欄位 ${miss.join(", ")}`);
}
const num = (v) => (typeof v === "number" ? v : v == null || v === "" ? 0 : Number(String(v).replace(/,/g, "")) || 0);
const str = (v) => (v == null ? "" : String(v).trim());
function toDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number") { const d = XLSX.SSF.parse_date_code(v); return d ? new Date(d.y, d.m - 1, d.d) : null; }
  if (typeof v === "string" && v) { const d = new Date(v.replace(/\//g, "-")); return isNaN(d) ? null : d; }
  return null;
}
function ymd(d) { const p = (n) => String(n).padStart(2, "0"); return d ? `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}` : ""; }

function load({ kind, buf, name }) {
  log(`讀取 ${name}…`);
  if (kind === "master") return loadMaster(buf, name);
  // auto mail 檔：明細在 View 工作表；只解析這一張，速度快很多
  const probe = XLSX.read(new Uint8Array(buf), { type: "array", bookSheets: true });
  const sheet = probe.SheetNames.includes("View") ? "View" : probe.SheetNames[0];
  const sheets = [sheet]; if (probe.SheetNames.includes("Mapping (3)")) sheets.push("Mapping (3)");
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, dense: true, sheets });
  const rows = sheetRows(wb, sheet);
  const h = headerIndex(rows[0]);
  if (kind === "css") return loadCss(rows, h, name);
  if (kind === "pick") return loadPick(rows, h, name);
  if (kind === "stock") return loadStock(rows, h, name, wb.Sheets["Mapping (3)"] ? sheetRows(wb, "Mapping (3)") : null);
  throw new Error("unknown kind " + kind);
}

function loadCss(rows, h, name) {
  need(h, ["FinRevEntity", "Warehouse_Nature", "Warehouse", "Total_Std_Cost_USD", "Date_Created", "Month"], name);
  const out = [];
  const months = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const ent = str(r[h.FinRevEntity]); if (!ent) continue;
    const d = toDate(r[h.Date_Created]);
    const m = num(r[h.Month]), y = h.Year != null ? num(r[h.Year]) : 0;
    const key = y ? `${y}-${String(m).padStart(2, "0")}` : String(m);
    months[key] = (months[key] || 0) + 1;
    out.push({
      ent, key, nat: str(r[h.Warehouse_Nature]), wh: str(r[h.Warehouse]),
      part: h.Part_Number != null ? str(r[h.Part_Number]) : "", qty: h.Inventory_On_Hand_Qty != null ? num(r[h.Inventory_On_Hand_Qty]) : 0,
      usd: num(r[h.Total_Std_Cost_USD]), loc: h.Total_Std_Cost != null ? num(r[h.Total_Std_Cost]) : 0,
      yr: d ? d.getFullYear() : null,
    });
  }
  S.css = { name, rows: out };
  const keys = Object.keys(months).sort();
  return { name, rows: out.length, months: keys.map((k) => ({ key: k, rows: months[k] })) };
}

function loadPick(rows, h, name) {
  need(h, ["CustCode", "SubType", "TotalAmount"], name);
  const part = h.ALLOCATED_PARTNUMBER != null ? h.ALLOCATED_PARTNUMBER : h.SOH_PartNumber;
  const out = [];
  let dmin = null, dmax = null;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const cust = str(r[h.CustCode]); if (!cust) continue;
    const d = h.PicklistDate != null ? toDate(r[h.PicklistDate]) : null;
    if (d) { if (!dmin || d < dmin) dmin = d; if (!dmax || d > dmax) dmax = d; }
    out.push({ cust, sub: str(r[h.SubType]), part: part != null ? str(r[part]) : "", amt: num(r[h.TotalAmount]),
               status: h.Status != null ? str(r[h.Status]) : "", date: d ? ymd(d) : "" });
  }
  S.pick = { name, rows: out, dmin: ymd(dmin), dmax: ymd(dmax) };
  return { name, rows: out.length, dmin: ymd(dmin), dmax: ymd(dmax) };
}

function loadStock(rows, h, name, mapRows) {
  need(h, ["WM_DESC", "Unit_Price", "Available_Qty"], name);
  const map = Object.assign({}, STOCK_MAP);
  if (mapRows) {
    // Mapping (3)：A 欄 Category、B 欄 WM_DESC
    for (const r of mapRows.slice(1)) { const c = str(r[0]), w = str(r[1]); if (c && w && c !== "Category") map[w] = c; }
    if (!("HUB DR" in map)) map["HUB DR"] = "Central";
  }
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const wm = str(r[h.WM_DESC]); if (!wm) continue;
    out.push({ wm, cat: map[wm] || "?", amt: num(r[h.Unit_Price]) * num(r[h.Available_Qty]), age: h.StockAging != null ? num(r[h.StockAging]) : null });
  }
  S.stock = { name, rows: out };
  return { name, rows: out.length };
}

/* 主檔 Stock Holding by Ageing：讀 include … with TBG 工作表的各 NS「季度列」與「月份列」、去年底與年度目標。
 * 主檔很大，只解析這一張工作表。 */
function loadMaster(buf, name) {
  const probe = XLSX.read(new Uint8Array(buf), { type: "array", bookSheets: true });
  const sheet = probe.SheetNames.find((n) => /^include.*with\s*TBG/i.test(n)) || probe.SheetNames.find((n) => /^include/i.test(n));
  if (!sheet) throw new Error(`${name}：找不到「include … with TBG」工作表`);
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, dense: true, sheets: [sheet] });
  const rows = sheetRows(wb, sheet);
  // 表頭列：B 欄是 "Entity"；同列找 Total、Grand Total、年份欄
  let hr = rows.findIndex((r) => r && str(r[1]) === "Entity");
  if (hr < 0) throw new Error(`${name}：${sheet} 找不到 Entity 表頭列`);
  const hdr = rows[hr], top = rows[hr - 1] || [];
  const col = (re, row) => (row || hdr).findIndex((v) => v != null && re.test(str(v)));
  const cTotal = col(/^Total$/), cGT = col(/^Grand Total$/);
  const years = [];
  hdr.forEach((v, i) => { const y = parseInt(str(v), 10); if (y >= 2000 && y <= 2100) years.push({ y, i }); });
  const cDec = col(/Dec\.?$/, top), cTarget = col(/Grand Target/, top), cLong = col(/Long Aging Target/, top);
  const ns = {};
  let cur = null;
  for (let r = hr + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (row.some((v) => /Ageing Stock by amt/i.test(str(v)))) break;   // 下方百分比表開始
    const lab = str(row[1]);
    if (lab && /^[A-Z]{2,5}$/.test(lab)) cur = NS_ALIAS[lab] || lab;
    if (!cur) continue;
    const per = row[2];
    const gt = num(row[cGT]);
    const rec = { gt, old: num(row[cTotal]), years: Object.fromEntries(years.map((y) => [y.y, num(row[y.i])])) };
    const d = toDate(per);
    if (/^Q[1-4]$/.test(str(per))) { (ns[cur] = ns[cur] || {}).q = Object.assign(rec, { label: str(per) }); }
    else if (d && d.getFullYear() > 2000) {
      const o = (ns[cur] = ns[cur] || {});
      o.cur = Object.assign(rec, { month: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}` });
      if (cDec >= 0) o.dec = num(row[cDec]);
      if (cTarget >= 0) o.target = num(row[cTarget]);
      if (cLong >= 0) o.longTarget = num(row[cLong]);
    }
  }
  const months = [...new Set(Object.values(ns).map((o) => o.cur && o.cur.month).filter(Boolean))];
  S.master = { name, sheet, ns, month: months[0] || "", decLabel: cDec >= 0 ? str(top[cDec]) : "" };
  return { name, sheet, month: S.master.month, entities: Object.keys(ns).length };
}

// ------------------------------------------------------------ 計算
function compute({ month }) {
  if (!S.css) throw new Error("請先選 CSS Inventory 檔");
  const all = S.css.rows;
  const mon = month || S.css.rows[0].key;
  const yearNow = parseInt(mon.slice(0, 4), 10) || new Date().getFullYear();
  const oldYear = yearNow - 5;            // 例：2026 年 → 2021 年（含）以前算舊庫存
  const yrs = []; for (let y = oldYear + 1; y <= yearNow; y++) yrs.push(y);
  const inMonth = all.filter((r) => r.key === mon);
  const checks = [];

  // 排除項目
  const exclNs = {}, exclNat = {};
  for (const r of inMonth) {
    if (EXCL_NS.has(r.ent)) exclNs[r.ent] = (exclNs[r.ent] || 0) + r.usd;
    else if (!KEEP_NATURE.has(r.nat)) exclNat[r.nat || "(空白)"] = (exclNat[r.nat || "(空白)"] || 0) + r.usd;
  }
  const rows = inMonth.filter((r) => !EXCL_NS.has(r.ent) && KEEP_NATURE.has(r.nat));

  // 各 NS：總額、依年份、舊庫存、當地幣
  const nsMap = {};
  for (const r of rows) {
    const o = nsMap[r.ent] || (nsMap[r.ent] = { ns: r.ent, total: 0, old: 0, loc: 0, years: {}, nat: {} });
    o.total += r.usd; o.loc += r.loc;
    o.nat[r.nat] = (o.nat[r.nat] || 0) + r.usd;
    if (r.yr == null || r.yr <= oldYear) o.old += r.usd;
    else o.years[r.yr] = (o.years[r.yr] || 0) + r.usd;
  }
  const nsList = Object.values(nsMap).sort((a, b) => a.ns.localeCompare(b.ns));
  const tot = nsList.reduce((s, o) => ({ total: s.total + o.total, old: s.old + o.old }), { total: 0, old: 0 });

  // 主檔比較：主檔月份 = 本月 → 核對；主檔月份 < 本月 → 上月比較
  const M = S.master;
  let mode = "";
  if (M && M.month) mode = M.month === mon ? "check" : M.month < mon ? "prev" : "later";
  const mrec = (ns) => (M && M.ns[ns]) || {};
  const nsTable = nsList.map((o) => {
    const m = mrec(o.ns);
    const row = { ns: o.ns, total: o.total, old: o.old, oldPct: o.total ? o.old / o.total : 0, years: o.years, loc: o.loc,
                  rate: o.total ? o.loc / o.total : null, defPct: o.total ? (o.nat.Defective || 0) / o.total : 0 };
    if (mode === "prev" && m.cur) { row.prev = m.cur.gt; row.prevOld = m.cur.old; }
    if (mode === "check" && m.cur) { row.master = m.cur.gt; row.q = m.q ? m.q.gt : null; }
    if (mode === "prev" && m.q) row.q = m.q.gt;
    if (m.dec) row.dec = m.dec;
    if (m.target) row.target = m.target;
    if (m.longTarget) row.longTarget = m.longTarget;
    return row;
  });
  if (mode === "check") {
    const bad = nsTable.filter((r) => r.master != null && Math.abs(r.master - r.total) > 1);
    checks.push(bad.length ? { level: "warn", key: "chk_master_bad", args: [bad.map((r) => `${r.ns}（${fmt0(r.master)} vs ${fmt0(r.total)}）`).join("、")] }
                           : { level: "ok", key: "chk_master_ok", args: [M.sheet] });
  } else if (mode === "prev") checks.push({ level: "info", key: "chk_master_prev", args: [M.month] });
  else if (mode === "later") checks.push({ level: "warn", key: "chk_master_later", args: [M.month, mon] });

  // 依倉別
  const whMap = {};
  for (const r of rows) {
    const k = r.wh || "?";
    const o = whMap[k] || (whMap[k] = { wh: k, nat: r.nat, total: 0, byNs: {} });
    o.total += r.usd; o.byNs[r.ent] = (o.byNs[r.ent] || 0) + r.usd;
  }
  const whTable = Object.values(whMap).sort((a, b) => a.wh.localeCompare(b.wh));
  const natMap = {};
  for (const r of rows) natMap[r.nat] = (natMap[r.nat] || 0) + r.usd;
  const natTable = ["Good", "FG", "Defective"].filter((k) => natMap[k] != null).map((k) => ({ nat: k, total: natMap[k], pct: tot.total ? natMap[k] / tot.total : 0 }));

  // 舊庫存前 20 料號
  const oldParts = {};
  for (const r of rows) if (r.yr == null || r.yr <= oldYear) {
    const k = r.ent + "|" + r.part;
    const o = oldParts[k] || (oldParts[k] = { ns: r.ent, part: r.part, qty: 0, usd: 0, wh: new Set() });
    o.qty += r.qty; o.usd += r.usd; o.wh.add((r.wh || "").slice(0, 3));
  }
  const topOld = Object.values(oldParts).sort((a, b) => b.usd - a.usd).slice(0, 20).map((o) => ({ ns: o.ns, part: o.part, qty: o.qty, usd: o.usd, wh: [...o.wh].sort().join("/") }));
  const top10 = topOld.slice(0, 10).reduce((s, o) => s + o.usd, 0);

  // Feili：PLS Stock + PLS Picklist
  let feili = null;
  if (S.stock || S.pick) {
    feili = { stock: {}, pick: {}, pickNs: {}, unmapped: {} };
    if (S.stock) for (const r of S.stock.rows) feili.stock[r.cat] = (feili.stock[r.cat] || 0) + r.amt;
    if (S.pick) {
      const exclP = {};
      for (const r of S.pick.rows) {
        if (EXCL_PICK.has(r.cust.toUpperCase())) { exclP[r.cust] = (exclP[r.cust] || 0) + r.amt; continue; }
        let cat;
        if (SUB_CTE.has(r.sub)) cat = "Open Picklist (CTE)";
        else if (SUB_FOC.has(r.sub)) cat = "Open Picklist (FOC)";
        else {
          const g = /^(NB\.|6M\.)/.test(r.part) ? "MBO & LCM" : "Others";
          let st = SUB_MP.has(r.sub) ? "MP" : "After MP";
          if (!SUB_MP.has(r.sub) && !SUB_AFTER.has(r.sub)) feili.unmapped[r.sub] = (feili.unmapped[r.sub] || 0) + r.amt;
          cat = `Open Picklist ${g} (${st})`;
        }
        feili.pick[cat] = (feili.pick[cat] || 0) + (cat.includes("FOC") ? 0 : r.amt);
        if (!cat.includes("FOC")) feili.pickNs[r.cust] = (feili.pickNs[r.cust] || 0) + r.amt;
      }
      feili.pickExcl = exclP;
      feili.pickDates = [S.pick.dmin, S.pick.dmax];
      if (Object.keys(feili.unmapped).length) checks.push({ level: "warn", key: "chk_pick_unmapped", args: [Object.keys(feili.unmapped).join("、")] });
      checks.push({ level: "info", key: "chk_pick_dates", args: [S.pick.dmin, S.pick.dmax] });
    }
    const hold = feili.stock["Hold In Feili"] || 0, central = feili.stock.Central || 0, cte = feili.stock.CTE || 0;
    const pickTotal = Object.entries(feili.pick).reduce((s, [k, v]) => s + (k.includes("FOC") ? 0 : v), 0);
    feili.summary = { hold, central, cte, pick: pickTotal, total: hold + central + cte + pickTotal, hasStock: !!S.stock, hasPick: !!S.pick };
    if (S.stock && feili.stock["?"]) checks.push({ level: "warn", key: "chk_stock_unmapped", args: [fmt0(feili.stock["?"])] });
  }

  // 其他檢查
  const monthsInFile = [...new Set(all.map((r) => r.key))].sort();
  if (monthsInFile.length > 1) checks.unshift({ level: "info", key: "chk_months", args: [monthsInFile.join("、"), mon] });
  const ex = Object.entries(exclNs).map(([k, v]) => `${k} ${fmt0(v)}`).join("、");
  if (ex) checks.push({ level: "info", key: "chk_excl_ns", args: [ex] });
  const en = Object.entries(exclNat).filter(([, v]) => Math.abs(v) > 0.5).map(([k, v]) => `${k} ${fmt0(v)}`).join("、");
  checks.push({ level: "info", key: "chk_excl_nat", args: [en || "0"] });
  const noYear = rows.filter((r) => r.yr == null).length;
  if (noYear) checks.push({ level: "warn", key: "chk_no_date", args: [noYear] });

  return {
    month: mon, oldYear, yrs, mode, masterMonth: M ? M.month : "", masterName: M ? M.name : "",
    total: tot.total, old: tot.old, oldPct: tot.total ? tot.old / tot.total : 0,
    defPct: tot.total ? (natMap.Defective || 0) / tot.total : 0,
    prevTotal: mode === "prev" ? nsTable.reduce((s, r) => s + (r.prev || 0), 0) : null,
    targetTotal: nsTable.reduce((s, r) => s + (r.target || 0), 0) || null,
    nsTable, whTable, natTable, topOld, top10, feili, checks,
    files: { css: S.css.name, pick: S.pick && S.pick.name, stock: S.stock && S.stock.name, master: M && M.name },
  };
}
function fmt0(v) { return Math.round(v).toLocaleString("en-US"); }

// ------------------------------------------------------------ 匯出 Excel
function exportXlsx({ sheets, name }) {
  const wb = XLSX.utils.book_new();
  for (const s of sheets) {
    const ws = XLSX.utils.aoa_to_sheet(s.rows);
    ws["!cols"] = (s.rows[0] || []).map((h, i) => ({ wch: Math.max(10, ...s.rows.slice(0, 60).map((r) => String(r[i] == null ? "" : r[i]).length + 2)) }));
    XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31));
  }
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return { buf: out.buffer || out, name };
}
