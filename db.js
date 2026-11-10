/* 資料庫存取（Supabase）：登入、匯入紀錄、案件快照。
 * 不用 supabase-js，直接呼叫 Supabase 的 Auth（GoTrue）與 REST（PostgREST）API，少一個外部套件。
 * 設定來自 config.js 的 APP_CONFIG.supabaseUrl / supabaseKey；登入狀態存在 localStorage（rma_db_session），過期會自動用 refresh token 換新。
 * 資料表與函式見 supabase/schema.sql。 */
(function () {
  "use strict";
  const SKEY = "rma_db_session";
  let BASE = "", KEY = "", sess = null;

  function loadSess() {
    try { const s = JSON.parse(localStorage.getItem(SKEY) || "null"); return s && s.refresh_token ? s : null; } catch (e) { return null; }
  }
  function saveSess(s) {
    sess = s;
    try { if (s) localStorage.setItem(SKEY, JSON.stringify(s)); else localStorage.removeItem(SKEY); } catch (e) { /* 無法存也沒關係 */ }
  }
  function configure(cfg) {
    // 從 Data API 頁複製來的網址會多 /rest/v1/，去掉只留專案網址
    BASE = String((cfg && cfg.supabaseUrl) || "").trim().replace(/\/+$/, "").replace(/\/(rest|auth)\/v1$/, "");
    KEY = String((cfg && cfg.supabaseKey) || "").trim();
    sess = configured() ? loadSess() : null;
  }
  // 另一個分頁登入／登出／換了 token：跟著用最新的（refresh token 只能用一次，兩個分頁各自換會互相踢掉）
  try { window.addEventListener("storage", (e) => { if (e.key === SKEY) sess = loadSess(); }); } catch (e) { /* 沒有 window 就算了 */ }
  function configured() { return !!(BASE && KEY); }
  function user() { return sess ? (sess.email || "") : null; }

  async function parse(r) {
    const text = await r.text();
    let j = null;
    try { j = text ? JSON.parse(text) : null; } catch (e) { j = null; }
    if (!r.ok) {
      const msg = (j && (j.error_description || j.msg || j.message || j.error || j.hint)) || text || "";
      throw new Error(`${msg || "HTTP error"} (HTTP ${r.status})`);
    }
    return j;
  }
  async function auth(path, body, bearer) {
    const headers = { apikey: KEY, "Content-Type": "application/json" };
    if (bearer) headers.Authorization = "Bearer " + bearer;
    return parse(await fetch(`${BASE}/auth/v1/${path}`, { method: "POST", headers, body: JSON.stringify(body || {}) }));
  }
  function setFromToken(j) {
    if (!j || !j.access_token) throw new Error("No access token in response");
    saveSess({ access_token: j.access_token, refresh_token: j.refresh_token, expires_at: Date.now() + Number(j.expires_in || 3600) * 1000,
               email: (j.user && j.user.email) || (sess && sess.email) || "" });
  }
  async function login(email, password) {
    setFromToken(await auth("token?grant_type=password", { email, password }));
    return user();
  }
  async function logout() {
    const s = sess;
    saveSess(null);
    if (s) { try { await auth("logout", {}, s.access_token); } catch (e) { /* 本機已登出即可 */ } }
  }
  /* 取得有效的 access token；快過期就換新。
   * 只有 Supabase 明確拒絕 refresh token（4xx）才算登出；網路不通或伺服器錯誤就把原本的錯誤丟回去，登入狀態保留。
   * 同時有多個請求時只換一次（refresh token 只能用一次，重複用會被拒絕）。 */
  let refreshing = null;
  async function token(force) {
    if (!sess) throw new Error("NOT_LOGGED_IN");
    if (force || Date.now() > sess.expires_at - 60000) {
      if (!refreshing) {
        const fresh = loadSess();   // 別的分頁可能已經換過 token
        if (fresh && fresh.refresh_token !== sess.refresh_token && fresh.expires_at > sess.expires_at) sess = fresh;
        if (!force && Date.now() <= sess.expires_at - 60000) return sess.access_token;
        refreshing = auth("token?grant_type=refresh_token", { refresh_token: sess.refresh_token })
          .then(setFromToken)
          .catch((e) => { if (/\(HTTP 4\d\d\)/.test(e.message)) { saveSess(null); throw new Error("SESSION_EXPIRED"); } throw e; })
          .finally(() => { refreshing = null; });
      }
      await refreshing;
      if (!sess) throw new Error("SESSION_EXPIRED");
    }
    return sess.access_token;
  }
  async function rest(method, path, body, prefer, retried) {
    const tk = await token();
    const headers = { apikey: KEY, Authorization: "Bearer " + tk, "Content-Type": "application/json" };
    if (prefer) headers.Prefer = prefer;
    try {
      return await parse(await fetch(`${BASE}/rest/v1/${path}`, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) }));
    } catch (e) {
      // token 被拒絕（例如密碼改了、在別處登出）：換一次新的再試；還是不行就視為登出
      if (/\(HTTP 401\)/.test(e.message) && !retried) { await token(true); return rest(method, path, body, prefer, true); }
      if (/\(HTTP 401\)/.test(e.message)) { saveSess(null); throw new Error("SESSION_EXPIRED"); }
      throw e;
    }
  }

  window.RMADB = {
    configure, configured, user, login, logout,
    /* 頁面開啟時：還原上次的登入（會順便確認 token 還有效） */
    async restore() { if (!sess) return null; try { await token(); return user(); } catch (e) { return null; } },
    /* 匯入紀錄（含沒上傳完的，complete=false），新的在前 */
    listImports() { return rest("GET", "imports?select=*&order=data_date.desc,id.desc"); },
    createImport(rec) { return rest("POST", "imports", rec, "return=representation").then((a) => (Array.isArray(a) ? a[0] : a)); },
    deleteImport(id) { return rest("DELETE", "imports?id=eq." + encodeURIComponent(id)); },
    /* 全部列上傳完成後：一次交易刪掉同一資料日期的舊匯入、標記完成（資料庫端的 finish_import） */
    finishImport(id) { return rest("POST", "rpc/finish_import", { p_import_id: id }); },
    /* 分批寫入案件（每批 1,000 列）；onProgress(已送, 全部) */
    async upsertCases(importId, rows, onProgress, batch) {
      const B = batch || 1000;
      let n = 0;
      for (let i = 0; i < rows.length; i += B) {
        n += Number(await rest("POST", "rpc/upsert_cases", { p_import_id: importId, p_rows: rows.slice(i, i + B) })) || 0;
        if (onProgress) onProgress(Math.min(rows.length, i + B), rows.length);
      }
      return n;
    },
    /* 讀取期間內的案件（每個 CaseID 最新版本），每頁 5,000 列，用上一頁最後的 CaseID 接著讀（鍵集分頁）；onProgress(已下載) */
    async fetchCases(from, to, onProgress, limit) {
      const L = limit || 5000;
      const out = [];
      let after = "";
      for (;;) {
        const page = await rest("POST", "rpc/get_cases", { p_from: from || null, p_to: to || null, p_after: after, p_limit: L });
        if (!Array.isArray(page) || !page.length) break;
        out.push(...page);
        if (onProgress) onProgress(out.length);
        const last = page[page.length - 1];
        const next = last && last._case_id != null ? String(last._case_id) : "";
        if (page.length < L || !next || next <= after) break;
        after = next;
      }
      return out;
    },
  };
})();
