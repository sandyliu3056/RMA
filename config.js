/* 網站設定（公開檔案，會隨網站部署）。
 * 資料庫：填入 Supabase 專案的 Project URL 與 anon public key，資料來源頁就會出現登入、「從資料庫載入」與「存入資料庫」
 *（建立步驟見 README 的「資料庫（Supabase）」）。anon key 本來就是公開的，資料靠資料庫的列級安全規則保護，沒登入的人讀不到。
 * 兩個都留空＝不用資料庫，只能載入單一檔案。本機測試可用 config.local.js 覆蓋這些值（該檔不要提交）。 */
window.APP_CONFIG = Object.assign(window.APP_CONFIG || {}, {
  supabaseUrl: "",   // 例如 "https://abcdefghijkl.supabase.co"
  supabaseKey: "",   // Project Settings → API → anon public
});
