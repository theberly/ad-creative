// 網頁 API：每人通行碼，只能操作自己看得到的客戶
// GET  ?action=me&code=…            我是誰、我的客戶
// GET  ?action=client&code=…&client=…   客戶設定、活動清單、最近需求
// GET  ?action=results&code=…&client=…&id=…   某次需求的產出（id 空白＝最新一筆）
// POST {action:'generate', code, client, request:{平台,活動名稱,格式,受眾,組數,產品／服務,活動目標,到達網址,補充說明}}
const API_REQ_FIELDS = ['平台', '活動名稱', '格式', '受眾', '組數', '產品／服務', '活動目標', '到達網址', '補充說明'];
const API_OUT_FIELDS = ['需求編號', '序號', '受眾', '文案', '標題', '說明', 'CTA', '到達網址', '字數（文案／標題／說明）', '是否超限', '審核警示', '確認狀態'];

function doGet(e) { return json_(handle_((e && e.parameter) || {})); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '請求格式錯誤' }); }
  return json_(handle_(p));
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function handle_(p) {
  try {
    const u = userByCode_(p.code);
    if (!u) return { ok: false, error: '通行碼錯誤' };
    if (p.action === 'me') return apiMe_(u);
    if (p.client && u.clients.indexOf(p.client) < 0) return { ok: false, error: '你沒有這個客戶的權限' };
    if (p.action === 'client') return apiClient_(p.client);
    if (p.action === 'results') return apiResults_(p.client, p.id);
    if (p.action === 'generate') return apiGenerate_(u, p.client, p.request || {});
    return { ok: false, error: '未知的動作' };
  } catch (err) { return { ok: false, error: err.message }; }
}

function apiMe_(u) {
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues(), h = v[0];
  const pick = function (r, k) { return r[h.indexOf(k)] || ''; };
  const clients = v.slice(1).filter(function (r) { return r[0] && u.clients.indexOf(r[0]) >= 0; }).map(function (r) {
    return { name: r[0], status: pick(r, '狀態'), meta: pick(r, 'Meta 負責'), google: pick(r, 'Google 負責'), reviewer: pick(r, '確認者'), sheet: pick(r, '工作 Sheet'), pending: pick(r, '待確認'), last: pick(r, '最近產出') };
  });
  return { ok: true, name: u.name, role: u.role, admin: u.admin, clients: clients };
}

function apiClient_(client) {
  const wid = workSheetIdOf_(client);
  if (!wid) return { ok: true, client: client, hasSheet: false };
  const ss = SpreadsheetApp.openById(wid);
  const set = {};
  ss.getSheetByName('設定').getDataRange().getDisplayValues().slice(1).forEach(function (r) { if (r[0]) set[r[0]] = r[1]; });
  const ref = tableOf_(ss.getSheetByName('參考資料'));
  const iA = ref.h.indexOf('活動'), counts = {}, acts = {};
  ref.rows.forEach(function (r) {
    if (!r[0] && !r[3]) return;
    counts[r[0] || '未分類'] = (counts[r[0] || '未分類'] || 0) + 1;
    if (iA >= 0) String(r[iA]).split('、').forEach(function (a) { a = a.trim(); if (a) acts[a] = 1; });
  });
  const rq = tableOf_(ss.getSheetByName('需求'));
  const recent = rq.rows.filter(function (r) { return r.join('').trim(); }).slice(-5).reverse().map(function (r) {
    const o = {}; rq.h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
  });
  return { ok: true, client: client, hasSheet: true, sheet: ss.getUrl(), settings: set, refCounts: counts, activities: Object.keys(acts), recent: recent };
}

function resultsOf_(wid, id) {
  const t = tableOf_(SpreadsheetApp.openById(wid).getSheetByName('產出'));
  const iId = t.h.indexOf('需求編號');
  if (!id) { const last = t.rows.filter(function (r) { return r[iId]; }).pop(); id = last ? last[iId] : ''; }
  return { id: id, rows: t.rows.filter(function (r) { return id && r[iId] === id; }).map(function (r) {
    const o = {}; API_OUT_FIELDS.forEach(function (k) { o[k] = r[t.h.indexOf(k)] || ''; }); return o;
  }) };
}

function apiResults_(client, id) {
  const wid = workSheetIdOf_(client);
  if (!wid) return { ok: false, error: '這個客戶還沒有工作 Sheet' };
  const r = resultsOf_(wid, id);
  return { ok: true, id: r.id, rows: r.rows };
}

function apiGenerate_(u, client, rq) {
  const wid = workSheetIdOf_(client);
  if (!wid) return { ok: false, error: '這個客戶還沒有工作 Sheet' };
  if (!rq['格式']) return { ok: false, error: '請選擇格式' };
  const sh = SpreadsheetApp.openById(wid).getSheetByName('需求');
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = H.map(function (h) { return h === '操作者' ? u.name : (API_REQ_FIELDS.indexOf(h) >= 0 ? String(rq[h] == null ? '' : rq[h]) : ''); });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let rowNum;
  try { sh.appendRow(row); rowNum = sh.getLastRow(); } finally { lock.releaseLock(); }
  Logger.clear();
  generate_(client, true, rowNum);
  const log = Logger.getLog();
  const id = sh.getRange(rowNum, H.indexOf('需求編號') + 1).getDisplayValue();
  if (!id) {
    const m = log.match(/產生失敗：([^\n]+)/);
    return { ok: false, error: m ? m[1] : '產生失敗，請稍後再試', row: rowNum };
  }
  const r = resultsOf_(wid, id);
  return { ok: true, id: id, rows: r.rows };
}

// 只讀測試：用人員分頁第一位「管理」的通行碼，跑「我是誰」和「客戶資料（鼎東策略）」
function testApi() {
  const v = SpreadsheetApp.getActive().getSheetByName('人員').getDataRange().getDisplayValues();
  const ic = v[0].indexOf('通行碼');
  const admin = v.slice(1).find(function (r) { return /管理/.test(r[2]) && r[ic]; });
  if (!admin) { Logger.log('人員分頁找不到有通行碼的管理者'); return; }
  const me = handle_({ action: 'me', code: admin[ic] });
  Logger.log('我是誰：' + me.name + '｜' + me.role + '｜客戶 ' + (me.clients || []).length + ' 家');
  const c = handle_({ action: 'client', code: admin[ic], client: '鼎東策略' });
  Logger.log('鼎東策略｜' + (c.ok ? '工作 Sheet ' + (c.hasSheet ? '有' : '無') + '｜參考資料 ' + JSON.stringify(c.refCounts) + '｜活動 ' + (c.activities || []).length + ' 個｜最近需求 ' + (c.recent || []).length + ' 筆' : c.error));
  const r = handle_({ action: 'results', code: admin[ic], client: '鼎東策略' });
  Logger.log('最新產出：' + (r.ok ? r.id + '｜' + r.rows.length + ' 組' : r.error));
  Logger.log('錯誤通行碼：' + handle_({ action: 'me', code: 'WRONG' }).error);
  Logger.log('沒權限的客戶：' + handle_({ action: 'client', code: 'WRONG', client: '鼎東策略' }).error);
}
