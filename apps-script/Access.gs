// 每人通行碼與可看客戶：角色含「管理」看全部，其他人看自己負責或確認的客戶
const CODE_CHARS = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

function newCode_() {
  let s = '';
  for (let i = 0; i < 8; i++) s += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  return s;
}

function clientsOf_(name, role) {
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const h = v[0], iM = h.indexOf('Meta 負責'), iG = h.indexOf('Google 負責'), iR = h.indexOf('確認者');
  const all = v.slice(1).filter(function (r) { return r[0]; });
  if (/管理/.test(role)) return all.map(function (r) { return r[0]; });
  return all.filter(function (r) { return splitNames_([r[iM], r[iG], r[iR]].join('、')).indexOf(name) >= 0; }).map(function (r) { return r[0]; });
}

// 用通行碼找人：回傳 {name, email, role, admin, clients} 或 null
function userByCode_(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) return null;
  const v = SpreadsheetApp.getActive().getSheetByName('人員').getDataRange().getDisplayValues();
  const h = v[0], ic = h.indexOf('通行碼');
  if (ic < 0) return null;
  const r = v.slice(1).find(function (x) { return x[0] && String(x[ic]).trim().toUpperCase() === code; });
  if (!r) return null;
  return { name: r[0], email: r[1], role: r[2], admin: /管理/.test(r[2]), clients: clientsOf_(r[0], r[2]) };
}

function peopleCodesPreview() { peopleCodes_(false); }
function peopleCodesApply() { peopleCodes_(true); }

function peopleCodes_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const sh = SpreadsheetApp.getActive().getSheetByName('人員');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let ic = h.indexOf('通行碼');
  if (ic < 0) {
    ic = h.length;
    Logger.log(tag + '人員分頁加欄：通行碼（第 ' + (ic + 1) + ' 欄）');
    if (apply) sh.getRange(1, ic + 1).setValue('通行碼').setFontWeight('bold');
  }
  const n = sh.getLastRow() - 1;
  if (n < 1) { Logger.log('人員分頁沒有資料'); return; }
  const v = sh.getRange(2, 1, n, ic + 1).getDisplayValues();
  const used = {};
  v.forEach(function (r) { if (r[ic]) used[String(r[ic]).toUpperCase()] = 1; });
  v.forEach(function (r, i) {
    if (!r[0]) return;
    const cl = clientsOf_(r[0], r[2]);
    const info = '｜' + (r[2] || '（無角色）') + '｜可看客戶 ' + cl.length + (/管理/.test(r[2]) ? '（全部）' : '：' + (cl.join('、') || '（無）'));
    if (r[ic]) { Logger.log(tag + r[0] + '｜已有通行碼' + info); return; }
    let c; do { c = newCode_(); } while (used[c]);
    used[c] = 1;
    Logger.log(tag + r[0] + '｜' + (apply ? '通行碼 ' + c : '會產生通行碼') + info);
    if (apply) sh.getRange(i + 2, ic + 1).setValue(c);
  });
}


// Google 登入：驗證網頁送來的 ID token，再用 Email 對總表「人員」
const OAUTH_CLIENT_ID = '226548623145-jts9mnei1saeue2rnv529h0dsfmnvlpo.apps.googleusercontent.com';

function verifyIdToken_(t) {
  if (!t) return { error: '請先登入' };
  const cache = CacheService.getScriptCache();
  const key = 'idt_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return { email: hit };
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(t), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return { error: '登入已過期，請重新登入' };
  const j = JSON.parse(res.getContentText());
  if (j.aud !== OAUTH_CLIENT_ID) return { error: '登入來源不正確' };
  if (String(j.email_verified) !== 'true') return { error: '這個 Google 帳號的 Email 未驗證' };
  const left = Number(j.exp) - Math.floor(Date.now() / 1000);
  if (!(left > 0)) return { error: '登入已過期，請重新登入' };
  const email = String(j.email).toLowerCase();
  cache.put(key, email, Math.min(left, 3000));
  return { email: email };
}

function userByEmail_(email) {
  email = String(email || '').toLowerCase();
  if (!email) return null;
  const v = SpreadsheetApp.getActive().getSheetByName('人員').getDataRange().getDisplayValues();
  const r = v.slice(1).find(function (x) {
    return x[0] && String(x[1]).toLowerCase().split(/[,，、\s]+/).indexOf(email) >= 0;
  });
  if (!r) return null;
  return { name: r[0], email: email, role: r[2], admin: /管理/.test(r[2]), clients: clientsOf_(r[0], r[2]) };
}
