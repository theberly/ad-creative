// 客戶工作 Sheet：建立、帶入設定、共用給負責人與確認者、連結寫回客戶清單
const WORK_DIR_ID = '140Ik9v9xKzmi-3-MbyTqvlmrFGbdKF40';
const WS_TABS = {
  '設定': ['欄位', '內容'],
  '參考資料': ['類型（成效好／客戶提供／僅供參考）', '平台', '格式', '內容', '來源／連結', '成效備註', '加入日期', '加入者'],
  '需求': ['需求編號', '建立時間', '操作者', '平台', '格式', '受眾', '組數', '產品／服務', '活動目標', '補充說明', '狀態'],
  '產出': ['需求編號', '產出時間', '平台', '格式', '欄位', '序號', '受眾', '文案', '字數（半形計）', '上限', '是否超限', '審核警示', '確認狀態', '確認者', '確認時間', '修正版文案']
};
const WS_SETTING_KEYS = ['客戶名稱', 'Google 帳號 ID', 'Meta 帳號 ID', '品牌語氣', '必提賣點', '禁用字詞', '素材資料夾連結', 'Meta 進稿表連結', 'Google 進稿表連結', '備註'];

function createCcWorkSheetPreview() { createWorkSheet_('創創', false); }
function createCcWorkSheetApply() { createWorkSheet_('創創', true); }

function splitNames_(s) {
  return String(s || '').split(/[、，,]+/).map(function (x) { return x.trim(); }).filter(String);
}

function createWorkSheet_(name, apply) {
  const tag = apply ? '' : '[預覽] ';
  const ss = SpreadsheetApp.getActive();

  const list = ss.getSheetByName('客戶清單');
  const lv = list.getDataRange().getDisplayValues(), lh = lv[0];
  const li = lv.findIndex(function (r, i) { return i > 0 && r[0] === name; });
  if (li < 0) { Logger.log('客戶清單找不到「' + name + '」，先停'); return; }
  const col = function (k) { return lv[li][lh.indexOf(k)]; };
  if (col('工作 Sheet')) { Logger.log(tag + '「' + name + '」已有工作 Sheet，略過：' + col('工作 Sheet')); return; }

  const set = {};
  const cs = ss.getSheetByName('客戶設定');
  if (cs) {
    const cv = cs.getDataRange().getDisplayValues();
    const row = cv.find(function (r, i) { return i > 0 && r[0] === name; });
    if (row) cv[0].forEach(function (k, i) { set[k] = row[i]; });
  }
  set['客戶名稱'] = name;

  const emailOf = {};
  ss.getSheetByName('人員').getDataRange().getDisplayValues().slice(1).forEach(function (r) {
    if (r[0]) emailOf[r[0].trim()] = String(r[1]).split(/[,，、\s]+/).filter(function (x) { return x.indexOf('@') > 0; });
  });
  const people = [];
  splitNames_([col('Meta 負責'), col('Google 負責'), col('確認者')].join('、')).forEach(function (n) { if (people.indexOf(n) < 0) people.push(n); });
  const emails = [], missing = [];
  people.forEach(function (n) {
    const e = emailOf[n] || [];
    if (!e.length) missing.push(n);
    e.forEach(function (x) { if (emails.indexOf(x) < 0) emails.push(x); });
  });

  const title = name + '_文案產生器';
  Logger.log(tag + '建立：' + title + '（放在 06_文案素材產生器）');
  WS_SETTING_KEYS.forEach(function (k) { Logger.log(tag + '設定｜' + k + '：' + (set[k] || '（空）')); });
  Object.keys(WS_TABS).forEach(function (t) { Logger.log(tag + '分頁：' + t + '｜' + WS_TABS[t].join('、')); });
  Logger.log(tag + '共用編輯權：' + people.join('、') + '｜' + (emails.join('、') || '（無）'));
  if (missing.length) Logger.log('注意：人員分頁找不到 Email：' + missing.join('、'));
  if (!apply) return;

  const nss = SpreadsheetApp.create(title);
  DriveApp.getFileById(nss.getId()).moveTo(DriveApp.getFolderById(WORK_DIR_ID));
  nss.setSpreadsheetTimeZone('Asia/Taipei');
  const first = nss.getSheets()[0];
  first.setName('設定');
  Object.keys(WS_TABS).forEach(function (t, i) {
    const sh = i === 0 ? first : nss.insertSheet(t);
    const h = WS_TABS[t];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });

  const st = nss.getSheetByName('設定');
  const n = WS_SETTING_KEYS.length;
  st.getRange(2, 2, n, 1).setNumberFormat('@');
  st.getRange(2, 1, n, 2).setValues(WS_SETTING_KEYS.map(function (k) { return [k, set[k] || '']; }));
  st.getRange(2, 2, n, 1).setWrap(true);
  st.setColumnWidth(1, 150); st.setColumnWidth(2, 640);

  const dv = function (vals) { return SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build(); };
  nss.getSheetByName('參考資料').getRange(2, 1, 999, 1).setDataValidation(dv(['成效好', '客戶提供', '僅供參考']));
  nss.getSheetByName('參考資料').getRange(2, 2, 999, 1).setDataValidation(dv(['Google', 'Meta', '共用']));
  nss.getSheetByName('產出').getRange(2, 13, 999, 1).setDataValidation(dv(['待確認', '已確認', '退回']));

  emails.forEach(function (e) {
    try { nss.addEditor(e); Logger.log('已共用：' + e); }
    catch (err) { Logger.log('共用失敗：' + e + '｜' + err.message); }
  });

  list.getRange(li + 1, lh.indexOf('工作 Sheet') + 1).setValue(nss.getUrl());
  Logger.log('已建立並寫回客戶清單：' + nss.getUrl());
}
