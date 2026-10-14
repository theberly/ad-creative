// 客戶工作 Sheet：建立、帶入設定、共用給負責人與確認者、連結寫回客戶清單
const WORK_DIR_ID = '140Ik9v9xKzmi-3-MbyTqvlmrFGbdKF40';
const WS_TABS = {
  '設定': ['欄位', '內容'],
  '參考資料': ['類型（成效好／客戶提供／僅供參考）', '平台', '格式', '內容', '來源／連結', '成效備註', '加入日期', '加入者', '活動'],
  '需求': ['需求編號', '建立時間', '操作者', '平台', '活動名稱', '格式', '受眾', '組數', '產品／服務', '活動目標', '到達網址', '補充說明', '狀態'],
  '產出': ['需求編號', '產出時間', '平台', '格式', '序號', '受眾', '文案', '標題', '說明', 'CTA', '到達網址', '字數（文案／標題／說明）', '是否超限', '審核警示', '確認狀態', '確認者', '確認時間', '修改備註']
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
  const mrow = masterRowOf_(name);
  if (!set['Google 帳號 ID'] && mrow['Google 帳號 ID']) set['Google 帳號 ID'] = mrow['Google 帳號 ID'];
  if (!set['Meta 帳號 ID'] && mrow['Meta 廣告帳號 ID']) set['Meta 帳號 ID'] = mrow['Meta 廣告帳號 ID'];

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
  Logger.log(tag + '之後要共用編輯權（這次不共用）：' + people.join('、') + '｜' + (emails.join('、') || '（無）'));
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
  nss.getSheetByName('產出').getRange(2, WS_TABS['產出'].indexOf('確認狀態') + 1, 999, 1).setDataValidation(dv(['待確認', '已確認', '退回']));

  Logger.log('開發中先不共用，之後會共用給：' + (emails.join('、') || '（無）'));

  list.getRange(li + 1, lh.indexOf('工作 Sheet') + 1).setValue(nss.getUrl());
  Logger.log('已建立並寫回客戶清單：' + nss.getUrl());
}


// 客戶主檔中某客戶那列，回傳 {欄名: 值}
function masterRowOf_(name) {
  const v = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔').getDataRange().getDisplayValues();
  const hi = v.findIndex(function (r) { return r[0] === '報表單位'; });
  if (hi < 0) return {};
  const row = v.slice(hi + 1).find(function (r) { return r[0] === name; });
  if (!row) return {};
  const o = {};
  v[hi].forEach(function (k, i) { if (k) o[k.trim()] = row[i]; });
  return o;
}

function createDingdongWorkSheetPreview() { createWorkSheet_('鼎東策略', false); }
function createDingdongWorkSheetApply() { createWorkSheet_('鼎東策略', true); }


// 把客戶清單裡所有工作 Sheet 升級成目前的範本欄位（只補缺的，不刪資料）
function upgradeWorkSheetsPreview() { upgradeWorkSheets_(false); }
function upgradeWorkSheetsApply() { upgradeWorkSheets_(true); }

function upgradeWorkSheets_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  v.slice(1).forEach(function (r) {
    const mm = String(r[iW]).match(/\/d\/([^/]+)/);
    if (!mm) return;
    const name = r[0], ss = SpreadsheetApp.openById(mm[1]);

    const ref = ss.getSheetByName('參考資料');
    const wantR = WS_TABS['參考資料'];
    const rh = ref.getRange(1, 1, 1, ref.getLastColumn()).getValues()[0];
    wantR.forEach(function (h) {
      if (rh.indexOf(h) >= 0) return;
      Logger.log(tag + name + '｜參考資料加欄：' + h + '（第 ' + (rh.length + 1) + ' 欄）');
      if (apply) ref.getRange(1, rh.length + 1).setValue(h).setFontWeight('bold');
      rh.push(h);
    });
    const n = ref.getLastRow() - 1;
    Logger.log(tag + name + '｜參考資料 ' + n + ' 列改成固定列高');
    if (apply && n > 0) {
      ref.getRange(2, 1, n, rh.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('top');
      ref.setRowHeightsForced(2, n, 21);
    }

    const po = ss.getSheetByName('產出');
    const wantP = WS_TABS['產出'];
    const ph = po.getRange(1, 1, 1, po.getLastColumn()).getValues()[0];
    if (ph.join('|') !== wantP.join('|')) {
      const hasData = po.getLastRow() > 1;
      Logger.log(tag + name + '｜產出改成新版欄位' + (hasData ? '（舊資料移到「產出_舊版」）' : ''));
      if (apply) {
        let target = po;
        if (hasData) {
          po.setName('產出_舊版' + (ss.getSheetByName('產出_舊版') ? '_' + Date.now() : ''));
          target = ss.insertSheet('產出', po.getIndex());
        } else {
          po.clear();
          po.getRange(1, 1, po.getMaxRows(), po.getMaxColumns()).clearDataValidations();
        }
        target.getRange(1, 1, 1, wantP.length).setValues([wantP]).setFontWeight('bold');
        target.setFrozenRows(1);
        target.getRange(2, wantP.indexOf('確認狀態') + 1, 999, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['待確認', '已確認', '退回'], true).setAllowInvalid(false).build());
      }
    } else { Logger.log(tag + name + '｜產出欄位已是最新'); }

    const dq = ss.getSheetByName('需求');
    const wantD = WS_TABS['需求'];
    const dh = dq.getRange(1, 1, 1, dq.getLastColumn()).getValues()[0];
    if (dh.join('|') === wantD.join('|')) { Logger.log(tag + name + '｜需求欄位已是最新'); return; }
    if (dq.getLastRow() > 1) { Logger.log('注意：' + name + ' 的需求已有資料，欄位不自動改'); return; }
    Logger.log(tag + name + '｜需求欄位改成：' + wantD.join('、'));
    if (apply) {
      dq.clear();
      dq.getRange(1, 1, 1, wantD.length).setValues([wantD]).setFontWeight('bold');
      dq.setFrozenRows(1);
    }
  });
}
