const SETTINGS_ID = '1bq9FednK81rj5-5dwknu8xL9q5zcVXeojY85l7PZclQ';
const CC_PLAN_ID = '165bctJJfIyK2b0BSoQQD4Bge70bjmH599yQ5WlJJYqE';
const CC_PLAN_GID = 1201308498;

// 只讀：列出成效摘要設定表「客戶主檔」的欄名與創創那列，以及創創媒體規劃表的內容
function inspectChuangchuang() {
  const master = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔');
  if (!master) { Logger.log('找不到分頁：客戶主檔'); return; }
  const v = master.getDataRange().getDisplayValues();
  Logger.log('客戶主檔 欄名：' + v[0].join('｜'));
  v.forEach(function (r, i) {
    if (i > 0 && r.join('').indexOf('創創') >= 0) Logger.log('第 ' + (i + 1) + ' 列：' + r.join('｜'));
  });

  const plan = SpreadsheetApp.openById(CC_PLAN_ID);
  Logger.log('媒體規劃表檔名：' + plan.getName() + '｜分頁：' + plan.getSheets().map(function (s) { return s.getName(); }).join('、'));
  const sh = plan.getSheets().filter(function (s) { return s.getSheetId() === CC_PLAN_GID; })[0];
  if (!sh) { Logger.log('找不到 gid ' + CC_PLAN_GID + ' 的分頁'); return; }
  Logger.log('gid 對應分頁：' + sh.getName());
  sh.getDataRange().getDisplayValues().forEach(function (r, i) {
    const t = r.filter(String).join('｜');
    if (t) Logger.log('第 ' + (i + 1) + ' 列：' + t.replace(/\n/g, ' / '));
  });
}


// 只讀：列出客戶主檔前 4 列（找欄名在哪一列），以及每位客戶的 A 欄名稱與狀態欄
function inspectMasterHeaders() {
  const sh = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔');
  const v = sh.getDataRange().getDisplayValues();
  Logger.log('共 ' + v.length + ' 列、' + (v[0] || []).length + ' 欄');
  for (let i = 0; i < Math.min(4, v.length); i++) {
    Logger.log('第 ' + (i + 1) + ' 列：' + v[i].map(function (x, j) { return (j + 1) + '=' + x; }).filter(function (t) { return !/=$/.test(t); }).join('｜'));
  }
  Logger.log('客戶 A 欄：' + v.slice(1).map(function (r) { return r[0]; }).filter(String).join('、'));
}


// 只讀：列出 A 欄＝報表單位 那一列的每個欄名，以及創創那列對應的值
function inspectMasterHeaderRow() {
  const v = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔').getDataRange().getDisplayValues();
  const hi = v.findIndex(function (r) { return r[0] === '報表單位'; });
  if (hi < 0) { Logger.log('找不到 A 欄＝報表單位'); return; }
  const cc = v.find(function (r) { return r[0] === '創創'; }) || [];
  Logger.log('欄名在第 ' + (hi + 1) + ' 列');
  v[hi].forEach(function (h, j) { if (h || cc[j]) Logger.log((j + 1) + '. ' + (h || '（無欄名）') + ' ＝ ' + (cc[j] || '')); });
}


const HUB_FOLDER_ID = '1KUHIBniAvIMyd4UDBlTrtMpOuMmnVWju';

// 只讀：檢查總表資料夾的名稱、擁有者（共用雲端硬碟的檔案沒有個人擁有者）、裡面已有的東西
function inspectHubFolder() {
  const fo = DriveApp.getFolderById(HUB_FOLDER_ID);
  const owner = fo.getOwner();
  Logger.log('資料夾：' + fo.getName());
  Logger.log('擁有者：' + (owner ? owner.getEmail() + '（在個人雲端硬碟）' : '無個人擁有者（在共用雲端硬碟）'));
  Logger.log('上層：' + (fo.getParents().hasNext() ? fo.getParents().next().getName() : '（無）'));
  const files = fo.getFiles(); const fl = [];
  while (files.hasNext() && fl.length < 20) fl.push(files.next().getName());
  const subs = fo.getFolders(); const sl = [];
  while (subs.hasNext() && sl.length < 20) sl.push(subs.next().getName());
  Logger.log('檔案：' + (fl.join('、') || '（空）'));
  Logger.log('子資料夾：' + (sl.join('、') || '（空）'));
}


// 只讀：列出創創 Meta／Google 進稿表每個分頁的欄名列與前 2 筆內容（每格截 30 字）
function inspectCcSubmissionSheets() {
  const ids = { 'Meta': '165bctJJfIyK2b0BSoQQD4Bge70bjmH599yQ5WlJJYqE', 'Google': '1N_Llrxyglz3Dd75DP-pExMXMIDbjUh72xPNpChhyTSo' };
  const kw = /文案|標題|Headline|說明|描述|Description/;
  const cut = function (x) { x = String(x).replace(/\s+/g, ' '); return x.length > 30 ? x.slice(0, 30) + '…' : x; };
  Object.keys(ids).forEach(function (pf) {
    const book = SpreadsheetApp.openById(ids[pf]);
    book.getSheets().forEach(function (sh) {
      const v = sh.getDataRange().getDisplayValues();
      const hi = v.findIndex(function (r) { return r.some(function (c) { return kw.test(c); }); });
      Logger.log('【' + pf + '】' + sh.getName() + '｜' + v.length + ' 列｜欄名列：' + (hi < 0 ? '找不到' : '第 ' + (hi + 1) + ' 列'));
      if (hi < 0) return;
      Logger.log('　欄名：' + v[hi].map(function (c, j) { return c ? (j + 1) + '=' + cut(c) : ''; }).filter(String).join('｜'));
      const rows = v.slice(hi + 1).filter(function (r) { return r.join('').trim(); });
      Logger.log('　有內容的列：' + rows.length);
      rows.slice(0, 2).forEach(function (r, k) {
        Logger.log('　範例 ' + (k + 1) + '：' + r.map(function (c, j) { return c ? (j + 1) + '=' + cut(c) : ''; }).filter(String).join('｜'));
      });
    });
  });
}
