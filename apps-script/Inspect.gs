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
