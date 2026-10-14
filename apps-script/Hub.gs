// 總表：客戶清單（同步自成效摘要設定表的客戶主檔）、人員、新客戶
const HUB_TABS = {
  '客戶清單': ['客戶', '狀態', 'Meta 負責', 'Google 負責', '確認者', '工作 Sheet', '待確認', '最近產出', '同步時間'],
  '人員': ['名字', 'Email', '角色（操作／確認／管理）', '備註'],
  '新客戶': ['客戶名稱', 'Meta 負責', 'Google 負責', '確認者', '建立狀態', '工作 Sheet']
};
const HUB_PEOPLE = [
  ['Keani', '', '管理／確認', ''], ['Kimberly', '', '確認', ''],
  ['Coco', '', '操作', ''], ['珍珠', '', '操作', ''], ['Lennard', '', '操作', ''], ['Clytie', '', '操作', ''], ['Emma', '', '操作', '']
];

function hubSetupPreview() { hubSetup_(false); }
function hubSetupApply() { hubSetup_(true); }
function syncClientsPreview() { syncClients_(false); }
function syncClientsApply() { syncClients_(true); }

function hubSetup_(apply) {
  const ss = SpreadsheetApp.getActive();
  const tag = apply ? '' : '[預覽] ';
  Object.keys(HUB_TABS).forEach(function (name) {
    if (ss.getSheetByName(name)) { Logger.log(tag + '已存在，略過：' + name); return; }
    const h = HUB_TABS[name];
    Logger.log(tag + '建立分頁：' + name + '｜' + h.join('、'));
    if (name === '人員') Logger.log(tag + '人員先填：' + HUB_PEOPLE.map(function (p) { return p[0] + '（' + p[2] + '）'; }).join('、'));
    if (!apply) return;
    const sh = ss.insertSheet(name, 0);
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (name === '人員') sh.getRange(2, 1, HUB_PEOPLE.length, 4).setValues(HUB_PEOPLE);
  });
  syncClients_(apply);
}

function syncClients_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const v = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔').getDataRange().getDisplayValues();
  const hi = v.findIndex(function (r) { return r[0] === '報表單位'; });
  if (hi < 0) { Logger.log('客戶主檔找不到 A 欄＝報表單位，先停'); return; }
  const H = v[hi];
  const iS = H.indexOf('狀態'), iM = H.indexOf('撰寫者(Meta)'), iG = H.indexOf('撰寫者(Google)');
  if (iS < 0 || iM < 0 || iG < 0) { Logger.log('客戶主檔找不到 狀態／撰寫者(Meta)／撰寫者(Google) 欄，先停'); return; }
  const src = v.slice(hi + 1).filter(function (r) { return r[0] && r[iS]; });

  const sh = SpreadsheetApp.getActive().getSheetByName('客戶清單');
  if (!sh && apply) { Logger.log('找不到分頁：客戶清單，先跑 hubSetupApply'); return; }
  const cur = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues() : [];
  const at = {}; cur.forEach(function (r, i) { at[r[0]] = { row: i + 2, v: r }; });
  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  let add = 0, upd = 0, skip = 0;

  src.forEach(function (r) {
    const name = r[0], st = r[iS], m = r[iM], g = r[iG];
    const e = at[name];
    if (!e) {
      if (st === '已結束') { skip++; return; }
      add++;
      Logger.log(tag + '新增：' + name + '｜' + st + '｜Meta：' + (m || '—') + '｜Google：' + (g || '—'));
      if (apply) sh.appendRow([name, st, m, g, '', '', '', '', now]);
      return;
    }
    if (e.v[1] !== st || e.v[2] !== m || e.v[3] !== g) {
      upd++;
      Logger.log(tag + '更新：' + name + '｜' + e.v.slice(1).join('／') + ' → ' + [st, m, g].join('／'));
      if (apply) { sh.getRange(e.row, 2, 1, 3).setValues([[st, m, g]]); sh.getRange(e.row, 9).setValue(now); }
    }
  });
  Logger.log(tag + '客戶清單：新增 ' + add + '、更新 ' + upd + '、略過已結束 ' + skip);
}
