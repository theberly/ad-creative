const CC_NAME = '創創';
const CC_META_URL = 'https://docs.google.com/spreadsheets/d/165bctJJfIyK2b0BSoQQD4Bge70bjmH599yQ5WlJJYqE/edit?gid=1201308498';
const CC_GOOGLE_ID = '1N_Llrxyglz3Dd75DP-pExMXMIDbjUh72xPNpChhyTSo';
const CC_GOOGLE_URL = 'https://docs.google.com/spreadsheets/d/' + CC_GOOGLE_ID + '/edit?gid=0';

function chuangchuangUpdatePreview() { ccUpdate_(false); }
function chuangchuangUpdateApply() { ccUpdate_(true); }

function ccUpdate_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const sh = SpreadsheetApp.getActive().getSheetByName('客戶設定');
  if (!sh) { Logger.log('找不到分頁：客戶設定'); return; }

  try {
    const g = SpreadsheetApp.openById(CC_GOOGLE_ID);
    Logger.log('Google 進稿表讀得到：' + g.getName() + '｜分頁：' + g.getSheets().map(function (s) { return s.getName(); }).join('、'));
  } catch (e) { Logger.log('讀不到 Google 進稿表，先停：' + e.message); return; }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const oldI = headers.indexOf('媒體規劃表連結');
  if (oldI >= 0) {
    Logger.log(tag + '欄名改：媒體規劃表連結 → Meta 進稿表連結');
    if (apply) sh.getRange(1, oldI + 1).setValue('Meta 進稿表連結');
    headers[oldI] = 'Meta 進稿表連結';
  }
  const mi = headers.indexOf('Meta 進稿表連結');
  if (mi < 0) { Logger.log('找不到 Meta 進稿表連結 欄，先停'); return; }
  if (headers.indexOf('Google 進稿表連結') < 0) {
    Logger.log(tag + '新增欄位：Google 進稿表連結（第 ' + (mi + 2) + ' 欄）');
    if (apply) {
      sh.insertColumnAfter(mi + 1);
      sh.getRange(1, mi + 2).setValue('Google 進稿表連結').setFontWeight('bold');
    }
    headers.splice(mi + 1, 0, 'Google 進稿表連結');
  }

  const names = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(function (r) { return r[0]; });
  const ri = names.indexOf(CC_NAME);
  if (ri < 0) { Logger.log('客戶設定找不到「' + CC_NAME + '」，先停'); return; }
  const rowNum = ri + 2;

  const upd = {
    'Meta 進稿表連結': CC_META_URL,
    'Google 進稿表連結': CC_GOOGLE_URL,
    '品牌語氣': '字少、有衝勁，可以浮誇（例：南部鄉親～計劃案專家來了！）',
    '禁用字詞': '連結、補助金、政府、實戰（改用：計劃案、企劃案、實操）',
    '備註': '只看前三行＋三個大重點；分段空行要清楚，字越少越好；素材不分區；Google、Meta 素材分開提供；客戶給的文案通常要重寫，產出後 @小愛 確認',
    '更新時間': Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm')
  };
  const miss = Object.keys(upd).filter(function (h) { return headers.indexOf(h) < 0; });
  if (miss.length) { Logger.log('找不到欄位，先停：' + miss.join('、')); return; }
  Object.keys(upd).forEach(function (h) {
    Logger.log(tag + '第 ' + rowNum + ' 列｜' + h + '：' + upd[h]);
    if (apply) sh.getRange(rowNum, headers.indexOf(h) + 1).setValue(upd[h]);
  });
  if (apply) Logger.log('已更新第 ' + rowNum + ' 列');
}
