const CC_NAME = '創創';
const CC_PLAN_URL = 'https://docs.google.com/spreadsheets/d/165bctJJfIyK2b0BSoQQD4Bge70bjmH599yQ5WlJJYqE/edit?gid=1201308498';

function chuangchuangPreview() { chuangchuang_(false); }
function chuangchuangApply() { chuangchuang_(true); }

function chuangchuang_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const sh = SpreadsheetApp.getActive().getSheetByName('客戶設定');
  if (!sh) { Logger.log('找不到分頁：客戶設定，先跑 Setup.gs 的 setupApply'); return; }

  // 1. 補欄位：媒體規劃表連結（放在素材資料夾連結後面）
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('媒體規劃表連結') < 0) {
    const at = headers.indexOf('素材資料夾連結');
    if (at < 0) { Logger.log('找不到「素材資料夾連結」欄，先停'); return; }
    Logger.log(tag + '新增欄位：媒體規劃表連結（第 ' + (at + 2) + ' 欄）');
    if (apply) {
      sh.insertColumnAfter(at + 1);
      sh.getRange(1, at + 2).setValue('媒體規劃表連結').setFontWeight('bold');
      headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    } else {
      headers.splice(at + 1, 0, '媒體規劃表連結');
    }
  } else {
    Logger.log(tag + '欄位已存在：媒體規劃表連結');
  }

  // 2. 已有創創就不重寫
  const names = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  if (names.indexOf(CC_NAME) >= 0) { Logger.log(tag + '客戶設定已有「' + CC_NAME + '」，略過'); return; }

  // 3. 從成效摘要設定表抓帳號 ID
  const mv = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔').getDataRange().getDisplayValues();
  const row = mv.filter(function (r) { return r[0] === CC_NAME; })[0];
  if (!row) { Logger.log('客戶主檔找不到 A 欄＝' + CC_NAME + '，先停'); return; }
  const meta = row.filter(function (x) { return /^act_\d+$/.test(x); })[0] || '';
  const google = row.filter(function (x) { return /^\d{3}-\d{3}-\d{4}$/.test(x); })[0] || '';
  if (!meta || !google) Logger.log('注意：帳號 ID 沒抓齊｜Meta=' + meta + '｜Google=' + google);

  const data = {
    '客戶名稱': CC_NAME,
    '是否新客戶': '否',
    'Google 帳號 ID': google,
    'Meta 帳號 ID': meta,
    '品牌語氣': '字少、有衝勁，可以浮誇（例：南部鄉親～計劃案專家來了！）',
    '必提賣點': '',
    '禁用字詞': '補助金、政府、實戰（改用：計劃案、企劃案、實操）',
    '素材資料夾連結': '',
    '媒體規劃表連結': CC_PLAN_URL,
    '備註': '重點寫在前三行＋三大重點＋地區字眼；Google、Meta 素材分開提供',
    '更新時間': Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm')
  };
  const out = headers.map(function (h) { return data.hasOwnProperty(h) ? data[h] : ''; });
  headers.forEach(function (h, i) { Logger.log(tag + h + '：' + (out[i] || '（空）')); });
  if (apply) {
    sh.appendRow(out);
    sh.getRange(sh.getLastRow(), headers.indexOf('Google 帳號 ID') + 1).setNumberFormat('@').setValue(google);
    Logger.log('已寫入第 ' + sh.getLastRow() + ' 列');
  }
}
