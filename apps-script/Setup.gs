const SETUP_TABS = [
  { name: '需求紀錄', headers: ['需求編號', '建立時間', '操作者', '客戶', '是否新客戶', '勾選格式', '產品／服務', '受眾', '活動目標', '品牌語氣', '必提賣點', '禁用字詞', '參考歷史成效', '備註'] },
  { name: '文案產出', headers: ['需求編號', '產出時間', '客戶', '格式', '欄位', '序號', '文案', '字數（半形計）', '上限', '是否超限', '審核警示', '確認狀態', '確認者', '確認時間', '修改後文案'] },
  { name: '客戶設定', headers: ['客戶名稱', '是否新客戶', 'Google 帳號 ID', 'Meta 帳號 ID', '品牌語氣', '必提賣點', '禁用字詞', '素材資料夾連結', '媒體規劃表連結', '備註', '更新時間'] },
  { name: '審核字庫', headers: ['類別', '字詞', '處理方式（擋下／提醒）', '說明', '適用客戶（空白＝全部）'] }
];

function setupPreview() { setup_(false); }
function setupApply() { setup_(true); }

function setup_(apply) {
  const ss = SpreadsheetApp.getActive();
  const tag = apply ? '' : '[預覽] ';
  SETUP_TABS.forEach(function (t) {
    if (ss.getSheetByName(t.name)) { Logger.log(tag + '已存在，略過：' + t.name); return; }
    Logger.log(tag + '建立分頁：' + t.name + '｜欄位：' + t.headers.join('、'));
    if (!apply) return;
    const sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) {
    Logger.log(tag + '刪除空白預設分頁：' + def.getName());
    if (apply && ss.getSheets().length > 1) ss.deleteSheet(def);
  }
}
