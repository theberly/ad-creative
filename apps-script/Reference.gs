// 把客戶 Meta 進稿表裡的舊文案匯入工作 Sheet 的「參考資料」
const CC_META_BOOK_ID = '165bctJJfIyK2b0BSoQQD4Bge70bjmH599yQ5WlJJYqE';
const CC_WORK_ID = '1Tb-QZsTwcyQxrOd3Ec2pGlIoKAWzJ8rkr3BBoDLw6ds';

function importCcMetaRefPreview() { importMetaRef_(CC_META_BOOK_ID, CC_WORK_ID, false); }
function importCcMetaRefApply() { importMetaRef_(CC_META_BOOK_ID, CC_WORK_ID, true); }

function importMetaRef_(bookId, workId, apply) {
  const tag = apply ? '' : '[預覽] ';
  const cut = function (x) { x = String(x).replace(/\s+/g, ' '); return x.length > 40 ? x.slice(0, 40) + '…' : x; };
  const book = SpreadsheetApp.openById(bookId);
  const ref = SpreadsheetApp.openById(workId).getSheetByName('參考資料');
  if (!ref) { Logger.log('工作 Sheet 找不到「參考資料」分頁，先停'); return; }

  const seen = {};
  if (ref.getLastRow() > 1) ref.getRange(2, 4, ref.getLastRow() - 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = 1; });
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const out = []; let dup = 0, empty = 0;

  book.getSheets().forEach(function (sh) {
    const v = sh.getDataRange().getDisplayValues();
    if (!v.length) return;
    const h = v[0];
    const iT = h.findIndex(function (x) { return /廣告文案/.test(x); });
    if (iT < 0) return;
    const iH = h.findIndex(function (x) { return /廣告標題/.test(x); });
    const iD = h.findIndex(function (x) { return /廣告說明/.test(x); });
    const iN = h.indexOf('素材');
    v.slice(1).forEach(function (r, k) {
      const t = String(r[iT] || '').trim();
      if (!t) { empty++; return; }
      const parts = ['【主要文字】\n' + t];
      if (iH >= 0 && String(r[iH]).trim()) parts.push('【標題】' + String(r[iH]).trim());
      if (iD >= 0 && String(r[iD]).trim()) parts.push('【說明】' + String(r[iD]).trim());
      const content = parts.join('\n');
      if (seen[content]) { dup++; return; }
      seen[content] = 1;
      const srcName = sh.getName() + '｜第 ' + (k + 2) + ' 列' + (iN >= 0 && r[iN] ? '｜' + r[iN] : '');
      out.push(['僅供參考', 'Meta', 'Meta 廣告', content, srcName, '', today, '程式匯入']);
    });
  });

  Logger.log(tag + '可匯入 ' + out.length + ' 則｜重複略過 ' + dup + '｜沒有文案略過 ' + empty);
  out.slice(0, 5).forEach(function (r) { Logger.log(tag + '範例｜' + r[4] + '｜' + cut(r[3])); });
  if (!apply || !out.length) return;
  ref.getRange(ref.getLastRow() + 1, 1, out.length, 8).setValues(out);
  ref.getRange(2, 4, ref.getLastRow() - 1, 1).setWrap(true);
  ref.setColumnWidth(4, 480);
  Logger.log('已寫入 ' + out.length + ' 則到參考資料');
}
