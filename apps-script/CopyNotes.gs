// 把廣告進度控管「帳戶狀況」的素材進稿注意事項，帶進各客戶工作 Sheet 設定的「備註」
// 只讀廣告進度控管；工作 Sheet 的備註已有內容就不覆蓋
const AP_FILE_NAME = 'theberly*_上下線進度 & 計算消耗';

function apAccounts_() {
  const it = DriveApp.getFilesByName(AP_FILE_NAME);
  if (!it.hasNext()) throw new Error('找不到廣告進度控管的試算表：' + AP_FILE_NAME);
  const sh = SpreadsheetApp.openById(it.next().getId()).getSheetByName('帳戶狀況');
  if (!sh) throw new Error('找不到分頁：帳戶狀況');
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(function (x) { return String(x).replace(/\s+/g, ''); });
  const col = function (k) { return h.findIndex(function (x) { return x.indexOf(k) === 0; }); };
  const iC = col('客戶'), iA = col('帳戶名稱'), iM = col('媒體'), iMat = col('素材進稿注意事項');
  if (iMat < 0) throw new Error('帳戶狀況找不到「素材進稿注意事項」欄');
  return v.slice(1).filter(function (r) { return (iC >= 0 && r[iC]) || (iA >= 0 && r[iA]); }).map(function (r) {
    return { client: iC >= 0 ? r[iC] : '', account: iA >= 0 ? r[iA] : '', media: iM >= 0 ? r[iM] : '', mat: String(r[iMat] || '').trim() };
  });
}

function syncCopyNotesPreview() { syncCopyNotes_(false); }
function syncCopyNotesApply() { syncCopyNotes_(true); }

function syncCopyNotes_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const rows = apAccounts_();
  Logger.log(tag + '帳戶狀況共 ' + rows.length + ' 列');
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  let filled = 0;
  v.slice(1).forEach(function (r) {
    const name = r[0];
    if (!name || !r[iW]) return;
    const mrow = masterRowOf_(name);
    const accts = Object.keys(mrow).filter(function (k) { return /帳戶|對照/.test(k); }).map(function (k) { return mrow[k]; }).filter(String);
    const hits = rows.filter(function (x) { return x.mat && (String(x.client).indexOf(name) >= 0 || accts.indexOf(x.account) >= 0); });
    const uniq = [];
    hits.forEach(function (x) { if (!uniq.some(function (u) { return u.mat === x.mat; })) uniq.push(x); });
    if (!uniq.length) { Logger.log('— ' + name + '：帳戶狀況沒有對應的注意事項'); return; }
    const text = uniq.length === 1 ? uniq[0].mat : uniq.map(function (x) { return '【' + (x.media || x.account) + '】\n' + x.mat; }).join('\n\n');

    const st = SpreadsheetApp.openById(workSheetIdOf_(name)).getSheetByName('設定');
    const sv = st.getDataRange().getDisplayValues();
    const ri = sv.findIndex(function (x) { return x[0] === '備註'; });
    if (ri < 0) { Logger.log('✗ ' + name + '：設定分頁找不到「備註」列'); return; }
    if (String(sv[ri][1]).trim()) { Logger.log('— ' + name + '：備註已有內容，略過'); return; }
    filled++;
    Logger.log(tag + '✓ ' + name + '（對到 ' + uniq.map(function (x) { return x.client + '／' + x.account; }).join('、') + '）\n' + text.slice(0, 300) + (text.length > 300 ? '…' : ''));
    if (apply) st.getRange(ri + 1, 2).setValue(text).setWrap(true);
  });
  Logger.log(tag + (apply ? '已寫入 ' : '會寫入 ') + filled + ' 家');
  if (!apply) Logger.log('帳戶狀況的客戶／帳戶：' + rows.map(function (x) { return x.client + '／' + x.account + (x.mat ? '' : '（無注意事項）'); }).join('、'));
}
