// Google RSA 參考資料：讀 Google Ads 指令碼（google-ads/export-rsa.js）匯出的試算表 → 依 Google 帳號 ID 寫進各家參考資料
// 成效好：花費 ≥ 500 的廣告中，有轉換的帳戶看每次轉換成本、沒轉換的帳戶看點擊率，取前 1/4
const GADS_EXPORT_ID = '1SgEyEyW-hQtVr4Bga7yKteDMUnFFhLVfpvrOBjpKhd8';

function importGoogleRefPreview() { importGoogleRef_(false); }
function importGoogleRefApply() { importGoogleRef_(true); }

function importGoogleRef_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const ev = SpreadsheetApp.openById(GADS_EXPORT_ID).getSheets()[0].getDataRange().getValues();
  const H = ev[0], c = function (k) { return H.indexOf(k); };
  const ads = ev.slice(1).filter(function (r) { return r[c('標題')]; });
  const digits = function (s) { return String(s || '').replace(/\D/g, ''); };
  const fmt = function (x) { return Math.round(x).toLocaleString('en-US'); };
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  Logger.log(tag + '匯出檔共 ' + ads.length + ' 則（期間 ' + (ads[0] ? ads[0][c('期間')] : '—') + '）');

  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  const used = {};
  v.slice(1).forEach(function (r) {
    const name = r[0];
    if (!name || !r[iW]) return;
    const gid = digits(masterRowOf_(name)['Google 帳號 ID']);
    if (!gid) return;
    used[gid] = 1;
    const mine = ads.filter(function (a) { return digits(a[c('帳戶 ID')]) === gid; });
    if (!mine.length) { Logger.log('— ' + name + '（' + gid + '）：匯出檔沒有 RSA'); return; }
    const sh = SpreadsheetApp.openById(workSheetIdOf_(name)).getSheetByName('參考資料');
    const has = sh.getDataRange().getDisplayValues().slice(1).some(function (x) { return x[1] === 'Google' && /Google 成效/.test(x[7]); });
    if (has) { Logger.log('— ' + name + '：已匯入過 Google 參考資料，略過'); return; }

    // 同樣文案合併
    const by = {};
    mine.forEach(function (a) {
      const content = '【標題】\n' + a[c('標題')] + '\n【說明】\n' + a[c('說明')];
      const m = by[content] || (by[content] = { content: content, cost: 0, clicks: 0, impr: 0, conv: 0, ids: [], camps: {} });
      m.cost += Number(a[c('花費')]) || 0; m.clicks += Number(a[c('點擊')]) || 0; m.impr += Number(a[c('曝光')]) || 0; m.conv += Number(a[c('轉換')]) || 0;
      m.ids.push(a[c('廣告 ID')]); m.camps[a[c('廣告活動')]] = 1;
    });
    const list = Object.keys(by).map(function (k) { return by[k]; });
    const useConv = list.some(function (m) { return m.conv > 0; });
    const score = function (m) { return useConv ? (m.conv ? m.cost / m.conv : Infinity) : (m.impr ? -m.clicks / m.impr : Infinity); };
    const elig = list.filter(function (m) { return m.cost >= 500 && score(m) < Infinity; }).sort(function (a, b) { return score(a) - score(b); });
    const good = {};
    elig.slice(0, Math.max(1, Math.ceil(elig.length / 4))).forEach(function (m) { good[m.content] = 1; });
    if (!elig.length) Object.keys(good).forEach(function (k) { delete good[k]; });

    const out = list.sort(function (a, b) { return (good[b.content] ? 1 : 0) - (good[a.content] ? 1 : 0) || b.cost - a.cost; }).map(function (m) {
      const note = '搜尋｜花費 ' + fmt(m.cost) + '｜點擊 ' + fmt(m.clicks) + '｜點擊率 ' + (m.impr ? (m.clicks / m.impr * 100).toFixed(1) : '0') + '%' +
        (m.conv ? '｜轉換 ' + (Math.round(m.conv * 10) / 10) + '｜每次轉換 ' + fmt(m.cost / m.conv) + ' 元' : '');
      return [good[m.content] ? '成效好' : '僅供參考', 'Google', 'Google RSA', m.content, '廣告 ID：' + m.ids.join('、'), note, today, '程式匯入（Google 成效）', Object.keys(m.camps).join('、')];
    });
    Logger.log(tag + '✓ ' + name + '（' + gid + '）：' + mine.length + ' 則廣告，合併後 ' + out.length + ' 則，成效好 ' + Object.keys(good).length + '（看' + (useConv ? '每次轉換成本' : '點擊率') + '）' +
      (out[0] ? '\n  例：' + out[0][0] + '｜' + out[0][5] + '\n  ' + String(out[0][3]).split('\n').slice(1, 4).join('／') : ''));
    if (!apply) return;
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, out.length, out[0].length).setValues(out).setWrap(true).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.setRowHeights(start, out.length, 21);
  });
  const rest = {};
  ads.forEach(function (a) { if (!used[digits(a[c('帳戶 ID')])]) rest[a[c('帳戶名稱')]] = (rest[a[c('帳戶名稱')]] || 0) + 1; });
  if (Object.keys(rest).length) Logger.log('對不到客戶清單（沒有工作 Sheet 或主檔沒填 Google 帳號 ID）：' + Object.keys(rest).map(function (k) { return k + ' ' + rest[k] + ' 則'; }).join('、'));
}
