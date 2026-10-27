/**
 * theberly 文案產生器：從 MCC 匯出各帳戶的回應式搜尋廣告（RSA）文案與成效
 * 用法：Google Ads MCC → 工具 → 批量操作 → 指令碼 → 新增 → 貼上 → 授權 → 執行
 * 只讀取廣告，不會修改任何帳戶。
 * 第一次執行會建立一份試算表，網址印在「記錄」裡；之後執行會寫回同一份（把網址貼到 SHEET_URL）。
 */
var SHEET_URL = '';     // 第一次留空；之後貼上第一次建立的試算表網址
var DAYS = 180;         // 取最近幾天的成效
var MIN_COST = 0;       // 花費低於這個金額的廣告不匯出

function main() {
  var end = new Date(), start = new Date(end.getTime() - DAYS * 86400000);
  var fmt = function (d) { return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd'); };
  var head = ['帳戶 ID', '帳戶名稱', '廣告活動', '廣告群組', '廣告 ID', '狀態', '標題', '說明', '花費', '點擊', '曝光', '轉換', '每次轉換成本', '期間'];
  var out = [];
  var it = AdsManagerApp.accounts().get();
  while (it.hasNext()) {
    var acc = it.next();
    AdsManagerApp.select(acc);
    var id = acc.getCustomerId(), name = acc.getName();
    var n = 0;
    try {
      var rows = AdsApp.search(
        'SELECT campaign.name, ad_group.name, ad_group_ad.ad.id, ad_group_ad.status, ' +
        'ad_group_ad.ad.responsive_search_ad.headlines, ad_group_ad.ad.responsive_search_ad.descriptions, ' +
        'metrics.cost_micros, metrics.clicks, metrics.impressions, metrics.conversions ' +
        'FROM ad_group_ad WHERE ad_group_ad.ad.type = RESPONSIVE_SEARCH_AD ' +
        "AND segments.date BETWEEN '" + fmt(start) + "' AND '" + fmt(end) + "' " +
        'AND metrics.cost_micros > ' + Math.round(MIN_COST * 1e6));
      while (rows.hasNext()) {
        var r = rows.next(), ad = r.adGroupAd.ad, rsa = ad.responsiveSearchAd || {};
        var txt = function (list) { return (list || []).map(function (a) { return a.text; }).join('\n'); };
        var cost = Number(r.metrics.costMicros || 0) / 1e6, conv = Number(r.metrics.conversions || 0);
        out.push([id, name, r.campaign.name, r.adGroup.name, String(ad.id), r.adGroupAd.status,
          txt(rsa.headlines), txt(rsa.descriptions), Math.round(cost), Number(r.metrics.clicks || 0),
          Number(r.metrics.impressions || 0), Math.round(conv * 100) / 100, conv ? Math.round(cost / conv) : '',
          fmt(start) + '～' + fmt(end)]);
        n++;
      }
    } catch (e) { Logger.log('✗ ' + name + '（' + id + '）：' + e.message); continue; }
    Logger.log((n ? '✓ ' : '— ') + name + '（' + id + '）：' + n + ' 則');
  }

  var ss = SHEET_URL ? SpreadsheetApp.openByUrl(SHEET_URL) : SpreadsheetApp.create('文案產生器_Google 廣告匯出');
  var sh = ss.getSheets()[0];
  sh.clear();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (out.length) sh.getRange(2, 1, out.length, head.length).setValues(out);
  sh.setFrozenRows(1);
  Logger.log('共 ' + out.length + ' 則，寫入：' + ss.getUrl());
}
