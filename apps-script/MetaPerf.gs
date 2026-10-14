// Meta 廣告成效：讀帳戶內每則廣告的文案與成效，對應到工作 Sheet 的「參考資料」
const META_API = 'https://graph.facebook.com/v23.0/';
const CC_META_ACT = 'act_2093852258108102';

function metaGet_(path, params) {
  const token = PropertiesService.getScriptProperties().getProperty('META_TOKEN');
  if (!token) throw new Error('找不到指令碼屬性 META_TOKEN');
  const q = Object.keys(params || {}).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  const url = path.indexOf('https://') === 0 ? path : META_API + path + '?' + q + '&access_token=' + encodeURIComponent(token);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText());
  if (body.error) throw new Error('Meta API：' + body.error.message);
  return body;
}

function norm_(s) { return String(s || '').replace(/[\s\u2800]+/g, ''); }

function adText_(c) {
  c = c || {};
  const s = c.object_story_spec || {}, ld = s.link_data || {}, vd = s.video_data || {}, af = c.asset_feed_spec || {};
  const first = function (arr) { return arr && arr[0] && arr[0].text || ''; };
  return {
    body: c.body || ld.message || vd.message || first(af.bodies),
    title: c.title || ld.name || vd.title || first(af.titles)
  };
}

function regOf_(list) {
  let n = 0;
  (list || []).forEach(function (a) { if (/complete_registration/.test(a.action_type)) n = Math.max(n, Number(a.value)); });
  return n;
}

function fetchMetaAds_(act) {
  const fields = 'name,effective_status,creative{body,title,object_story_spec{link_data{message,name},video_data{message,title}},asset_feed_spec{bodies,titles}},insights.date_preset(maximum){spend,impressions,clicks,actions}';
  let page = metaGet_(act + '/ads', { fields: fields, limit: 100 });
  const ads = [];
  for (let i = 0; i < 10; i++) {
    (page.data || []).forEach(function (ad) {
      const ins = ad.insights && ad.insights.data && ad.insights.data[0] || {};
      const t = adText_(ad.creative);
      const spend = Number(ins.spend || 0), reg = regOf_(ins.actions);
      ads.push({ name: ad.name, status: ad.effective_status, body: t.body, title: t.title, spend: spend, clicks: Number(ins.clicks || 0), reg: reg, cpa: reg ? spend / reg : null });
    });
    if (!page.paging || !page.paging.next) break;
    page = metaGet_(page.paging.next);
  }
  return ads;
}

// 只讀：列出創創每則廣告的成效，並看能對到參考資料的幾則
function inspectCcMetaPerf() {
  const ads = fetchMetaAds_(CC_META_ACT);
  const cut = function (x) { x = String(x).replace(/[\s\u2800]+/g, ' ').trim(); return x.length > 30 ? x.slice(0, 30) + '…' : x; };
  Logger.log('廣告數：' + ads.length + '｜有花費：' + ads.filter(function (a) { return a.spend > 0; }).length + '｜有文案：' + ads.filter(function (a) { return a.body; }).length);
  ads.filter(function (a) { return a.spend > 0; }).sort(function (a, b) { return b.reg - a.reg; }).slice(0, 15).forEach(function (a) {
    Logger.log(a.name + '｜' + a.status + '｜花費 ' + Math.round(a.spend) + '｜註冊 ' + a.reg + '｜成本 ' + (a.cpa ? Math.round(a.cpa) : '—') + '｜' + cut(a.body));
  });

  const ref = SpreadsheetApp.openById(CC_WORK_ID).getSheetByName('參考資料');
  const rows = ref.getLastRow() > 1 ? ref.getRange(2, 4, ref.getLastRow() - 1, 1).getValues() : [];
  let hit = 0;
  rows.forEach(function (r) {
    const m = String(r[0]).split('【主要文字】')[1] || '';
    const main = norm_(m.split('\n【')[0]).slice(0, 40);
    if (main && ads.some(function (a) { return norm_(a.body).slice(0, 40) === main; })) hit++;
  });
  Logger.log('參考資料 ' + rows.length + ' 則，能對到廣告成效的：' + hit + ' 則');
}


// 從客戶主檔查 Meta 帳號 ID
function metaActOf_(name) {
  const v = SpreadsheetApp.openById(SETTINGS_ID).getSheetByName('客戶主檔').getDataRange().getDisplayValues();
  const hi = v.findIndex(function (r) { return r[0] === '報表單位'; });
  const iM = v[hi].indexOf('Meta 廣告帳號 ID');
  const row = v.slice(hi + 1).find(function (r) { return r[0] === name; });
  return row && iM >= 0 ? row[iM] : '';
}

// 只讀：用客戶名稱查帳號，列出成效前 15 則與文案抓取情況
function inspectMetaPerfByClient_(name) {
  const act = metaActOf_(name);
  if (!/^act_\d+$/.test(act)) { Logger.log('客戶主檔找不到「' + name + '」的 Meta 帳號 ID，先停'); return; }
  Logger.log(name + '｜' + act);
  const ads = fetchMetaAds_(act);
  const cut = function (x) { x = String(x).replace(/[\s\u2800]+/g, ' ').trim(); return x.length > 30 ? x.slice(0, 30) + '…' : x; };
  const spent = ads.filter(function (a) { return a.spend > 0; });
  Logger.log('廣告數：' + ads.length + '｜有花費：' + spent.length + '｜有文案：' + ads.filter(function (a) { return a.body; }).length + '｜有註冊數：' + ads.filter(function (a) { return a.reg > 0; }).length);
  spent.sort(function (a, b) { return b.spend - a.spend; }).slice(0, 15).forEach(function (a) {
    Logger.log(a.name + '｜' + a.status + '｜花費 ' + Math.round(a.spend) + '｜點擊 ' + a.clicks + '｜註冊 ' + a.reg + '｜' + (cut(a.body) || '（沒抓到文案）'));
  });
}

function inspectDingdongMetaPerf() { inspectMetaPerfByClient_('鼎東策略'); }
