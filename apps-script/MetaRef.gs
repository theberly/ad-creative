// 從 Meta 廣告直接匯入參考資料：文案＋成效，依活動目標分組比成本，前 1/4 標「成效好」
// 銷售／名單會自動挑這個客戶實際有數字的成果（購買 > 完成註冊 > 名單 > 瀏覽內容）
// 其他目標（觸及、互動、影片觀看…）也匯入，但只標「僅供參考」、不排名
// 廣告組合設定的轉換事件 → 用來比的成果
const REF_EVENT = { PURCHASE: 'purchase', COMPLETE_REGISTRATION: 'complete_registration', LEAD: 'lead', CONTENT_VIEW: 'view_content', ADD_TO_CART: 'add_to_cart', INITIATED_CHECKOUT: 'initiate_checkout', CONTACT: 'contact', SUBMIT_APPLICATION: 'submit_application', SUBSCRIBE: 'subscribe' };
const REF_METRIC_LABEL = { purchase: '購買', complete_registration: '完成註冊', lead: '名單', view_content: '瀏覽內容', add_to_cart: '加入購物車', initiate_checkout: '開始結帳', contact: '聯絡', submit_application: '送出申請', subscribe: '訂閱', landing_page_view: '頁面瀏覽' };
const REF_GOALS = {
  OUTCOME_TRAFFIC: { label: '流量', metrics: [['landing_page_view', '頁面瀏覽']] },
  OUTCOME_SALES: { label: '銷售', metrics: [['purchase', '購買'], ['complete_registration', '完成註冊'], ['lead', '名單'], ['view_content', '瀏覽內容']] },
  OUTCOME_LEADS: { label: '名單', metrics: [['lead', '名單'], ['complete_registration', '完成註冊']] }
};
const REF_MIN_SPEND = 500;

function importDingdongMetaRefPreview() { importMetaPerfRef_('鼎東策略', /香積麵|悅善廚/, false); }
function importDingdongMetaRefApply() { importMetaPerfRef_('鼎東策略', /香積麵|悅善廚/, true); }

function workSheetIdOf_(name) {
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const i = v[0].indexOf('工作 Sheet');
  const row = v.find(function (r, k) { return k > 0 && r[0] === name; });
  const m = row && String(row[i]).match(/\/d\/([^/]+)/);
  return m ? m[1] : '';
}

// action_type 正規化：同一成果的像素、omni、站內版本取最大值，不重複加總
function normActs_(list) {
  const o = {};
  (list || []).forEach(function (a) {
    let k = String(a.action_type).replace(/^offsite_conversion\.fb_pixel_/, '').replace(/^omni_/, '').replace(/^onsite_conversion\./, '');
    if (/^lead(gen)?_grouped$/.test(k)) k = 'lead';
    o[k] = Math.max(o[k] || 0, Number(a.value) || 0);
  });
  return o;
}

function importMetaPerfRef_(name, exclude, apply) {
  const tag = apply ? '' : '[預覽] ';
  const act = metaActOf_(name);
  const wid = workSheetIdOf_(name);
  if (!/^act_\d+$/.test(act)) { Logger.log('找不到「' + name + '」的 Meta 帳號 ID，先停'); return; }
  if (!wid) { Logger.log('客戶清單裡「' + name + '」還沒有工作 Sheet，先建工作 Sheet'); return; }
  const ref = SpreadsheetApp.openById(wid).getSheetByName('參考資料');

  const fields = 'name,campaign{name,objective},creative{body,title,object_story_spec{link_data{message,name},video_data{message,title}},asset_feed_spec{bodies,titles}},insights.date_preset(maximum){spend,actions},adset{promoted_object}';
  let page = metaGet_(act + '/ads', { fields: fields, limit: 100 });
  const merged = {};
  let excluded = 0;
  for (let i = 0; i < 30; i++) {
    (page.data || []).forEach(function (ad) {
      const camp = ad.campaign || {};
      if (exclude && (exclude.test(ad.name) || exclude.test(camp.name || ''))) { excluded++; return; }
      const goal = REF_GOALS[camp.objective] ? camp.objective : 'OTHER';
      const ins = ad.insights && ad.insights.data && ad.insights.data[0] || {};
      const na = normActs_(ins.actions);
      const t = adText_(ad.creative);
      const po = ad.adset && ad.adset.promoted_object || {};
      const ev = goal === 'OTHER' ? '' : (REF_EVENT[po.custom_event_type] || '');
      const key = goal + '｜' + ev + '｜' + ad.name;
      const m = merged[key] || (merged[key] = { key: key, name: ad.name, goal: goal, ev: ev, obj: camp.objective || '', spend: 0, acts: {}, body: '', title: '', camps: {} });
      m.spend += Number(ins.spend || 0);
      Object.keys(na).forEach(function (k) { m.acts[k] = (m.acts[k] || 0) + na[k]; });
      if (camp.name) m.camps[camp.name] = 1;
      if (!m.body && t.body) { m.body = t.body; m.title = t.title || ''; }
    });
    if (!page.paging || !page.paging.next) break;
    page = metaGet_(page.paging.next);
  }

  const list = Object.keys(merged).map(function (k) { return merged[k]; });
  const noText = list.filter(function (m) { return !m.body; }).length;
  const good = {};
  const groups = {};
  list.filter(function (m) { return m.goal !== 'OTHER'; }).forEach(function (m) { (groups[m.goal + '｜' + m.ev] || (groups[m.goal + '｜' + m.ev] = [])).push(m); });
  Object.keys(groups).forEach(function (gk) {
    const inG = groups[gk], goal = inG[0].goal, ev = inG[0].ev, g = REF_GOALS[goal];
    let mt;
    if (ev) mt = [ev, REF_METRIC_LABEL[ev] || ev];
    else mt = g.metrics.find(function (x) { return inG.some(function (m) { return (m.acts[x[0]] || 0) > 0; }); }) || g.metrics[0];
    inG.forEach(function (m) { m.mt = mt; m.res = m.acts[mt[0]] || 0; });
    const ok = inG.filter(function (m) { return m.body && m.spend >= REF_MIN_SPEND && m.res > 0; })
      .sort(function (a, b) { return a.spend / a.res - b.spend / b.res; });
    const top = Math.ceil(ok.length / 4);
    ok.slice(0, top).forEach(function (m) { good[m.key] = 1; });
    Logger.log(tag + '【' + g.label + '｜看' + mt[1] + (ev ? '' : '（推測）') + '】合併後 ' + inG.length + ' 則｜可比較 ' + ok.length + '｜成效好 ' + top);
  });
  const others = list.filter(function (m) { return m.goal === 'OTHER'; });
  if (others.length) Logger.log(tag + '【其他目標】' + others.length + ' 則（只匯入，不排名）');

  const seen = {};
  if (ref.getLastRow() > 1) ref.getRange(2, 4, ref.getLastRow() - 1, 1).getValues().forEach(function (r) { seen[String(r[0])] = 1; });
  const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const fmt = function (x) { return Math.round(x).toLocaleString('en-US'); };
  const out = []; let dup = 0;
  list.filter(function (m) { return m.body; }).sort(function (a, b) {
    const ga = good[a.key] ? 0 : 1, gb = good[b.key] ? 0 : 1;
    return ga - gb || b.spend - a.spend;
  }).forEach(function (m) {
    const content = '【主要文字】\n' + String(m.body).trim() + (m.title ? '\n【標題】' + String(m.title).trim() : '');
    if (seen[content]) { dup++; return; }
    seen[content] = 1;
    let note;
    if (m.goal === 'OTHER') note = '其他目標（' + String(m.obj).replace('OUTCOME_', '') + '）｜花費 ' + fmt(m.spend);
    else {
      const g = REF_GOALS[m.goal], mt = m.mt;
      note = g.label + '｜花費 ' + fmt(m.spend) + '｜' + mt[1] + ' ' + fmt(m.res) + (m.res ? '｜每次' + String(mt[1]).replace(/（.*?）/g, '') + ' ' + fmt(m.spend / m.res) + ' 元' : '');
    }
    out.push([good[m.key] ? '成效好' : '僅供參考', 'Meta', 'Meta 廣告', content, '廣告名稱：' + m.name, note, today, '程式匯入（Meta 成效）', Object.keys(m.camps).join('、')]);
  });

  Logger.log(tag + '排除品牌 ' + excluded + ' 則｜沒抓到文案 ' + noText + ' 則｜重複略過 ' + dup + ' 則');
  Logger.log(tag + '可寫入 ' + out.length + ' 則（成效好 ' + out.filter(function (r) { return r[0] === '成效好'; }).length + '）');
  out.filter(function (r) { return r[0] === '成效好'; }).slice(0, 5).forEach(function (r) { Logger.log(tag + '成效好｜' + r[4] + '｜' + r[5]); });
  if (!apply || !out.length) return;
  ref.getRange(ref.getLastRow() + 1, 1, out.length, 9).setValues(out);
  const n = ref.getLastRow() - 1;
  ref.getRange(2, 1, n, 9).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('top');
  ref.setRowHeightsForced(2, n, 21);
  ref.setColumnWidth(4, 480);
  Logger.log('已寫入 ' + out.length + ' 則到參考資料');
}

// 已匯入的參考資料補填「活動」：用來源欄的廣告名稱，查 Meta 上它屬於哪個行銷活動
function fillDingdongRefCampaignPreview() { fillRefCampaign_('鼎東策略', /香積麵|悅善廚/, false); }
function fillDingdongRefCampaignApply() { fillRefCampaign_('鼎東策略', /香積麵|悅善廚/, true); }

function fillRefCampaign_(name, exclude, apply) {
  const tag = apply ? '' : '[預覽] ';
  const act = metaActOf_(name), wid = workSheetIdOf_(name);
  if (!/^act_\d+$/.test(act) || !wid) { Logger.log('找不到「' + name + '」的 Meta 帳號或工作 Sheet，先停'); return; }
  const ref = SpreadsheetApp.openById(wid).getSheetByName('參考資料');
  const h = ref.getRange(1, 1, 1, ref.getLastColumn()).getValues()[0];
  const iA = h.indexOf('活動'), iS = h.indexOf('來源／連結');
  if (iA < 0) { Logger.log('參考資料還沒有「活動」欄，先跑 WorkSheet.gs 的 upgradeWorkSheetsApply'); return; }

  const map = {};
  let page = metaGet_(act + '/ads', { fields: 'name,campaign{name}', limit: 200 });
  for (let i = 0; i < 20; i++) {
    (page.data || []).forEach(function (ad) {
      const c = ad.campaign && ad.campaign.name || '';
      if (!c || (exclude && (exclude.test(ad.name) || exclude.test(c)))) return;
      (map[ad.name] || (map[ad.name] = {}))[c] = 1;
    });
    if (!page.paging || !page.paging.next) break;
    page = metaGet_(page.paging.next);
  }

  const n = ref.getLastRow() - 1;
  if (n < 1) { Logger.log('參考資料沒有資料'); return; }
  const rows = ref.getRange(2, 1, n, h.length).getDisplayValues();
  let k = 0, miss = 0;
  rows.forEach(function (r, j) {
    if (r[iA]) return;
    const mm = String(r[iS]).match(/^廣告名稱：(.+)$/);
    if (!mm) return;
    const c = map[mm[1]];
    if (!c) { miss++; return; }
    const val = Object.keys(c).join('、');
    k++;
    if (k <= 5) Logger.log(tag + mm[1] + ' → ' + val);
    if (apply) ref.getRange(j + 2, iA + 1).setValue(val);
  });
  Logger.log(tag + '補填活動 ' + k + ' 列｜找不到對應 ' + miss + ' 列');
}


// 各客戶要排除的品牌（廣告或活動名稱符合就不匯入）
const REF_EXCLUDE = { '鼎東策略': /香積麵|悅善廚/ };

// 一次處理全部：有工作 Sheet、有 Meta 帳號、參考資料還是空的客戶
function importAllMetaRefPreview() { importAllMetaRef_(false); }
function importAllMetaRefApply() { importAllMetaRef_(true); }

function importAllMetaRef_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  const t0 = Date.now();
  let stop = false;
  v.slice(1).forEach(function (r) {
    const name = r[0];
    if (!name || !r[iW]) return;
    const act = metaActOf_(name);
    if (!/^act_\d+$/.test(act)) { Logger.log('— ' + name + '：沒有 Meta 帳號，略過'); return; }
    const wid = workSheetIdOf_(name);
    const ref = SpreadsheetApp.openById(wid).getSheetByName('參考資料');
    if (ref.getLastRow() > 1) { Logger.log('— ' + name + '：參考資料已有 ' + (ref.getLastRow() - 1) + ' 列，略過'); return; }
    if (stop) return;
    if (Date.now() - t0 > 270000) { Logger.log('⏸ 時間快到了，' + name + ' 之後的客戶請再跑一次'); stop = true; return; }
    Logger.log('=== ' + name + '｜' + act);
    try {
      if (!apply) {
        let page = metaGet_(act + '/campaigns', { fields: 'name,objective,insights.date_preset(maximum){spend}', limit: 100 });
        const camps = [];
        for (let i = 0; i < 5; i++) {
          (page.data || []).forEach(function (c) { const ins = c.insights && c.insights.data && c.insights.data[0] || {}; camps.push({ n: c.name, o: c.objective, s: Number(ins.spend || 0) }); });
          if (!page.paging || !page.paging.next) break;
          page = metaGet_(page.paging.next);
        }
        camps.sort(function (a, b) { return b.s - a.s; });
        Logger.log(tag + '活動 ' + camps.length + ' 個（花費前 8）：' + camps.slice(0, 8).map(function (c) { return c.n + '〔' + c.o.replace('OUTCOME_', '') + '〕'; }).join('｜'));
      }
      importMetaPerfRef_(name, REF_EXCLUDE[name] || null, apply);
    } catch (e) {
      Logger.log('✗ ' + name + '：' + e.message);
    }
  });
  Logger.log('全部處理完');
}
