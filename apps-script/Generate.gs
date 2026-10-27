
// 產生文案：讀工作 Sheet 的需求 → 組提示（設定＋參考資料）→ AI → 檢查 → 寫進產出（一組一列）
const GEN_FORMATS = {
  'Meta 廣告': {
    fields: [['文案', null], ['標題', 40], ['說明', 40]],
    cta: ['瞭解詳情', '立即報名', '立即申請', '註冊', '傳送訊息', '立即購買', '索取優惠']
  },
  // 回應式搜尋廣告：一組＝15 則標題＋4 則說明，每則一行寫在同一格
  'Google RSA': {
    multi: [['標題', 15, 30], ['說明', 4, 90]]
  }
};
// Google 廣告政策常見退件：標題不能有驚嘆號、不能用表情符號、不能重複標點
function googleWarn_(o) {
  const w = [];
  if (/[!！]/.test(o['標題'])) w.push('Google：標題不可有驚嘆號');
  if (/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/u.test(o['標題'] + o['說明'])) w.push('Google：不可用表情符號');
  if (/([!?！？。，,.])\1/.test(o['標題'] + o['說明'])) w.push('Google：標點不可重複');
  const hs = o['標題'].split('\n');
  if (hs.some(function (x, i) { return x && hs.indexOf(x) !== i; })) w.push('Google：標題有重複');
  return w;
}
const GEN_RISKY = ['第一', '100%', '百分百', '保證', '最強', '最好', '最便宜', '最有效', '唯一', '根治', '零風險', '永久有效'];
const GEN_FACT_WORDS = ['補助', '審核', '資格', '證書', '名額', '限額', '免費', '折扣', '優惠', '截止', '倒數'];
const GEN_MAX_REFS = 8;

function generateDingdongPreview() { generate_('鼎東策略', false); }
function generateDingdongApply() { generate_('鼎東策略', true); }

function halfLen_(s) {
  return Array.from(String(s || '')).reduce(function (n, ch) { return n + (ch.charCodeAt(0) > 0xFF ? 2 : 1); }, 0);
}
function tableOf_(sh) { const v = sh.getDataRange().getDisplayValues(); return { h: v[0], rows: v.slice(1) }; }
function bannedOf_(text) {
  return String(text || '').split(/[（(]/)[0].split(/[、，,\s]+/).map(function (x) { return x.trim(); }).filter(String);
}
function markPending_(text) {
  text = String(text || '');
  const b = SpreadsheetApp.newRichTextValue().setText(text);
  const red = SpreadsheetApp.newTextStyle().setForegroundColor('#c62828').setBold(true).build();
  const re = /【待確認[^】]*】/g;
  let m;
  while ((m = re.exec(text))) b.setTextStyle(m.index, m.index + m[0].length, red);
  return b.build();
}

function generate_(name, apply, onlyRow) {
  const tag = apply ? '' : '[預覽] ';
  const OUT = WS_TABS['產出'];
  const wid = workSheetIdOf_(name);
  if (!wid) { Logger.log('客戶清單裡「' + name + '」沒有工作 Sheet，先停'); return; }
  const ss = SpreadsheetApp.openById(wid);
  const po = ss.getSheetByName('產出');
  const ph = po.getRange(1, 1, 1, po.getLastColumn()).getValues()[0];
  if (ph.join('|') !== OUT.join('|')) { Logger.log('產出分頁還是舊欄位，先跑 WorkSheet.gs 的 upgradeWorkSheetsApply'); return; }

  const set = {};
  ss.getSheetByName('設定').getDataRange().getDisplayValues().slice(1).forEach(function (r) { if (r[0]) set[r[0]] = r[1]; });

  const req = ss.getSheetByName('需求'), rq = tableOf_(req), H = rq.h;
  const col = function (r, k) { return r[H.indexOf(k)] || ''; };
  const todo = [];
  rq.rows.forEach(function (r, i) { if (!col(r, '狀態') && col(r, '格式') && (!onlyRow || i + 2 === onlyRow)) todo.push({ r: r, row: i + 2 }); });
  if (!todo.length) { Logger.log('需求分頁沒有待產的列（狀態是空的才會產）'); return; }

  const ref = tableOf_(ss.getSheetByName('參考資料')), RH = ref.h;
  const iType = 0, iPf = RH.indexOf('平台'), iC = RH.indexOf('內容'), iNote = RH.indexOf('成效備註'), iAct = RH.indexOf('活動');

  const banned = bannedOf_(set['禁用字詞']);
  const hubRisk = SpreadsheetApp.getActive().getSheetByName('審核字庫');
  const extra = hubRisk && hubRisk.getLastRow() > 1 ? hubRisk.getRange(2, 2, hubRisk.getLastRow() - 1, 1).getDisplayValues().map(function (x) { return x[0]; }).filter(String) : [];
  const risky = GEN_RISKY.concat(extra);

  const out = [];
  todo.slice(0, 3).forEach(function (t) {
    const r = t.r, fmt = col(r, '格式'), spec = GEN_FORMATS[fmt];
    if (!spec) { Logger.log('第 ' + t.row + ' 列格式「' + fmt + '」尚未支援，略過'); return; }
    const n = Math.max(1, Math.min(5, Number(col(r, '組數')) || 1));
    const act = col(r, '活動名稱');
    const inputs = Object.keys(set).map(function (k) { return set[k]; }).concat(r).join('\n');

    const pf = col(r, '平台');
    const refs = ref.rows.filter(function (x) { return x[iC]; })
      .map(function (x) { return { x: x, s: (x[iType] === '成效好' ? 3 : 0) + (act && iAct >= 0 && x[iAct].indexOf(act) >= 0 ? 2 : 0) + (x[iType] === '客戶提供' ? 1 : 0) + (!pf || x[iPf] === pf || x[iPf] === '共用' ? 2 : 0) }; })
      .sort(function (a, b) { return b.s - a.s; }).slice(0, GEN_MAX_REFS).map(function (o) { return o.x; });

    const fieldSpec = spec.multi ? spec.multi.map(function (f) { return f[1] + ' 則「' + f[0] + '」（每則上限 ' + f[2] + ' 半形字元，中文算 2，也就是約 ' + Math.floor(f[2] / 2) + ' 個中文字）'; }).join('、') : spec.fields.map(function (f) { return '「' + f[0] + '」' + (f[1] ? '（上限 ' + f[1] + ' 半形字元，中文與 emoji 算 2）' : ''); }).join('、') + '、「CTA」（只能從這些選一個：' + spec.cta.join('、') + '）';
    const prompt = [
      '請為以下需求寫 ' + n + ' 組 ' + fmt + ' 文案，每組包含：' + fieldSpec + '。',
      '',
      '【客戶設定】',
      '客戶：' + name,
      '品牌語氣：' + (set['品牌語氣'] || '（未提供）'),
      '必提賣點：' + (set['必提賣點'] || '（未提供）'),
      '禁用字詞：' + (set['禁用字詞'] || '（無）'),
      '備註：' + (set['備註'] || '（無）'),
      '',
      '【這次的需求】',
      '活動：' + (act || '（未提供）'),
      '受眾：' + (col(r, '受眾') || '（未提供）'),
      '產品／服務：' + (col(r, '產品／服務') || '（未提供）'),
      '活動目標：' + (col(r, '活動目標') || '（未提供）'),
      '補充說明：' + (col(r, '補充說明') || '（無）'),
      '',
      '【參考文案】（標「成效好」的是實際成效較好的，優先學它的寫法；只學語氣和結構，不要照抄，也不要沿用裡面的事實）',
      refs.map(function (x, j) { return '--- 參考 ' + (j + 1) + '｜' + x[iType] + (x[iNote] ? '｜' + x[iNote] : '') + '\n' + String(x[iC]).slice(0, 700); }).join('\n') || '（無）',
      '',
      '【寫作規則】',
      '1. 繁體中文、台灣用語。',
      spec.multi ? '2. 這是 Google 回應式搜尋廣告：Google 會任意挑 3 則標題和 2 則說明組合，所以每則都要能單獨成立、不能接續上一則；標題之間不可重複或只差一兩字；切角要分散（品牌名 2～3 則、賣點、優惠或價格、適合誰、行動呼籲各幾則）；參考文案若是 Meta 的長文，只取它的賣點和用字，不要照長文結構寫。標題不可用驚嘆號，標題與說明都不可用表情符號、不可重複標點。說明要是完整句子，可以帶行動呼籲。'
        : '2. 文案：前三行要抓住注意力；中段用 3 個重點條列；分段之間空一行；結尾一句行動呼籲。',
      '3. 不可出現禁用字詞，也不可用誇大或保證式用語（例如：第一、100%、保證、最有效）。',
      '4. 不要寫網址或「點連結」之類的字。',
      '5. 只能使用【客戶設定】與【這次的需求】裡寫到的事實（價格、日期、名額、資格、補助、審核、證書、課程內容、數字）。參考文案裡的事實不可沿用，只學寫法。',
      '6. 需要但沒提供的事實，寫成【待確認：項目】，例如【待確認：課程內容】，不要自己補。',
      '7. ' + n + ' 組的切角要不同。',
      '',
      spec.multi ? '8. 每則一定要在字數上限內，寧可短。' : '',
      '只回傳 JSON，格式：' + (spec.multi ? '{"sets":[{' + spec.multi.map(function (f) { return '"' + f[0] + '":["…"]'; }).join(',') + '}]}'
        : '{"sets":[{' + spec.fields.map(function (f) { return '"' + f[0] + '":"…"'; }).join(',') + ',"CTA":"…"}]}')
    ].join('\n');

    Logger.log(tag + '第 ' + t.row + ' 列｜' + act + '｜' + fmt + '｜' + n + ' 組｜參考 ' + refs.length + ' 則（成效好 ' + refs.filter(function (x) { return x[iType] === '成效好'; }).length + '）｜提示 ' + prompt.length + ' 字');

    let sets;
    try {
      const txt = callAI_(prompt, { json: true, system: '你是台灣的資深廣告文案，熟悉 Meta 與 Google 廣告。' });
      sets = JSON.parse(txt.replace(/^\s*```(?:json)?|```\s*$/g, '')).sets || [];
    } catch (e) { Logger.log('第 ' + t.row + ' 列產生失敗：' + e.message); return; }

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    const id = name + '-' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd') + '-' + ('0' + t.row).slice(-2);
    sets.slice(0, n).forEach(function (st, si) {
      const o = { '需求編號': id, '產出時間': now, '平台': col(r, '平台'), '格式': fmt, '序號': si + 1, '受眾': col(r, '受眾'), '到達網址': col(r, '到達網址'), '確認狀態': '待確認' };
      const lens = [], over = [];
      if (spec.multi) {
        o['文案'] = '';
        spec.multi.forEach(function (f) {
          const items = (Array.isArray(st[f[0]]) ? st[f[0]] : String(st[f[0]] || '').split('\n')).map(function (x) { return String(x).trim(); }).filter(String);
          items.forEach(function (x, k) { const len = halfLen_(x); if (len > f[2]) over.push(f[0] + (k + 1) + ' ' + len + '/' + f[2]); });
          o[f[0]] = items.join('\n');
          lens.push(f[0] + ' ' + items.length + '/' + f[1] + ' 則，最長 ' + Math.max.apply(null, items.map(halfLen_).concat([0])));
        });
        o['CTA'] = '';
      } else {
        spec.fields.forEach(function (f) {
          o[f[0]] = String(st[f[0]] || '').trim();
          const len = halfLen_(o[f[0]]);
          lens.push(len);
          if (f[1] && len > f[1]) over.push(f[0] + ' ' + len + '/' + f[1]);
        });
        o['CTA'] = spec.cta.indexOf(String(st['CTA'] || '').trim()) >= 0 ? String(st['CTA']).trim() : '瞭解詳情';
      }
      o['字數（文案／標題／說明）'] = lens.join('／');
      o['是否超限'] = over.join('、');
      const all = ['文案', '標題', '說明'].map(function (h) { return o[h] || ''; }).join('\n');
      const nums = (all.match(/\d+(?:\.\d+)?/g) || []).filter(function (x, i, a) { return a.indexOf(x) === i && inputs.indexOf(x) < 0; });
      o['審核警示'] = banned.filter(function (x) { return all.indexOf(x) >= 0; }).map(function (x) { return '禁用：' + x; })
        .concat(risky.filter(function (x) { return all.indexOf(x) >= 0; }).map(function (x) { return '誇大：' + x; }))
        .concat(GEN_FACT_WORDS.filter(function (x) { return all.indexOf(x) >= 0 && inputs.indexOf(x) < 0; }).map(function (x) { return '需確認：' + x; }))
        .concat(nums.map(function (x) { return '需確認數字：' + x; }))
        .concat(all.indexOf('【待確認') >= 0 ? ['有待確認項目'] : [])
        .concat(spec.multi ? googleWarn_(o) : []).join('、');
      out.push(OUT.map(function (h) { return o[h] == null ? '' : o[h]; }));
      Logger.log(tag + '第 ' + (si + 1) + ' 組｜字數 ' + o['字數（文案／標題／說明）'] + (o['CTA'] ? '｜CTA ' + o['CTA'] : '') + (o['是否超限'] ? '｜超限：' + o['是否超限'] : '') + (o['審核警示'] ? '｜' + o['審核警示'] : '') +
        '\n【文案】\n' + o['文案'] + '\n【標題】' + o['標題'] + '\n【說明】' + o['說明']);
    });
    if (apply) {
      req.getRange(t.row, H.indexOf('需求編號') + 1).setValue(id);
      req.getRange(t.row, H.indexOf('建立時間') + 1).setValue(now);
      if (!col(r, '操作者')) req.getRange(t.row, H.indexOf('操作者') + 1).setValue(Session.getActiveUser().getEmail());
      req.getRange(t.row, H.indexOf('狀態') + 1).setValue('已產出（待確認）');
    }
  });

  if (!apply || !out.length) return;
  const start = po.getLastRow() + 1;
  po.getRange(start, 1, out.length, OUT.length).setValues(out).setVerticalAlignment('top');
  ['文案', '標題', '說明'].forEach(function (h) {
    const c = OUT.indexOf(h);
    po.getRange(start, c + 1, out.length, 1).setRichTextValues(out.map(function (row) { return [markPending_(row[c])]; })).setWrap(true);
  });
  const iW = OUT.indexOf('審核警示');
  out.forEach(function (row, k) { if (row[iW]) po.getRange(start + k, iW + 1).setBackground('#fde7e7').setWrap(true); });
  po.setColumnWidth(OUT.indexOf('文案') + 1, 420);
  po.setColumnWidth(OUT.indexOf('標題') + 1, 200);
  po.setColumnWidth(OUT.indexOf('說明') + 1, 200);
  po.setColumnWidth(iW + 1, 200);
  Logger.log('已寫入產出 ' + out.length + ' 組');
}
