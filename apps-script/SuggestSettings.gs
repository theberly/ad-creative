// 用各家「參考資料」的既有廣告文案，請 AI 歸納設定建議（品牌語氣、必提賣點、備註）
// 主模型這時段常卡好幾分鐘，這裡直接用備用（較快）模型
// 只寫在設定分頁的 C 欄「AI 建議」，不動 B 欄內容；C 欄已有建議就略過
const SUGGEST_KEYS = ['品牌語氣', '必提賣點', '備註'];

function suggestSettingsPreview() { suggestSettings_(false, 'ebs'); }
function suggestSettingsApply() { suggestSettings_(true); }

function suggestSettings_(apply, only) {
  const tag = apply ? '' : '[預覽] ';
  const t0 = Date.now();
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  let done = 0, slow = 0;
  const cache = CacheService.getScriptCache();
  for (let k = 1; k < v.length; k++) {
    const name = v[k][0];
    if (!name || !v[k][iW] || (only && name !== only)) continue;
    if (Date.now() - t0 + slow * 1.2 > 330000) { Logger.log('時間快到，先停在「' + name + '」之前，再執行一次會接著做'); break; }
    const ss = SpreadsheetApp.openById(workSheetIdOf_(name));
    const st = ss.getSheetByName('設定');
    const sv = st.getDataRange().getDisplayValues();
    if (sv.some(function (r) { return String(r[2] || '').trim() && r[0] !== '欄位'; })) { Logger.log('— ' + name + '：已有 AI 建議，略過'); continue; }

    const rv = ss.getSheetByName('參考資料').getDataRange().getDisplayValues().slice(1).filter(function (r) { return r[3]; });
    if (!rv.length) { Logger.log('— ' + name + '：參考資料是空的，略過'); continue; }
    const rank = { '成效好': 0, '客戶提供': 1, '僅供參考': 2 };
    rv.sort(function (a, b) { return (rank[a[0]] == null ? 3 : rank[a[0]]) - (rank[b[0]] == null ? 3 : rank[b[0]]); });
    const ads = rv.slice(0, 30).map(function (r, i) { return (i + 1) + '.【' + r[0] + '】' + String(r[3]).slice(0, 400); }).join('\n\n');

    const prompt = [
      '以下是客戶「' + name + '」過去實際投放的廣告文案（【成效好】是成效最好的一批）。',
      '請只根據這些文案歸納，不要編造文案裡沒有的事實，給負責人當填寫設定的草稿：',
      '1. 品牌語氣：一兩句描述語氣與用字風格（例如口吻、稱呼客人的方式、是否用表情符號）。',
      '2. 必提賣點：條列 3～6 點反覆出現、或成效好文案共同強調的賣點。',
      '3. 備註：品牌名／產品名的固定寫法、常用的優惠或活動說法、常見 CTA。看不出來的就不要寫。',
      '用繁體中文。只回傳 JSON：{"品牌語氣":"…","必提賣點":"• …\\n• …","備註":"…"}',
      '', ads
    ].join('\n');

    let out;
    const t1 = Date.now();
    const hit = cache.get('suggest|' + name);
    if (hit) { out = JSON.parse(hit); Logger.log('（' + name + ' 沿用剛才預覽的結果）'); }
    else try {
      const txt = callAI_(prompt, { json: true, temperature: 0.3, model: PropertiesService.getScriptProperties().getProperty('GEMINI_FALLBACK_MODEL') || 'gemini-3.5-flash-lite', system: '你是台灣的資深廣告文案，擅長從既有廣告歸納品牌調性。' });
      out = JSON.parse(txt.replace(/^\s*```(?:json)?|```\s*$/g, ''));
      cache.put('suggest|' + name, JSON.stringify(out), 21600);
    } catch (e) { Logger.log('✗ ' + name + '：' + e.message); continue; }
    slow = Math.max(slow, Date.now() - t1);

    done++;
    Logger.log(tag + '✓ ' + name + '（參考 ' + Math.min(rv.length, 30) + ' 則，AI 花 ' + Math.round((Date.now() - t1) / 1000) + ' 秒）\n' + SUGGEST_KEYS.map(function (x) { return '【' + x + '】' + (out[x] || '（無）'); }).join('\n'));
    if (!apply) continue;
    st.getRange(1, 3).setValue('AI 建議（從過去文案歸納，確認後請搬到 B 欄）').setFontWeight('bold');
    SUGGEST_KEYS.forEach(function (x) {
      const ri = sv.findIndex(function (r) { return r[0] === x; });
      if (ri > 0 && out[x]) st.getRange(ri + 1, 3).setValue(out[x]).setWrap(true).setFontColor('#666666');
    });
    st.setColumnWidth(3, 420);
  }
  Logger.log(tag + (apply ? '已寫入 ' : '預覽 ') + done + ' 家');
}


// 把 C 欄 AI 建議搬進 B 欄（B 欄已有內容的不動），搬過的 C 欄清掉
function moveSuggestToBPreview() { moveSuggestToB_(false); }
function moveSuggestToBApply() { moveSuggestToB_(true); }

function moveSuggestToB_(apply) {
  const tag = apply ? '' : '[預覽] ';
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  v.slice(1).forEach(function (r) {
    const name = r[0];
    if (!name || !r[iW]) return;
    const st = SpreadsheetApp.openById(workSheetIdOf_(name)).getSheetByName('設定');
    const sv = st.getDataRange().getDisplayValues();
    const moved = [], kept = [];
    sv.forEach(function (x, i) {
      if (i === 0 || !String(x[2] || '').trim()) return;
      if (String(x[1]).trim()) { kept.push(x[0]); return; }
      moved.push(x[0]);
      if (apply) { st.getRange(i + 1, 2).setValue(x[2]).setWrap(true); st.getRange(i + 1, 3).clearContent(); }
    });
    if (!moved.length && !kept.length) return;
    if (apply && !kept.length) st.getRange(1, 3).clearContent();
    Logger.log(tag + name + '：搬入 ' + (moved.join('、') || '（無）') + (kept.length ? '｜B 欄已有內容、保留 C 欄建議：' + kept.join('、') : ''));
  });
}

// 舊的成效備註「每次 3 元」改成「每次頁面瀏覽 3 元」
function fixRefCostLabelPreview() { fixRefCostLabel_(false); }
function fixRefCostLabelApply() { fixRefCostLabel_(true); }

function fixRefCostLabel_(apply) {
  const v = SpreadsheetApp.getActive().getSheetByName('客戶清單').getDataRange().getDisplayValues();
  const iW = v[0].indexOf('工作 Sheet');
  let total = 0, sample = '';
  v.slice(1).forEach(function (r) {
    if (!r[0] || !r[iW]) return;
    const sh = SpreadsheetApp.openById(workSheetIdOf_(r[0])).getSheetByName('參考資料');
    const last = sh.getLastRow();
    if (last < 2) return;
    const rg = sh.getRange(2, 6, last - 1, 1);
    let n = 0;
    const vals = rg.getValues().map(function (x) {
      const s = String(x[0]);
      const t = s.replace(/｜([^｜]+?) ([\d,.]+)｜每次 ([\d,.]+) 元/, function (all, label, num, cost) {
        return '｜' + label + ' ' + num + '｜每次' + label.replace(/（.*?）/g, '') + ' ' + cost + ' 元';
      });
      if (t !== s) { n++; if (!sample) sample = s + '  →  ' + t; }
      return [t];
    });
    total += n;
    if (n && apply) rg.setValues(vals);
    if (n) Logger.log((apply ? '' : '[預覽] ') + r[0] + '：' + n + ' 列');
  });
  Logger.log((apply ? '已改 ' : '[預覽] 會改 ') + total + ' 列｜例：' + sample);
}
