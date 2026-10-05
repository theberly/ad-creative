// 網頁應用程式入口：?action=clients&key=通行碼
function doGet(e) {
  const p = (e && e.parameter) || {};
  const key = PropertiesService.getScriptProperties().getProperty('ACCESS_KEY');
  if (!key || p.key !== key) return json_({ ok: false, error: '通行碼錯誤' });
  if (p.action === 'clients') return json_({ ok: true, clients: readClients_() });
  return json_({ ok: false, error: '未知的動作' });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readClients_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('客戶設定');
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0];
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    h.forEach(function (k, i) { if (k) o[k] = r[i]; });
    return o;
  });
}

// 只讀：看網頁會拿到什麼
function testClients() {
  Logger.log(JSON.stringify(readClients_(), null, 2));
}

// 產生通行碼（已有就不覆蓋，只顯示）
function setAccessKey() {
  const props = PropertiesService.getScriptProperties();
  const old = props.getProperty('ACCESS_KEY');
  if (old) { Logger.log('通行碼已存在：' + old); return; }
  const k = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  props.setProperty('ACCESS_KEY', k);
  Logger.log('已建立通行碼：' + k);
}
