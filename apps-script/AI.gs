// AI 呼叫：目前用 Gemini；之後換 Claude 只改這支
// 指令碼屬性：GEMINI_API_KEY（必填）、GEMINI_MODEL（選填，預設 gemini-2.5-flash）
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/';

function geminiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!k) throw new Error('找不到指令碼屬性 GEMINI_API_KEY');
  return k;
}
function geminiModel_() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || 'gemini-2.5-flash';
}

// 送出提示，回傳文字。opts.json=true 時要求模型回傳 JSON
function callAI_(prompt, opts) {
  opts = opts || {};
  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: opts.temperature == null ? 0.9 : opts.temperature }
  };
  if (opts.system) body.systemInstruction = { parts: [{ text: opts.system }] };
  if (opts.json) body.generationConfig.responseMimeType = 'application/json';
  const res = UrlFetchApp.fetch(GEMINI_BASE + 'models/' + geminiModel_() + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': geminiKey_() }, payload: JSON.stringify(body)
  });
  const j = JSON.parse(res.getContentText());
  if (j.error) throw new Error('Gemini：' + j.error.message);
  const c = j.candidates && j.candidates[0];
  const text = c && c.content && c.content.parts ? c.content.parts.map(function (p) { return p.text || ''; }).join('') : '';
  if (!text) throw new Error('Gemini 沒有回傳內容（' + (c && c.finishReason || '未知原因') + '）');
  return text;
}

// 只讀測試：列出可用模型，並用目前模型回一句話
function testGemini() {
  const res = UrlFetchApp.fetch(GEMINI_BASE + 'models?pageSize=100', { muteHttpExceptions: true, headers: { 'x-goog-api-key': geminiKey_() } });
  const j = JSON.parse(res.getContentText());
  if (j.error) { Logger.log('金鑰有問題：' + j.error.message); return; }
  const names = (j.models || []).filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .map(function (m) { return m.name.replace('models/', ''); });
  Logger.log('可用模型（' + names.length + '）：' + names.filter(function (n) { return /flash|pro/.test(n); }).join('、'));
  Logger.log('目前使用：' + geminiModel_() + (names.indexOf(geminiModel_()) >= 0 ? '' : '（不在清單裡，要改 GEMINI_MODEL）'));
  if (names.indexOf(geminiModel_()) < 0) return;
  Logger.log('測試回覆：' + callAI_('用一句繁體中文介紹你自己，20 字以內。', { temperature: 0.2 }));
}
