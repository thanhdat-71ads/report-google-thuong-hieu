/* =====================================================================================
   APPS SCRIPT — REPORT GOOGLE THƯƠNG HIỆU — ĐĂNG NHẬP BẰNG FIREBASE AUTHENTICATION (71 ADS)
   - Dùng chung dự án Firebase "report-chan-doan-tu-khoa" với dashboard Chẩn đoán Từ khoá.
   - Dashboard đăng nhập bằng Firebase, gửi kèm "ID token" khi xin dữ liệu; Apps Script hỏi
     Google xác minh token + kiểm tra email có trong danh sách được phép của RIÊNG dashboard này.
   - Cài đặt ở  Cài đặt dự án → Thuộc tính tập lệnh:
       FIREBASE_API_KEY = AIzaSyCd9OoQ5uW-LfoFMnST5qYBBbVAljqG-Zo  (Web API Key dự án Firebase)
       ALLOWED_EMAILS   = các email được xem Report Thương Hiệu, cách nhau dấu phẩy
   - Mở đường dẫn /exec trực tiếp (GET) KHÔNG thấy dữ liệu.
   ===================================================================================== */
const SHEET_NAMES = ['FILE THÔ', 'GDN 1', 'GDN 2'];

const CACHE_KEY = 'brand_payload_v1';
const CACHE_TTL_SECONDS = 600;     // 10 phút; nút "Cập nhật ngay" luôn đọc mới từ Sheet
const CHUNK_SIZE = 30000;
const VERIFY_CACHE_SECONDS = 300;

// ---------- GET: không trả dữ liệu ----------
function doGet(e) {
  return json_({ success: false, error: 'AUTH_REQUIRED', message: 'Cần đăng nhập từ Dashboard để xem dữ liệu.' });
}

// ---------- POST: {action:'data', idToken, nocache} ----------
function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ success: false, error: 'BAD_REQUEST' }); }
  if (String(req.action || '') !== 'data') return json_({ success: false, error: 'BAD_REQUEST' });

  const cache = CacheService.getScriptCache();
  const v = verifyFirebaseUser_(cache, req.idToken);
  if (!v.ok) return json_({ success: false, error: v.error });

  let text = req.nocache ? null : readChunkedCache_(cache);
  if (!text) {
    text = buildPayloadJson_();
    writeChunkedCache_(cache, text);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

// Xác minh Firebase ID token qua Google Identity Toolkit + kiểm tra danh sách email được phép
function verifyFirebaseUser_(cache, idToken) {
  if (!idToken) return { ok: false, error: 'AUTH_REQUIRED' };
  const key = 'fb_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(idToken), Utilities.Charset.UTF_8));
  const hit = readJson_(cache, key);
  if (hit) return hit;

  const props = PropertiesService.getScriptProperties();
  const apiKey = String(props.getProperty('FIREBASE_API_KEY') || '').trim();
  const allowed = String(props.getProperty('ALLOWED_EMAILS') || '')
    .split(/[\s,;]+/).map(function (s) { return s.trim().toLowerCase(); }).filter(String);
  if (!apiKey || !allowed.length) return { ok: false, error: 'NOT_CONFIGURED' };

  let user = null;
  try {
    const res = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(apiKey),
      { method: 'post', contentType: 'application/json', payload: JSON.stringify({ idToken: String(idToken) }), muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, error: 'INVALID_TOKEN' };
    const body = JSON.parse(res.getContentText());
    user = body && body.users && body.users[0];
  } catch (err) {
    return { ok: false, error: 'VERIFY_FAILED' };
  }
  if (!user || !user.email) return { ok: false, error: 'INVALID_TOKEN' };

  let result;
  const email = String(user.email).toLowerCase();
  if (user.disabled) result = { ok: false, error: 'USER_DISABLED' };
  else if (allowed.indexOf(email) < 0) result = { ok: false, error: 'NOT_ALLOWED' };
  else if (!user.emailVerified) result = { ok: false, error: 'EMAIL_NOT_VERIFIED' };
  else result = { ok: true, email: email };
  if (result.ok) safeCachePut_(cache, key, JSON.stringify(result), VERIFY_CACHE_SECONDS);
  return result;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// =================== ĐỌC 3 SHEET (giữ logic cũ, gửi dạng gọn: tên cột 1 lần + mảng giá trị) ===================
function buildPayloadJson_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {};
  SHEET_NAMES.forEach(function (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheets[sheetName] = { cols: [], data: [] }; return; }
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) { sheets[sheetName] = { cols: [], data: [] }; return; }
    const headers = values.shift().map(function (h) { return String(h).trim(); });
    const keepIdx = [];
    headers.forEach(function (h, i) { if (h) keepIdx.push(i); });
    sheets[sheetName] = {
      cols: keepIdx.map(function (i) { return headers[i]; }),
      data: values
        .filter(function (row) { return row.some(function (cell) { return cell !== '' && cell !== null; }); })
        .map(function (row) { return keepIdx.map(function (i) { return cellValue_(row[i]); }); })
    };
  });
  return JSON.stringify({
    success: true,
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    sheets: sheets
  });
}

function cellValue_(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
}

// =================== CACHE CHIA MẢNH + HELPERS AN TOÀN ===================
function readChunkedCache_(cache) {
  try {
    const n = Number(cache.get(CACHE_KEY + '_n') || 0);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    const got = cache.getAll(keys);
    let text = '';
    for (let i = 0; i < keys.length; i++) {
      if (got[keys[i]] === undefined || got[keys[i]] === null) return null;  // thiếu mảnh → đọc lại Sheet
      text += got[keys[i]];
    }
    return text;
  } catch (err) { return null; }
}
function writeChunkedCache_(cache, text) {
  try {
    const obj = {};
    let n = 0;
    for (let i = 0; i < text.length; i += CHUNK_SIZE) { obj[CACHE_KEY + '_' + n] = text.slice(i, i + CHUNK_SIZE); n++; }
    obj[CACHE_KEY + '_n'] = String(n);
    cache.putAll(obj, CACHE_TTL_SECONDS);
  } catch (err) { /* cache lỗi thì bỏ qua, vẫn trả dữ liệu bình thường */ }
}
function readJson_(cache, key) {
  try { const v = cache.get(key); return v ? JSON.parse(v) : null; } catch (err) { return null; }
}
function safeCacheGet_(cache, key) { try { return cache.get(key); } catch (err) { return null; } }
function safeCachePut_(cache, key, text, ttl) { try { cache.put(key, text, ttl); } catch (err) {} }
function safeCacheRemove_(cache, key) { try { cache.remove(key); } catch (err) {} }
