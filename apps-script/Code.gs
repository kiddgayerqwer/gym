/**
 * 教練學員管理 — Google Apps Script 後端
 *
 * 用途：把學員、上課紀錄、簽到、合約的資料存在「這份 Google 試算表」，
 *       合約檔案存在你的 Google 雲端硬碟「學員合約」資料夾。
 *
 * 部署方式（只要做一次）：
 *   1. 在 Google 雲端硬碟新增一份試算表，命名「學員管理」。
 *   2. 開啟它 → 上方選單「擴充功能」→「Apps Script」。
 *   3. 把這整段程式貼進去（取代原本內容），並把下面的 TOKEN 改成你自己的一長串英數字。
 *   4. 上方函式選單選 setup，按「執行」，依畫面授權（會要求存取試算表與雲端硬碟）。
 *   5. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」：
 *        執行身分：我；誰可以存取：所有人 → 按「部署」。
 *   6. 複製「網頁應用程式」網址（結尾是 /exec），到 APP 右上「連線」填入網址與金鑰。
 *
 * 之後如果修改這段程式，要「部署 → 管理部署作業 → 編輯 → 新版本」才會生效。
 * 金鑰請不要外流；知道網址又知道金鑰的人才能讀寫資料。
 */

const TOKEN = '請改成你自己的金鑰（一長串英數字，不要外流）';
const FOLDER_NAME = '學員合約';
const TZ = 'Asia/Taipei';
const VERSION = 1;

// 每張工作表的欄位：[程式用的欄位名, 試算表上顯示的標題]。順序固定，請勿隨意調整。
const SCHEMA = {
  members: {
    sheet: '學員',
    cols: [['id', 'ID'], ['name', '姓名'], ['phone', '電話'], ['gender', '性別'], ['birthday', '生日'],
      ['emergency', '緊急聯絡人'], ['emergencyPhone', '緊急聯絡電話'], ['coach', '負責教練'], ['joinDate', '入會日期'],
      ['plan', '課程方案'], ['total', '購買堂數'], ['goal', '訓練目標'], ['health', '傷病／健康注意事項'],
      ['note', '其他備註'], ['createdAt', '建立時間(ms)']]
  },
  classes: {
    sheet: '上課紀錄',
    cols: [['id', 'ID'], ['memberId', '學員ID'], ['date', '上課日期'], ['time', '上課時間'], ['coach', '授課教練'],
      ['content', '課程內容'], ['note', '學員狀況／下次建議'], ['createdAt', '建立時間(ms)']]
  },
  checkins: {
    sheet: '簽到',
    cols: [['id', 'ID'], ['memberId', '學員ID'], ['date', '日期'], ['time', '時間'], ['ts', '時間戳(ms)'], ['coach', '教練']]
  },
  contracts: {
    sheet: '合約',
    cols: [['id', 'ID'], ['memberId', '學員ID'], ['title', '合約名稱'], ['start', '起日'], ['end', '到期日'],
      ['note', '備註'], ['fileName', '檔名'], ['type', '檔案類型'], ['size', '大小(bytes)'],
      ['fileId', '雲端硬碟檔案ID'], ['url', '檔案連結'], ['createdAt', '建立時間(ms)']]
  }
};
const NUM_KEYS = ['ts', 'createdAt', 'size'];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() {
  const s = SpreadsheetApp.getActiveSpreadsheet();
  if (!s) throw new Error('請從試算表的「擴充功能 → Apps Script」建立這支程式。');
  return s;
}

function sheet_(col) {
  const def = SCHEMA[col];
  if (!def) throw new Error('未知的資料表：' + col);
  const ss = ss_();
  let sh = ss.getSheetByName(def.sheet);
  if (!sh) {
    sh = ss.insertSheet(def.sheet);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols.map(function (c) { return c[1]; })]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function cell_(k, v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (NUM_KEYS.indexOf(k) >= 0 && !isNaN(Number(v))) return Number(v);
  return String(v);
}

function readAll_(col) {
  const def = SCHEMA[col];
  const sh = sheet_(col);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const n = def.cols.length;
  const vals = sh.getRange(2, 1, last - 1, n).getValues();
  const out = [];
  vals.forEach(function (r) {
    if (r[0] === '' || r[0] === null) return;
    const o = {};
    def.cols.forEach(function (c, i) { o[c[0]] = cell_(c[0], r[i]); });
    out.push(o);
  });
  return out;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function saveRow_(col, row) {
  if (!row || !row.id) throw new Error('缺少 id');
  const def = SCHEMA[col];
  const sh = sheet_(col);
  const arr = def.cols.map(function (c) {
    const v = row[c[0]];
    return (v === null || v === undefined) ? '' : (NUM_KEYS.indexOf(c[0]) >= 0 && v !== '' && !isNaN(Number(v)) ? Number(v) : String(v));
  });
  let r = findRow_(sh, row.id);
  if (r < 0) r = Math.max(sh.getLastRow(), 1) + 1;
  // 非數字欄位一律當純文字，避免日期、電話前面的 0 被 Google 自動轉換
  const range = sh.getRange(r, 1, 1, def.cols.length);
  range.setNumberFormat('@');
  range.setValues([arr]);
}

function deleteRow_(col, id) {
  const sh = sheet_(col);
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function handle_(body) {
  if (!body || body.token !== TOKEN) return { ok: false, error: '金鑰不正確' };
  const a = body.action;
  if (a === 'ping') return { ok: true, version: VERSION };
  if (a === 'all') {
    const data = {};
    Object.keys(SCHEMA).forEach(function (c) { data[c] = readAll_(c); });
    return { ok: true, data: data };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (a === 'save') { saveRow_(body.col, body.row); return { ok: true }; }
    if (a === 'delete') { deleteRow_(body.col, body.id); return { ok: true }; }
    if (a === 'upload') {
      const bytes = Utilities.base64Decode(body.b64);
      const blob = Utilities.newBlob(bytes, body.type || 'application/octet-stream', body.name || 'contract');
      const f = folder_().createFile(blob);
      return { ok: true, fileId: f.getId(), url: f.getUrl() };
    }
    if (a === 'deleteFile') {
      if (body.fileId) DriveApp.getFileById(body.fileId).setTrashed(true);
      return { ok: true };
    }
    return { ok: false, error: '未知的操作：' + a };
  } finally {
    lock.releaseLock();
  }
}

/** 所有讀寫都用 POST（金鑰放在內容裡，不會出現在網址）。 */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(handle_(body));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** 直接用瀏覽器打開網址時，只顯示程式有在運作，不會給任何資料。 */
function doGet() {
  return json_({ ok: true, service: '教練學員管理', version: VERSION });
}

/** 第一次部署前手動執行一次，用來建立工作表、資料夾並完成授權。 */
function setup() {
  Object.keys(SCHEMA).forEach(function (c) { sheet_(c); });
  folder_();
  Logger.log('完成：已建立工作表與「' + FOLDER_NAME + '」資料夾。');
}
