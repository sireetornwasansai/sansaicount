/**
 * Backend: Google Apps Script + Google Sheet  (แดชบอร์ดดอกไม้จันทน์ รพ.สันทราย)
 * Deploy > New deployment > Web app  |  Execute as: Me  |  Who has access: Anyone
 *
 * Sheet 3 แท็บ (แถวแรกเป็นหัวคอลัมน์ ตัวพิมพ์เล็ก):
 *  Daily    : date | count | note | photo
 *  Config   : key | value   (updated_at, target, deadline, total[ไม่บังคับ ใช้แทนผลรวมรายวัน])
 *  Comments : timestamp | name | message | approved   (ติ๊ก approved = TRUE ถึงจะแสดงบนเว็บ)
 *
 * ครั้งแรก: เลือกฟังก์ชัน setup แล้วกด Run 1 ครั้ง เพื่อสร้างแท็บ/หัวคอลัมน์ที่ยังไม่มี
 *
 * ระบบแอดมิน: หน้าเว็บส่งรหัสผ่านมาตรวจที่นี่เท่านั้น (รหัสไม่อยู่ในไฟล์ index.html)
 *  - ล็อกอินสำเร็จ จะได้ token อายุ 2 ชั่วโมง ใช้เรียกคำสั่งแก้ไขข้อมูล
 *  - กรอกรหัสผิดครบ 5 ครั้ง ระบบล็อก 10 นาที
 */

const SHEET_ID = '1HRtOt5JzETig1HuncMxONXoUQDUWEkWhccWxYhoyM5E';
const TZ = 'Asia/Bangkok';
const HEADERS = {
  Daily:    ['date', 'count', 'note', 'photo'],
  Config:   ['key', 'value'],
  Comments: ['timestamp', 'name', 'message', 'approved']
};

// รหัสผ่านแอดมิน: แนะนำให้ย้ายไปเก็บที่ Project Settings > Script properties
// ชื่อคีย์ ADMIN_PASSWORD (ถ้ามีค่าตรงนั้น จะใช้ค่านั้นแทนค่าด้านล่าง)
const ADMIN_PASSWORD_DEFAULT = 'SANSAICOUNT2570';
const TOKEN_TTL_SEC = 7200;
const MAX_FAILS = 5;
const LOCK_SEC = 600;

function adminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD_DEFAULT;
}

function ss_() {
  // สคริปต์แบบ standalone ไม่มี getActive() จึงเปิดด้วย ID เสมอ
  return SpreadsheetApp.openById(SHEET_ID);
}

function setup() {
  const ss = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
  });
}

function doGet() {
  try {
    const daily = readRows_('Daily').map(function (r) {
      return { date: fmtDate_(r.date), count: toNum_(r.count), note: String(r.note || ''), photo: String(r.photo || '') };
    }).filter(function (r) { return r.date; });

    const config = {};
    readRows_('Config').forEach(function (r) {
      if (r.key) config[String(r.key).trim()] = r.value instanceof Date ? fmtDateTime_(r.value) : r.value;
    });

    const comments = readRows_('Comments')
      .filter(function (r) { return String(r.approved).toUpperCase() === 'TRUE'; })
      .map(function (r) { return { timestamp: fmtDateTime_(r.timestamp), name: String(r.name || ''), message: String(r.message || '') }; })
      .slice(-30).reverse();

    return json_({ daily: daily, config: config, comments: comments });
  } catch (err) {
    return json_({ daily: [], config: {}, comments: [], error: String(err) });
  }
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'bad_request' }); }
  if (d && d.action) return admin_(d);
  return addComment_(d || {});
}

/* ---------- ความคิดเห็นจากผู้เข้าชม ---------- */
function addComment_(d) {
  const lock = LockService.getScriptLock();
  try {
    const name = clean_(d.name, 60);
    const message = clean_(d.message, 500);
    if (!name || !message) return json_({ ok: false, error: 'missing' });

    // กันส่งรัวเกินไป (ทั้งระบบ 1 รายการ / 5 วินาที)
    const cache = CacheService.getScriptCache();
    if (cache.get('last_comment')) return json_({ ok: false, error: 'too_fast' });

    lock.waitLock(5000);
    const sh = ss_().getSheetByName('Comments');
    if (!sh) return json_({ ok: false, error: 'no_sheet' });
    sh.appendRow([new Date(), name, message, false]); // รอแอดมินอนุมัติ
    cache.put('last_comment', '1', 5);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- แอดมิน ---------- */
function admin_(d) {
  const cache = CacheService.getScriptCache();

  if (d.action === 'login') {
    const fails = Number(cache.get('admin_fails') || 0);
    if (fails >= MAX_FAILS) return json_({ ok: false, error: 'locked' });
    if (String(d.password || '') !== adminPassword_()) {
      cache.put('admin_fails', String(fails + 1), LOCK_SEC);
      return json_({ ok: false, error: 'wrong_password', left: MAX_FAILS - fails - 1 });
    }
    cache.remove('admin_fails');
    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('tok_' + token, '1', TOKEN_TTL_SEC);
    return json_({ ok: true, token: token });
  }

  const tokKey = 'tok_' + String(d.token || '');
  if (!d.token || !cache.get(tokKey)) return json_({ ok: false, error: 'unauthorized' });
  if (d.action === 'logout') { cache.remove(tokKey); return json_({ ok: true }); }
  if (d.action === 'ping') return json_({ ok: true });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    switch (d.action) {
      case 'save_daily':   return saveDaily_(d);
      case 'delete_daily': return deleteDaily_(d);
      case 'save_config':  return saveConfig_(d);
      case 'list_pending': return listPending_();
      case 'moderate':     return moderate_(d);
      default:             return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function saveDaily_(d) {
  const sh = ss_().getSheetByName('Daily');
  if (!sh) return json_({ ok: false, error: 'no_sheet' });
  const date = fmtDate_(d.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'bad_date' });
  const count = Math.round(toNum_(d.count));
  if (!(count >= 0 && count <= 10000000)) return json_({ ok: false, error: 'bad_count' });
  const note = clean_(d.note, 200);
  let photo = String(d.photo || '').trim().slice(0, 500);
  if (photo && !/^https?:\/\//i.test(photo)) return json_({ ok: false, error: 'bad_photo' });

  let row = findDailyRow_(sh, date);
  const isNew = !row;
  if (!row) row = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(row, 1).setNumberFormat('@').setValue(date);
  sh.getRange(row, 2, 1, 3).setValues([[count, note, photo]]);
  touch_();
  return json_({ ok: true, mode: isNew ? 'added' : 'updated' });
}

function deleteDaily_(d) {
  const sh = ss_().getSheetByName('Daily');
  const date = fmtDate_(d.date);
  const row = sh ? findDailyRow_(sh, date) : 0;
  if (!row) return json_({ ok: false, error: 'not_found' });
  sh.deleteRow(row);
  touch_();
  return json_({ ok: true });
}

function findDailyRow_(sh, date) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (fmtDate_(vals[i][0]) === date) return i + 2;
  }
  return 0;
}

function saveConfig_(d) {
  const out = {};
  if ('target' in d) {
    const t = Math.round(toNum_(d.target));
    if (!(t >= 1 && t <= 10000000)) return json_({ ok: false, error: 'bad_target' });
    out.target = t;
  }
  if ('deadline' in d) {
    const s = String(d.deadline || '').trim();
    if (s && isNaN(new Date(s).getTime())) return json_({ ok: false, error: 'bad_deadline' });
    out.deadline = s;
  }
  if ('total' in d) {
    const raw = String(d.total == null ? '' : d.total).trim();
    if (raw === '') out.total = '';           // ว่าง = ลบ ให้กลับไปใช้ผลรวมรายวัน
    else {
      const n = Math.round(toNum_(raw));
      if (!(n >= 0 && n <= 10000000)) return json_({ ok: false, error: 'bad_total' });
      out.total = n;
    }
  }
  Object.keys(out).forEach(function (k) { setConfig_(k, out[k]); });
  touch_();
  return json_({ ok: true });
}

function setConfig_(key, value) {
  const sh = ss_().getSheetByName('Config');
  const last = sh.getLastRow();
  let row = 0;
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === key) { row = i + 2; break; }
  }
  if (value === '') { if (row) sh.deleteRow(row); return; }
  if (!row) row = Math.max(last, 1) + 1;
  sh.getRange(row, 1).setValue(key);
  sh.getRange(row, 2).setNumberFormat('@').setValue(String(value));
}

function touch_() {
  setConfig_('updated_at', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
}

function listPending_() {
  const sh = ss_().getSheetByName('Comments');
  if (!sh || sh.getLastRow() < 2) return json_({ ok: true, items: [] });
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const items = [];
  vals.forEach(function (r, i) {
    if (String(r[3]).toUpperCase() === 'TRUE') return;
    items.push({ row: i + 2, timestamp: fmtDateTime_(r[0]), name: String(r[1] || ''), message: String(r[2] || '') });
  });
  return json_({ ok: true, items: items.slice(-50).reverse() });
}

function moderate_(d) {
  const sh = ss_().getSheetByName('Comments');
  const row = Number(d.row);
  if (!sh || !(row >= 2 && row <= sh.getLastRow())) return json_({ ok: false, error: 'bad_row' });
  // กันแถวเลื่อน: ข้อความในแถวต้องตรงกับที่แอดมินเห็น
  if (String(sh.getRange(row, 3).getValue()) !== String(d.message)) return json_({ ok: false, error: 'stale' });
  if (d.decision === 'approve') sh.getRange(row, 4).setValue(true);
  else if (d.decision === 'delete') sh.deleteRow(row);
  else return json_({ ok: false, error: 'bad_decision' });
  return json_({ ok: true });
}

/* ---------- helpers ---------- */
function readRows_(sheetName) {
  const sh = ss_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const head = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  return values.slice(1).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { if (h) o[h] = r[i]; });
    return o;
  });
}
function toNum_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v == null ? '' : v).replace(/[,\s]/g, ''));
  return isFinite(n) ? n : 0;
}
function clean_(v, max) {
  let t = String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t; // กัน formula injection ในชีต
  return t;
}
function fmtDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); // d/m/yyyy (รองรับ พ.ศ.)
  if (m) { const y = Number(m[3]) > 2400 ? Number(m[3]) - 543 : Number(m[3]); return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); }
  return s;
}
function fmtDateTime_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v || '');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
