/**
 * MAP Math & Science — Google Sheets backend
 * วิธีใช้: วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * (Execute as: Me · Who has access: Anyone)
 */

// ① เปลี่ยนรหัสสถาบันตรงนี้ (ครูต้องใช้รหัสนี้ตอนเชื่อมต่อในเว็บ) — ใช้ตัวอักษรภาษาอังกฤษหรือตัวเลข อย่างน้อย 8 ตัว
const ACCESS_KEY = 'CHANGE-ME-2026';

// ② รหัสผู้ดูแล — ใช้เข้าเมนู ⚙️ ตั้งค่า ในเว็บ (ตั้งให้ต่างจาก ACCESS_KEY และไม่บอกครูคนอื่น)
const ADMIN_KEY = 'CHANGE-ADMIN-2026';

const SHEETS = {
  students: ['id', 'updatedAt', 'name', 'nickname', 'grade', 'code', 'json'],
  tests:    ['id', 'updatedAt', 'studentName', 'studentId', 'subject', 'grade', 'finishedAt', 'score', 'se', 'level', 'correct', 'items', 'json'],
  practice: ['id', 'updatedAt', 'studentName', 'studentId', 'createdAt', 'strands', 'score', 'total', 'json']
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key)) return out_({ ok: false, error: 'unauthorized' });
  if (p.action === 'ping') return out_({ ok: true, time: Date.now() });
  if (p.action === 'admin') {
    const ok = ADMIN_KEY !== 'CHANGE-ADMIN-2026' && typeof p.admin === 'string' && p.admin === ADMIN_KEY;
    return out_({ ok: ok, admin: true, error: ok ? '' : (ADMIN_KEY === 'CHANGE-ADMIN-2026' ? 'admin key not set' : 'wrong admin key') });
  }
  return out_({
    ok: true,
    students: readAll_('students'),
    tests: readAll_('tests'),
    practices: readAll_('practice'),
    time: Date.now()
  });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, error: 'bad json' }); }
  if (!checkKey_(b.key)) return out_({ ok: false, error: 'unauthorized' });
  if (!SHEETS[b.col] || !b.id) return out_({ ok: false, error: 'bad request' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(b.col);
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    const idx = ids.indexOf(String(b.id));
    if (b.action === 'del') {
      if (idx >= 0) sh.deleteRow(idx + 2);
      return out_({ ok: true });
    }
    const row = toRow_(b.col, String(b.id), b.data || {});
    if (idx >= 0) sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- helpers ---------- */
function checkKey_(k) { return typeof k === 'string' && k.length > 0 && k === ACCESS_KEY; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold').setBackground('#e4edfb');
  }
  return sh;
}

function readAll_(name) {
  const sh = sheet_(name);
  const v = sh.getDataRange().getValues();
  const j = v[0].indexOf('json');
  const res = [];
  for (let i = 1; i < v.length; i++) {
    try { const o = JSON.parse(v[i][j]); o.id = String(v[i][0]); res.push(o); } catch (x) {}
  }
  return res;
}

function studentName_(id) {
  const sh = sheet_('students');
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return v[i][2];
  return '';
}

function gradeLabel_(g) { g = Number(g); return g <= 6 ? 'ป.' + g : 'ม.' + (g - 6); }
function levelLabel_(score) {
  const L = Math.round((score - 140) / 8);
  if (L < 1) return 'ต่ำกว่าต้น ป.1';
  if (L > 18) return 'สูงกว่า ม.3';
  return (L % 2 ? 'ต้น ' : 'ปลาย ') + gradeLabel_(Math.ceil(L / 2));
}

function toRow_(col, id, d) {
  const now = new Date();
  const J = JSON.stringify(d);
  if (col === 'students') return [id, now, d.name || '', d.nickname || '', gradeLabel_(d.grade), d.code || '', J];
  if (col === 'tests') {
    const rs = d.responses || [];
    const k = rs.filter(function (r) { return r.c; }).length;
    return [id, now, studentName_(d.studentId), d.studentId || '', d.subject === 'sci' ? 'Science' : 'Math', gradeLabel_(d.grade),
      d.finishedAt ? new Date(d.finishedAt) : '', Math.round(d.theta || 0), d.se || '', levelLabel_(d.theta || 0), k, rs.length, J];
  }
  return [id, now, studentName_(d.studentId), d.studentId || '', d.createdAt ? new Date(d.createdAt) : '', (d.strands || []).join(', '), d.score, d.total, J];
}

/** กด Run ฟังก์ชันนี้ 1 ครั้งหลังวางโค้ด เพื่อสร้างแท็บ students / tests / practice และอนุญาตสิทธิ์ */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  if (ACCESS_KEY === 'CHANGE-ME-2026') Logger.log('⚠️ อย่าลืมเปลี่ยน ACCESS_KEY เป็นรหัสของสถาบันก่อน Deploy');
  if (ADMIN_KEY === 'CHANGE-ADMIN-2026') Logger.log('⚠️ อย่าลืมเปลี่ยน ADMIN_KEY (รหัสผู้ดูแล)');
  else Logger.log('พร้อมใช้งาน — ไปที่ Deploy > New deployment > Web app');
}
