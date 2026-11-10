/**
 * نظام مباشرة الطلبة المقبولين — كلية علوم الحاسوب والرياضيات
 * الواجهة الخلفية (Google Apps Script) — تعمل على جوجل شيت الطلبة
 * تصميم: OMAR INAAM
 *
 * الأوراق:
 *  - ورقة1      : التسلسل | القسم | اسم الطالب | مباشرة | وقت المباشرة (يُضاف تلقائياً)
 *  - الأقسام    : القسم | الرمز   (تُنشأ تلقائياً برموز عشوائية عند تشغيل setup)
 */

const STUDENTS_SHEET = 'ورقة1';
const DEPTS_SHEET = 'الأقسام';
const DONE_TEXT = 'تمت المباشرة';
const TZ = 'Asia/Baghdad';

const H_ID = 'التسلسل';
const H_DEPT = 'القسم';
const H_NAME = 'اسم الطالب';
const H_DONE = 'مباشرة';
const H_TIME = 'وقت المباشرة';

/* ======================= نقطة الدخول ======================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    switch (p.action) {
      case 'depts':   out = { ok: true, depts: getDepts_().map(d => d.name) }; break;
      case 'login':   out = login_(p.dept, p.pin); break;
      case 'list':    auth_(p.dept, p.pin); out = { ok: true, students: listDept_(p.dept) }; break;
      case 'lookup':  auth_(p.dept, p.pin); out = lookup_(p.id); break;
      case 'confirm': auth_(p.dept, p.pin); out = confirm_(p.dept, p.id); break;
      default:        out = { ok: true, msg: 'نظام المباشرة يعمل' };
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================= الإعداد الأولي =======================
 * شغّل هذه الدالة مرة واحدة من محرر السكربت (Run ▶ setup)
 * تنشئ ورقة "الأقسام" برمز سري لكل قسم وتضيف عمود "وقت المباشرة".
 */
function setup() {
  const sh = studentsSheet_();
  ensureTimeColumn_(sh);
  const depts = getDepts_();
  Logger.log('الأقسام ورموزها:\n' + depts.map(d => d.name + ' : ' + d.pin).join('\n'));
}

/* ======================= منطق العمل ======================= */
function login_(dept, pin) {
  auth_(dept, pin);
  return { ok: true, dept: dept, students: listDept_(dept) };
}

function auth_(dept, pin) {
  if (!dept || !pin) throw new Error('يرجى اختيار القسم وإدخال الرمز');
  const d = getDepts_().find(x => norm_(x.name) === norm_(dept));
  if (!d) throw new Error('القسم غير موجود');
  if (String(d.pin).trim() !== String(pin).trim()) throw new Error('الرمز السري غير صحيح');
}

function listDept_(dept) {
  const t = readTable_();
  const nd = norm_(dept);
  return t.rows
    .filter(r => norm_(r.dept) === nd)
    .map(r => ({ id: r.id, name: r.name, done: r.done, time: r.time }));
}

function lookup_(id) {
  const key = normId_(id);
  const r = readTable_().rows.find(x => normId_(x.id) === key);
  if (!r) return { ok: true, found: false };
  return { ok: true, found: true, student: { id: r.id, name: r.name, dept: r.dept, done: r.done, time: r.time } };
}

function confirm_(dept, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = studentsSheet_();
    const t = readTable_(sh);
    const key = normId_(id);
    const r = t.rows.find(x => normId_(x.id) === key);
    if (!r) return { ok: false, error: 'الطالب غير موجود في القائمة' };
    if (norm_(r.dept) !== norm_(dept)) {
      return { ok: false, error: 'الطالب مقبول في قسم آخر: ' + r.dept, otherDept: r.dept };
    }
    if (r.done) {
      return { ok: true, already: true, student: { id: r.id, name: r.name, done: true, time: r.time } };
    }
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm');
    sh.getRange(r.row, t.col.done + 1).setValue(DONE_TEXT);
    sh.getRange(r.row, t.col.time + 1).setValue(now);
    SpreadsheetApp.flush();
    return { ok: true, already: false, student: { id: r.id, name: r.name, done: true, time: now } };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= أدوات القراءة ======================= */
function studentsSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(STUDENTS_SHEET) || ss.getSheets()[0];
}

function ensureTimeColumn_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  if (head.indexOf(H_TIME) === -1) sh.getRange(1, lastCol + 1).setValue(H_TIME);
}

function readTable_(sh) {
  sh = sh || studentsSheet_();
  ensureTimeColumn_(sh);
  const data = sh.getDataRange().getDisplayValues();
  const head = data[0].map(h => String(h).trim());
  const col = {
    id: head.indexOf(H_ID), dept: head.indexOf(H_DEPT), name: head.indexOf(H_NAME),
    done: head.indexOf(H_DONE), time: head.indexOf(H_TIME)
  };
  ['id', 'dept', 'name', 'done'].forEach(k => {
    if (col[k] === -1) throw new Error('عمود مفقود في الشيت: ' + { id: H_ID, dept: H_DEPT, name: H_NAME, done: H_DONE }[k]);
  });
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const v = data[i];
    const id = String(v[col.id]).trim();
    if (!id) continue;
    rows.push({
      row: i + 1, id: id, dept: String(v[col.dept]).trim(), name: String(v[col.name]).trim(),
      done: String(v[col.done]).trim() !== '', time: col.time > -1 ? String(v[col.time]).trim() : ''
    });
  }
  return { col: col, rows: rows };
}

function getDepts_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(DEPTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DEPTS_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['القسم', 'الرمز']]).setFontWeight('bold');
    sh.setRightToLeft(true);
  }
  // إضافة أي قسم جديد يظهر في قائمة الطلبة تلقائياً برمز عشوائي
  const existing = sh.getDataRange().getDisplayValues().slice(1)
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), pin: String(r[1]).trim() }));
  const known = existing.map(d => norm_(d.name));
  const fromStudents = readTableNoDepts_();
  const toAdd = [];
  fromStudents.forEach(n => {
    if (known.indexOf(norm_(n)) === -1) {
      known.push(norm_(n));
      const pin = String(Math.floor(1000 + Math.random() * 9000));
      toAdd.push([n, pin]);
      existing.push({ name: n, pin: pin });
    }
  });
  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setNumberFormat('@').setValues(toAdd);
  }
  return existing;
}

function readTableNoDepts_() {
  const sh = studentsSheet_();
  const data = sh.getDataRange().getDisplayValues();
  const c = data[0].map(h => String(h).trim()).indexOf(H_DEPT);
  if (c === -1) return [];
  const seen = {}, out = [];
  for (let i = 1; i < data.length; i++) {
    const n = String(data[i][c]).trim();
    if (n && !seen[norm_(n)]) { seen[norm_(n)] = 1; out.push(n); }
  }
  return out;
}

/* توحيد الكتابة العربية للمقارنة (الامن = الأمن ...) */
function norm_(s) {
  return String(s || '').trim()
    .replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي')
    .replace(/[ً-ْـ]/g, '').replace(/\s+/g, ' ');
}
function normId_(s) {
  return String(s || '').trim()
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/^0+(?=\d)/, '');
}
