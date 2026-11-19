/**
 * דף יום הורים – שרת קטן בגוגל שיטס
 * התקנה: גיליון → תוספים (Extensions) → Apps Script → מדביקים את כל הקובץ → שמירה
 *        → Deploy → New deployment → Web app | Execute as: Me | Who has access: Anyone
 * עדכון גרסה: מדביקים מחדש → שמירה → Deploy → Manage deployments → ✏️ → Version: New version → Deploy
 * הקודים לא כתובים כאן: הם בלשונית "הגדרות" בגיליון (נוצרת לבד בפתיחה הראשונה של כתובת ה-/exec).
 */
const CLASSES = 'כיתות';
const ENTRIES = 'הערכות';
const TEACHERS = 'מורים';
const SETTINGS = 'הגדרות';
const COORDS = 'רכזי שכבה';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
const classesSheet_ = () => sheet_(CLASSES, ['מזהה', 'מפתח מחנך', 'כיתה', 'מחנך/ת', 'נוצר', 'נתונים', 'דוחות']);
const entriesSheet_ = () => sheet_(ENTRIES, ['מזהה כיתה', 'מקצוע', 'תלמיד', 'הערכה', 'חוזקות', 'מענים', 'עודכן', 'נראה', 'מאת']);
const coordsSheet_ = () => sheet_(COORDS, ['שכבה', 'רכז/ת', 'קוד קישור', 'נשלח', 'נוצר']);
const teachersSheet_ = () => sheet_(TEACHERS, ['שם המורה', 'קוד קישור', 'נשלח', 'נוצר']);

/* the codes live in the "הגדרות" tab, so pasting a new version of this script never needs editing */
function settings_() {
  let sh = ss_().getSheetByName(SETTINGS);
  if (!sh) {
    sh = ss_().insertSheet(SETTINGS); sh.setRightToLeft(true);
    sh.getRange(1, 1, 3, 3).setValues([
      ['הגדרה', 'ערך', 'הסבר'],
      ['קוד בית ספרי', '', 'קוד שנמסר לכל המורים בעל פה. נדרש בכניסה הראשונה ובפתיחת כיתה'],
      ['קוד מנהל', '', 'קוד לדף מנהל המערכת (כל הכיתות, המורים ורכזי השכבות). רק לך']]);
    sh.getRange('B2:B3').setNumberFormat('@'); sh.setColumnWidth(3, 420);
  }
  const v = sh.getRange(2, 2, 2, 1).getValues();
  return { school: String(v[0][0]).trim(), admin: String(v[1][0]).trim() };
}

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
/* text that starts with = + - @ would turn into a formula */
function txt_(s) { s = s == null ? '' : String(s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function token_(n) { const a = 'abcdefghijkmnpqrstuvwxyz23456789'; let s = ''; for (let i = 0; i < n; i++) s += a[Math.floor(Math.random() * a.length)]; return s; }
const clsKey_ = s => String(s == null ? '' : s).replace(/["'`׳״\s]/g, '');
const norm_ = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

function allClasses_() {
  const sh = classesSheet_(), v = sh.getDataRange().getValues(), out = [];
  for (let r = 1; r < v.length; r++) out.push({ sh, row: r + 1, v: v[r], id: v[r][0], data: JSON.parse(v[r][5] || '{}') });
  return out;
}
function findClass_(id) { return id ? allClasses_().find(c => c.id === id) || null : null; }
function saveData_(c) { c.sh.getRange(c.row, 6).setValue(JSON.stringify(c.data)); }
function entries_(id, subjectId) {
  const v = entriesSheet_().getDataRange().getValues(), out = [];
  for (let r = 1; r < v.length; r++) if (v[r][0] === id && (!subjectId || v[r][1] === subjectId))
    out.push({ subjectId: v[r][1], studentKey: v[r][2], eval: v[r][3], strength: v[r][4], response: v[r][5], updated: v[r][6], seen: v[r][7] === true || v[r][7] === 'TRUE', by: v[r][8] });
  return out;
}
/* how many students already have an evaluation, per class + subject */
function doneCounts_() {
  const v = entriesSheet_().getDataRange().getValues(), n = {};
  for (let r = 1; r < v.length; r++) if (norm_(v[r][3])) { const k = v[r][0] + '|' + v[r][1]; n[k] = (n[k] || 0) + 1; }
  return n;
}
function upsert_(q, subjectId, by) {
  const sh = entriesSheet_(), v = sh.getDataRange().getValues();
  let row = -1;
  for (let r = 1; r < v.length; r++) if (v[r][0] === q.id && v[r][1] === subjectId && v[r][2] === q.studentKey) { row = r + 1; break; }
  const now = new Date(), hr = subjectId === '_hr', extra = !!(q.strength || q.response);
  if (row < 0) sh.appendRow([q.id, subjectId, q.studentKey, txt_(q.eval), txt_(q.strength), txt_(q.response), now, hr || !extra, by]);
  else {
    const old = v[row - 1];
    const changed = (q.strength || '') !== String(old[4]) || (q.response || '') !== String(old[5]);
    const seen = hr ? true : (changed ? !extra : old[7]);
    sh.getRange(row, 4, 1, 6).setValues([[txt_(q.eval), txt_(q.strength), txt_(q.response), now, seen, by]]);
  }
  return { ok: true, at: now };
}

/* ---------- teachers: one personal link per teacher (by name), for all of their classes ---------- */
function teachers_() {
  const sh = teachersSheet_(), v = sh.getDataRange().getValues(), list = [];
  for (let r = 1; r < v.length; r++) list.push({ row: r + 1, name: norm_(v[r][0]), token: String(v[r][1]), sent: v[r][2] === true || v[r][2] === 'TRUE' });
  return { sh, list };
}
function ensureTeachers_(names) {
  const t = teachers_(), out = {};
  for (const raw of names) {
    const name = norm_(raw); if (!name || out[name]) continue;
    let x = t.list.find(y => y.name === name);
    if (!x) { x = { name, token: token_(14), sent: false }; t.sh.appendRow([txt_(name), x.token, false, new Date()]); x.row = t.sh.getLastRow(); t.list.push(x); }
    out[name] = x;
  }
  return out;
}
function teacherByToken_(tok) { return tok ? teachers_().list.find(x => x.token === tok) || null : null; }
/* a teacher's classes, from every class in the sheet */
function teacherClasses_(name, done) {
  const out = [];
  for (const c of allClasses_()) for (const s of (c.data.subjects || []))
    if (norm_(s.teacher) === name) out.push({ id: c.id, cls: c.v[2], homeroom: c.v[3], subjectId: s.id, subject: s.subject, total: s.students.length, done: done[c.id + '|' + s.id] || 0 });
  return out;
}

/* ---------- grades and grade coordinators ---------- */
const gradeOf_ = name => (String(name || '').match(/[\u05D0-\u05EA]/) || [''])[0];
function coords_() {
  const sh = coordsSheet_(), v = sh.getDataRange().getValues(), list = [];
  for (let r = 1; r < v.length; r++) list.push({ row: r + 1, grade: String(v[r][0]).trim(), name: norm_(v[r][1]), token: String(v[r][2]), sent: v[r][3] === true || v[r][3] === 'TRUE' });
  return { sh, list };
}
function coordByToken_(tok) { return tok ? coords_().list.find(x => x.token === tok) || null : null; }
/* classes (with progress) and their teachers, for the whole school or for one grade */
function overview_(grade) {
  const classes = allClasses_().filter(c => !grade || gradeOf_(c.v[2]) === grade), done = doneCounts_();
  const per = {}, names = [];
  const cls = classes.map(c => {
    let d = 0, t = 0;
    for (const s of (c.data.subjects || [])) {
      const n = Math.min(done[c.id + '|' + s.id] || 0, s.students.length); d += n; t += s.students.length;
      const name = norm_(s.teacher); names.push(name);
      const p = (per[name] = per[name] || { classes: {}, done: 0, total: 0 }); p.classes[c.id] = 1; p.done += n; p.total += s.students.length;
    }
    return { id: c.id, key: c.v[1], name: c.v[2], homeroom: c.v[3], grade: gradeOf_(c.v[2]), done: d, total: t };
  });
  const map = ensureTeachers_(names);
  const teachers = Object.values(map).map(x => ({ name: x.name, token: x.token, sent: x.sent, classes: Object.keys(per[x.name].classes).length, done: per[x.name].done, total: per[x.name].total }))
    .sort((a, b) => a.name.localeCompare(b.name, 'he'));
  return { classes: cls, teachers };
}

function doGet() { settings_(); return out_({ ok: true, msg: 'דף יום הורים – השרת עובד' }); }

function doPost(e) {
  let q; try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad request' }); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return out_(handle_(q)); }
  catch (err) { return out_({ error: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle_(q) {
  const S = settings_();
  if (!S.school) return { error: 'עוד לא נקבע קוד בית ספרי (בגיליון, לשונית "הגדרות").' };
  const codeOk = String(q.code || '').trim() === S.school;

  const adminOk = !!S.admin && String(q.admin || '').trim() === S.admin;
  const coord = q.g ? coordByToken_(q.g) : null;

  /* ---- the system manager: every class, teacher and grade coordinator ---- */
  if (q.action && q.action.indexOf('admin') === 0) {
    if (!adminOk) return { error: 'קוד מנהל שגוי', adminError: true };
    if (q.action === 'adminGet') {
      const o = overview_('');
      return { teachers: o.teachers, classes: o.classes, coords: coords_().list.map(x => ({ grade: x.grade, name: x.name, token: x.token, sent: x.sent })) };
    }
    if (q.action === 'adminSetCoord') { /* appoint (or rename) the coordinator of a grade */
      const g = gradeOf_(q.grade); if (!g) return { error: 'חסרה שכבה' };
      const c = coords_(), x = c.list.find(y => y.grade === g);
      if (x) { c.sh.getRange(x.row, 2).setValue(txt_(norm_(q.name))); if (q.newToken) c.sh.getRange(x.row, 3, 1, 2).setValues([[token_(14), false]]); }
      else c.sh.appendRow([g, txt_(norm_(q.name)), token_(14), false, new Date()]);
      return { ok: true };
    }
    if (q.action === 'adminCoordSent') { const x = coords_().list.find(y => y.grade === gradeOf_(q.grade)); if (x) coordsSheet_().getRange(x.row, 4).setValue(true); return { ok: true }; }
    if (q.action === 'adminRetoken') {
      const x = teachers_().list.find(y => y.name === norm_(q.name)); if (!x) return { error: 'לא נמצא' };
      x.token = token_(14); teachersSheet_().getRange(x.row, 2, 1, 2).setValues([[x.token, false]]); return { token: x.token };
    }
    if (q.action === 'adminSent') { return markSent_(q.names || []); }
    return { error: 'unknown action' };
  }

  /* ---- a grade coordinator: personal link + school code, only their grade ---- */
  if (q.g && q.action !== 'create') {
    if (!codeOk) return { error: 'קוד בית ספר שגוי', codeError: true };
    if (!coord) return { error: 'הקישור האישי הזה כבר לא בתוקף. בקשו קישור חדש ממנהל המערכת.' };
    if (q.action === 'gget') { const o = overview_(coord.grade); return { name: coord.name, grade: coord.grade, classes: o.classes, teachers: o.teachers }; }
    if (q.action === 'gsent') { const mine = overview_(coord.grade).teachers.map(t => t.name); return markSent_((q.names || []).filter(n => mine.indexOf(norm_(n)) >= 0)); }
    return { error: 'אין הרשאה' };
  }

  /* ---- opening a class: the system manager, or a grade coordinator for their own grade ---- */
  if (q.action === 'create') {
    if (!adminOk) {
      if (!q.g) return { error: 'רק מנהל המערכת או רכז/ת השכבה פותחים כיתות.' };
      if (!codeOk) return { error: 'קוד בית ספר שגוי', codeError: true };
      if (!coord) return { error: 'הקישור האישי הזה כבר לא בתוקף.' };
      if (gradeOf_(q.name) !== coord.grade) return { error: 'הכיתה ' + q.name + ' לא בשכבה שלך.', notMine: true };
    }
    const ck = clsKey_(q.name);
    if (!ck) return { error: 'חסר שם כיתה' };
    if (allClasses_().some(c => clsKey_(c.v[2]) === ck)) return { error: 'הכיתה ' + q.name + ' כבר נפתחה.', exists: true };
    const id = token_(10), key = token_(16), data = q.data || {};
    classesSheet_().appendRow([id, key, txt_(q.name), txt_(q.homeroom), new Date(), JSON.stringify(data), JSON.stringify(q.reports || {})]);
    ensureTeachers_((data.subjects || []).map(s => s.teacher));
    return { id, key };
  }

  /* ---- a subject teacher: personal link (one per teacher, all classes) + school code ---- */
  if (q.m) {
    if (!codeOk) return { error: 'קוד בית ספר שגוי', codeError: true };
    const me = teacherByToken_(q.m);
    if (!me) return { error: 'הקישור האישי הזה כבר לא בתוקף. בקשו קישור חדש מהרכז/ת.' };
    if (q.action === 'mine') return { name: me.name, classes: teacherClasses_(me.name, doneCounts_()) };
    const c = findClass_(q.id); if (!c) return { error: 'הכיתה לא נמצאה' };
    const sub = (c.data.subjects || []).find(s => s.id === q.subjectId && norm_(s.teacher) === me.name);
    if (!sub) return { error: 'אין לך הרשאה למקצוע הזה' };
    if (q.action === 'mclass') {
      const mine = (c.data.students || []).filter(s => sub.students.indexOf(s.key) >= 0);
      return { cls: { id: c.id, name: c.v[2], homeroom: c.v[3] }, subject: { id: sub.id, subject: sub.subject, teacher: sub.teacher }, students: mine, entries: entries_(c.id, sub.id) };
    }
    if (q.action === 'save') {
      if (sub.students.indexOf(q.studentKey) < 0) return { error: 'התלמיד/ה לא ברשימה שלך' };
      return upsert_(q, sub.id, sub.teacher);
    }
    return { error: 'אין הרשאה' };
  }

  const c = findClass_(q.id);
  if (!c) return { error: 'הכיתה לא נמצאה. בדקו את הקישור.' };

  /* ---- everything else: only the homeroom teacher (secret key) ---- */
  if (!q.key || q.key !== c.v[1]) return { error: 'אין הרשאה' };
  if (q.action === 'get') {
    const map = ensureTeachers_((c.data.subjects || []).map(s => s.teacher));
    const teachers = Object.values(map).map(t => ({ name: t.name, token: t.token, sent: t.sent }));
    return { cls: { id: q.id, name: c.v[2], homeroom: c.v[3], data: c.data, reports: JSON.parse(c.v[6] || '{}'), admin: true }, entries: entries_(q.id), teachers };
  }
  if (q.action === 'sent') {
    const mine = (c.data.subjects || []).map(s => norm_(s.teacher));
    return markSent_((q.names || []).filter(n => mine.indexOf(norm_(n)) >= 0));
  }
  if (q.action === 'save') { if (q.subjectId !== '_hr') return { error: 'אין הרשאה' }; return upsert_(q, '_hr', 'מחנך/ת'); }
  if (q.action === 'review') {
    const sh = entriesSheet_(), v = sh.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) if (v[r][0] === q.id && v[r][1] === q.subjectId && v[r][2] === q.studentKey) {
      sh.getRange(r + 1, 5, 1, 2).setValues([[txt_(q.strength != null ? q.strength : v[r][4]), txt_(q.response != null ? q.response : v[r][5])]]);
      sh.getRange(r + 1, 8).setValue(true);
      return { ok: true };
    }
    return { error: 'לא נמצא' };
  }
  if (q.action === 'update') {
    if (q.reports) c.sh.getRange(c.row, 7).setValue(JSON.stringify(q.reports));
    if (q.name != null) c.sh.getRange(c.row, 3).setValue(txt_(q.name));
    if (q.homeroom != null) c.sh.getRange(c.row, 4).setValue(txt_(q.homeroom));
    return { ok: true };
  }
  return { error: 'unknown action' };
}
function markSent_(names) {
  const t = teachers_();
  for (const n of names) { const x = t.list.find(y => y.name === norm_(n)); if (x && !x.sent) { t.sh.getRange(x.row, 3).setValue(true); x.sent = true; } }
  return { ok: true };
}
