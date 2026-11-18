/**
 * דף יום הורים – שרת קטן בגוגל שיטס
 * 1. מדביקים את כל הקובץ הזה ב: גיליון חדש → תוספים (Extensions) → Apps Script
 * 2. משנים את הקוד הבית ספרי בשורה SCHOOL_CODE למטה (מספר שתמסרו למורים בעל פה)
 * 3. פריסה (Deploy) → פריסה חדשה → אפליקציית אינטרנט (Web app)
 *    הרצה בתור: אני | מי יכול לגשת: כל אחד (Anyone)
 * את הכתובת שמתקבלת (מסתיימת ב- /exec) שולחים לאריק / מדביקים בראש app.js בשורה  const API = '...'
 */
const SCHOOL_CODE = 'CHANGE-ME';   // ← כאן כותבים את הקוד הבית ספרי, למשל '4821'

const CLASSES = 'כיתות';
const ENTRIES = 'הערכות';

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); sh.setRightToLeft(true); }
  return sh;
}
const classesSheet_ = () => sheet_(CLASSES, ['מזהה', 'מפתח מחנך', 'כיתה', 'מחנך/ת', 'נוצר', 'נתונים', 'דוחות']);
const entriesSheet_ = () => sheet_(ENTRIES, ['מזהה כיתה', 'מקצוע', 'תלמיד', 'הערכה', 'חוזקות', 'מענים', 'עודכן', 'נראה', 'מאת']);

function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
/* text that starts with = + - @ would turn into a formula */
function txt_(s) { s = s == null ? '' : String(s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function token_(n) { const a = 'abcdefghijkmnpqrstuvwxyz23456789'; let s = ''; for (let i = 0; i < n; i++) s += a[Math.floor(Math.random() * a.length)]; return s; }

function findClass_(id) {
  if (!id) return null;
  const sh = classesSheet_(), v = sh.getDataRange().getValues();
  for (let r = 1; r < v.length; r++) if (v[r][0] === id) return { sh, row: r + 1, v: v[r], data: JSON.parse(v[r][5] || '{}') };
  return null;
}
function saveData_(c) { c.sh.getRange(c.row, 6).setValue(JSON.stringify(c.data)); }
function entries_(id, subjectId) {
  const v = entriesSheet_().getDataRange().getValues(), out = [];
  for (let r = 1; r < v.length; r++) if (v[r][0] === id && (!subjectId || v[r][1] === subjectId))
    out.push({ subjectId: v[r][1], studentKey: v[r][2], eval: v[r][3], strength: v[r][4], response: v[r][5], updated: v[r][6], seen: v[r][7] === true || v[r][7] === 'TRUE', by: v[r][8] });
  return out;
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

function doGet() { return out_({ ok: true, msg: 'דף יום הורים – השרת עובד' }); }

function doPost(e) {
  let q; try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad request' }); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return out_(handle_(q)); }
  catch (err) { return out_({ error: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle_(q) {
  if (SCHOOL_CODE === 'CHANGE-ME' || !SCHOOL_CODE) return { error: 'עוד לא נקבע קוד בית ספרי בסקריפט (השורה SCHOOL_CODE).' };
  const codeOk = String(q.code || '').trim() === String(SCHOOL_CODE);

  /* ---- homeroom teacher opens a class (needs the school code) ---- */
  if (q.action === 'create') {
    if (!codeOk) return { error: 'קוד בית ספר שגוי', codeError: true };
    const id = token_(10), key = token_(16), data = q.data || {};
    (data.subjects || []).forEach(s => { s.token = token_(12); });
    classesSheet_().appendRow([id, key, txt_(q.name), txt_(q.homeroom), new Date(), JSON.stringify(data), JSON.stringify(q.reports || {})]);
    return { id, key };
  }

  const c = findClass_(q.id);
  if (!c) return { error: 'הכיתה לא נמצאה. בדקו את הקישור.' };
  const admin = !!q.key && q.key === c.v[1];

  /* ---- a subject teacher: personal link (token) + school code. Gets only their own students and entries ---- */
  if (q.t) {
    if (!codeOk) return { error: 'קוד בית ספר שגוי', codeError: true };
    const sub = (c.data.subjects || []).find(s => s.token && s.token === q.t);
    if (!sub) return { error: 'הקישור האישי הזה כבר לא בתוקף. בקשו מהמחנך/ת קישור חדש.' };
    if (q.action === 'tget') {
      const mine = (c.data.students || []).filter(s => sub.students.indexOf(s.key) >= 0);
      return { cls: { id: q.id, name: c.v[2], homeroom: c.v[3] }, subject: { id: sub.id, subject: sub.subject, teacher: sub.teacher }, students: mine, entries: entries_(q.id, sub.id) };
    }
    if (q.action === 'save') {
      if (sub.students.indexOf(q.studentKey) < 0) return { error: 'התלמיד/ה לא ברשימה שלך' };
      return upsert_(q, sub.id, sub.teacher);
    }
    return { error: 'אין הרשאה' };
  }

  /* ---- everything else: only the homeroom teacher (secret key) ---- */
  if (!admin) return { error: 'אין הרשאה' };
  if (q.action === 'get') return { cls: { id: q.id, name: c.v[2], homeroom: c.v[3], data: c.data, reports: JSON.parse(c.v[6] || '{}'), admin: true }, entries: entries_(q.id) };
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
  if (q.action === 'retoken') { /* a new personal link for one teacher; the old one stops working */
    const sub = (c.data.subjects || []).find(s => s.id === q.subjectId); if (!sub) return { error: 'לא נמצא' };
    sub.token = token_(12); saveData_(c); return { token: sub.token };
  }
  if (q.action === 'update') {
    if (q.reports) c.sh.getRange(c.row, 7).setValue(JSON.stringify(q.reports));
    if (q.name != null) c.sh.getRange(c.row, 3).setValue(txt_(q.name));
    if (q.homeroom != null) c.sh.getRange(c.row, 4).setValue(txt_(q.homeroom));
    return { ok: true };
  }
  return { error: 'unknown action' };
}
