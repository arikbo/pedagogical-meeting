/**
 * דף יום הורים – שרת קטן בגוגל שיטס
 * מדביקים את כל הקובץ הזה ב: גיליון חדש → תוספים (Extensions) → Apps Script
 * ואז: פריסה (Deploy) → פריסה חדשה → אפליקציית אינטרנט (Web app)
 *      הרצה בתור: אני | מי יכול לגשת: כל אחד (Anyone)
 * את הכתובת שמתקבלת (מסתיימת ב- /exec) מדביקים בראש index.html בשורה  const API = '...'
 */
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
function t_(s) { s = s == null ? '' : String(s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function token_(n) { const a = 'abcdefghijkmnpqrstuvwxyz23456789'; let s = ''; for (let i = 0; i < n; i++) s += a[Math.floor(Math.random() * a.length)]; return s; }

function findClass_(id) {
  const sh = classesSheet_(), v = sh.getDataRange().getValues();
  for (let r = 1; r < v.length; r++) if (v[r][0] === id) return { sh, row: r + 1, v: v[r] };
  return null;
}
function classOut_(c, admin) {
  return { id: c.v[0], name: c.v[2], homeroom: c.v[3], data: JSON.parse(c.v[5] || '{}'), reports: JSON.parse(c.v[6] || '{}'), admin: admin };
}
function entries_(id) {
  const v = entriesSheet_().getDataRange().getValues(), out = [];
  for (let r = 1; r < v.length; r++) if (v[r][0] === id)
    out.push({ subjectId: v[r][1], studentKey: v[r][2], eval: v[r][3], strength: v[r][4], response: v[r][5], updated: v[r][6], seen: v[r][7] === true || v[r][7] === 'TRUE', by: v[r][8] });
  return out;
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
  if (q.action === 'create') {
    const id = token_(10), key = token_(14);
    classesSheet_().appendRow([id, key, t_(q.name), t_(q.homeroom), new Date(), JSON.stringify(q.data || {}), JSON.stringify(q.reports || {})]);
    return { id, key };
  }
  const c = findClass_(q.id);
  if (!c) return { error: 'הכיתה לא נמצאה. בדקו את הקישור.' };
  const admin = !!q.key && q.key === c.v[1];

  if (q.action === 'get') return { cls: classOut_(c, admin), entries: entries_(q.id) };

  if (q.action === 'save') { /* a subject teacher, or the homeroom teacher (subjectId "_hr") */
    if (q.subjectId === '_hr' && !admin) return { error: 'אין הרשאה' };
    const sh = entriesSheet_(), v = sh.getDataRange().getValues();
    let row = -1;
    for (let r = 1; r < v.length; r++) if (v[r][0] === q.id && v[r][1] === q.subjectId && v[r][2] === q.studentKey) { row = r + 1; break; }
    const now = new Date();
    if (row < 0) {
      const extra = !!(q.strength || q.response);
      sh.appendRow([q.id, q.subjectId, q.studentKey, t_(q.eval), t_(q.strength), t_(q.response), now, q.subjectId === '_hr' || !extra, q.by || '']);
    } else {
      const old = v[row - 1];
      const changed = (q.strength || '') !== old[4] || (q.response || '') !== old[5];
      const seen = q.subjectId === '_hr' ? true : (changed ? !(q.strength || q.response) : old[7]);
      sh.getRange(row, 4, 1, 6).setValues([[t_(q.eval), t_(q.strength), t_(q.response), now, seen, q.by || old[8]]]);
    }
    return { ok: true, at: now };
  }

  if (!admin) return { error: 'אין הרשאה' };

  if (q.action === 'review') { /* the homeroom teacher approves / edits / removes a teacher's addition */
    const sh = entriesSheet_(), v = sh.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) if (v[r][0] === q.id && v[r][1] === q.subjectId && v[r][2] === q.studentKey) {
      sh.getRange(r + 1, 5, 1, 2).setValues([[t_(q.strength ?? v[r][4]), t_(q.response ?? v[r][5])]]);
      sh.getRange(r + 1, 8).setValue(true);
      return { ok: true };
    }
    return { error: 'לא נמצא' };
  }
  if (q.action === 'update') {
    if (q.data) c.sh.getRange(c.row, 6).setValue(JSON.stringify(q.data));
    if (q.reports) c.sh.getRange(c.row, 7).setValue(JSON.stringify(q.reports));
    if (q.name != null) c.sh.getRange(c.row, 3).setValue(q.name);
    if (q.homeroom != null) c.sh.getRange(c.row, 4).setValue(q.homeroom);
    return { ok: true };
  }
  return { error: 'unknown action' };
}
