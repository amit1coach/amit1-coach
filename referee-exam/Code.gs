/**
 * מבחן שופט מתלמד – צד השרת (Google Apps Script)
 * מגיש את דף המבחן לנבחנים ושומר כל ניסיון בגיליון שאליו הסקריפט מחובר.
 * נבחן מזוהה לפי מספר תעודת הזהות שלו.
 *
 * גיליונות שנוצרים אוטומטית:
 *   "סיכום לפי נבחן" – שורה אחת לכל נבחן: עבר/טרם עבר, ציון, מספר ניסיונות. נבנה מחדש בכל שמירה.
 *   "כל הניסיונות"   – שורה לכל ניסיון, לפי סדר ההגשה.
 */

const PASS_MARK = 80;
const LOG_SHEET = 'כל הניסיונות';
const SUMMARY_SHEET = 'סיכום לפי נבחן';
const LOG_HEADERS = ['תאריך ושעה', 'ת.ז.', 'שם הנבחן', "ניסיון מס'", 'ציון', 'סטטוס',
  'נכונות', 'חלקיות', 'שגויות / ללא מענה', 'משך (דקות)', 'נושאים לחיזוק'];
const SUMMARY_HEADERS = ['ת.ז.', 'שם הנבחן', 'סטטוס', 'ציון במעבר', "עבר בניסיון מס'", 'תאריך מעבר',
  'סה"כ ניסיונות', 'הציון הגבוה ביותר', 'ניסיון אחרון'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('מבחן שופט מתלמד')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** להרצה מתוך העורך: יוצר את הגיליונות, מעדכן כותרות ומבקש את ההרשאות. */
function setup() {
  const log = getSheet_(LOG_SHEET, LOG_HEADERS);
  rebuildSummary_(log);
}

/** מספר הניסיון הבא של הנבחן (לתצוגה בתחילת המבחן). */
function getNextAttempt(id) {
  id = normId_(id);
  if (!validId_(id)) return 1;
  return countAttempts_(getSheet_(LOG_SHEET, LOG_HEADERS), id) + 1;
}

/** שומר ניסיון אחד ומעדכן את הסיכום. מחזיר את מספר הניסיון הרשמי. */
function saveResult(r) {
  const name = normName_(r && r.name);
  const id = normId_(r && r.id);
  if (!name) throw new Error('חסר שם נבחן');
  if (!validId_(id)) throw new Error('מספר ת.ז. לא תקין');
  const score = Math.max(0, Math.min(100, Math.round(Number(r.score) || 0)));
  const passed = score >= PASS_MARK;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = getSheet_(LOG_SHEET, LOG_HEADERS);
    const attempt = countAttempts_(log, id) + 1;
    log.appendRow([new Date(), "'" + id, name, attempt, score, passed ? 'עבר' : 'לא עבר',
      Number(r.full) || 0, Number(r.part) || 0, Number(r.zero) || 0,
      Number(r.minutes) || '', (r.weak || []).join(', ')]);
    log.getRange(log.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    rebuildSummary_(log);
    return { attempt: attempt, score: score, passed: passed };
  } finally {
    lock.releaseLock();
  }
}

function normName_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim();
}

/** ספרות בלבד, עם אפסים מובילים עד 9 ספרות (הגיליון עלול להשמיט אפסים מובילים). */
function normId_(id) {
  const d = String(id == null ? '' : id).replace(/\D/g, '');
  return d && d.length <= 9 ? ('000000000' + d).slice(-9) : d;
}

function validId_(id) {
  if (!/^\d{9}$/.test(id) || /^0+$/.test(id)) return false;
  let sum = 0;
  for (let i = 0; i < 9; i++) {
    let x = Number(id[i]) * (i % 2 ? 2 : 1);
    sum += x > 9 ? x - 9 : x;
  }
  return sum % 10 === 0;
}

function getSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.setRightToLeft(true);
  }
  if (headers) {
    const head = sh.getRange(1, 1, 1, headers.length);
    if (head.getValues()[0].join('|') !== headers.join('|')) {
      head.setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function countAttempts_(log, id) {
  const last = log.getLastRow();
  if (last < 2) return 0;
  return log.getRange(2, 2, last - 1, 1).getValues()
    .filter(function (row) { return normId_(row[0]) === id; }).length;
}

function rebuildSummary_(log) {
  const last = log.getLastRow();
  const rows = last < 2 ? [] : log.getRange(2, 1, last - 1, 6).getValues();
  const byId = {};
  rows.forEach(function (row) {
    const when = row[0], id = normId_(row[1]), name = normName_(row[2]);
    const attempt = row[3], score = Number(row[4]) || 0, status = row[5];
    if (!validId_(id)) return;
    const s = byId[id] || (byId[id] = {
      id: id, name: name, attempts: 0, best: 0, lastWhen: '', passedWhen: '', passedScore: '', passedAttempt: ''
    });
    s.name = name || s.name;   // השם האחרון שהנבחן הקליד
    s.attempts++;
    s.best = Math.max(s.best, score);
    s.lastWhen = when;
    if (status === 'עבר' && !s.passedWhen) {
      s.passedWhen = when; s.passedScore = score; s.passedAttempt = attempt;
    }
  });

  const out = Object.keys(byId).map(function (k) { return byId[k]; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'he'); })
    .map(function (s) {
      return ["'" + s.id, s.name, s.passedWhen ? 'עבר' : 'טרם עבר', s.passedScore, s.passedAttempt,
        s.passedWhen, s.attempts, s.best, s.lastWhen];
    });

  const sum = getSheet_(SUMMARY_SHEET, null);
  sum.clear();
  sum.setRightToLeft(true);
  sum.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]).setFontWeight('bold');
  sum.setFrozenRows(1);
  if (out.length) {
    sum.getRange(2, 1, out.length, SUMMARY_HEADERS.length).setValues(out);
    sum.getRange(2, 6, out.length, 1).setNumberFormat('dd/MM/yyyy');
    sum.getRange(2, 9, out.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }
  sum.autoResizeColumns(1, SUMMARY_HEADERS.length);
}
