/**
 * מבחן שופט מתלמד – צד השרת (Google Apps Script)
 * מגיש את דף המבחן לנבחנים ושומר כל ניסיון בגיליון שאליו הסקריפט מחובר.
 *
 * כל ניסיון נרשם בשורה ברגע שהנבחן מתחיל (סטטוס "לא הוגש"),
 * והשורה מתעדכנת בהגשה: שעת סיום, משך, ציון. הזמנים נלקחים מהשעון של Google.
 *
 * גיליונות שנוצרים אוטומטית:
 *   "סיכום לפי נבחן" – שורה אחת לכל נבחן: עבר/טרם עבר, ציון, מספר ניסיונות. נבנה מחדש בכל שמירה.
 *   "כל הניסיונות"   – שורה לכל ניסיון, לפי סדר ההתחלה.
 *   "תרגולים"        – אותו מבנה, לתרגולים בלבד (הקישור עם ?mode=practice). לא נכנסים לסיכום.
 */

const PASS_MARK = 80;
const TIME_LIMIT_MIN = 20;      // חייב להתאים ל-TIME_LIMIT_MIN שב-Index.html
const LATE_GRACE_SEC = 60;      // מרווח לעיכובי רשת לפני שהגשה נחשבת באיחור
const LOG_SHEET = 'כל הניסיונות';
const SUMMARY_SHEET = 'סיכום לפי נבחן';
const LOG_HEADERS = ['שם הנבחן', "ניסיון מס'", 'שעת התחלה', 'שעת סיום', 'משך (דק:שנ)', 'ציון', 'סטטוס',
  'אופן ההגשה', 'נכונות', 'חלקיות', 'שגויות / ללא מענה', 'נושאים לחיזוק', 'מזהה ניסיון'];
const COL = { name: 1, attempt: 2, start: 3, end: 4, duration: 5, score: 6, status: 7, token: 13 };
const SUMMARY_HEADERS = ['שם הנבחן', 'סטטוס', 'ציון במעבר', "עבר בניסיון מס'", 'תאריך מעבר',
  'סה"כ ניסיונות', 'הציון הגבוה ביותר', 'ניסיון אחרון'];
const NOT_SUBMITTED = 'לא הוגש';
const PRACTICE_SHEET = 'תרגולים';
const PRACTICE_HEADERS = LOG_HEADERS.slice(0, 1).concat(["תרגול מס'"], LOG_HEADERS.slice(2), ['זמן שנבחר (דקות)']);
const COL_LIMIT = 14;                 // עמודת "זמן שנבחר" – רק בגיליון התרגולים
const PRACTICE_LIMITS = [20, 25, 30]; // חייב להתאים ל-PRACTICE_LIMITS שב-Index.html
const NO_LIMIT = 'ללא הגבלה';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('מבחן שופט מתלמד')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** להרצה מתוך העורך: יוצר את הגיליונות, ממיר נתונים מגרסה קודמת ומבקש את ההרשאות. */
function setup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    rebuildSummary_(getLog_());
    getPractice_();
  } finally {
    lock.releaseLock();
  }
}

function getPractice_() {
  const sh = getSheet_(PRACTICE_SHEET, PRACTICE_HEADERS);
  if (!sh.isColumnHiddenByUser(COL.token)) sh.hideColumns(COL.token);
  return sh;
}

function isPractice_(x) {
  return !!(x && x.mode === 'practice');
}

/** זמן התרגול שנבחר: 20/25/30, או '' לללא הגבלה. */
function practiceLimit_(value) {
  const n = Number(value);
  return PRACTICE_LIMITS.indexOf(n) >= 0 ? n : '';
}

/** גיליון הניסיונות. אם הוא עדיין במבנה של גרסה קודמת, ממיר אותו קודם (עם גיבוי). */
function getLog_() {
  const old = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
  if (old && old.getRange(1, 1).getValue() === 'תאריך ושעה') migrateOldLog_(old);
  const log = getSheet_(LOG_SHEET, LOG_HEADERS);
  if (!log.isColumnHiddenByUser(COL.token)) log.hideColumns(COL.token);
  return log;
}

/**
 * ממיר שורות מהגרסה שלפני הגבלת הזמן. בגרסה ההיא "תאריך ושעה" היה זמן ההגשה,
 * ו"משך (דקות)" היה מעוגל לדקות, לכן שעת ההתחלה המשוחזרת מדויקת רק לדקה.
 */
function migrateOldLog_(sh) {
  const ss = sh.getParent();
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH-mm');
  sh.copyTo(ss).setName('גיבוי – ' + LOG_SHEET + ' ' + stamp);

  const width = sh.getLastColumn(), last = sh.getLastRow();
  const head = sh.getRange(1, 1, 1, width).getValues()[0];
  const off = head[1] === 'ת.ז.' ? 1 : 0;   // גרסה שבה הייתה עמודת ת.ז.
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, width).getValues();
  const out = rows.filter(function (r) { return normName_(r[1 + off]); }).map(function (r) {
    const end = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const min = Number(r[8 + off]) || 0;
    return [normName_(r[1 + off]), r[2 + off], new Date(end.getTime() - min * 60000), end, min * 60 / 86400,
      r[3 + off], r[4 + off], 'לפני הגבלת הזמן', r[5 + off], r[6 + off], r[7 + off], r[9 + off] || '',
      Utilities.getUuid()];
  });

  sh.clear();
  sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (out.length) {
    sh.getRange(2, 1, out.length, LOG_HEADERS.length).setValues(out);
    sh.getRange(2, COL.start, out.length, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getRange(2, COL.duration, out.length, 1).setNumberFormat('[m]:ss');
  }
}

/**
 * פותח ניסיון חדש: רושם שורה עם שעת ההתחלה ומחזיר מזהה ומספר ניסיון.
 * opts: { mode: 'practice', limitMin: 20|25|30|null } בתרגול; בלי opts – מבחן.
 */
function startAttempt(name, opts) {
  name = normName_(name);
  if (!name) throw new Error('חסר שם נבחן');
  const practice = isPractice_(opts);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = practice ? getPractice_() : getLog_();
    const attempt = countAttempts_(log, name) + 1;
    const token = Utilities.getUuid();
    const row = [name, attempt, new Date(), '', '', '', NOT_SUBMITTED, '', '', '', '', '', token];
    const limit = practice ? practiceLimit_(opts.limitMin) : TIME_LIMIT_MIN;
    if (practice) row.push(limit === '' ? NO_LIMIT : limit);
    log.appendRow(row);
    formatRow_(log, log.getLastRow());
    if (!practice) rebuildSummary_(log);
    return { token: token, attempt: attempt, limitMin: limit === '' ? null : limit };
  } finally {
    lock.releaseLock();
  }
}

/** שומר הגשה: מעדכן את שורת הניסיון ואת הסיכום. */
function saveResult(r) {
  const name = normName_(r && r.name);
  if (!name) throw new Error('חסר שם נבחן');
  const score = Math.max(0, Math.min(100, Math.round(Number(r.score) || 0)));
  const passed = score >= PASS_MARK;
  const end = new Date();
  const practice = isPractice_(r);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const log = practice ? getPractice_() : getLog_();
    let row = findRow_(log, r.token);
    let attempt, start, limit;
    if (row) {
      const values = log.getRange(row, 1, 1, practice ? PRACTICE_HEADERS.length : LOG_HEADERS.length).getValues()[0];
      if (values[COL.status - 1] !== NOT_SUBMITTED) {   // כבר נשמר (למשל לחיצה כפולה)
        return { attempt: values[COL.attempt - 1], score: values[COL.score - 1], passed: values[COL.status - 1] === 'עבר' };
      }
      attempt = values[COL.attempt - 1];
      start = values[COL.start - 1];
      limit = practice ? practiceLimit_(values[COL_LIMIT - 1]) : TIME_LIMIT_MIN;
    } else {
      // לא נמצאה שורת התחלה (למשל אם הפתיחה נכשלה): יוצרים שורה לפי הזמן שמדד הדפדפן
      attempt = countAttempts_(log, name) + 1;
      const elapsed = Math.max(0, Number(r.elapsedSec) || 0);
      start = new Date(end.getTime() - elapsed * 1000);
      limit = practice ? practiceLimit_(r.limitMin) : TIME_LIMIT_MIN;
      log.appendRow([name, attempt, start]);
      row = log.getLastRow();
      if (practice) log.getRange(row, COL_LIMIT).setValue(limit === '' ? NO_LIMIT : limit);
    }

    const durationSec = Math.max(0, Math.round((end - start) / 1000));
    let how;
    if (limit === '') {
      how = 'ללא הגבלת זמן';
    } else {
      const lateSec = durationSec - limit * 60;
      how = lateSec > LATE_GRACE_SEC ? 'באיחור של ' + formatMinSec_(lateSec) + ' דק\''
        : (r.auto ? 'אוטומטית בתום הזמן' : 'בזמן');
    }

    log.getRange(row, 1, 1, LOG_HEADERS.length).setValues([[
      name, attempt, start, end, durationSec / 86400, score, passed ? 'עבר' : 'לא עבר', how,
      Number(r.full) || 0, Number(r.part) || 0, Number(r.zero) || 0, (r.weak || []).join(', '),
      r.token || Utilities.getUuid()
    ]]);
    formatRow_(log, row);
    if (!practice) rebuildSummary_(log);
    return { attempt: attempt, score: score, passed: passed, durationSec: durationSec };
  } finally {
    lock.releaseLock();
  }
}

function normName_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim();
}

function formatMinSec_(sec) {
  const m = Math.floor(sec / 60), s = sec % 60;
  return m + ':' + (s < 10 ? '0' : '') + s;
}

function formatRow_(log, row) {
  log.getRange(row, COL.start, 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  log.getRange(row, COL.duration).setNumberFormat('[m]:ss');
}

function getSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.setRightToLeft(true);
  }
  if (headers) {
    const width = Math.max(headers.length, sh.getLastColumn());
    const current = sh.getRange(1, 1, 1, width).getValues()[0];
    const wanted = headers.concat(new Array(width - headers.length).fill(''));
    if (current.join('|') !== wanted.join('|')) {
      sh.getRange(1, 1, 1, width).setValues([wanted]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function findRow_(log, token) {
  if (!token || log.getLastRow() < 2) return 0;
  const hit = log.getRange(2, COL.token, log.getLastRow() - 1, 1)
    .createTextFinder(String(token)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function countAttempts_(log, name) {
  const last = log.getLastRow();
  if (last < 2) return 0;
  const key = name.toLowerCase();
  return log.getRange(2, COL.name, last - 1, 1).getValues()
    .filter(function (row) { return normName_(row[0]).toLowerCase() === key; }).length;
}

function rebuildSummary_(log) {
  const last = log.getLastRow();
  const rows = last < 2 ? [] : log.getRange(2, 1, last - 1, COL.status).getValues();
  const byName = {};
  rows.forEach(function (row) {
    const name = normName_(row[COL.name - 1]);
    if (!name) return;
    const attempt = row[COL.attempt - 1], start = row[COL.start - 1], end = row[COL.end - 1];
    const score = Number(row[COL.score - 1]) || 0, status = row[COL.status - 1];
    const key = name.toLowerCase();
    const s = byName[key] || (byName[key] = {
      name: name, attempts: 0, best: '', lastWhen: '', passedWhen: '', passedScore: '', passedAttempt: ''
    });
    s.attempts++;
    s.lastWhen = start;
    if (status !== NOT_SUBMITTED) s.best = s.best === '' ? score : Math.max(s.best, score);
    if (status === 'עבר' && !s.passedWhen) {
      s.passedWhen = end || start; s.passedScore = score; s.passedAttempt = attempt;
    }
  });

  const out = Object.keys(byName).map(function (k) { return byName[k]; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'he'); })
    .map(function (s) {
      return [s.name, s.passedWhen ? 'עבר' : 'טרם עבר', s.passedScore, s.passedAttempt,
        s.passedWhen, s.attempts, s.best, s.lastWhen];
    });

  const sum = getSheet_(SUMMARY_SHEET, null);
  sum.clear();
  sum.setRightToLeft(true);
  sum.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]).setFontWeight('bold');
  sum.setFrozenRows(1);
  if (out.length) {
    sum.getRange(2, 1, out.length, SUMMARY_HEADERS.length).setValues(out);
    sum.getRange(2, 5, out.length, 1).setNumberFormat('dd/MM/yyyy');
    sum.getRange(2, 8, out.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }
  sum.autoResizeColumns(1, SUMMARY_HEADERS.length);
}
