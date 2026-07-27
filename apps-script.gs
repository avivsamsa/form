/**
 * Google Apps Script — מקבל שליחות מהטופס וכותב אותן ל-Google Sheet.
 *
 * מבנה הטבלה: לאורך.
 *   עמודה A = השאלות (שורה לכל שאלה).
 *   כל שליחה חדשה נכנסת בעמודה הבאה (B, C, D ...), התשובה ליד השאלה שלה.
 *   ככה קוראים "שאלה | תשובה" באותה שורה.
 *
 * ── התקנה / עדכון ──
 * 1. פתח את ה-Sheet: "שאלון התאמה לצמיחה עסקית — תשובות"
 * 2. Extensions → Apps Script
 * 3. מחק הכל, הדבק את כל הקובץ הזה, ושמור (אייקון הדיסקט).
 * 4. Deploy → Manage deployments → בחר את הפריסה הקיימת → עיפרון (Edit)
 *      → Version: New version → Deploy.
 *    (ככה הכתובת /exec נשארת אותו דבר.)
 *    אם אין עדיין פריסה: Deploy → New deployment → Web app.
 * 5. חובה! הגדרות הפריסה:
 *      - Execute as: Me
 *      - Who has access: Anyone   ← בלי זה שום שליחה לא נכנסת
 * 6. אשר הרשאות אם מבקשים (Advanced → Go to project → Allow).
 *
 * ── בדיקה מהירה ──
 * פתח את כתובת ה-/exec בדפדפן. אם מופיע הכיתוב "החיבור לשאלון פעיל ✓" —
 * הפריסה ציבורית ותקינה. אם מופיע מסך התחברות/הרשאה של גוגל — הגישה
 * לא Anyone, חזור לשלב 5.
 */

// [מפתח בטופס, טקסט השאלה כפי שיוצג בעמודה A]
var COLUMNS = [
  ['timestamp',    'מועד מילוי'],
  ['ownerName',    'שם מלא'],
  ['phone',        'טלפון'],
  ['email',        'מייל'],
  ['bizName',      'שם העסק'],
  ['bizField',     'תחום העסק'],
  ['bizAbout',     'ספר על העסק בכמה מילים'],
  ['product',      'מה המוצר או השירות שנמכר'],
  ['channel',      'אונליין / אופליין / גם וגם'],
  ['revenue',      'מחזור חודשי משוער'],
  ['margin',       'אחוזי רווח משוערים'],
  ['audience',     'לאיזה קהל מוכר הכי הרבה'],
  ['bestseller',   'מה הדבר הכי נמכר'],
  ['ads',          'איפה מפרסם את העסק'],
  ['adsOtherText', 'פרסום אחר (פירוט)'],
  ['goal',         'לאן היית רוצה לקחת את העסק'],
  ['working',      'מה כן עובד טוב בעסק'],
  ['notWorking',   'מה לא עובד טוב בעסק']
];

function doGet() {
  return ContentService
    .createTextOutput('החיבור לשאלון פעיל ✓  — הטופס יכתוב לכאן את השליחות.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // עמודת השאלות (A) — נכתבת פעם אחת
    if (sheet.getRange(1, 1).getValue() === '') {
      var labels = COLUMNS.map(function (c) { return [c[1]]; });
      sheet.getRange(1, 1, COLUMNS.length, 1).setValues(labels);
      sheet.getRange(1, 1, COLUMNS.length, 1).setFontWeight('bold');
      sheet.setColumnWidth(1, 220);
      sheet.setFrozenColumns(1);
    }

    var data = JSON.parse(e.postData.contents);
    data.timestamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm');

    var values = COLUMNS.map(function (c) {
      var v = data[c[0]];
      if (Array.isArray(v)) v = v.join(', ');
      return [ (v === undefined || v === null) ? '' : v ];
    });

    // עמודה חדשה לשליחה הזו
    var col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col, COLUMNS.length, 1).setValues(values);
    sheet.setColumnWidth(col, 280);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
