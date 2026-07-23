/**
 * Google Apps Script — מקבל שליחות מהטופס וכותב אותן ל-Google Sheet.
 *
 * התקנה (פעם אחת):
 * 1. פתח את ה-Sheet: "שאלון התאמה לצמיחה עסקית — תשובות"
 * 2. תפריט: Extensions → Apps Script
 * 3. מחק את הקוד שיש שם והדבק את כל הקובץ הזה
 * 4. לחץ Deploy → New deployment → Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Authorize / אשר את ההרשאות
 * 6. העתק את כתובת ה-Web app (מסתיימת ב-/exec)
 * 7. הדבק אותה בקובץ index.html במשתנה ENDPOINT
 */

// סדר העמודות בגיליון: [מפתח בטופס, כותרת בעברית]
var COLUMNS = [
  ['timestamp',    'תאריך ושעה'],
  ['ownerName',    'שם מלא'],
  ['phone',        'טלפון'],
  ['email',        'מייל'],
  ['bizName',      'שם העסק'],
  ['bizField',     'תחום העסק'],
  ['bizAbout',     'על העסק'],
  ['product',      'מוצר / שירות'],
  ['channel',      'אונליין / אופליין'],
  ['revenue',      'מחזור חודשי'],
  ['margin',       'אחוזי רווח'],
  ['audience',     'קהל עיקרי'],
  ['bestseller',   'הכי נמכר'],
  ['ads',          'איפה מפרסם'],
  ['adsOtherText', 'פרסום — אחר'],
  ['goal',         'מטרות'],
  ['working',      'מה עובד טוב'],
  ['notWorking',   'מה לא עובד טוב']
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // כותרות — נכתבות פעם אחת בשורה הראשונה
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
      sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var data = JSON.parse(e.postData.contents);
    data.timestamp = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm');

    var row = COLUMNS.map(function (c) {
      var v = data[c[0]];
      if (Array.isArray(v)) return v.join(', ');
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.appendRow(row);

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
