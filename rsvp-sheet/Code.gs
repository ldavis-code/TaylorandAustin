/**
 * Wedding RSVP → Google Sheet
 *
 * Paste this whole file into a Google Sheet's Apps Script editor
 * (Extensions → Apps Script), then deploy it as a web app:
 *   Deploy → New deployment → type "Web app"
 *   Execute as: Me    Who has access: Anyone
 * Copy the web app URL (ends in /exec) into SHEET_URL in rsvp.html.
 *
 * Each RSVP sent from the website becomes one new row.
 */

var SHEET_NAME = 'RSVPs';

var COLUMNS = [
  'Submitted',
  'Name',
  'Email',
  'Attending',
  'Number of guests',
  'Guest 1 entrée',
  'Guest 2 name', 'Guest 2 entrée',
  'Guest 3 name', 'Guest 3 entrée',
  'Guest 4 name', 'Guest 4 entrée',
  'Dietary restrictions',
  'Note for the couple'
];

// Form field name for each column (after the timestamp)
var FIELDS = [
  'name',
  'email',
  'attending',
  'guests',
  'Guest 1 entrée',
  'Guest 2 name', 'Guest 2 entrée',
  'Guest 3 name', 'Guest 3 entrée',
  'Guest 4 name', 'Guest 4 entrée',
  'dietary',
  'note'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = (e && e.parameter) || {};

    // Hidden "_honey" field is only filled in by spam bots
    if (p._honey) return reply({ result: 'success' });

    var sheet = getSheet();
    var row = [new Date()].concat(FIELDS.map(function (f) {
      var v = p[f] || '';
      // Stop values like "=..." from being treated as spreadsheet formulas
      return /^[=+\-@]/.test(v) ? "'" + v : v;
    }));
    sheet.appendRow(row);

    return reply({ result: 'success' });
  } catch (err) {
    return reply({ result: 'error', error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function reply(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
