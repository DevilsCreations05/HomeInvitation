/****************************************************************
 * Gruha Pravesham — RSVP backend (Google Apps Script)
 *
 * SETUP (one time):
 *  1. Create a new Google Sheet (this is your "excel").
 *  2. In the Sheet: Extensions ▸ Apps Script.
 *  3. Delete any sample code, paste THIS whole file, click Save.
 *  4. Deploy ▸ New deployment ▸ (gear) Web app.
 *       - Description : RSVP
 *       - Execute as  : Me
 *       - Who has access : Anyone
 *     Click Deploy, authorise access when prompted.
 *  5. Copy the Web app URL (it ends with /exec).
 *  6. Paste that URL into RSVP_API in BOTH index.html and admin.html.
 *
 * The "RSVPs" tab and its header row are created automatically on
 * the first submission — you don't need to make columns yourself.
 ****************************************************************/

var SHEET_NAME = 'RSVPs';
var ADMIN_PWD  = 'kalyan950';   // must match the password used by admin.html

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'rsvp') {
      var sheet = getSheet_();
      sheet.appendRow([
        new Date(),
        clip_(p.name, 100),
        (String(p.attending).toLowerCase() === 'no') ? 'No' : 'Yes',
        Math.max(0, Math.min(50, parseInt(p.guests, 10) || 0)),
        clip_(p.phone, 25),
        clip_(p.message, 600)
      ]);
      out = { ok: true };

    } else if (p.action === 'list') {
      out = (p.pwd === ADMIN_PWD)
        ? { ok: true, rows: readRows_() }
        : { ok: false, error: 'unauthorized' };

    } else {
      out = { ok: true, message: 'Gruha Pravesham RSVP API is live.' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return reply_(out, p.callback);
}

/* Return JSONP when a ?callback= is supplied (used by the website),
   otherwise plain JSON. */
function reply_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Name', 'Attending', 'Guests', 'Phone', 'Message']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1').setFontWeight('bold');
  }
  return sheet;
}

function readRows_() {
  var values = getSheet_().getDataRange().getValues();
  return values.slice(1).map(function (r) {
    return {
      time:      r[0],
      name:      r[1],
      attending: r[2],
      guests:    r[3],
      phone:     r[4],
      message:   r[5]
    };
  });
}

function clip_(v, n) {
  return String(v == null ? '' : v).slice(0, n);
}
