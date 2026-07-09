/**
 * O-No-Stop order webhook.
 *
 * Setup:
 * 1. Create a new Google Sheet (this will be your orders database).
 * 2. Extensions > Apps Script, delete any starter code, paste this file in.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployment URL and paste it into SHEET_WEBHOOK_URL in index.html.
 * 5. Reopen the Sheet once — onOpen() below creates the header row automatically,
 *    or it will be created on the first form submission if you skip this step.
 */

var SHEET_NAME = 'Orders';

function doPost(e) {
  var sheet = getOrdersSheet_();
  var data = parseRequestBody_(e);

  sheet.appendRow([
    new Date(),
    data.name || '',
    data.phone || '',
    data.address || '',
    data.city || '',
    data.pincode || '',
    data.pack || '',
    data.source || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'O-No-Stop order webhook is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  getOrdersSheet_();
}

function getOrdersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Name', 'Phone', 'Address', 'City', 'Pincode', 'Pack', 'Source']);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// The order form sends JSON as a text/plain body (to avoid a CORS preflight),
// so it always lands in e.postData.contents rather than e.parameter.
function parseRequestBody_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      // fall through to form-encoded parsing below
    }
  }
  return (e && e.parameter) || {};
}
