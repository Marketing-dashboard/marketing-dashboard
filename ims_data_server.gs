/**
 * IMS Campaign Data Server — Google Apps Script Web App
 *
 * HOW TO DEPLOY:
 *  1. Open Google Sheets → Extensions → Apps Script
 *  2. Paste this entire file (replace any existing code)
 *  3. Click "Deploy" → "New deployment"
 *  4. Type: Web app
 *  5. Execute as: Me  (your Google account — gives access to the private sheet)
 *  6. Who has access: Anyone
 *  7. Click "Deploy" → Copy the Web app URL
 *  8. Paste that URL into the dashboard setup screen
 */

const SHEET_ID   = '10KKt1iJlAgiTlFj1tiPkuHuAZ72tvVmxm0Yv0BhiRvw';
const SHEET_NAME = 'Raw';

function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return jsonOut({ status: 'error', message: 'Sheet "' + SHEET_NAME + '" not found.' });
    }

    const values  = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return jsonOut({ status: 'ok', rows: [], timestamp: new Date().toISOString() });
    }

    const headers = values[0].map(h => String(h).trim());
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    const rows = values.slice(1)
      .filter(row => row.some(v => v !== '' && v !== null && v !== undefined))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          const v = row[i];
          // Serialize dates as YYYY-MM-DD in the sheet's timezone (avoids UTC midnight shift)
          obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v;
        });
        return obj;
      });

    const result = { status: 'ok', rows: rows, timestamp: new Date().toISOString() };
    return respond(e, result);

  } catch (err) {
    return respond(e, { status: 'error', message: err.message });
  }
}

// Supports both plain JSON and JSONP (callback= param) so fetch() and <script> both work.
// Workspace GAS URLs (/a/macros/domain/) block cross-origin fetch() due to auth redirects;
// JSONP via <script> tag bypasses CORS entirely.
function respond(e, obj) {
  const output = JSON.stringify(obj);
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + output + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}
