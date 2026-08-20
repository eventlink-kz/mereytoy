const SHEET_NAME = 'RSVP';
const ADMIN_KEY = '8080';

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '');
  const callback = String((e && e.parameter && e.parameter.callback) || '');

  if (action === 'ping') {
    return output({ ok: true }, callback);
  }

  if (action === 'list') {
    if (String(e.parameter.key || '') !== ADMIN_KEY) {
      return output({ ok: false, error: 'Unauthorized' }, callback);
    }

    const sh = getSheet();
    const values = sh.getDataRange().getValues();

    const entries = values.slice(1)
      .filter(r => r[0])
      .map(r => ({
        id: String(r[0] || ''),
        createdAt: String(r[1] || ''),
        name: String(r[2] || ''),
        attendance: String(r[3] || ''),
        guests: Number(r[4] || 0),
        phone: String(r[5] || ''),
        comment: String(r[6] || '')
      }))
      .reverse();

    return output({ ok: true, entries: entries }, callback);
  }

  return output({ ok: false, error: 'Unknown action' }, callback);
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const x = body.entry || {};

    if (body.action !== 'submit' || !x.name || !['yes', 'no'].includes(x.attendance)) {
      return output({ ok: false, error: 'Validation error' });
    }

    getSheet().appendRow([
      safeCell(String(x.id || Utilities.getUuid())),
      safeCell(String(x.createdAt || new Date().toISOString())),
      safeCell(String(x.name).slice(0, 80)),
      x.attendance,
      x.attendance === 'yes' ? Math.max(1, Math.min(20, Number(x.guests || 1))) : 0,
      safeCell(String(x.phone || '').slice(0, 30)),
      safeCell(String(x.comment || '').slice(0, 300))
    ]);

    return output({ ok: true });
  } catch (err) {
    return output({ ok: false, error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet not found. Open Apps Script from Extensions > Apps Script inside the RSVP Google Sheet.');

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['id', 'createdAt', 'name', 'attendance', 'guests', 'phone', 'comment']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function safeCell(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function output(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}