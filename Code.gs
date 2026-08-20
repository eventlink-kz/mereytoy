const SHEET_NAME = 'RSVP';
const ADMIN_KEY = '8080';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || '');
  const callback = String(p.callback || '');

  if (action === 'ping') {
    return output({ ok: true, service: 'rsvp', time: new Date().toISOString() }, callback);
  }

  if (action === 'check') {
    const id = String(p.id || '');
    if (!id) return output({ ok: false, found: false, error: 'Missing id' }, callback);
    return output({ ok: true, found: hasId(id) }, callback);
  }

  if (action === 'list') {
    if (String(p.key || '') !== ADMIN_KEY) {
      return output({ ok: false, error: 'Unauthorized' }, callback);
    }
    const sh = getSheet();
    const values = sh.getDataRange().getValues();
    const entries = values.slice(1).filter(r => r[0]).map(r => ({
      id: String(r[0] || ''),
      createdAt: String(r[1] || ''),
      name: String(r[2] || ''),
      attendance: String(r[3] || ''),
      guests: Number(r[4] || 0),
      phone: String(r[5] || ''),
      comment: String(r[6] || '')
    })).reverse();
    return output({ ok: true, entries: entries }, callback);
  }

  return output({ ok: false, error: 'Unknown action' }, callback);
}

function doPost(e) {
  try {
    let body = {};
    if (e && e.parameter && Object.keys(e.parameter).length) {
      body = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents || '{}');
      if (body.entry) body = Object.assign({ action: body.action }, body.entry);
    }

    if (String(body.action || '') !== 'submit') {
      return output({ ok: false, error: 'Unknown action' });
    }

    const entry = normalizeEntry(body);
    if (!entry.name || !['yes', 'no'].includes(entry.attendance)) {
      return output({ ok: false, error: 'Validation error' });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (!hasId(entry.id)) {
        getSheet().appendRow([
          safeCell(entry.id),
          safeCell(entry.createdAt),
          safeCell(entry.name),
          entry.attendance,
          entry.guests,
          safeCell(entry.phone),
          safeCell(entry.comment)
        ]);
      }
    } finally {
      lock.releaseLock();
    }

    return output({ ok: true, id: entry.id });
  } catch (err) {
    return output({ ok: false, error: String(err) });
  }
}

function normalizeEntry(x) {
  const attendance = String(x.attendance || '');
  return {
    id: String(x.id || Utilities.getUuid()).slice(0, 100),
    createdAt: String(x.createdAt || new Date().toISOString()).slice(0, 80),
    name: String(x.name || '').trim().slice(0, 80),
    attendance: attendance,
    guests: attendance === 'yes' ? Math.max(1, Math.min(20, Number(x.guests || 1))) : 0,
    phone: String(x.phone || '').trim().slice(0, 30),
    comment: String(x.comment || '').trim().slice(0, 300)
  };
}

function hasId(id) {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return false;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
  return ids.indexOf(String(id)) !== -1;
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
  value = String(value == null ? '' : value);
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function output(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
