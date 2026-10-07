/* BWMS Picking - shared store.
   Paste ALL of this into Extensions > Apps Script of your Google Sheet. */

const SHEET_JOBS = 'Jobs';
const SHEET_LINES = 'Lines';
const JOB_HEAD = ['Pick list', 'Lines', 'Total qty', 'UOM', 'Picker', 'Status', 'Issued', 'Start', 'End', 'Lines done', 'Qty picked', 'Short', 'Last update'];
const LINE_HEAD = ['Saved at', 'Pick list', 'Picker', 'Line', 'Location', 'Pallet', 'Product code', 'Task qty', 'Picked', 'Status', 'Alt location', 'Alt qty', 'Alt status', 'Total picked', 'Seconds taken'];

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}

/* Run this once after pasting. It creates the two tabs. */
function setup() {
  sheet_(SHEET_JOBS, JOB_HEAD);
  sheet_(SHEET_LINES, LINE_HEAD);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function findRow_(sh, p) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === String(p)) return i + 2;
  return 0;
}

/* The picking page sends here: issue, start, line, finish. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.p || !d.action) return out_({ ok: false, error: 'Missing pick list or action' });
    const now = new Date();
    const jobs = sheet_(SHEET_JOBS, JOB_HEAD);
    let r = findRow_(jobs, d.p);
    if (!r) {
      jobs.appendRow(["'" + d.p, d.lines || '', d.total || '', d.uom || '', '', 'Not started', now, '', '', 0, 0, '', now]);
      r = jobs.getLastRow();
    }
    if (d.action === 'issue') {
      jobs.getRange(r, 2, 1, 3).setValues([[d.lines || '', d.total || '', d.uom || '']]);
    } else if (d.action === 'start') {
      jobs.getRange(r, 5, 1, 8).setValues([[d.picker || '', 'Picking', jobs.getRange(r, 7).getValue(), now, '', 0, 0, '']]);
    } else if (d.action === 'line') {
      const L = d.line || {};
      sheet_(SHEET_LINES, LINE_HEAD).appendRow([now, "'" + d.p, d.picker || '', L.n, L.loc, L.pallet, "'" + (L.sku || ''), L.task, L.q, L.st, L.alt || '', L.alt ? L.aq : '', L.ast || '', L.t, L.sec]);
      jobs.getRange(r, 10, 1, 2).setValues([[d.done || 0, d.qty || 0]]);
    } else if (d.action === 'finish') {
      jobs.getRange(r, 6).setValue('Completed');
      jobs.getRange(r, 9, 1, 4).setValues([[now, d.done || 0, d.qty || 0, d.short || 0]]);
    } else {
      return out_({ ok: false, error: 'Unknown action' });
    }
    jobs.getRange(r, 13).setValue(now);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* The dashboard reads here. */
function doGet() {
  const sh = sheet_(SHEET_JOBS, JOB_HEAD);
  const n = sh.getLastRow();
  const ms = v => (v instanceof Date ? v.getTime() : null);
  const jobs = [];
  if (n > 1) {
    sh.getRange(2, 1, n - 1, JOB_HEAD.length).getValues().forEach(function (v) {
      if (v[0] === '') return;
      jobs.push({ p: String(v[0]), lines: Number(v[1]) || 0, total: Number(v[2]) || 0, uom: String(v[3] || 'PCS'), picker: String(v[4] || ''),
        status: String(v[5] || 'Not started'), issued: ms(v[6]), start: ms(v[7]), end: ms(v[8]),
        done: Number(v[9]) || 0, qty: Number(v[10]) || 0, short: Number(v[11]) || 0, updated: ms(v[12]) });
    });
  }
  return out_({ ok: true, now: Date.now(), jobs: jobs });
}
