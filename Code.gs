/**
 * Subscriber Lookup - Google Sheet backend
 *
 * Creates/uses two tabs in the spreadsheet this script is attached to:
 *   "Activity Log"  - one row per event (login, search, view...)
 *   "For Payment"   - one row per subscriber (latest action wins)
 *
 * The page POSTs JSON (as text/plain) and reads with:
 *   GET ?read=1    -> { rows: [ activity log rows ] }
 *   GET ?read=pay  -> { rows: [ for payment rows ] }
 */

var LOG_SHEET = 'Activity Log';
var PAY_SHEET = 'For Payment';
var LOG_HEAD = ['time', 'email', 'code', 'name', 'action', 'where', 'search', 'device', 'visitor'];
var PAY_HEAD = ['key', 'name', 'account', 'agent', 'label', 'action', 'date', 'remarks', 'by', 'code', 'at'];
var MAX_LOG_ROWS = 5000; // rows returned to the page

function getSheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.type === 'pay') {
      savePay_(d);
    } else {
      saveLog_(d);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveLog_(d) {
  var sh = getSheet_(LOG_SHEET, LOG_HEAD);
  sh.appendRow(LOG_HEAD.map(function (k) { return d[k] == null ? '' : String(d[k]); }));
}

function savePay_(d) {
  var sh = getSheet_(PAY_SHEET, PAY_HEAD);
  var row = PAY_HEAD.map(function (k) { return d[k] == null ? '' : String(d[k]); });
  var last = sh.getLastRow();
  if (last > 1) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(d.key)) {
        sh.getRange(i + 2, 1, 1, PAY_HEAD.length).setValues([row]); // update existing subscriber
        return;
      }
    }
  }
  sh.appendRow(row);
}

function doGet(e) {
  var mode = e && e.parameter && e.parameter.read;
  if (mode === 'pay') return json_({ rows: readRows_(PAY_SHEET, PAY_HEAD, 0) });
  if (mode) return json_({ rows: readRows_(LOG_SHEET, LOG_HEAD, MAX_LOG_ROWS) });
  return json_({ ok: true, message: 'Subscriber Lookup backend is running' });
}

function readRows_(name, head, limit) {
  var sh = getSheet_(name, head);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = 2;
  if (limit && last - 1 > limit) start = last - limit + 1;
  var vals = sh.getRange(start, 1, last - start + 1, head.length).getDisplayValues();
  return vals.map(function (r) {
    var o = {};
    head.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}
