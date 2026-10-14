/*
  sir.nxs tracker (Google Apps Script)
  1) Change ADMIN_PASSWORD below to your own password.
  2) Deploy > New deployment > type: Web app
     Execute as: Me      Who has access: Anyone
  3) Copy the Web app URL (ends with /exec) into:
     - index.html:  const API="PASTE_URL_HERE";
     - the program: Admin button > setapi PASTE_URL_HERE
  Stores only: time, event, random browser id, device name, language, screen size, referrer.
  No IP address and no names are saved.
*/
var ADMIN_PASSWORD = 'CHANGE-THIS-PASSWORD';
var SHEET_NAME = 'events';
var TYPES = { visit: 1, download_click: 1, download_ok: 1 };

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function text_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['time', 'type', 'visitor', 'device', 'language', 'screen', 'referrer']);
  }
  return sh;
}
// visitors control these strings: cut them down and stop spreadsheet formulas
function clean_(v, n) {
  var s = String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').slice(0, n);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  return out_({ ok: true, service: 'sir.nxs tracker' });
}

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false }); }
  if (!d || typeof d !== 'object') return out_({ ok: false });

  if (d.action === 'stats' || d.action === 'report') {
    var asText = d.action === 'report';
    if (ADMIN_PASSWORD === 'CHANGE-THIS-PASSWORD' || String(d.pw) !== ADMIN_PASSWORD) {
      return asText
        ? text_('ERROR: wrong password (or ADMIN_PASSWORD was not changed in tracker.gs).')
        : out_({ ok: false, error: 'wrong password' });
    }
    var s = stats_();
    return asText ? text_(report_(s)) : out_(s);
  }

  if (Object.prototype.hasOwnProperty.call(TYPES, d.type)) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet_().appendRow([
        new Date(), d.type,
        clean_(d.vid, 40), clean_(d.device, 60), clean_(d.lang, 20), clean_(d.screen, 20), clean_(d.ref, 150)
      ]);
    } finally {
      lock.releaseLock();
    }
    return out_({ ok: true });
  }
  return out_({ ok: false });
}

function stats_() {
  var sh = sheet_();
  var n = sh.getLastRow();
  var rows = n > 1 ? sh.getRange(2, 1, n - 1, 7).getValues() : [];
  var tz = Session.getScriptTimeZone();
  var by = {};
  var visits = 0;
  rows.forEach(function (r) {
    var id = String(r[2]);
    var t = r[0] instanceof Date ? r[0] : new Date(r[0]);
    var v = by[id] || (by[id] = { id: id, first: t, last: t, visits: 0, device: '', clicked: false, saved: false });
    if (t < v.first) v.first = t;
    if (t > v.last) v.last = t;
    if (r[3]) v.device = String(r[3]);
    if (r[1] === 'visit') { v.visits++; visits++; }
    if (r[1] === 'download_click') v.clicked = true;
    if (r[1] === 'download_ok') v.saved = true;
  });
  var list = Object.keys(by).map(function (k) { return by[k]; })
    .sort(function (a, b) { return b.last - a.last; });
  var fmt = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm'); };
  return {
    ok: true,
    visits: visits,
    unique: list.length,
    clickers: list.filter(function (v) { return v.clicked || v.saved; }).length,
    downloaders: list.filter(function (v) { return v.saved; }).length,
    visitors: list.slice(0, 300).map(function (v) {
      return { id: v.id.slice(0, 12), first: fmt(v.first), last: fmt(v.last), visits: v.visits,
               device: v.device, clicked: v.clicked, saved: v.saved };
    })
  };
}

function pad_(v, n) {
  v = String(v);
  while (v.length < n) v += ' ';
  return v;
}
function report_(s) {
  var L = [];
  L.push('Visits: ' + s.visits + '   Unique visitors: ' + s.unique +
         '   Pressed download: ' + s.clickers + '   Downloaded: ' + s.downloaders);
  L.push('');
  L.push(pad_('Visitor', 14) + pad_('Last seen', 18) + pad_('Visits', 8) + pad_('Device', 22) + 'Downloaded');
  s.visitors.slice(0, 40).forEach(function (v) {
    L.push(pad_(v.id, 14) + pad_(v.last, 18) + pad_(v.visits, 8) + pad_(v.device, 22) +
           (v.saved ? 'YES' : (v.clicked ? 'clicked only' : 'no')));
  });
  return L.join('\n');
}
