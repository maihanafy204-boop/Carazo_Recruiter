// Recruiter assessment: server side (scoring + Google Sheet). Paste into Code.gs
var SHEET_NAME = 'Test Results';
var SECTIONS = {
  P: {n: 'Prioritization', w: 25},
  S: {n: 'Sourcing', w: 20},
  J: {n: 'Screening', w: 20},
  C: {n: 'Compliance', w: 15},
  K: {n: 'Stakeholders', w: 10},
  D: {n: 'Data & KPIs', w: 10}
};
var ORDER = ['P', 'S', 'J', 'C', 'K', 'D'];
// [section, points for each option in its original order]
var KEY = [["P",[1,1,0,2]],["P",[0,1,2,1]],["P",[0,1,1,2]],["P",[1,1,0,2]],["S",[1,2,0,1]],["S",[2,0,1,1]],["S",[0,2,1,1]],["C",[0,2,1,1]],["J",[1,1,0,2]],["J",[2,0,1,1]],["J",[1,0,2,1]],["J",[1,0,2,1]],["K",[1,1,2,0]],["K",[0,2,1,1]],["K",[1,0,1,2]],["C",[1,1,2,0]],["C",[0,1,0,2]],["C",[0,2,1,1]],["D",[0,2,1,1]],["D",[0,0,2,0]]];

function doGet() {
  return ContentService.createTextOutput('Recruiter test API is running.');
}

function doPost(e) {
  var out;
  try {
    out = submit(JSON.parse(e.postData.contents));
  } catch (err) {
    out = {ok: false, error: String(err.message || err)};
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function score_(answers) {
  var got = {}, max = {};
  for (var i = 0; i < KEY.length; i++) {
    var s = KEY[i][0];
    max[s] = (max[s] || 0) + 2;
    var a = answers[i];
    got[s] = (got[s] || 0) + (a === null ? 0 : KEY[i][1][a]);
  }
  var pct = {}, total = 0;
  ORDER.forEach(function (s) {
    pct[s] = Math.round(got[s] / max[s] * 100);
    total += pct[s] * SECTIONS[s].w / 100;
  });
  total = Math.round(total);
  var gate = pct.J >= 60 && pct.C >= 60;
  var status = (total >= 70 && gate) ? 'PASS' : (total >= 60 ? 'REVIEW' : 'FAIL');
  var note = (total >= 70 && !gate) ? 'Gate failed (Screening or Compliance below 60%)' : '';
  return {pct: pct, total: total, status: status, note: note};
}

function safe_(v) {
  v = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function ensureSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var head = ['Date', 'Name', 'Email', 'Language', 'Country', 'Attempt'];
    ORDER.forEach(function (s) { head.push(SECTIONS[s].n + ' %'); });
    head = head.concat(['Overall %', 'Status', 'Note', 'Timeouts', 'Avg sec/question']);
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8691F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 5).getValue() !== 'Country') {
    sh.insertColumnBefore(5);
    sh.getRange(1, 5).setValue('Country').setFontWeight('bold').setBackground('#E8691F').setFontColor('#FFFFFF');
  }
  return sh;
}

// Run this once from the editor to create the empty "Test Results" tab with its headers.
function setup() {
  ensureSheet_(SpreadsheetApp.getActiveSpreadsheet());
}

function submit(p) {
  if (!p || typeof p.name !== 'string' || typeof p.email !== 'string' ||
      !Array.isArray(p.answers) || p.answers.length !== KEY.length ||
      !Array.isArray(p.times) || !p.name.trim() || p.email.indexOf('@') < 1) {
    throw new Error('Bad input');
  }
  var answers = p.answers.map(function (a) {
    if (a === null) return null;
    if (typeof a !== 'number' || a % 1 !== 0 || a < 0 || a > 3) throw new Error('Bad answer');
    return a;
  });
  var r = score_(answers);
  var timeouts = answers.filter(function (a) { return a === null; }).length;
  var times = p.times.map(Number).filter(function (x) { return isFinite(x); });
  var avg = times.length ? Math.round(times.reduce(function (x, y) { return x + y; }, 0) / times.length) : 0;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cache = CacheService.getScriptCache();
    var sidKey = 'sid_' + String(p.sid || '').slice(0, 80);
    if (p.sid && cache.get(sidKey)) return {ok: true, dup: true};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureSheet_(ss);
    var email = safe_(p.email.trim().toLowerCase());
    var attempt = 1;
    if (sh.getLastRow() > 1) {
      var emails = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
      emails.forEach(function (e) { if (String(e[0]).toLowerCase() === email) attempt++; });
    }
    var row = [
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'),
      safe_(p.name.trim()), email, p.lang === 'es' ? 'Spanish' : 'English', safe_(p.country || ''), attempt
    ];
    ORDER.forEach(function (s) { row.push(r.pct[s]); });
    row = row.concat([r.total, r.status, r.note, timeouts, avg]);
    sh.appendRow(row);
    var last = sh.getLastRow();
    var statusCol = 6 + ORDER.length + 2;
    var colors = {PASS: '#C6EFCE', REVIEW: '#FFEB9C', FAIL: '#FFC7CE'};
    sh.getRange(last, statusCol).setBackground(colors[r.status]).setFontWeight('bold');
    if (p.sid) cache.put(sidKey, '1', 21600);
  } finally {
    lock.releaseLock();
  }
  return {ok: true};
}
