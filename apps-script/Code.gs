/**
 * Client Protection Analysis: saves each client as a NEW tab in this spreadsheet.
 *
 * Setup (once):
 *  1. In your Google Sheet: Extensions > Apps Script. Paste this file, save.
 *  2. Project Settings (gear icon) > Script properties > Add property:
 *       name  = SECRET_KEY
 *       value = a long random password you make up (paste the same one in the app)
 *  3. Deploy > New deployment > type "Web app":
 *       Execute as: Me      Who has access: Anyone
 *     Authorise when asked, then copy the Web app URL (ends with /exec) into the app.
 *  After you change this code later, use Deploy > Manage deployments > Edit > New version.
 */

function doGet() {
  return json_({ ok: true, msg: 'Client Protection Analysis endpoint is running.' });
}

function doPost(e) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('SECRET_KEY');
    if (!key) return json_({ ok: false, error: 'SECRET_KEY is not set in Script properties.' });
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.key !== key) return json_({ ok: false, error: 'Wrong secret key.' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.action === 'ping') return json_({ ok: true, title: ss.getName() });
    if (!p.rows || !p.rows.length) return json_({ ok: false, error: 'Nothing to save.' });
    if (p.rows.length > 1000) return json_({ ok: false, error: 'Too many rows.' });

    var name = uniqueName_(ss, cleanName_(p.name));
    var sh = ss.insertSheet(name);
    writeRows_(sh, p.rows);
    return json_({ ok: true, sheet: name, url: ss.getUrl() + '#gid=' + sh.getSheetId() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Tab names: max 100 chars, no [ ] * ? : / \ and no leading/trailing apostrophe.
function cleanName_(s) {
  var n = String(s || '').replace(/[\[\]*?:\/\\]/g, ' ').replace(/\s+/g, ' ').replace(/^'+|'+$/g, '').trim();
  return (n || 'Client').slice(0, 90);
}

// A new tab every time: "Name", then "Name (2)", "Name (3)", ...
function uniqueName_(ss, base) {
  if (!ss.getSheetByName(base)) return base;
  for (var i = 2; i < 1000; i++) {
    var cand = base + ' (' + i + ')';
    if (!ss.getSheetByName(cand)) return cand;
  }
  return base + ' ' + Date.now();
}

// Stop typed text turning into formulas, dates or numbers (NRIC, DOB, "=...", "-...").
function clean_(v) {
  if (typeof v !== 'string') return v == null ? '' : v;
  if (/^[=+\-@]/.test(v) || /^[\d\s\-\/:.,]+$/.test(v) || /^\d{1,2}\s+[A-Za-z]+\s+\d{4}$/.test(v)) return "'" + v;
  return v;
}

function col_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function writeRows_(sh, rows) {
  var w = 2;
  rows.forEach(function (r) { if (r.c.length > w) w = r.c.length; });
  var vals = rows.map(function (r) {
    var a = r.c.map(clean_);
    while (a.length < w) a.push('');
    return a;
  });
  sh.getRange(1, 1, vals.length, w).setValues(vals);

  var g = { h1: [], sec: [], sub: [], th: [], tot: [], note: [], money: [], pct: [], ok: [], gap: [] };
  rows.forEach(function (r, i) {
    var rn = i + 1, whole = 'A' + rn + ':' + col_(w) + rn;
    if (g[r.t]) g[r.t].push(whole);
    (r.m || []).forEach(function (ci) { if (typeof r.c[ci] === 'number') g.money.push(col_(ci + 1) + rn); });
    (r.p || []).forEach(function (ci) { if (typeof r.c[ci] === 'number') g.pct.push(col_(ci + 1) + rn); });
    r.c.forEach(function (v, ci) {
      var good = (v === 'OK' || v === 'Surplus' || v === 'Fully covered');
      var bad = (v === 'GAP' || v === 'Shortfall');
      if (good) g.ok.push(col_(ci + 1) + rn);
      else if (bad) g.gap.push(col_(ci + 1) + rn);
      // Shortfall / surplus: also colour the amount in the cell just before the status word
      if ((v === 'Surplus' || v === 'Shortfall') && ci > 0 && typeof r.c[ci - 1] === 'number') {
        (good ? g.ok : g.gap).push(col_(ci) + rn);
      }
    });
  });
  function on(list, fn) { if (list.length) fn(sh.getRangeList(list)); }
  on(g.h1, function (r) { r.setFontSize(16).setFontWeight('bold').setFontColor('#14284b'); });
  on(g.sec, function (r) { r.setBackground('#14284b').setFontColor('#ffffff').setFontWeight('bold'); });
  on(g.sub, function (r) { r.setFontWeight('bold').setFontColor('#a67c2e'); });
  on(g.th, function (r) { r.setBackground('#e8edf5').setFontWeight('bold'); });
  on(g.tot, function (r) { r.setFontWeight('bold'); });
  on(g.note, function (r) { r.setFontStyle('italic').setFontColor('#666666'); });
  on(g.money, function (r) { r.setNumberFormat('"RM"#,##0;[Red]-"RM"#,##0'); });
  on(g.pct, function (r) { r.setNumberFormat('0%'); });
  on(g.ok, function (r) { r.setFontColor('#2b7a4b').setFontWeight('bold'); });
  on(g.gap, function (r) { r.setFontColor('#b4472b').setFontWeight('bold'); });
  sh.setColumnWidth(1, 360);
  for (var c = 2; c <= w; c++) sh.setColumnWidth(c, 150);
  sh.setHiddenGridlines(true);
  sh.setTabColor('#14284b');
}
