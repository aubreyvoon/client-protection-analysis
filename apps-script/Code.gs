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
 *
 * Version 2 adds "Open saved client": each save also stores an openable copy of the client in a
 * hidden tab called _cpa_data (same sheet, same access as the visible tabs). To update from
 * version 1: paste this file over the old code, save, then Deploy > Manage deployments > the pencil
 * (Edit) > Version: New version > Deploy. The web app URL and the secret key stay the same.
 */

var VERSION = 2;
var DATA_SHEET = '_cpa_data';

function doGet() {
  return json_({ ok: true, msg: 'Client Protection Analysis endpoint is running.' });
}

function doPost(e) {
  var lock = null;
  try {
    var key = PropertiesService.getScriptProperties().getProperty('SECRET_KEY');
    if (!key) return json_({ ok: false, error: 'SECRET_KEY is not set in Script properties.' });
    var p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (p.key !== key) return json_({ ok: false, error: 'Wrong secret key.' });
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.action === 'ping') return json_({ ok: true, title: ss.getName(), v: VERSION });
    if (p.action === 'list') return json_({ ok: true, items: list_(ss) });
    if (p.action === 'load') return load_(ss, p.tab);
    if (!p.rows || !p.rows.length) return json_({ ok: false, error: 'Nothing to save.' });
    if (p.rows.length > 1000) return json_({ ok: false, error: 'Too many rows.' });
    var data = typeof p.data === 'string' ? p.data : '';
    if (data.length > 45000) return json_({ ok: false, error: 'Client data is too large to store.' });

    lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var name = uniqueName_(ss, cleanName_(p.name));
    var sh = ss.insertSheet(name);
    writeRows_(sh, p.rows);
    var stored = data ? store_(ss, sh, p.name, data) : false;
    return json_({ ok: true, sheet: name, openable: stored, url: ss.getUrl() + '#gid=' + sh.getSheetId() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

// Openable copy of each saved client: one row per save in a hidden tab.
function store_(ss, clientSheet, client, data) {
  var sh = ss.getSheetByName(DATA_SHEET), fresh = false;
  if (!sh) {
    fresh = true;
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange(1, 1, 1, 4).setValues([['Tab', 'Saved (UTC)', 'Client', 'Data (JSON), used by Open saved client; do not edit']]);
    sh.setFrozenRows(1);
  }
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 4).setNumberFormat('@').setValues([[clientSheet.getName(), new Date().toISOString(), String(client || '').slice(0, 200), data]]);
  ss.setActiveSheet(clientSheet);
  if (fresh) sh.hideSheet();
  return true;
}

// Saved clients that still have their tab, newest first (one entry per tab name).
function list_(ss) {
  var sh = ss.getSheetByName(DATA_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var tabs = {};
  ss.getSheets().forEach(function (s) { tabs[s.getName()] = 1; });
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var seen = {}, out = [];
  for (var i = vals.length - 1; i >= 0 && out.length < 300; i--) {
    var tab = String(vals[i][0]);
    if (!tab || seen[tab] || !tabs[tab]) continue;
    seen[tab] = 1;
    out.push({ tab: tab, saved: String(vals[i][1]), name: String(vals[i][2]) });
  }
  return out;
}

function load_(ss, tab) {
  var sh = ss.getSheetByName(DATA_SHEET);
  if (sh && sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(tab)) {
        return json_({ ok: true, tab: String(tab), saved: String(vals[i][1]), name: String(vals[i][2]), data: String(vals[i][3]) });
      }
    }
  }
  return json_({ ok: false, error: 'This tab has no openable copy (it was saved before the update). Open the client some other way and save it again.' });
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
  if (base !== DATA_SHEET && !ss.getSheetByName(base)) return base;
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
