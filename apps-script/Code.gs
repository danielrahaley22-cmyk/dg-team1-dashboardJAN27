/**
 * DG Team 1 dashboard – Bob's weekly outreach.
 *
 * Saves to the "DG Team 1 Outreach (Bob)" sheet (OUTREACH_SHEET_ID). It:
 *  - sends the dashboard Bob's team (code, name, Instagram only), their
 *    Melbourne sales and this week's ticks (doGet)
 *  - saves each Focus / Sent / Replied tick and comment as a row in the
 *    "Outreach" tab, one row per promoter per week (doPost)
 *  - emails you a summary of the week every Sunday at 8pm (weeklySummary)
 *
 * Setup (once):
 *  1. Open script.new (or Extensions > Apps Script in the sheet), paste this file over everything, Save.
 *  2. Pick "setup" in the function menu at the top and press Run. Allow access
 *     when Google asks. This makes the Outreach tab and the Sunday email.
 *  3. Deploy > New deployment > Web app:
 *       Execute as: Me
 *       Who has access: Anyone
 *     Copy the Web app URL (ends in /exec) into SCRIPT_URL in index.html.
 *
 * If you change this file later: Deploy > Manage deployments > edit (pencil) >
 * Version: New version > Deploy. The /exec link stays the same.
 */

// Bob's team list: MELB G1 BOB - FINAL 2026, tab "Team 1". Only A:C is read
// (Code, Name, Instagram), so phone numbers and emails never leave the sheet.
var TEAM_SHEET_ID = '1QAZl8A0arrzHinRZXLAwY0ciN3f3i8M7_Vn2Lsnu1As';
var TEAM_RANGE = "'Team 1'!A1:C400";

// DG Discount Code Sales, Data tab (replaced with each Megatix export).
var SALES_SHEET_ID = '1Duvr9aJTaH5AMA4l6slkpIb5CQDR2j48I18FQAEaGr8';
var SALES_RANGE = 'Data!A1:J3000';
var SALES_STATE = 'Melbourne';

// The "DG Team 1 Outreach (Bob)" sheet, where the ticks are saved.
var OUTREACH_SHEET_ID = '1oNlAL7z2pokQ2mLW8Rgp2b13cfI7eG7ScA_D6HOBEgY';
var OUTREACH_TAB = 'Outreach';
var HEADERS = ['Week', 'Code', 'Name', 'Instagram', 'Focus', 'Sent', 'Replied', 'Comment', 'Updated'];
var TZ = 'Australia/Melbourne';

// Leave blank to email the account that deployed the script.
var NOTIFY_EMAIL = '';
var TEAM_LEAD = 'Bob';

/* ---------- web app ---------- */

function doGet(e) {
  try {
    var week = cleanWeek((e && e.parameter && e.parameter.week) || '') || thisWeek();
    return json({ ok: true, data: buildData(week) });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var week = cleanWeek(d.week) || thisWeek();
    var code = String(d.code || '').trim().toUpperCase();
    if (!code) throw new Error('No code');
    var team = cachedTeam();
    var who = team.filter(function (r) { return r.code === code; })[0];
    if (!who) throw new Error('Code ' + code + " isn't on the Team 1 list");

    var sheet = outreachSheet();
    // Only read Week and Code to find the row, then just that one row.
    var last = sheet.getLastRow();
    var keys = last > 1 ? sheet.getRange(2, 1, last - 1, 2).getValues() : [];
    var rowIndex = -1;
    for (var i = keys.length - 1; i >= 0; i--) {
      if (weekKey(keys[i][0]) === week && String(keys[i][1]).toUpperCase() === code) { rowIndex = i + 1; break; }
    }
    var row = rowIndex > -1 ? sheet.getRange(rowIndex + 1, 1, 1, HEADERS.length).getValues()[0]
      : [week, code, who.name, who.instagram, false, false, false, '', ''];
    row[0] = week;
    row[2] = who.name;
    row[3] = who.instagram;
    if (d.focus !== undefined) row[4] = d.focus === true;
    if (d.sent !== undefined) row[5] = d.sent === true;
    if (d.replied !== undefined) row[6] = d.replied === true;
    if (d.comment !== undefined) row[7] = clean(d.comment, 500);
    row[8] = new Date();

    var target = rowIndex > -1 ? rowIndex + 1 : last + 1;
    if (rowIndex < 0) sheet.getRange(target, 1).setNumberFormat('@');
    sheet.getRange(target, 1, 1, HEADERS.length).setValues([row]);
    if (rowIndex < 0) sheet.getRange(target, 5, 1, 3).insertCheckboxes();
    return json({ ok: true, tick: tickFromRow(row) });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- data ---------- */

// Team and sales change rarely, so they are kept for 5 minutes to make the
// page and every tick faster. Ticks are always read fresh.
function cachedTeam() {
  var c = CacheService.getScriptCache(), v = c.get('team');
  if (v) return JSON.parse(v);
  var t = readTeam();
  c.put('team', JSON.stringify(t), 300);
  return t;
}

function cachedPeople() {
  var c = CacheService.getScriptCache(), v = c.get('people');
  if (v) return JSON.parse(v);
  var p = buildPeople();
  try { c.put('people', JSON.stringify(p), 300); } catch (e) {}
  return p;
}

function buildPeople() {
  var team = cachedTeam();
  var sales = readSales();
  var people = team.map(function (r) {
    var s = sales.byCode[r.code];
    return {
      code: r.code,
      name: r.name,
      instagram: r.instagram,
      tickets: s ? s.tickets : 0,
      last7: s ? s.last7 : 0,
      lastSale: s ? s.lastSale : ''
    };
  });
  return { people: people, melTotal: sales.total, salesTo: sales.latest };
}

function buildData(week) {
  var p = cachedPeople();
  var ticks = readTicks();
  var weeks = {};
  weeks[thisWeek()] = true;
  ticks.forEach(function (t) { weeks[t.week] = true; });
  return {
    week: week,
    thisWeek: thisWeek(),
    weeks: Object.keys(weeks).sort().reverse(),
    people: p.people,
    ticks: ticks.filter(function (t) { return t.week === week; }),
    melTotal: p.melTotal,
    salesTo: p.salesTo
  };
}

function readTeam() {
  var v = SpreadsheetApp.openById(TEAM_SHEET_ID).getRange(TEAM_RANGE).getDisplayValues();
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var ci = h.indexOf('code'), ni = h.indexOf('name'), ii = h.indexOf('instagram');
  if (ci < 0) ci = 0; if (ni < 0) ni = 1; if (ii < 0) ii = 2;
  var seen = {}, out = [];
  v.slice(1).forEach(function (r) {
    var code = String(r[ci] || '').trim().toUpperCase();
    if (!code || seen[code]) return;
    seen[code] = true;
    out.push({ code: code, name: String(r[ni] || '').trim(), instagram: String(r[ii] || '').trim() });
  });
  return out;
}

function readSales() {
  var v = SpreadsheetApp.openById(SALES_SHEET_ID).getRange(SALES_RANGE).getValues();
  var h = v[0].map(function (x) { return String(x).trim(); });
  var col = function (name) { return h.indexOf(name); };
  var c = col('Code'), st = col('Event state'), t = col('Tickets'), l7 = col('Last 7 days tix'), ls = col('Last sale');
  var byCode = {}, total = 0, latest = '';
  v.slice(1).forEach(function (r) {
    if (String(r[st]).trim() !== SALES_STATE) return;
    var tix = num(r[t]);
    total += tix;
    var last = dateStr(r[ls]);
    if (last && last > latest) latest = last;
    var code = String(r[c] || '').trim().toUpperCase();
    if (code) byCode[code] = { tickets: tix, last7: num(r[l7]), lastSale: last };
  });
  return { byCode: byCode, total: total, latest: latest };
}

function readTicks() {
  var sheet = outreachSheet();
  var v = sheet.getDataRange().getValues();
  return v.slice(1).filter(function (r) { return r[1]; }).map(tickFromRow);
}

function tickFromRow(r) {
  return {
    week: weekKey(r[0]),
    code: String(r[1]).toUpperCase(),
    focus: r[4] === true,
    sent: r[5] === true,
    replied: r[6] === true,
    comment: String(r[7] || ''),
    updated: r[8] instanceof Date ? Utilities.formatDate(r[8], TZ, "yyyy-MM-dd'T'HH:mm") : String(r[8] || '')
  };
}

/* ---------- weekly email ---------- */

function weeklySummary() {
  var week = thisWeek();
  var team = readTeam();
  var names = {};
  team.forEach(function (r) { names[r.code] = r; });
  var ticks = readTicks().filter(function (t) { return t.week === week && names[t.code]; });
  var focus = ticks.filter(function (t) { return t.focus; });
  var sent = ticks.filter(function (t) { return t.sent; });
  var replied = ticks.filter(function (t) { return t.replied; });
  var notSent = focus.filter(function (t) { return !t.sent; });
  var noReply = sent.filter(function (t) { return !t.replied; });
  var line = function (t) { return '  ' + t.code + (t.comment ? ' – ' + t.comment : ''); };
  var list = function (arr) { return arr.length ? arr.map(line).join('\n') : '  (none)'; };

  var body =
    TEAM_LEAD + "'s Team 1 outreach, week of " + prettyWeek(week) + '\n\n' +
    'In focus: ' + focus.length + '\n' +
    'Messaged: ' + sent.length + '\n' +
    'Replied: ' + replied.length + (sent.length ? ' (' + Math.round(100 * replied.length / sent.length) + '% of messaged)' : '') + '\n\n' +
    'Replied:\n' + list(replied) + '\n\n' +
    'Messaged, no reply yet:\n' + list(noReply) + '\n\n' +
    'In focus but not messaged yet:\n' + list(notSent) + '\n\n' +
    'Full log: ' + SpreadsheetApp.openById(OUTREACH_SHEET_ID).getUrl();

  var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: TEAM_LEAD + "'s outreach: " + sent.length + ' messaged, ' + replied.length + ' replied (week of ' + prettyWeek(week) + ')',
    body: body
  });
}

/* ---------- setup ---------- */

function setup() {
  outreachSheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'weeklySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklySummary').timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(20).inTimezone(TZ).create();
  // Touch the other sheets so Google asks for access to them now.
  readTeam();
  readSales();
}

function outreachSheet() {
  var ss = SpreadsheetApp.openById(OUTREACH_SHEET_ID);
  var sheet = ss.getSheetByName(OUTREACH_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(OUTREACH_TAB, 0);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(8, 320);
  }
  return sheet;
}

/* ---------- helpers ---------- */

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Monday of the current week, Melbourne time, as yyyy-MM-dd.
function thisWeek() {
  var now = new Date();
  var dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1 = Monday ... 7 = Sunday
  var monday = new Date(now.getTime() - (dow - 1) * 864e5);
  return Utilities.formatDate(monday, TZ, 'yyyy-MM-dd');
}

function weekKey(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return cleanWeek(v);
}

function cleanWeek(v) {
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function prettyWeek(w) {
  var p = w.split('-');
  return Utilities.formatDate(new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12), TZ, 'EEE d MMM');
}

function dateStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];
  var d = new Date(s.replace(/\s+\d{1,2}:\d{2}.*/, ''));
  return isNaN(d) ? '' : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function num(x) {
  var n = parseFloat(String(x == null ? '' : x).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function clean(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max || 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
