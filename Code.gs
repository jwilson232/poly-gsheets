const API_BASE     = 'https://dev.app.polylogger.com';
const USERNAME_KEY  = 'POLYLOGGER_USERNAME';
const RAW_SHEET     = 'polylogger_raw_logs';
const STATS_SHEET   = 'polylogger_stats';
const TOP_ACTIVITIES = 15;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Polylogger')
    .addItem('Pull my logs + stats', 'pullLogs')
    .addSeparator()
    .addItem('Set / change username', 'setUsername')
    .addItem('Show saved username', 'showUsername')
    .addItem('Clear saved username', 'clearUsername')
    .addToUi();
}

function props_() { return PropertiesService.getUserProperties(); }

function getUsername_(promptIfMissing) {
  const saved = props_().getProperty(USERNAME_KEY);
  if (saved) return saved;
  return promptIfMissing ? setUsername() : null;
}

function setUsername() {
  const ui = SpreadsheetApp.getUi();
  const current = props_().getProperty(USERNAME_KEY) || '';
  const msg = current
    ? 'Current username: ' + current + '\n\nEnter a new Polylogger username:'
    : 'Enter your Polylogger username:';
  const res = ui.prompt('Polylogger username', msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const value = res.getResponseText().trim().replace(/^@/, '');
  if (!value) { ui.alert('No username entered - nothing changed.'); return null; }
  props_().setProperty(USERNAME_KEY, value);
  return value;
}

function showUsername() {
  const u = props_().getProperty(USERNAME_KEY);
  SpreadsheetApp.getUi().alert(u ? 'Saved username: ' + u : 'No username saved yet.');
}

function clearUsername() {
  props_().deleteProperty(USERNAME_KEY);
  SpreadsheetApp.getUi().alert('Saved username cleared.');
}

function apiGet_(path) {
  const resp = UrlFetchApp.fetch(API_BASE + path, {
    method: 'get', muteHttpExceptions: true, headers: { 'Accept': 'application/json' }
  });
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code === 404) throw new Error('User not found. Check the username (menu -> Set / change username).');
  if (code === 401) throw new Error('That endpoint needs a login token - this script only reads public data.');
  if (code !== 200) throw new Error('Polylogger API error (HTTP ' + code + '): ' + body.slice(0, 200));
  try { return JSON.parse(body); }
  catch (e) { throw new Error('Unexpected (non-JSON) response from Polylogger.'); }
}

function pullLogs() {
  const ui = SpreadsheetApp.getUi();
  const username = getUsername_(true);
  if (!username) return;

  let rows;
  try {
    rows = apiGet_('/user/' + encodeURIComponent(username) + '/getAllLogsWithLanguages');
  } catch (e) {
    ui.alert('Could not fetch logs', String(e.message || e), ui.ButtonSet.OK);
    return;
  }
  if (!Array.isArray(rows) || rows.length === 0) {
    ui.alert('No logs found for "' + username + '".');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeRaw_(ss, rows);
  writeStats_(ss, username, rows);
}

function writeRaw_(ss, rows) {
  const sheet = getOrCreateSheet_(ss, RAW_SHEET);
  sheet.clear();

  const header = ['Date', 'Month', 'Language', 'Type', 'Activity', 'Minutes', 'Hours'];
  const data = rows.map(function (r) {
    const min = Number(r.time) || 0;
    return [toDate_(r.date), yyyymm_(r.date), r.language, r.type, r.name, min, round2_(min / 60)];
  });

  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.getRange(2, 1, data.length, header.length).setValues(data);
  sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy-mm-dd'); // Date column
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
}

function writeStats_(ss, username, rows) {
  const sheet = getOrCreateSheet_(ss, STATS_SHEET);
  sheet.clear();

  const totalMin = sum_(rows, 'time');
  const dates = uniqueSorted_(rows.map(function (r) { return r.date; }));
  const first = dates[0], last = dates[dates.length - 1];
  const span = dayDiff_(first, last) + 1;
  const streaks = streaks_(dates);

  const out = [];
  const bold = [];
  function push(arr) { while (arr.length < 5) arr.push(''); out.push(arr); return out.length; }
  function section(title) { bold.push(push([title])); }

  bold.push(push(['Polylogger stats']));
  push(['Generated', new Date(), 'user: ' + username]);
  push([]);

  section('Overview');
  push(['Total sessions', rows.length]);
  push(['Total minutes', totalMin]);
  push(['Total hours', round1_(totalMin / 60)]);
  push(['Active days', dates.length]);
  push(['First log', first]);
  push(['Last log', last]);
  push(['Days in span', span]);
  push(['Avg min / active day', round1_(totalMin / dates.length)]);
  push(['Avg min / day (span)', round1_(totalMin / span)]);
  push(['Longest streak (days)', streaks.longest]);
  push(['Streak ending ' + last, streaks.current]);
  push([]);

  writeBreakdown_(push, section, bold, 'By type', groupBy_(rows, 'type'), totalMin, true);
  push([]);
  writeBreakdown_(push, section, bold, 'By language', groupBy_(rows, 'language'), totalMin, true);
  push([]);
  writeBreakdown_(push, section, bold, 'By month', groupByMonth_(rows), totalMin, false);
  push([]);

  const acts = groupBy_(rows, 'name').slice(0, TOP_ACTIVITIES);
  writeBreakdown_(push, section, bold, 'Top ' + TOP_ACTIVITIES + ' activities', acts, totalMin, false);

  sheet.getRange(1, 1, out.length, 5).setValues(out);
  bold.forEach(function (r) { sheet.getRange(r, 1, 1, 5).setFontWeight('bold'); });
  sheet.getRange(1, 1, 1, 1).setFontSize(13);
  sheet.getRange(2, 2, 1, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);
}

function writeBreakdown_(push, section, bold, title, groups, totalMin, withShare) {
  section(title);
  const header = withShare
    ? ['', 'Sessions', 'Minutes', 'Hours', '% of min']
    : ['', 'Sessions', 'Minutes', 'Hours'];
  bold.push(push(header.slice()));
  groups.forEach(function (g) {
    const row = [g.name, g.sessions, g.minutes, round1_(g.minutes / 60)];
    if (withShare) row.push(totalMin ? round1_(100 * g.minutes / totalMin) : 0);
    push(row);
  });
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function toDate_(s) {
  const p = String(s).split('-');
  if (p.length === 3) { return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
  return s;
}

function sum_(rows, key) {
  return rows.reduce(function (s, r) { return s + (Number(r[key]) || 0); }, 0);
}

function uniqueSorted_(arr) {
  const seen = {}, out = [];
  arr.forEach(function (x) { if (!seen[x]) { seen[x] = 1; out.push(x); } });
  out.sort();
  return out;
}

function dayDiff_(a, b) {
  return Math.round((toDate_(b).getTime() - toDate_(a).getTime()) / 86400000);
}

function streaks_(sortedDates) {
  if (!sortedDates.length) return { longest: 0, current: 0 };
  let longest = 1, run = 1;
  for (let i = 1; i < sortedDates.length; i++) {
    run = (dayDiff_(sortedDates[i - 1], sortedDates[i]) === 1) ? run + 1 : 1;
    if (run > longest) longest = run;
  }
  let current = 1;
  for (let i = sortedDates.length - 1; i > 0; i--) {
    if (dayDiff_(sortedDates[i - 1], sortedDates[i]) === 1) current++; else break;
  }
  return { longest: longest, current: current };
}

function groupBy_(rows, key) {
  const m = {};
  rows.forEach(function (r) {
    const k = r[key]; if (!m[k]) m[k] = { sessions: 0, minutes: 0 };
    m[k].sessions++; m[k].minutes += Number(r.time) || 0;
  });
  return Object.keys(m).map(function (k) {
    return { name: k, sessions: m[k].sessions, minutes: m[k].minutes };
  }).sort(function (a, b) { return b.minutes - a.minutes; });
}

function groupByMonth_(rows) {
  const m = {};
  rows.forEach(function (r) {
    const k = String(r.date).slice(0, 7); if (!m[k]) m[k] = { sessions: 0, minutes: 0 };
    m[k].sessions++; m[k].minutes += Number(r.time) || 0;
  });
  return Object.keys(m).sort().map(function (k) {
    return { name: k, sessions: m[k].sessions, minutes: m[k].minutes };
  });
}

function yyyymm_(value) {
  if (typeof value === 'string') {
    var m = value.match(/^(\d{4})-(\d{2})/);
    if (m) return m[1] + '-' + m[2];
  }
  var d = new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
}

function round1_(n) { return Math.round(n * 10) / 10; }
function round2_(n) { return Math.round(n * 100) / 100; }
