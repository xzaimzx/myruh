/**
 * Malaysiano Tour – Google Sheets backend
 *
 * Paste this whole file into Extensions → Apps Script in your Google Sheet.
 * 1. Choose "setup" in the function dropdown and click Run (once only).
 * 2. Deploy → New deployment → Web app, Execute as: Me, Who has access: Anyone.
 * 3. Copy the Web app URL (ends in /exec) into index.html.
 *
 * Tabs: Players, Stops, Matches, ScoreLog, Settings, Rules.
 * Running setup again later is safe: it only adds tabs or columns that are missing.
 * PINs live in the Settings tab and are never sent to the website.
 */

var TABS = {
  Players: ['id', 'name', 'gender', 'joined'],
  Stops: ['id', 'name', 'date', 'venue', 'total', 'created', 'time'],
  Matches: ['id', 'stop', 'cup', 'round', 'time', 'court', 'a1', 'a2', 'b1', 'b2', 'sa', 'sb', 'seq', 'enteredBy', 'updated'],
  ScoreLog: ['time', 'name', 'matchId', 'match', 'oldScore', 'newScore'],
  Settings: ['key', 'value', 'note'],
  Rules: ['rule']
};

var DEFAULT_RULES = [
  'Every game is 21 points in total (e.g. 12–9). No draws.',
  'Serve rotates every 4 points.',
  'Each player collects their own points from the team score.',
  'When your court finishes, start the next round straight away.',
  'Back-to-back games stay on the same court.',
  'Injured or tired? A resting player fills in.',
  "At 10–10, if there's a lady in the match, men can't serve to her.",
  'Highest points wins each cup. Top 3 in each cup get gold, silver and bronze medals.'
];

/* ------------------------------------------------------------------ */
/* One-time setup                                                     */
/* ------------------------------------------------------------------ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    // Writes the header row, which also adds any new columns to tabs made by an older version.
    sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Plain text everywhere so Sheets doesn't turn dates, times or ids into something else.
    sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat('@');
  });
  var first = ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  var settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() < 2) {
    settings.getRange(2, 1, 2, 3).setValues([
      ['ScorePIN', '1710', 'Share with the group. Needed to enter scores. Change it any week.'],
      ['AdminPIN', '2468', 'Only you. Needed to add players, stops and matches, and delete matches.']
    ]);
  }

  var rules = ss.getSheetByName('Rules');
  if (rules.getLastRow() < 2) {
    rules.getRange(2, 1, DEFAULT_RULES.length, 1).setValues(DEFAULT_RULES.map(function (r) { return [r]; }));
    rules.setColumnWidth(1, 700);
  }

  var now = new Date().toISOString();
  var players = ss.getSheetByName('Players');
  if (players.getLastRow() < 2) {
    var roster = [
      ['omar', 'Omar H', 'M'], ['miza', 'Amiza', 'F'], ['faisal', 'Faisal J', 'M'], ['ain', 'NF Ain', 'F'],
      ['zaim', 'Zaim Z', 'M'], ['farah', 'Farah AN', 'F'], ['shazwan', 'Shazwan AR', 'M'], ['fakhri', 'Fakri G', 'M'],
      ['qayyum', 'Qayyum Z', 'M'], ['amir', 'Amir AJ', 'M'], ['derek', 'Derek P', 'M'], ['du', 'Du', 'M'],
      ['amirul', 'Amirul I', 'M']
    ];
    players.getRange(2, 1, roster.length, 4).setValues(roster.map(function (r) { return r.concat([now]); }));
  }

  var stops = ss.getSheetByName('Stops');
  if (stops.getLastRow() < 2) {
    stops.getRange(2, 1, 1, 7).setValues([['stop-1', 'Tour Stop 1', '2026-10-17', 'Padel In Ar Rabwah', '21', now, '8:45am – 12:00pm']]);
  }

  var matches = ss.getSheetByName('Matches');
  if (matches.getLastRow() < 2) {
    var cups = [
      ['Mixed Mayhem Cup', 0, ['9:00', '9:15', '9:30', '9:45', '10:00', '10:15'], [
        'omar,miza|faisal,ain', 'shazwan,fakhri|qayyum,amir',
        'omar,miza|du,derek', 'zaim,farah|qayyum,amirul',
        'shazwan,du|fakhri,amir', 'faisal,ain|zaim,farah',
        'omar,miza|fakhri,amirul', 'faisal,ain|qayyum,derek',
        'omar,miza|zaim,farah', 'shazwan,qayyum|amir,derek',
        'zaim,farah|fakhri,du', 'faisal,ain|shazwan,amirul'
      ]],
      ['Smash Showdown Cup', 1, ['10:30', '10:45', '11:00', '11:15', '11:30', '11:45'], [
        'omar,miza|fakhri,du', 'amir,amirul|qayyum,zaim',
        'omar,miza|shazwan,faisal', 'amir,derek|qayyum,amirul',
        'shazwan,fakhri|faisal,zaim', 'amir,du|amirul,derek',
        'omar,miza|fakhri,zaim', 'amir,qayyum|derek,du',
        'omar,miza|shazwan,amirul', 'qayyum,derek|du,faisal',
        'shazwan,amir|amirul,fakhri', 'derek,faisal|du,zaim'
      ]]
    ];
    var rows = [];
    cups.forEach(function (c) {
      c[3].forEach(function (pair, i) {
        var round = Math.floor(i / 2) + 1, court = (i % 2) + 1;
        var t = pair.split('|'), a = t[0].split(','), b = t[1].split(',');
        rows.push(['s1-c' + (c[1] + 1) + '-r' + round + '-' + court, 'stop-1', c[0], String(round), c[2][round - 1],
          String(court), a[0], a[1], b[0], b[1], '', '', String(c[1] * 1000 + round * 10 + court), '', '']);
      });
    });
    matches.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// Run once from the Apps Script editor to give players their full display names.
// Matches by id, so every score stays linked to the right person.
function updateNames() {
  var names = {
    omar: 'Omar H', miza: 'Amiza', faisal: 'Faisal J', ain: 'NF Ain', zaim: 'Zaim Z', farah: 'Farah AN',
    shazwan: 'Shazwan AR', fakhri: 'Fakri G', qayyum: 'Qayyum Z', amir: 'Amir AJ', derek: 'Derek P',
    du: 'Du', amirul: 'Amirul I'
  };
  var sh = sheet('Players'), n = sh.getLastRow() - 1;
  if (n < 1) return;
  var range = sh.getRange(2, 1, n, 2), vals = range.getValues();
  vals.forEach(function (r) { if (names[String(r[0])]) r[1] = names[String(r[0])]; });
  range.setValues(vals);
}

/* ------------------------------------------------------------------ */
/* Web app                                                            */
/* ------------------------------------------------------------------ */

function doGet() {
  return json(readAll());
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'Bad request.' }); }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json({ ok: false, error: 'The sheet is busy. Try again in a moment.' });
  try {
    var res = handle(req || {});
    SpreadsheetApp.flush();
    return json(res);
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function handle(req) {
  switch (req.action) {
    case 'score': return saveScore(req);
    case 'checkScorePin': return checkPin('ScorePIN', req.pin) ? { ok: true } : { ok: false, error: "That Score PIN didn't work." };
    case 'checkAdmin': return admin(req) || { ok: true };
    case 'addPlayer': return admin(req) || addPlayer(req);
    case 'addStop': return admin(req) || addStop(req);
    case 'addMatch': return admin(req) || addMatch(req);
    case 'deleteMatch': return admin(req) || deleteMatch(req);
  }
  return { ok: false, error: 'Unknown action.' };
}

/* ------------------------------------------------------------------ */
/* Actions                                                            */
/* ------------------------------------------------------------------ */

function saveScore(req) {
  var name = cleanText(req.name, 40);
  if (!name) return { ok: false, error: 'Type your name first.' };
  if (!checkPin('ScorePIN', req.pin)) return { ok: false, error: "That Score PIN didn't work.", badPin: true };
  var sa = scoreVal(req.sa), sb = scoreVal(req.sb);
  if (sa === undefined || sb === undefined) return { ok: false, error: 'Scores must be whole numbers.' };
  if ((sa === '') !== (sb === '')) return { ok: false, error: 'Enter both scores, or clear both.' };

  var sh = sheet('Matches'), row = findRow(sh, req.matchId);
  if (!row) return { ok: false, error: 'That match no longer exists. Refresh the page.' };
  var cols = TABS.Matches, vals = sh.getRange(row, 1, 1, cols.length).getValues()[0];
  var m = {}; cols.forEach(function (c, i) { m[c] = String(vals[i]); });
  var oldScore = m.sa === '' ? '–' : m.sa + '–' + m.sb;
  var newScore = sa === '' ? '–' : sa + '–' + sb;
  if (oldScore === newScore) return { ok: true, match: m, unchanged: true };

  var now = new Date().toISOString();
  m.sa = String(sa); m.sb = String(sb); m.enteredBy = sa === '' ? '' : name; m.updated = now;
  sh.getRange(row, cols.indexOf('sa') + 1, 1, 2).setValues([[m.sa, m.sb]]);
  sh.getRange(row, cols.indexOf('enteredBy') + 1, 1, 2).setValues([[m.enteredBy, now]]);

  var names = playerNames();
  var desc = m.cup + ' R' + m.round + ' C' + m.court + ': ' + names(m.a1) + '/' + names(m.a2) + ' v ' + names(m.b1) + '/' + names(m.b2);
  sheet('ScoreLog').appendRow([now, name, m.id, desc, oldScore, newScore]);
  return { ok: true, match: m };
}

function addPlayer(req) {
  var name = cleanText(req.name, 30);
  if (!name) return { ok: false, error: 'Type a name first.' };
  var gender = req.gender === 'F' ? 'F' : 'M';
  var existing = rows('Players');
  if (existing.some(function (p) { return p.name.toLowerCase() === name.toLowerCase(); })) return { ok: false, error: name + ' is already on the list.' };
  var base = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'p', id = base, i = 2;
  var ids = existing.map(function (p) { return p.id; });
  while (ids.indexOf(id) >= 0) id = base + '-' + (i++);
  sheet('Players').appendRow([id, name, gender, new Date().toISOString()]);
  return { ok: true, id: id };
}

function addStop(req) {
  var name = cleanText(req.name, 40);
  if (!name) return { ok: false, error: 'Give the stop a name.' };
  var date = /^\d{4}-\d{2}-\d{2}$/.test(req.date || '') ? req.date : '';
  var total = parseInt(req.total, 10);
  var id = 'stop-' + Date.now().toString(36);
  sheet('Stops').appendRow([id, name, date, cleanText(req.venue, 60), isNaN(total) ? '' : String(total), new Date().toISOString(), cleanText(req.time, 40)]);
  return { ok: true, id: id };
}

function addMatch(req) {
  var ids = [req.a1, req.a2, req.b1, req.b2].map(String);
  var known = rows('Players').map(function (p) { return p.id; });
  if (ids.some(function (x) { return known.indexOf(x) < 0; })) return { ok: false, error: 'Pick all four players.' };
  if (ids.filter(function (x, i) { return ids.indexOf(x) === i; }).length !== 4) return { ok: false, error: 'Each player can only appear once in a match.' };
  if (!rows('Stops').some(function (s) { return s.id === String(req.stop); })) return { ok: false, error: 'That tour stop no longer exists.' };
  var cup = cleanText(req.cup, 40) || 'Cup';
  var round = parseInt(req.round, 10) || 1, court = parseInt(req.court, 10) || 1;
  var cupIdx = parseInt(req.cupIdx, 10) || 0;
  var id = 'm-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  sheet('Matches').appendRow([id, String(req.stop), cup, String(round), cleanText(req.time, 10), String(court),
    ids[0], ids[1], ids[2], ids[3], '', '', String(cupIdx * 1000 + round * 10 + court), '', '']);
  return { ok: true, id: id };
}

function deleteMatch(req) {
  var sh = sheet('Matches'), row = findRow(sh, req.matchId);
  if (!row) return { ok: true };
  var vals = sh.getRange(row, 1, 1, TABS.Matches.length).getValues()[0];
  sh.deleteRow(row);
  sheet('ScoreLog').appendRow([new Date().toISOString(), 'Organiser', String(vals[0]), 'Deleted match ' + vals[2] + ' R' + vals[3] + ' C' + vals[5],
    vals[10] === '' ? '–' : vals[10] + '–' + vals[11], 'deleted']);
  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* Helpers                                                            */
/* ------------------------------------------------------------------ */

function readAll() {
  var players = rows('Players').map(function (p) { return { id: p.id, name: p.name, gender: p.gender === 'F' ? 'F' : 'M' }; });
  var stops = rows('Stops').map(function (s) {
    return { id: s.id, name: s.name, date: s.date, venue: s.venue, total: num(s.total), created: s.created, time: s.time };
  });
  var matches = rows('Matches').map(function (m) {
    return {
      id: m.id, stop: m.stop, cup: m.cup, round: num(m.round) || 1, time: m.time, court: num(m.court),
      a: [m.a1, m.a2], b: [m.b1, m.b2], sa: num(m.sa), sb: num(m.sb), seq: num(m.seq) || 0,
      enteredBy: m.enteredBy, updated: m.updated
    };
  });
  var ruleTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  var rules = ruleTab ? rows('Rules').map(function (r) { return r.rule; }).filter(String) : null;
  return { ok: true, players: players, stops: stops, matches: matches, rules: rules };
}

function rows(tab) {
  var sh = sheet(tab), cols = TABS[tab], n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      var o = {};
      cols.forEach(function (c, i) {
        var v = r[i];
        if (v instanceof Date) v = c === 'time' ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'H:mm')
          : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        o[c] = String(v).trim();
      });
      return o;
    });
}

function sheet(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('The ' + name + ' tab is missing. Run setup in Apps Script.');
  return sh;
}

function findRow(sh, id) {
  var n = sh.getLastRow() - 1;
  if (n < 1 || !id) return 0;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function playerNames() {
  var map = {};
  rows('Players').forEach(function (p) { map[p.id] = p.name; });
  return function (id) { return map[id] || id; };
}

function setting(key) {
  var r = rows('Settings');
  for (var i = 0; i < r.length; i++) if (r[i].key === key) return r[i].value;
  return '';
}

function checkPin(key, pin) {
  var want = setting(key);
  return want !== '' && String(pin || '').trim() === want;
}

// Slows down guessing: after 10 wrong Admin PINs in 15 minutes, organiser tools pause for 15 minutes.
function admin(req) {
  var cache = CacheService.getScriptCache();
  var fails = parseInt(cache.get('adminFails') || '0', 10);
  if (fails >= 10) return { ok: false, error: 'Too many wrong Admin PINs. Wait 15 minutes and try again.' };
  if (checkPin('AdminPIN', req.adminPin)) return null;
  cache.put('adminFails', String(fails + 1), 900);
  return { ok: false, error: "That Admin PIN didn't work.", badPin: true };
}

function scoreVal(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(v);
  if (!isFinite(n) || n < 0 || n > 999 || Math.floor(n) !== n) return undefined;
  return n;
}

function num(v) {
  if (v === '' || v === null || v === undefined) return null;
  var n = Number(v);
  return isFinite(n) ? n : null;
}

// Trims, shortens and stops text from being read as a spreadsheet formula.
function cleanText(v, max) {
  var s = String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
