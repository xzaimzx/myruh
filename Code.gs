/**
 * MYRUH Padelers Tour – Google Sheets backend
 *
 * Paste this whole file into Extensions → Apps Script in your Google Sheet.
 * 1. Choose "setup" in the function dropdown and click Run (once only).
 * 2. Deploy → New deployment → Web app, Execute as: Me, Who has access: Anyone.
 * 3. Copy the Web app URL (ends in /exec) into index.html.
 *
 * Tabs: Players, Stops, Matches, ScoreLog, Settings, Rules, Comments, Captains.
 * Comments: delete a row in the Comments tab to remove that comment from the website.
 * Captains: one row per Court Captain with their name and their own Host PIN.
 *   Delete a row to take away that captain's access.
 * Running setup again later is safe: it only adds tabs or columns that are missing.
 * PINs live in the Settings tab and are never sent to the website.
 */

var TABS = {
  Players: ['id', 'name', 'gender', 'joined'],
  Stops: ['id', 'name', 'date', 'venue', 'total', 'created', 'time', 'level', 'format', 'host', 'courts', 'rounds', 'players', 'setup', 'closes', 'link'],
  Matches: ['id', 'stop', 'cup', 'round', 'time', 'court', 'a1', 'a2', 'b1', 'b2', 'sa', 'sb', 'seq', 'enteredBy', 'updated', 'sets', 'subs', 'wo'],
  ScoreLog: ['time', 'name', 'matchId', 'match', 'oldScore', 'newScore'],
  Settings: ['key', 'value', 'note'],
  Rules: ['rule'],
  Comments: ['id', 'article', 'time', 'name', 'text'],
  Visitors: ['time', 'visitor', 'city', 'country', 'lat', 'lon'],
  Captains: ['name', 'pin', 'note']
};

// Event levels a Court Captain can pick, and the formats each one allows. Only the organiser can run a Major.
var LEVELS = {
  social: ['americano', 'mexicano'],
  team: ['team', 'mexicano'],
  match: ['short', 'full'],
  major: ['americano', 'mexicano', 'team', 'short', 'full']
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
      ['qayyum', 'Qayyum Z', 'M'], ['amir', 'Amir AJ', 'M'], ['derek', 'Derek P', 'M'], ['du', 'Redzuan (Du)', 'M'],
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
    du: 'Redzuan (Du)', amirul: 'Amirul I'
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
    case 'comment': return addComment(req);
    case 'checkScorePin': return checkPin('ScorePIN', req.pin) ? { ok: true } : { ok: false, error: "That Score PIN didn't work." };
    case 'checkAdmin': return admin(req) || { ok: true };
    case 'addPlayer': return (captain(req).name ? null : admin(req)) || addPlayer(req);
    case 'checkHost': return checkHost(req);
    case 'createEvent': return createEvent(req);
    case 'addRound': return addRound(req);
    case 'setPlayers': return setPlayers(req);
    case 'swapPlayer': return swapPlayer(req);
    case 'walkover': return walkover(req);
    case 'visit': return addVisit(req);
    case 'deleteEvent': return deleteEvent(req);
    case 'linkEvents': return linkEvents(req);
    case 'setCourtNames': return setCourtNames(req);
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
  var sets = req.sets ? String(req.sets).replace(/[^0-9\- ]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40) : '';
  if (sa === undefined || sb === undefined) return { ok: false, error: 'Scores must be whole numbers.' };
  if ((sa === '') !== (sb === '')) return { ok: false, error: 'Enter both scores, or clear both.' };

  var sh = sheet('Matches'), row = findRow(sh, req.matchId);
  if (!row) return { ok: false, error: 'That match no longer exists. Refresh the page.' };
  var cols = TABS.Matches, vals = sh.getRange(row, 1, 1, cols.length).getValues()[0];
  var m = {}; cols.forEach(function (c, i) { m[c] = String(vals[i]); });
  if (m.wo === 'a' || m.wo === 'b') return { ok: false, error: 'This game is a walkover. Ask the Court Captain to undo it first.' };
  var oldScore = m.sa === '' ? '–' : m.sa + '–' + m.sb;
  var newScore = sa === '' ? '–' : sa + '–' + sb;
  if (sets) newScore += ' (' + sets + ')';
  if (oldScore === newScore && (m.sets || '') === sets) return { ok: true, match: m, unchanged: true };

  var now = new Date().toISOString();
  m.sa = String(sa); m.sb = String(sb); m.enteredBy = sa === '' ? '' : name; m.updated = now; m.sets = sa === '' ? '' : sets;
  sh.getRange(row, cols.indexOf('sa') + 1, 1, 2).setValues([[m.sa, m.sb]]);
  sh.getRange(row, cols.indexOf('enteredBy') + 1, 1, 3).setValues([[m.enteredBy, now, m.sets]]);

  var names = playerNames();
  var desc = m.cup + ' R' + m.round + ' C' + m.court + ': ' + names(m.a1) + '/' + names(m.a2) + ' v ' + names(m.b1) + '/' + names(m.b2);
  sheet('ScoreLog').appendRow([now, name, m.id, desc, oldScore, newScore]);
  return { ok: true, match: m };
}

/* Court Captains --------------------------------------------------- */

// Returns {name, admin} for a valid Host PIN (or the Admin PIN), otherwise {}.
function captain(req) {
  if (req.adminPin && checkPin('AdminPIN', req.adminPin)) return { name: 'Organiser', admin: true };
  var pin = String(req.hostPin || '').trim();
  if (!pin) return {};
  var list = rows('Captains');
  for (var i = 0; i < list.length; i++) if (list[i].pin !== '' && list[i].pin === pin) return { name: list[i].name, admin: false };
  return {};
}

// Slows down guessing Host PINs the same way as the Admin PIN.
function host(req) {
  var cache = CacheService.getScriptCache();
  var fails = parseInt(cache.get('hostFails') || '0', 10);
  if (fails >= 20) return { error: { ok: false, error: 'Too many wrong Host PINs. Wait 15 minutes and try again.' } };
  var c = captain(req);
  if (c.name) return c;
  cache.put('hostFails', String(fails + 1), 900);
  return { error: { ok: false, error: "That Host PIN didn't work. Ask the organiser for yours.", badPin: true } };
}

function checkHost(req) {
  var c = host(req);
  return c.error || { ok: true, name: c.name, admin: c.admin };
}

function ownsStop(c, stop) {
  return c.admin || (stop && String(stop.host).toLowerCase() === String(c.name).toLowerCase());
}

function matchRows(stopId, list, playerIds, cupIdx) {
  var out = [], now = Date.now().toString(36);
  for (var i = 0; i < list.length; i++) {
    var m = list[i], ids = [m.a1, m.a2, m.b1, m.b2].map(String);
    if (ids.some(function (x) { return playerIds.indexOf(x) < 0; })) throw new Error('Every player in a game must be on the event list.');
    if (ids.filter(function (x, j) { return ids.indexOf(x) === j; }).length !== 4) throw new Error('A player can only appear once in a game.');
    var round = parseInt(m.round, 10) || 1, court = parseInt(m.court, 10) || 1;
    out.push(['m-' + now + i.toString(36) + Math.random().toString(36).slice(2, 5), stopId, cleanText(m.cup, 40) || 'Cup', String(round),
      cleanText(m.time, 10), String(court), ids[0], ids[1], ids[2], ids[3], '', '', String(cupIdx * 1000 + round * 10 + court), '', '', '', '', '']);
  }
  return out;
}

function createEvent(req) {
  var c = host(req);
  if (c.error) return c.error;
  var level = String(req.level || ''), format = String(req.format || '');
  if (!LEVELS[level] || LEVELS[level].indexOf(format) < 0) return { ok: false, error: 'Pick a level and a format.' };
  if (level === 'major' && !c.admin) return { ok: false, error: 'Only the organiser can run a Major.' };
  var name = cleanText(req.name, 60);
  if (!name) return { ok: false, error: 'Give the event a name.' };
  if (rows('Stops').some(function (s) { return s.name.toLowerCase() === name.toLowerCase(); })) return { ok: false, error: name + ' already exists. Change the number.' };
  var known = rows('Players').map(function (p) { return p.id; });
  var players = (req.players || []).map(String).filter(function (x, i, a) { return known.indexOf(x) >= 0 && a.indexOf(x) === i; });
  var list = req.matches || [];
  if (list.length && players.length < 6) return { ok: false, error: 'A ranked event needs at least 6 players.' };
  var date = /^\d{4}-\d{2}-\d{2}$/.test(req.date || '') ? req.date : '';
  var total = parseInt(req.total, 10), courts = parseInt(req.courts, 10) || 1, planned = parseInt(req.rounds, 10) || '';
  var id = 'stop-' + Date.now().toString(36);
  var closes = /^\d{4}-\d{2}-\d{2}$/.test(req.closes || '') ? req.closes : '';
  var out;
  try { out = matchRows(id, list, players, 0); } catch (err) { return { ok: false, error: err.message }; }
  sheet('Stops').appendRow([id, name, date, cleanText(req.venue, 60), isNaN(total) ? '' : String(total), new Date().toISOString(),
    cleanText(req.time, 40), level, format, c.name, String(courts), String(planned), players.join(','), setupText(req.setup, players), closes]);
  if (out.length) {
    var sh = sheet('Matches');
    sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  }
  sheet('ScoreLog').appendRow([new Date().toISOString(), c.name, id, 'Created ' + name + (out.length ? ' (' + out.length + ' games)' : ' (sign-ups open)'), '–', 'created']);
  return { ok: true, id: id };
}

// How players are grouped: split into men and ladies, and pairs that stay together all event.
function setupText(setup, players) {
  setup = setup || {};
  var used = {}, fixed = [];
  (setup.fixed || []).forEach(function (pr) {
    var a = String(pr[0]), b = String(pr[1]);
    if (a !== b && players.indexOf(a) >= 0 && players.indexOf(b) >= 0 && !used[a] && !used[b]) { used[a] = used[b] = 1; fixed.push([a, b]); }
  });
  var hm = /^\d{1,2}:\d{2}$/, times = (setup.courtTimes || []).slice(0, 8).filter(function (t) { return t && hm.test(t[0]) && hm.test(t[1]); }).map(function (t) { return [t[0], t[1]]; });
  var sets = setup.sets === '1' || setup.sets === '3' ? setup.sets : '';
  var names = (setup.courtNames || []).slice(0, 4).map(function (n) { return cleanText(n, 12); });
  if (!names.some(String)) names = [];
  if (!setup.split && !setup.mixed && !setup.rotate && !fixed.length && !times.length && !sets && !setup.tbd && !names.length) return '';
  var o = { split: !!setup.split, mixed: !setup.split && !!setup.mixed, rotate: !!setup.rotate, fixed: fixed };
  if (times.length) o.courtTimes = times;
  if (sets) o.sets = sets;
  if (setup.tbd) o.tbd = true;   // format to be picked once sign-ups close
  if (names.length) o.courtNames = names;   // the courts' real numbers or names at the venue
  return JSON.stringify(o);
}

function parseSetup(text) {
  try { var o = JSON.parse(text); return { split: !!o.split, mixed: !!o.mixed, rotate: !!o.rotate, fixed: o.fixed || [], courtTimes: o.courtTimes || null, sets: o.sets || '', tbd: !!o.tbd, courtNames: o.courtNames || [] }; } catch (e) { return { split: false, fixed: [] }; }
}

// One row per visit: a random id for the browser (not a person), and the rough city from the visitor's connection. No IP addresses are stored.
function addVisit(req) {
  var vid = String(req.visitor || '').replace(/[^a-z0-9]/gi, '').slice(0, 24);
  if (!vid) return { ok: false, error: 'Bad request.' };
  var num = function (x, lim) { var n = parseFloat(x); return isNaN(n) || Math.abs(n) > lim ? '' : String(Math.round(n * 100) / 100); };
  ensureTab('Visitors').appendRow([new Date().toISOString(), vid, cleanText(req.city, 40), cleanText(req.country, 40), num(req.lat, 90), num(req.lon, 180)]);
  return { ok: true };
}

// Totals for the homepage: visits, different visitors, and cities with 2 or more visitors (so nobody can be singled out).
function visitStats() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visitors');
  if (!sh) return { visits: 0, visitors: 0, cities: [] };
  var list = rows('Visitors'), seen = {}, byCity = {};
  list.forEach(function (v) {
    seen[v.visitor] = 1;
    if (!v.city) return;
    var k = v.city + '|' + v.country;
    var c = byCity[k] || (byCity[k] = { city: v.city, country: v.country, lat: parseFloat(v.lat), lon: parseFloat(v.lon), who: {} });
    c.who[v.visitor] = 1;
  });
  var cities = Object.keys(byCity).map(function (k) { var c = byCity[k]; return { city: c.city, country: c.country, lat: c.lat, lon: c.lon, n: Object.keys(c.who).length }; })
    .filter(function (c) { return c.n >= 2 && !isNaN(c.lat) && !isNaN(c.lon); })
    .sort(function (a, b) { return b.n - a.n; });
  return { visits: list.length, visitors: Object.keys(seen).length, cities: cities };
}

// Sets or changes an event's players and schedule. Scored games are kept; unscored games from fromRound on are replaced.
function setPlayers(req) {
  var c = host(req);
  if (c.error) return c.error;
  var sh = sheet('Stops'), row = findRow(sh, req.stop);
  var stop = rows('Stops').filter(function (s) { return s.id === String(req.stop); })[0];
  if (!row || !stop) return { ok: false, error: 'That event no longer exists.' };
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can change this event.' };
  var known = rows('Players').map(function (p) { return p.id; });
  var players = (req.players || []).map(String).filter(function (x, i, a) { return known.indexOf(x) >= 0 && a.indexOf(x) === i; });
  if (players.length < 6) return { ok: false, error: 'A ranked event needs at least 6 players.' };
  var list = req.matches || [], from = parseInt(req.fromRound, 10) || 1;
  if (!list.length) return { ok: false, error: 'Make the schedule first.' };
  // A last-minute change of level or format, allowed until the first game is scored
  var level = req.level ? String(req.level) : '', format = req.format ? String(req.format) : '', newName = '';
  if (level || format) {
    if (!LEVELS[level] || LEVELS[level].indexOf(format) < 0) return { ok: false, error: 'Pick a level and a format.' };
    if (level === 'major' && stop.level !== 'major' && !c.admin) return { ok: false, error: 'Only the organiser can run a Major.' };
    if (rows('Matches').some(function (m) { return m.stop === stop.id && String(m.sa) !== ''; })) return { ok: false, error: 'A game is already scored, so the format can no longer change. Refresh the page.' };
    newName = cleanText(req.name, 60);
    if (newName && newName.toLowerCase() !== stop.name.toLowerCase() && rows('Stops').some(function (s) { return s.name.toLowerCase() === newName.toLowerCase(); })) newName = '';
  }
  var ms = sheet('Matches'), cols = TABS.Matches, n = ms.getLastRow() - 1;
  var vals = n > 0 ? ms.getRange(2, 1, n, cols.length).getValues() : [];
  var iStop = cols.indexOf('stop'), iRound = cols.indexOf('round'), iSa = cols.indexOf('sa'), drop = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][iStop]) !== stop.id || (parseInt(vals[i][iRound], 10) || 1) < from) continue;
    if (String(vals[i][iSa]) !== '') return { ok: false, error: 'Round ' + vals[i][iRound] + ' already has scores. Refresh the page.' };
    drop.push(i + 2);
  }
  var out;
  try { out = matchRows(stop.id, list, players, 0); } catch (err) { return { ok: false, error: err.message }; }
  for (var j = drop.length - 1; j >= 0; j--) ms.deleteRow(drop[j]);
  ms.getRange(ms.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  var sc = TABS.Stops, set = function (k, v) { sh.getRange(row, sc.indexOf(k) + 1).setValue(v); };
  set('players', players.join(','));
  set('setup', setupText(req.setup, players));
  if (req.rounds !== undefined) set('rounds', String(parseInt(req.rounds, 10) || ''));
  if (req.courts) set('courts', String(parseInt(req.courts, 10) || 1));
  if (req.time) set('time', cleanText(req.time, 40));
  if (level) {
    set('level', level); set('format', format);
    var total = parseInt(req.total, 10); set('total', isNaN(total) ? '' : String(total));
    if (newName) set('name', newName);
  }
  sheet('ScoreLog').appendRow([new Date().toISOString(), c.name, stop.id, 'Set ' + players.length + ' players, ' + out.length + ' games from round ' + from, '–', 'players']);
  return { ok: true };
}

// Swaps an injured or tired player out of unscored games. Only those games change; the fill-in is marked so the
// game counts for the Tour Rankings but not for their event total. With leave, the player is also taken off the event.
function swapPlayer(req) {
  var c = host(req);
  if (c.error) return c.error;
  var sh = sheet('Stops'), row = findRow(sh, req.stop);
  var stop = rows('Stops').filter(function (s) { return s.id === String(req.stop); })[0];
  if (!row || !stop) return { ok: false, error: 'That event no longer exists.' };
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can change this event.' };
  var out = String(req.out || ''), swaps = req.swaps || [];
  var known = rows('Players').map(function (p) { return p.id; });
  if (!swaps.length) return { ok: false, error: 'Pick who comes in.' };
  var ms = sheet('Matches'), cols = TABS.Matches, n = ms.getLastRow() - 1;
  var vals = n > 0 ? ms.getRange(2, 1, n, cols.length).getValues() : [];
  var ix = function (k) { return cols.indexOf(k); }, P = ['a1', 'a2', 'b1', 'b2'].map(ix);
  var plan = [];
  for (var k = 0; k < swaps.length; k++) {
    var sw = swaps[k], inId = String(sw['in'] || '');
    if (known.indexOf(inId) < 0) return { ok: false, error: 'Pick who comes in for every game.' };
    var i = -1;
    for (var j = 0; j < vals.length; j++) if (String(vals[j][0]) === String(sw.id) && String(vals[j][ix('stop')]) === stop.id) i = j;
    if (i < 0) return { ok: false, error: 'A game has changed. Refresh the page.' };
    if (String(vals[i][ix('sa')]) !== '') return { ok: false, error: 'Round ' + vals[i][ix('round')] + ' is already scored. Refresh the page.' };
    var pos = P.filter(function (p) { return String(vals[i][p]) === out; })[0];
    if (pos === undefined) return { ok: false, error: 'A game has changed. Refresh the page.' };
    var round = String(vals[i][ix('round')]);
    var busy = vals.some(function (v, j) { return j !== i && String(v[ix('stop')]) === stop.id && String(v[ix('round')]) === round && P.some(function (p) { return String(v[p]) === inId; }); });
    if (busy || P.some(function (p) { return String(vals[i][p]) === inId; })) return { ok: false, error: 'The fill-in for round ' + round + ' is already playing that round.' };
    plan.push({ row: i + 2, pos: pos, inId: inId, subs: String(vals[i][ix('subs')] || '') });
  }
  plan.forEach(function (x) {
    ms.getRange(x.row, x.pos + 1).setValue(x.inId);
    var subs = x.subs ? x.subs.split(',') : []; if (subs.indexOf(x.inId) < 0) subs.push(x.inId);
    ms.getRange(x.row, ix('subs') + 1).setValue(subs.join(','));
  });
  var players = String(stop.players || '').split(',').filter(String);
  plan.forEach(function (x) { if (players.indexOf(x.inId) < 0) players.push(x.inId); });
  if (req.leave) players = players.filter(function (id) { return id !== out; });
  sh.getRange(row, TABS.Stops.indexOf('players') + 1).setValue(players.join(','));
  sheet('ScoreLog').appendRow([new Date().toISOString(), c.name, stop.id, 'Swapped ' + out + ' out of ' + plan.length + (plan.length === 1 ? ' game' : ' games'), '–', 'swap']);
  return { ok: true };
}

// A walkover: one pair didn't turn up. The pair that did wins the full score to 0 (24–0, 6–0, or 6-0 6-0), which
// counts for the event standings; the website leaves the game out of the Tour Rankings. side '' undoes it.
function walkover(req) {
  var c = host(req);
  if (c.error) return c.error;
  var sh = sheet('Matches'), row = findRow(sh, req.matchId);
  if (!row) return { ok: false, error: 'That game no longer exists. Refresh the page.' };
  var cols = TABS.Matches, vals = sh.getRange(row, 1, 1, cols.length).getValues()[0];
  var m = {}; cols.forEach(function (k, i) { m[k] = String(vals[i]); });
  var stop = rows('Stops').filter(function (s) { return s.id === m.stop; })[0];
  if (!stop) return { ok: false, error: 'That event no longer exists.' };
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can change this event.' };
  var side = req.side === 'a' || req.side === 'b' ? req.side : '';
  var sa = '', sb = '', sets = '';
  if (side) {
    var full = stop.format === 'full', games = full || stop.format === 'short';
    var win = games ? 6 : (parseInt(stop.total, 10) || 21);
    if (full) { var setup = parseSetup(stop.setup), n = setup.sets === '1' ? 1 : 2; win = 6 * n; sets = n === 1 ? '6-0' : '6-0 6-0'; }
    if (side === 'a') { sa = 0; sb = win; if (sets) sets = sets.replace(/6-0/g, '0-6'); } else { sa = win; sb = 0; }
  } else if (m.wo !== 'a' && m.wo !== 'b') return { ok: true, unchanged: true };
  var now = new Date().toISOString(), old = m.sa === '' ? '–' : m.sa + '–' + m.sb;
  sh.getRange(row, cols.indexOf('sa') + 1, 1, 2).setValues([[String(sa), String(sb)]]);
  sh.getRange(row, cols.indexOf('enteredBy') + 1, 1, 3).setValues([[side ? c.name : '', now, sets]]);
  sh.getRange(row, cols.indexOf('wo') + 1).setValue(side);
  var names = playerNames(), no = side ? (side === 'a' ? [m.a1, m.a2] : [m.b1, m.b2]) : [];
  sheet('ScoreLog').appendRow([now, c.name, m.id, side ? 'Walkover: ' + no.map(names).join('/') + " didn't show" : 'Walkover undone', old, side ? sa + '–' + sb : '–']);
  return { ok: true };
}

// Adds the next round of a Mexicano event once the captain has made it on the website.
function addRound(req) {
  var c = host(req);
  if (c.error) return c.error;
  var stop = rows('Stops').filter(function (s) { return s.id === String(req.stop); })[0];
  if (!stop) return { ok: false, error: 'That event no longer exists.' };
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can add rounds to this event.' };
  var round = parseInt(req.round, 10);
  if (rows('Matches').some(function (m) { return m.stop === stop.id && parseInt(m.round, 10) === round; })) return { ok: false, error: 'Round ' + round + ' is already there. Refresh the page.' };
  var out;
  try { out = matchRows(stop.id, req.matches || [], String(stop.players).split(','), 0); } catch (err) { return { ok: false, error: err.message }; }
  if (!out.length) return { ok: false, error: 'No games to add.' };
  var sh = sheet('Matches');
  sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
  return { ok: true };
}

// Renames an event's courts (e.g. Court 5 and Court 7) without touching its schedule.
function setCourtNames(req) {
  var c = host(req);
  if (c.error) return c.error;
  var stop = rows('Stops').filter(function (s) { return s.id === String(req.stop); })[0];
  if (!stop) return { ok: false, error: 'That event no longer exists.' };
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can rename the courts.' };
  var setup = {}; try { setup = JSON.parse(stop.setup) || {}; } catch (e) {}
  setup.courtNames = req.courtNames || [];
  var sh = sheet('Stops');
  sh.getRange(findRow(sh, stop.id), TABS.Stops.indexOf('setup') + 1).setValue(setupText(setup, stop.players ? stop.players.split(',') : []));
  return { ok: true };
}

// Links events on the same day so they share Day champions. Unlinking clears the whole group.
function linkEvents(req) {
  var c = host(req);
  if (c.error) return c.error;
  ensureHeader('Stops');
  var all = rows('Stops'), ids = (req.stops || []).map(String);
  if (req.unlink) {
    var one = all.filter(function (s) { return s.id === ids[0]; })[0];
    if (!one || !one.link) return { ok: true };
    ids = all.filter(function (s) { return s.link === one.link; }).map(function (s) { return s.id; });
  }
  var picked = all.filter(function (s) { return ids.indexOf(s.id) >= 0; });
  if (!req.unlink && picked.length < 2) return { ok: false, error: 'Pick at least 2 events to link.' };
  if (picked.some(function (s) { return s.date !== picked[0].date; })) return { ok: false, error: 'Only events on the same day can be linked.' };
  var notMine = picked.filter(function (s) { return !ownsStop(c, s); })[0];
  if (notMine) return { ok: false, error: 'Only ' + notMine.host + ' or the organiser can link ' + notMine.name + '.' };
  var sh = sheet('Stops'), col = TABS.Stops.indexOf('link') + 1, link = req.unlink ? '' : picked[0].id;
  picked.forEach(function (s) { sh.getRange(findRow(sh, s.id), col).setValue(link); });
  sheet('ScoreLog').appendRow([new Date().toISOString(), c.name, picked[0].id, (req.unlink ? 'Unlinked ' : 'Linked ') + picked.map(function (s) { return s.name; }).join(' + '), '–', 'link']);
  return { ok: true };
}

function deleteEvent(req) {
  var c = host(req);
  if (c.error) return c.error;
  var sh = sheet('Stops'), row = findRow(sh, req.stop);
  if (!row) return { ok: true };
  var stop = rows('Stops').filter(function (s) { return s.id === String(req.stop); })[0];
  if (!ownsStop(c, stop)) return { ok: false, error: 'Only ' + stop.host + ' or the organiser can delete this event.' };
  var ms = sheet('Matches'), n = ms.getLastRow() - 1;
  if (n > 0) {
    var stops = ms.getRange(2, 2, n, 1).getValues();
    for (var i = stops.length - 1; i >= 0; i--) if (String(stops[i][0]) === stop.id) ms.deleteRow(i + 2);
  }
  sh.deleteRow(row);
  sheet('ScoreLog').appendRow([new Date().toISOString(), c.name, stop.id, 'Deleted event ' + stop.name, '–', 'deleted']);
  return { ok: true };
}

function addComment(req) {
  var name = cleanText(req.name, 40);
  if (!name) return { ok: false, error: 'Type your name first.' };
  if (!checkPin('ScorePIN', req.pin)) return { ok: false, error: "That Score PIN didn't work.", badPin: true };
  var article = cleanText(req.article, 60), text = cleanText(req.text, 500);
  if (!article) return { ok: false, error: 'Bad request.' };
  if (!text) return { ok: false, error: 'Write something first.' };
  var c = { id: 'c-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), article: article, time: new Date().toISOString(), name: name, text: text };
  ensureTab('Comments').appendRow([c.id, c.article, c.time, c.name, c.text]);
  return { ok: true, comment: c };
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
  ensureHeader('Stops'); ensureHeader('Matches'); ensureTab('Captains');
  var players = rows('Players').map(function (p) { return { id: p.id, name: p.name, gender: p.gender === 'F' ? 'F' : 'M' }; });
  var stops = rows('Stops').map(function (s) {
    return {
      id: s.id, name: s.name, date: s.date, venue: s.venue, total: num(s.total), created: s.created, time: s.time,
      level: s.level, format: s.format, host: s.host, courts: num(s.courts), rounds: num(s.rounds), players: s.players ? s.players.split(',') : [], setup: parseSetup(s.setup), closes: s.closes || '', link: s.link || ''
    };
  });
  var matches = rows('Matches').map(function (m) {
    return {
      id: m.id, stop: m.stop, cup: m.cup, round: num(m.round) || 1, time: m.time, court: num(m.court),
      a: [m.a1, m.a2], b: [m.b1, m.b2], sa: num(m.sa), sb: num(m.sb), seq: num(m.seq) || 0,
      enteredBy: m.enteredBy, updated: m.updated, sets: m.sets, subs: m.subs ? String(m.subs).split(',') : [], wo: m.wo === 'a' || m.wo === 'b' ? m.wo : ''
    };
  });
  var ruleTab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rules');
  var rules = ruleTab ? rows('Rules').map(function (r) { return r.rule; }).filter(String) : null;
  ensureTab('Comments');
  var comments = rows('Comments').map(function (c) { return { id: c.id, article: c.article, time: c.time, name: c.name, text: c.text.replace(/^'(?=[=+\-@])/, '') }; });
  return { ok: true, players: players, stops: stops, matches: matches, rules: rules, comments: comments, visits: visitStats() };
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

// Adds a tab with its header row if it doesn't exist yet, so new features work without running setup again.
function ensureTab(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(1, 1, sh.getMaxRows(), TABS[name].length).setNumberFormat('@');
  return sh;
}

// Adds header names for columns that newer versions of this script use.
function ensureHeader(name) {
  var sh = sheet(name), cols = TABS[name];
  if (sh.getLastColumn() >= cols.length) return;
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
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
