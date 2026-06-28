/**
 * World Cup 2026 Pool — backend.
 *
 * The sheet literally named "Sheet1" (see MAIN_SHEET_NAME/getMainSheet_)
 * holds one row per participant:
 *   Timestamp | Name | Group A..L | Total | BracketPicks (JSON) | BracketSubmittedAt
 *
 * More sheets are created automatically as needed:
 *   Results               — Round | Idx | Winner   (actual knockout outcomes)
 *   GroupAdvancers        — Team                   (the 32 real Round of 32 teams)
 *   R32Setup              — Idx | TeamA | TeamB     (live Round of 32 matchups)
 *   BracketPicksReadable  — decoded bracket picks per person (Pool Admin menu)
 *
 * Deploy: paste this whole file over Code.gs, save, then
 * Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 * Keep the same deployment so the existing web app URL keeps working.
 *
 * R32Setup auto-refreshes from ESPN (see refreshR32Setup_) — no manual edits
 * needed as placeholder slots (1G, 3RD A/E/H/I/J, etc.) resolve to real teams.
 * R32_SEED below is only the one-time bootstrap value used the first time the
 * R32Setup sheet is created; after that the sheet is the source of truth.
 * bracket.html independently fetches the same live setup via doGet?type=setup.
 */

const GROUP_LETTERS = ['A','B','C','D','E','F','G','H','I','J','K','L'];

const HEADERS = ['Timestamp','Name',
  'Group A','Group B','Group C','Group D','Group E','Group F',
  'Group G','Group H','Group I','Group J','Group K','Group L','Total',
  'BracketPicks','BracketSubmittedAt'];

const PREV = { r16:'r32', qf:'r16', sf:'qf', final:'sf' };
const ROUND_COUNTS = { r32:16, r16:8, qf:4, sf:2, third:1, final:1 };
const ROUND_POINTS = { r32:2, r16:4, qf:8, sf:16, third:8, final:32 };

// One-time bootstrap for the R32Setup sheet — see file header. Not used again
// once that sheet exists.
const R32_SEED = [
  ['Germany','Paraguay'],
  ['France','Sweden'],
  ['South Africa','Canada'],
  ['Netherlands','Morocco'],
  ['2K','2L'],
  ['Spain','2J'],
  ['United States','Bosnia and Herzegovina'],
  ['Belgium','3RD A/E/H/I/J'],
  ['Brazil','Japan'],
  ['Ivory Coast','Norway'],
  ['Mexico','3RD C/E/F/H/I'],
  ['1L','3RD E/H/I/J/K'],
  ['Argentina','Cape Verde'],
  ['Australia','Egypt'],
  ['Switzerland','3RD E/F/G/I/J'],
  ['1K','3RD D/E/I/J/L']
];

const TEAM_NAME_ALIASES = {
  'Bosnia-Herzegovina': 'Bosnia and Herzegovina',
  'Bosnia & Herzegovina': 'Bosnia and Herzegovina',
  'Bosnia-Herz.': 'Bosnia and Herzegovina',
  'Czechia': 'Czech Republic',
  'USA': 'United States',
  'U.S.': 'United States',
  'Cabo Verde': 'Cape Verde',
  'Türkiye': 'Turkey',
  'Congo DR': 'DR Congo'
};

function canonical_(name) {
  if (!name) return '';
  const trimmed = String(name).trim();
  return TEAM_NAME_ALIASES[trimmed] || trimmed;
}

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pool Admin')
    .addItem('🔄 Refresh Results from ESPN', 'refreshResultsFromEspn')
    .addItem('📋 Refresh Readable Bracket Picks', 'refreshReadableBracketPicks')
    .addToUi();
}

// ---------- Sheet helpers ----------

// The main participant data lives in this sheet specifically — never rely on
// getActiveSheet() for it, since that resolves to whichever tab a person last
// clicked into, not necessarily this one (this previously caused
// refreshReadableBracketPicks to wipe its own output when run while sitting
// on the BracketPicksReadable tab instead of this one).
const MAIN_SHEET_NAME = 'Sheet1';

function getMainSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(MAIN_SHEET_NAME) || ss.getSheets()[0];
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function ensureSchema_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    return;
  }
  const lastCol = sheet.getLastColumn();
  if (lastCol < HEADERS.length) {
    sheet.getRange(1, lastCol + 1, 1, HEADERS.length - lastCol).setValues([HEADERS.slice(lastCol)]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
}

function findRowByName_(sheet, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const target = String(name).trim().toLowerCase();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim().toLowerCase() === target) return i + 2;
  }
  return -1;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Submissions ----------

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sheet = getMainSheet_();
  ensureSchema_(sheet);

  const name = String(data.name || '').trim();
  if (!name) {
    return jsonOutput_({ success: false, error: 'Name required' });
  }

  const type = data.type || 'group';
  const rowNum = findRowByName_(sheet, name);

  if (type === 'bracket') {
    const bracketJSON = JSON.stringify(data.picks || {});
    if (rowNum > 0) {
      sheet.getRange(rowNum, 16, 1, 2).setValues([[bracketJSON, new Date()]]);
    } else {
      const row = new Array(HEADERS.length).fill('');
      row[0] = new Date();
      row[1] = name;
      row[15] = bracketJSON;
      row[16] = new Date();
      sheet.appendRow(row);
    }
  } else {
    const groupVals = GROUP_LETTERS.map(g => (data.picks[g] || []).join(', '));
    const total = GROUP_LETTERS.reduce((s, g) => s + (data.picks[g] || []).length, 0);
    if (rowNum > 0) {
      sheet.getRange(rowNum, 1, 1, 1).setValue(new Date());
      sheet.getRange(rowNum, 3, 1, 12).setValues([groupVals]);
      sheet.getRange(rowNum, 15, 1, 1).setValue(total);
    } else {
      const row = new Array(HEADERS.length).fill('');
      row[0] = new Date();
      row[1] = name;
      groupVals.forEach((v, i) => row[2 + i] = v);
      row[14] = total;
      sheet.appendRow(row);
    }
  }

  return jsonOutput_({ success: true });
}

// ---------- Leaderboard ----------

function doGet(e) {
  if (e && e.parameter && e.parameter.type === 'setup') {
    const setup = getR32Setup_();
    return jsonOutput_(setup.map((pair, idx) => ({ idx, teamA: pair[0], teamB: pair[1] })));
  }

  const sheet = getMainSheet_();
  ensureSchema_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonOutput_([]);

  const numCols = Math.max(sheet.getLastColumn(), HEADERS.length);
  const rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const results = getResultsMap_();
  const advancers = getGroupAdvancers_();

  const participants = rows.filter(r => r[1]).map(r => {
    const p = { name: r[1], submitted: r[0] };
    let groupPoints = 0;
    GROUP_LETTERS.forEach((g, i) => {
      const val = r[2 + i] || '';
      p[g] = val;
      String(val).split(',').map(s => s.trim()).filter(Boolean).forEach(team => {
        if (advancers.has(canonical_(team))) groupPoints++;
      });
    });

    let bracketPicks = {};
    try { bracketPicks = r[15] ? JSON.parse(r[15]) : {}; } catch (err) { bracketPicks = {}; }
    const breakdown = computeBracketBreakdown_(bracketPicks, results);
    const bracketPoints = Object.values(breakdown).reduce((s, v) => s + v, 0);

    p.groupPoints = groupPoints;
    p.r32Points = breakdown.r32;
    p.r16Points = breakdown.r16;
    p.qfPoints = breakdown.qf;
    p.sfPoints = breakdown.sf;
    p.thirdPoints = breakdown.third;
    p.finalPoints = breakdown.final;
    p.bracketPoints = bracketPoints;
    p.totalPoints = groupPoints + bracketPoints;
    return p;
  });

  participants.sort((a, b) => b.totalPoints - a.totalPoints);
  return jsonOutput_(participants);
}

function computeBracketBreakdown_(picks, results) {
  const breakdown = { r32: 0, r16: 0, qf: 0, sf: 0, third: 0, final: 0 };
  Object.keys(ROUND_COUNTS).forEach(round => {
    for (let idx = 0; idx < ROUND_COUNTS[round]; idx++) {
      const key = `${round}_${idx}`;
      const actual = results[key];
      const pick = picks[key];
      if (actual && pick && canonical_(pick) === actual) breakdown[round] += ROUND_POINTS[round];
    }
  });
  return breakdown;
}

// ---------- Readable bracket picks (Pool Admin menu) ----------
// BracketPicks is stored as raw JSON for scoring; this decodes it into a
// plain-language sheet so you can actually read what everyone picked.

const READABLE_ROUNDS = ['r32', 'r16', 'qf', 'sf', 'third', 'final'];
const READABLE_HEADERS = ['Name', 'R32 Picks', 'R16 Picks', 'QF Picks', 'SF Picks', '3rd Place Pick', 'Champion Pick'];

function refreshReadableBracketPicks() {
  const sheet = getMainSheet_();
  ensureSchema_(sheet);
  const outSheet = getOrCreateSheet_('BracketPicksReadable', READABLE_HEADERS);

  // Re-assert the header row and wipe any stray leftover columns past it
  // (this sheet previously got corrupted by a getActiveSheet() mixup that
  // wrote the main sheet's headers/data onto it — this makes the fix
  // self-healing instead of requiring a manual cleanup).
  outSheet.getRange(1, 1, 1, READABLE_HEADERS.length).setValues([READABLE_HEADERS]).setFontWeight('bold');
  const outLastRow = outSheet.getLastRow();
  const outLastCol = outSheet.getLastColumn();
  if (outLastCol > READABLE_HEADERS.length) {
    outSheet.getRange(1, READABLE_HEADERS.length + 1, Math.max(outLastRow, 1), outLastCol - READABLE_HEADERS.length).clearContent();
  }
  if (outLastRow > 1) {
    outSheet.getRange(2, 1, outLastRow - 1, READABLE_HEADERS.length).clearContent();
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const numCols = Math.max(sheet.getLastColumn(), HEADERS.length);
  const rows = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  const out = [];
  rows.forEach(r => {
    const name = r[1];
    if (!name) return;
    let bracketPicks = {};
    try { bracketPicks = r[15] ? JSON.parse(r[15]) : {}; } catch (err) { bracketPicks = {}; }
    if (Object.keys(bracketPicks).length === 0) return; // no bracket submission yet

    const cols = READABLE_ROUNDS.map(round => {
      const list = [];
      for (let idx = 0; idx < ROUND_COUNTS[round]; idx++) {
        const v = bracketPicks[`${round}_${idx}`];
        if (v) list.push(v);
      }
      return list.join(', ');
    });
    out.push([name, ...cols]);
  });

  if (out.length) {
    outSheet.getRange(2, 1, out.length, READABLE_HEADERS.length).setValues(out);
  }
}

// ---------- Bracket tree (mirrors bracket.html's getTeam/getLoser) ----------

function getTeamServer_(round, idx, pos, results, r32Setup) {
  if (round === 'r32') return (r32Setup[idx] || ['', ''])[pos] || '';
  if (round === 'third') return getLoserServer_('sf', pos, results, r32Setup);
  const prev = PREV[round];
  return results[`${prev}_${idx * 2 + pos}`] || '';
}

function getLoserServer_(round, idx, results, r32Setup) {
  const t0 = getTeamServer_(round, idx, 0, results, r32Setup);
  const t1 = getTeamServer_(round, idx, 1, results, r32Setup);
  const winner = results[`${round}_${idx}`];
  if (!winner) return '';
  return winner === t0 ? t1 : t0;
}

function getResultsMap_() {
  const sheet = getOrCreateSheet_('Results', ['Round', 'Idx', 'Winner']);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(r => {
      if (r[0] !== '' && r[2]) map[`${r[0]}_${r[1]}`] = canonical_(r[2]);
    });
  }
  return map;
}

function getR32Setup_() {
  const sheet = getOrCreateSheet_('R32Setup', ['Idx', 'TeamA', 'TeamB']);
  if (sheet.getLastRow() <= 1) {
    const rows = R32_SEED.map((pair, idx) => [idx, pair[0], pair[1]]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  const lastRow = sheet.getLastRow();
  const setup = R32_SEED.map(pair => pair.slice());
  sheet.getRange(2, 1, lastRow - 1, 3).getValues().forEach(r => {
    const idx = Number(r[0]);
    if (idx >= 0 && idx < setup.length) setup[idx] = [String(r[1] || ''), String(r[2] || '')];
  });
  return setup;
}

function getGroupAdvancers_() {
  const sheet = getOrCreateSheet_('GroupAdvancers', ['Team']);
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(r => {
      if (r[0]) set.add(canonical_(r[0]));
    });
  }
  return set;
}

// ---------- ESPN auto-pull ----------
// Run manually from the Pool Admin menu, or add a time-driven trigger
// (Triggers icon in the left sidebar > Add Trigger > refreshResultsFromEspn).

function refreshResultsFromEspn() {
  refreshGroupAdvancers_();
  refreshR32Setup_();
  refreshBracketResults_();
}

// ESPN shows still-undetermined Round of 32 sides as generic text like
// "Group K Winner" or "Third Place Group A/E/H/I/J" — translate that into the
// same placeholder-code notation bracket.html/R32_SEED already use, so a
// fixture can be matched to our slot whether or not either side has resolved
// to a real team name yet.
function translateEspnPlaceholder_(name) {
  if (!name) return '';
  let m = name.match(/^Group ([A-L]) Winner$/);
  if (m) return `1${m[1]}`;
  m = name.match(/^Group ([A-L]) 2nd Place$/);
  if (m) return `2${m[1]}`;
  m = name.match(/^Third Place Group (.+)$/);
  if (m) return `3RD ${m[1]}`;
  return canonical_(name);
}

// Keep whichever side already matches our current TeamA in position A, so a
// slot's left/right (top/bottom) assignment stays stable across refreshes.
function alignPair_(fa, fb, ca, cb) {
  if (fa === ca || fb === cb) return [fa, fb];
  if (fa === cb || fb === ca) return [fb, fa];
  return null;
}

function refreshR32Setup_() {
  const sheet = getOrCreateSheet_('R32Setup', ['Idx', 'TeamA', 'TeamB']);
  const current = getR32Setup_();

  const res = UrlFetchApp.fetch(
    'https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard?dates=20260628-20260703',
    { muteHttpExceptions: true }
  );
  const json = JSON.parse(res.getContentText());

  (json.events || []).forEach(ev => {
    const comp = ev.competitions && ev.competitions[0];
    const competitors = comp && comp.competitors;
    if (!competitors || competitors.length !== 2) return;

    const fa = translateEspnPlaceholder_(competitors[0].team.displayName);
    const fb = translateEspnPlaceholder_(competitors[1].team.displayName);

    for (let idx = 0; idx < current.length; idx++) {
      const [ca, cb] = current[idx];
      const aligned = alignPair_(fa, fb, ca, cb);
      if (aligned) { current[idx] = aligned; break; }
    }
  });

  sheet.getRange(2, 1, current.length, 3).setValues(current.map((pair, idx) => [idx, pair[0], pair[1]]));
}

function refreshGroupAdvancers_() {
  const res = UrlFetchApp.fetch(
    'https://site.api.espn.com/apis/v2/sports/soccer/fifa.world/standings',
    { muteHttpExceptions: true }
  );
  const json = JSON.parse(res.getContentText());
  const advancers = [];
  (json.children || []).forEach(group => {
    const entries = (group.standings && group.standings.entries) || [];
    entries.forEach(entry => {
      // Only trust ESPN's explicit "advanced" stat (0 or 1) — note.description
      // text like "Best 8 advance" describes a team still in contention for a
      // 3rd-place slot, not a confirmed result, even though it contains the
      // word "advance".
      const advancedStat = (entry.stats || []).find(s => s.name === 'advanced');
      if (advancedStat && advancedStat.value === 1) {
        advancers.push(canonical_(entry.team.displayName));
      }
    });
  });

  const sheet = getOrCreateSheet_('GroupAdvancers', ['Team']);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
  if (advancers.length) sheet.getRange(2, 1, advancers.length, 1).setValues(advancers.map(a => [a]));
}

function refreshBracketResults_() {
  const resultsSheet = getOrCreateSheet_('Results', ['Round', 'Idx', 'Winner']);
  const results = getResultsMap_();
  const r32Setup = getR32Setup_();

  const res = UrlFetchApp.fetch(
    'https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard?dates=20260628-20260720',
    { muteHttpExceptions: true }
  );
  const json = JSON.parse(res.getContentText());
  const events = (json.events || []).slice().sort((a, b) => new Date(a.date) - new Date(b.date));

  const fixtures = [];
  events.forEach(ev => {
    const comp = ev.competitions && ev.competitions[0];
    if (!comp || !comp.status || !comp.status.type || !comp.status.type.completed) return;
    const competitors = comp.competitors || [];
    if (competitors.length !== 2) return;
    const winner = competitors.find(c => c.winner);
    if (!winner) return;
    fixtures.push({
      teams: competitors.map(c => canonical_(c.team.displayName)),
      winner: canonical_(winner.team.displayName)
    });
  });

  const rounds = ['r32', 'r16', 'qf', 'sf', 'third', 'final'];
  let changed = true;
  let guard = 0;
  while (changed && guard < 6) {
    changed = false;
    guard++;
    rounds.forEach(round => {
      for (let idx = 0; idx < ROUND_COUNTS[round]; idx++) {
        const key = `${round}_${idx}`;
        if (results[key]) continue;
        const t0 = getTeamServer_(round, idx, 0, results, r32Setup);
        const t1 = getTeamServer_(round, idx, 1, results, r32Setup);
        if (!t0 || !t1) continue;
        const match = fixtures.find(f => f.teams.includes(t0) && f.teams.includes(t1));
        if (match) {
          results[key] = match.winner;
          resultsSheet.appendRow([round, idx, match.winner]);
          changed = true;
        }
      }
    });
  }
}
