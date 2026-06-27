/**
 * World Cup 2026 Pool — backend.
 *
 * One sheet (whichever is "active" when this runs — matches the original
 * script's behavior) holds one row per participant:
 *   Timestamp | Name | Group A..L | Total | BracketPicks (JSON) | BracketSubmittedAt
 *
 * More sheets are created automatically as needed:
 *   Results               — Round | Idx | Winner   (actual knockout outcomes)
 *   GroupAdvancers        — Team                   (the 32 real Round of 32 teams)
 *   BracketPicksReadable  — decoded bracket picks per person (Pool Admin menu)
 *
 * Deploy: paste this whole file over Code.gs, save, then
 * Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 * Keep the same deployment so the existing web app URL keeps working.
 *
 * IMPORTANT: R32_SETUP below must be kept in sync with the DEFAULT array in
 * bracket.html — update both whenever the official Round of 32 matchups change.
 */

const GROUP_LETTERS = ['A','B','C','D','E','F','G','H','I','J','K','L'];

const HEADERS = ['Timestamp','Name',
  'Group A','Group B','Group C','Group D','Group E','Group F',
  'Group G','Group H','Group I','Group J','Group K','Group L','Total',
  'BracketPicks','BracketSubmittedAt'];

const PREV = { r16:'r32', qf:'r16', sf:'qf', final:'sf' };
const ROUND_COUNTS = { r32:16, r16:8, qf:4, sf:2, third:1, final:1 };
const ROUND_POINTS = { r32:2, r16:4, qf:8, sf:16, third:8, final:32 };

// Mirror of bracket.html's DEFAULT array — keep these two in sync.
const R32_SETUP = [
  ['Germany','Paraguay'],
  ['France','Sweden'],
  ['South Africa','Canada'],
  ['Netherlands','Morocco'],
  ['2K','2L'],
  ['Spain','2J'],
  ['United States','Bosnia and Herzegovina'],
  ['1G','3RD A/E/H/I/J'],
  ['Brazil','Japan'],
  ['Ivory Coast','Norway'],
  ['Mexico','3RD C/E/F/H/I'],
  ['1L','3RD E/H/I/J/K'],
  ['Argentina','Cape Verde'],
  ['Australia','2G'],
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureSchema_(sheet);
  const outSheet = getOrCreateSheet_('BracketPicksReadable', READABLE_HEADERS);

  const outLastRow = outSheet.getLastRow();
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

function getTeamServer_(round, idx, pos, results) {
  if (round === 'r32') return (R32_SETUP[idx] || ['', ''])[pos] || '';
  if (round === 'third') return getLoserServer_('sf', pos, results);
  const prev = PREV[round];
  return results[`${prev}_${idx * 2 + pos}`] || '';
}

function getLoserServer_(round, idx, results) {
  const t0 = getTeamServer_(round, idx, 0, results);
  const t1 = getTeamServer_(round, idx, 1, results);
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
  refreshBracketResults_();
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
        const t0 = getTeamServer_(round, idx, 0, results);
        const t1 = getTeamServer_(round, idx, 1, results);
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
