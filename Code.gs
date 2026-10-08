// ============================================================
//  INITIATIVE ROLLER — Google Apps Script Backend
//  Stores all data in the bound Google Sheet
// ============================================================

const SHEET_NAMES = {
  CONFIG:            'Config',
  MEMBERS:           'Members',
  SESSIONS:          'Sessions',
  ROLLS:             'Rolls',
  SESSION_STATE:     'SessionState',
  SESSION_ROLLS:     'SessionRolls',
  PAST_LEADERBOARDS: 'PastLeaderboards',
  ROLLOFF_STATS:     'RolloffStats'
};

const BAYESIAN_C = 8;

// ── doGet ────────────────────────────────────────────────────
function doGet(e) {
  const view = e && e.parameter && e.parameter.view;
  if (view === 'leaderboard') {
    return HtmlService.createHtmlOutputFromFile('Leaderboard')
      .setTitle('Initiative Leaderboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (view === 'history') {
    return HtmlService.createHtmlOutputFromFile('History')
      .setTitle('Past Leaderboards')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Roll for Initiative')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Sheet bootstrap ──────────────────────────────────────────
function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

function bootstrapSheets() {
  getOrCreateSheet_(SHEET_NAMES.MEMBERS,           ['name', 'defaultPresent']);
  getOrCreateSheet_(SHEET_NAMES.SESSIONS,          ['sessionId', 'date', 'status', 'finalOrder']);
  getOrCreateSheet_(SHEET_NAMES.ROLLS,             ['sessionId', 'name', 'initialRoll', 'rolloffRoll', 'present', 'effectiveRoll']);
  getOrCreateSheet_(SHEET_NAMES.SESSION_STATE,     ['key', 'value']);
  getOrCreateSheet_(SHEET_NAMES.SESSION_ROLLS,     ['sessionId', 'name', 'roll', 'rolloffRoll', 'roundIndex', 'status']);
  getOrCreateSheet_(SHEET_NAMES.PAST_LEADERBOARDS, ['quarterKey', 'quarterLabel', 'archivedAt', 'rank', 'name', 'bayesian', 'rawAvg', 'count']);
  getOrCreateSheet_(SHEET_NAMES.ROLLOFF_STATS,     ['sessionId', 'name', 'participated', 'won', 'lost']);

  const membersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEMBERS);
  if (membersSheet.getLastRow() <= 1) {
    const defaults = [
      ['Member 1', true],
      ['Member 2', true],
      ['Member 3', true],
      ['Member 4', true],
      ['Member 5', true],
      ['Occasional 1', false],
      ['Occasional 2', false]
    ];
    defaults.forEach(row => membersSheet.appendRow(row));
  }
}

// ── Member CRUD ──────────────────────────────────────────────
function getMembers_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEMBERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(r => ({ name: String(r[0]), defaultPresent: r[1] === true || r[1] === 'TRUE' || r[1] === 'true' }));
}

function addMember(name, defaultPresent) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEMBERS);
  const members = getMembers_();
  if (members.find(m => m.name.toLowerCase() === name.toLowerCase())) {
    return { success: false, error: 'Name already exists' };
  }
  sheet.appendRow([name, defaultPresent !== false]);
  return { success: true };
}

function setMemberDefaultPresence(name, defaultPresent) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEMBERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(defaultPresent);
      return { success: true };
    }
  }
  return { success: false, error: 'Member not found' };
}

function removeMember(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MEMBERS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Member not found' };
}

// ── Session State ─────────────────────────────────────────────
function getStateValue_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSION_STATE);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

function setStateValue_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSION_STATE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function clearSessionState_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSION_STATE);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
}

function getCurrentSession() {
  const raw = getStateValue_('currentSession');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function setCurrentSession_(session) {
  setStateValue_('currentSession', JSON.stringify(session));
}

// ── SessionRolls helpers ──────────────────────────────────────
function getSessionRollsSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSION_ROLLS);
}

function initSessionRollRows_(sessionId, presentNames) {
  const sheet = getSessionRollsSheet_();
  presentNames.forEach(name => {
    sheet.appendRow([sessionId, name, '', '', 0, 'pending']);
  });
}

function initRolloffRows_(sessionId, groups, roundIndex) {
  const sheet = getSessionRollsSheet_();
  groups.flat().forEach(name => {
    sheet.appendRow([sessionId, name, '', '', roundIndex, 'pending']);
  });
}

function findRollRow_(sessionId, name, roundIndex) {
  const sheet = getSessionRollsSheet_();
  if (sheet.getLastRow() <= 1) return null;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === sessionId &&
        String(data[i][1]).toLowerCase() === name.toLowerCase() &&
        parseInt(data[i][4]) === roundIndex &&
        String(data[i][5]) === 'pending') {
      return i + 2;
    }
  }
  return null;
}

function getSubmittedRolls_(sessionId, roundIndex) {
  const sheet = getSessionRollsSheet_();
  if (sheet.getLastRow() <= 1) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const result = {};
  data.forEach(r => {
    if (String(r[0]) === sessionId &&
        parseInt(r[4]) === roundIndex &&
        String(r[5]) === 'submitted') {
      result[String(r[1])] = parseInt(r[2]);
    }
  });
  return result;
}

function clearSessionRollRows_(sessionId) {
  const sheet = getSessionRollsSheet_();
  if (sheet.getLastRow() <= 1) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]) === sessionId) {
      sheet.deleteRow(i + 2);
    }
  }
}

// ── Start a new roll session ──────────────────────────────────
function startSession(presentNames) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = getCurrentSession();
    if (existing && existing.status !== 'complete') {
      return { success: false, error: 'A session is already in progress' };
    }
    const sessionId = Utilities.getUuid();
    const session = {
      sessionId,
      date: new Date().toISOString(),
      status: 'rolling',
      presentNames,
      rolloffRounds: [],
      finalOrder: []
    };
    setCurrentSession_(session);
    initSessionRollRows_(sessionId, presentNames);
    setStateValue_('startingClaim', ''); // session now exists; clear the starting claim
    return { success: true, session: buildClientSession_(session) };
  } finally {
    lock.releaseLock();
  }
}

// ── Submit a roll — NO LOCK ───────────────────────────────────
function submitRoll(name, roll) {
  roll = parseInt(roll);
  if (isNaN(roll) || roll < 1 || roll > 20) return { success: false, error: 'Invalid roll' };

  const sheet = getSessionRollsSheet_();
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, error: 'No active session' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === name.toLowerCase() &&
        parseInt(data[i][4]) === 0 &&
        String(data[i][5]) === 'pending') {
      const rowNum = i + 2;
      sheet.getRange(rowNum, 3).setValue(roll);
      sheet.getRange(rowNum, 6).setValue('submitted');
      return { success: true };
    }
  }
  return { success: false, error: 'Roll row not found' };
}

// ── Submit a rolloff roll — NO LOCK ──────────────────────────
function submitRolloffRoll(name, roll) {
  roll = parseInt(roll);
  if (isNaN(roll) || roll < 1 || roll > 20) return { success: false, error: 'Invalid roll' };

  const sheet = getSessionRollsSheet_();
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, error: 'No active session' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  let targetRow = null;
  let targetRoundIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).toLowerCase() === name.toLowerCase() &&
        parseInt(data[i][4]) > 0 &&
        String(data[i][5]) === 'pending') {
      const ri = parseInt(data[i][4]);
      if (ri > targetRoundIndex) {
        targetRoundIndex = ri;
        targetRow = i + 2;
      }
    }
  }
  if (!targetRow) return { success: false, error: 'Rolloff row not found' };

  sheet.getRange(targetRow, 3).setValue(roll);
  sheet.getRange(targetRow, 6).setValue('submitted');
  return { success: true };
}

// ── Advance after all rolls submitted ────────────────────────
function advanceFromRolling() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session || session.status !== 'rolling') return { success: false, error: 'Not in rolling phase' };

    const rolls = getSubmittedRolls_(session.sessionId, 0);
    if (!session.presentNames.every(n => rolls[n] !== undefined)) {
      return { success: false, error: 'Not everyone has rolled' };
    }

    const tiedGroups = findTiedGroups_(session.presentNames, rolls);
    if (tiedGroups.length > 0) {
      session.status = 'rolloff';
      session.rolloffRounds.push({ groups: tiedGroups });
      setCurrentSession_(session);
      initRolloffRows_(session.sessionId, tiedGroups, 1);
      return { success: true, session: buildClientSession_(session) };
    } else {
      return finalizeSession_(session, rolls, []);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Advance after rolloff round ──────────────────────────────
function advanceFromRolloff() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session || session.status !== 'rolloff') return { success: false, error: 'Not in rolloff phase' };

    const roundIndex = session.rolloffRounds.length;
    const currentRound = session.rolloffRounds[roundIndex - 1];
    const allRolloffNames = currentRound.groups.flat();
    const rolloffRolls = getSubmittedRolls_(session.sessionId, roundIndex);

    if (!allRolloffNames.every(n => rolloffRolls[n] !== undefined)) {
      return { success: false, error: 'Not everyone in rolloff has rolled' };
    }

    const stillTied = [];
    currentRound.groups.forEach(group => {
      findTiedGroups_(group, rolloffRolls).forEach(t => stillTied.push(t));
    });

    if (stillTied.length > 0) {
      const nextRoundIndex = roundIndex + 1;
      session.rolloffRounds.push({ groups: stillTied });
      setCurrentSession_(session);
      initRolloffRows_(session.sessionId, stillTied, nextRoundIndex);
      return { success: true, session: buildClientSession_(session) };
    } else {
      const initialRolls = getSubmittedRolls_(session.sessionId, 0);
      const allRolloffRounds = [];
      for (let i = 1; i <= roundIndex; i++) {
        allRolloffRounds.push(getSubmittedRolls_(session.sessionId, i));
      }
      return finalizeSession_(session, initialRolls, allRolloffRounds);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Finalize ──────────────────────────────────────────────────
function finalizeSession_(session, initialRolls, allRolloffRounds) {
  const order = computeFinalOrder_(session.presentNames, initialRolls, allRolloffRounds);
  session.status = 'complete';
  session.finalOrder = order;
  setCurrentSession_(session);
  persistSession_(session, initialRolls, allRolloffRounds);
  clearSessionRollRows_(session.sessionId);
  return { success: true, session: buildClientSession_(session, initialRolls, allRolloffRounds) };
}

function persistSession_(session, initialRolls, allRolloffRounds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const rollsSheet = ss.getSheetByName(SHEET_NAMES.ROLLS);

  sessionsSheet.appendRow([session.sessionId, session.date, 'complete', session.finalOrder.join(', ')]);

  const members = getMembers_();
  members.forEach(m => {
    const present = session.presentNames.includes(m.name);
    if (!present) {
      rollsSheet.appendRow([session.sessionId, m.name, '', '', false, '']);
      return;
    }
    const initialRoll = initialRolls[m.name] || '';
    const rolloffRoll = getLastRolloffRollFromRounds_(m.name, allRolloffRounds);
    const effective = initialRoll; // rolloffs are tiebreakers only, not counted in averages
    rollsSheet.appendRow([session.sessionId, m.name, initialRoll, rolloffRoll || '', true, effective]);
  });

  persistRolloffStats_(session);
}

// Records, per person who was in any rolloff this session, how many rolloff
// contests they participated in / won / lost. A "contest" is one group within
// one rolloff round. You win a contest if you finished ahead of everyone else
// in that group in the final speaking order, otherwise it counts as a loss.
// Win/loss tracking is only meaningful going forward (older sessions have no rows).
function persistRolloffStats_(session) {
  if (!session.rolloffRounds || session.rolloffRounds.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName(SHEET_NAMES.ROLLOFF_STATS);
  if (!statsSheet) return;

  const finalPos = {};
  session.finalOrder.forEach((name, i) => { finalPos[name] = i; });

  const tally = {}; // name -> { participated, won, lost }
  function ensure(name) {
    if (!tally[name]) tally[name] = { participated: 0, won: 0, lost: 0 };
    return tally[name];
  }

  session.rolloffRounds.forEach(round => {
    round.groups.forEach(group => {
      group.forEach(name => {
        const rec = ensure(name);
        rec.participated++;
        const myPos = finalPos[name];
        const opponents = group.filter(n => n !== name);
        const wonContest = opponents.every(opp => (finalPos[opp] === undefined ? true : finalPos[opp] > myPos));
        if (wonContest) rec.won++; else rec.lost++;
      });
    });
  });

  Object.keys(tally).forEach(name => {
    const rec = tally[name];
    statsSheet.appendRow([session.sessionId, name, rec.participated, rec.won, rec.lost]);
  });
}

function getLastRolloffRollFromRounds_(name, allRolloffRounds) {
  for (let i = allRolloffRounds.length - 1; i >= 0; i--) {
    if (allRolloffRounds[i][name] !== undefined) return allRolloffRounds[i][name];
  }
  return null;
}

// ── Order computation ─────────────────────────────────────────
function computeFinalOrder_(presentNames, initialRolls, allRolloffRounds) {
  function getRolloffScore(name) {
    for (let i = allRolloffRounds.length - 1; i >= 0; i--) {
      if (allRolloffRounds[i][name] !== undefined) return allRolloffRounds[i][name];
    }
    return -1;
  }
  return [...presentNames].sort((a, b) => {
    const ar = initialRolls[a] || 0;
    const br = initialRolls[b] || 0;
    if (ar !== br) return br - ar;
    return getRolloffScore(b) - getRolloffScore(a);
  });
}

// ── Build client session object ───────────────────────────────
function buildClientSession_(session, initialRolls, allRolloffRounds) {
  const sessionId = session.sessionId;
  const rounds = session.rolloffRounds || [];

  if (!initialRolls) {
    initialRolls = getSubmittedRolls_(sessionId, 0);
  }

  if (!allRolloffRounds) {
    allRolloffRounds = [];
    for (let i = 1; i <= rounds.length; i++) {
      allRolloffRounds.push(getSubmittedRolls_(sessionId, i));
    }
  }

  const clientRolloffRounds = rounds.map((round, i) => ({
    groups: round.groups,
    rolls: allRolloffRounds[i] || {}
  }));

  return {
    sessionId: session.sessionId,
    date: session.date,
    status: session.status,
    presentNames: session.presentNames,
    rolls: initialRolls,
    rolloffRounds: clientRolloffRounds,
    finalOrder: session.finalOrder || []
  };
}

// ── Helpers ───────────────────────────────────────────────────
function findTiedGroups_(names, rollMap) {
  const groups = {};
  names.forEach(n => {
    const r = rollMap[n];
    if (r === undefined) return;
    if (!groups[r]) groups[r] = [];
    groups[r].push(n);
  });
  return Object.values(groups).filter(g => g.length > 1);
}

function sanitizeSession_(session) {
  return JSON.parse(JSON.stringify(session));
}

// ── Leaderboard ───────────────────────────────────────────────
function getLeaderboard() {
  const members = getMembers_();
  const rollsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROLLS);
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSIONS);

  if (!rollsSheet || rollsSheet.getLastRow() <= 1) {
    return { members: members.map(m => ({ name: m.name, avg: null, bayesian: null, count: 0 })), quarter: currentQuarterLabel_() };
  }

  const rollData = rollsSheet.getRange(2, 1, rollsSheet.getLastRow() - 1, 6).getValues();
  const sessData = sessionsSheet && sessionsSheet.getLastRow() > 1
    ? sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 4).getValues()
    : [];

  const sessDateMap = {};
  sessData.forEach(r => { sessDateMap[r[0]] = r[1]; });

  const qinfo = currentQuarter_();
  const totals = {}, counts = {};
  members.forEach(m => { totals[m.name] = 0; counts[m.name] = 0; });

  rollData.forEach(r => {
    const sessionId = r[0];
    const name = r[1];
    const effectiveRoll = parseFloat(r[5]);
    const present = r[4] === true || r[4] === 'TRUE' || r[4] === 'true';
    if (!present || isNaN(effectiveRoll)) return;
    if (!totals.hasOwnProperty(name)) return;
    const sessDate = sessDateMap[sessionId];
    if (!sessDate) return;
    const dq = quarterOf_(new Date(sessDate));
    if (dq.q !== qinfo.q || dq.y !== qinfo.y) return;
    totals[name] += effectiveRoll;
    counts[name]++;
  });

  const allRolls = Object.keys(totals).filter(n => counts[n] > 0);
  const globalMean = allRolls.length > 0
    ? allRolls.reduce((s, n) => s + totals[n], 0) / allRolls.reduce((s, n) => s + counts[n], 0)
    : 10.5;

  // Dynamic C: use the highest session count on the team this quarter
  // so that low-attendance members are always weighted against whoever showed up the most
  const maxCount = allRolls.length > 0 ? Math.max(...allRolls.map(n => counts[n])) : BAYESIAN_C;

  // Attendance penalty: anyone who has attended fewer than 33% of the top
  // roller's session count gets a flat -3 on their Bayesian score. Keeps
  // occasional attendees off the podium without removing them from the board.
  // Always active — as each sync happens, the people showing up most get the
  // most accurate standings relative to the leader.
  const penaltyThreshold = Math.round(maxCount * 0.33);

  const results = members.map(m => {
    const n = m.name;
    const c = counts[n];
    const rawAvg = c > 0 ? totals[n] / c : null;
    let bayesian = c > 0 ? ((maxCount * globalMean + totals[n]) / (maxCount + c)) : null;
    const penalized = c > 0 && c < penaltyThreshold;
    if (penalized && bayesian !== null) bayesian -= 3;
    return { name: n, avg: rawAvg, bayesian, count: c, defaultPresent: m.defaultPresent, penalized };
  }).sort((a, b) => {
    if (a.bayesian === null && b.bayesian === null) return 0;
    if (a.bayesian === null) return 1;
    if (b.bayesian === null) return -1;
    return b.bayesian - a.bayesian;
  });

  return { members: results, quarter: currentQuarterLabel_(), globalMean };
}

function getLastOrder() {
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SESSIONS);
  if (!sessionsSheet || sessionsSheet.getLastRow() <= 1) return [];
  const lastRow = sessionsSheet.getLastRow();
  const row = sessionsSheet.getRange(lastRow, 1, 1, 4).getValues()[0];
  const orderStr = row[3];
  if (!orderStr) return [];
  return String(orderStr).split(', ').map(s => s.trim()).filter(Boolean);
}

// ── Quarter helpers ───────────────────────────────────────────
function currentQuarter_() {
  return quarterOf_(new Date());
}

function quarterOf_(d) {
  const m = d.getMonth();
  return { q: Math.floor(m / 3) + 1, y: d.getFullYear() };
}

function currentQuarterLabel_() {
  const { q, y } = currentQuarter_();
  return `Q${q} ${y}`;
}

// ── Poll endpoints ────────────────────────────────────────────
// ── "Starting session" claim ──────────────────────────────────
// Bridges the gap between a leader entering the attendance screen and the
// actual session being created. While the claim is held, everyone else's
// "Start Session" button on the home screen greys out so only one person can
// kick off a session at a time. The claim carries the claimer's clientId so
// their own button doesn't grey out, and a timestamp so a stale claim (leader
// wandered off mid-attendance) auto-expires after STARTING_CLAIM_TTL_MS.
const STARTING_CLAIM_TTL_MS = 120000; // 2 minutes

function getStartingClaim_() {
  const raw = getStateValue_('startingClaim');
  if (!raw) return null;
  let claim;
  try { claim = JSON.parse(raw); } catch(e) { return null; }
  if (!claim || !claim.at) return null;
  if (Date.now() - claim.at > STARTING_CLAIM_TTL_MS) {
    // Expired — clear it
    setStateValue_('startingClaim', '');
    return null;
  }
  return claim;
}

function claimStartingSession(clientId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    // If a real session already exists, no claim needed
    const session = getCurrentSession();
    if (session && session.status !== 'complete') {
      return { success: false, error: 'A session is already in progress' };
    }
    const existing = getStartingClaim_();
    if (existing && existing.clientId !== clientId) {
      return { success: false, error: 'Someone else is already starting a session' };
    }
    setStateValue_('startingClaim', JSON.stringify({ clientId: clientId, at: Date.now() }));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function releaseStartingSession() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    setStateValue_('startingClaim', '');
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// Returns the clientId currently holding the start claim, or '' if none/expired.
function getStartingClaimId_() {
  const claim = getStartingClaim_();
  return claim ? String(claim.clientId) : '';
}

function poll() {
  const session = getCurrentSession();
  return {
    session: session ? buildClientSession_(session) : null,
    leaderboard: getLeaderboard(),
    lastOrder: getLastOrder(),
    members: getMembers_(),
    startingClaimId: getStartingClaimId_()
  };
}

function pollSession() {
  const session = getCurrentSession();
  return {
    session: session ? buildClientSession_(session) : null,
    members: getMembers_(),
    startingClaimId: getStartingClaimId_()
  };
}

// ── Remove a person from the active session ───────────────────
function removeFromSession(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session || session.status === 'complete') return { success: false, error: 'No active session' };
    const idx = session.presentNames.indexOf(name);
    if (idx === -1) return { success: false, error: 'Not in session' };
    session.presentNames.splice(idx, 1);

    const sheet = getSessionRollsSheet_();
    if (sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      const roundIndex = session.status === 'rolloff' ? session.rolloffRounds.length : 0;
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === session.sessionId &&
            String(data[i][1]).toLowerCase() === name.toLowerCase() &&
            parseInt(data[i][4]) === roundIndex &&
            String(data[i][5]) === 'pending') {
          sheet.deleteRow(i + 2);
          break;
        }
      }
    }

    if (session.rolloffRounds && session.rolloffRounds.length > 0) {
      const currentRound = session.rolloffRounds[session.rolloffRounds.length - 1];
      currentRound.groups = currentRound.groups
        .map(g => g.filter(n => n !== name))
        .filter(g => g.length > 1);
    }

    setCurrentSession_(session);
    return { success: true, session: buildClientSession_(session) };
  } finally {
    lock.releaseLock();
  }
}

// ── Cancel session ────────────────────────────────────────────
function cancelSession() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const session = getCurrentSession();
    if (session) clearSessionRollRows_(session.sessionId);
    clearSessionState_();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Quarter-end archive ───────────────────────────────────────
function archiveQuarterLeaderboard() {
  const now = new Date();
  const month = now.getMonth();
  const quarterEndMonths = [2, 5, 8, 11];
  if (!quarterEndMonths.includes(month)) {
    Logger.log('archiveQuarterLeaderboard: not a quarter-end month, skipping.');
    return;
  }

  const qinfo = currentQuarter_();
  const quarterKey = `Q${qinfo.q}_${qinfo.y}`;
  const quarterLabel = `Q${qinfo.q} ${qinfo.y} Final Leaderboard`;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  if (sheet.getLastRow() > 1) {
    const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    if (existing.includes(quarterKey)) {
      Logger.log(`Already archived ${quarterKey}, skipping.`);
      return;
    }
  }

  const lb = getLeaderboard();
  if (!lb || !lb.members) return;

  const archivedAt = now.toISOString();
  lb.members.filter(m => m.count > 0).forEach((m, i) => {
    sheet.appendRow([
      quarterKey, quarterLabel, archivedAt, i + 1, m.name,
      m.bayesian !== null ? parseFloat(m.bayesian.toFixed(2)) : '',
      m.avg !== null ? parseFloat(m.avg.toFixed(2)) : '',
      m.count
    ]);
  });

  sendArchiveEmail_(quarterLabel, quarterKey);
  Logger.log(`Archived ${quarterLabel}`);
}

function sendArchiveEmail_(quarterLabel, quarterKey) {
  try {
    const deploymentUrl = getDeploymentUrl_();
    const historyUrl = deploymentUrl ? `${deploymentUrl}?view=history` : '(open the app and click Past Quarters)';
    const subject = `Initiative Roller — ${quarterLabel} is ready`;
    const body = `Hi,\n\nThe ${quarterLabel} has been automatically saved.\n\nView it here:\n${historyUrl}\n\n— Initiative Roller`;
    GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
  } catch(e) {
    Logger.log('Failed to send archive email: ' + e.message);
  }
}

function getDeploymentUrl_() {
  return PropertiesService.getScriptProperties().getProperty('DEPLOYMENT_URL') || '';
}

// ── Get past leaderboards ─────────────────────────────────────
function getPastLeaderboards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const quarters = {}, quarterOrder = [];

  data.forEach(r => {
    const quarterKey = String(r[0]);
    if (!quarters[quarterKey]) {
      quarters[quarterKey] = { quarterKey, quarterLabel: String(r[1]), archivedAt: String(r[2]), members: [] };
      quarterOrder.push(quarterKey);
    }
    quarters[quarterKey].members.push({
      rank: parseInt(r[3]), name: String(r[4]),
      bayesian: r[5] !== '' ? parseFloat(r[5]) : null,
      rawAvg: r[6] !== '' ? parseFloat(r[6]) : null,
      count: parseInt(r[7])
    });
  });

  return quarterOrder
    .sort((a, b) => {
      const [qa, ya] = parseQuarterKey_(a);
      const [qb, yb] = parseQuarterKey_(b);
      // Newest quarter first
      return yb !== ya ? yb - ya : qb - qa;
    })
    .map(key => {
      const q = quarters[key];
      q.achievements = getQuarterStats_(key);
      return q;
    });
}

function parseQuarterKey_(key) {
  const m = key.match(/Q(\d)_(\d{4})/);
  return m ? [parseInt(m[1]), parseInt(m[2])] : [0, 0];
}

// ── Quarter achievement stats ─────────────────────────────────
// Computes the fun end-of-quarter superlatives for a given quarter.
// Reads live from the Rolls sheet (nat 20s, nat 1s, sessions attended) and
// the RolloffStats sheet (rolloff participation / wins / losses).
// Each stat returns { names: [...], value: n } — names is an array so ties
// surface everyone tied for the lead. Returns null if there's no data.
function getQuarterStats_(quarterKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rollsSheet = ss.getSheetByName(SHEET_NAMES.ROLLS);
  const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const statsSheet = ss.getSheetByName(SHEET_NAMES.ROLLOFF_STATS);

  if (!rollsSheet || rollsSheet.getLastRow() <= 1) return null;

  const [targetQ, targetY] = parseQuarterKey_(quarterKey);

  // Which session IDs belong to this quarter?
  const quarterSessionIds = {};
  if (sessionsSheet && sessionsSheet.getLastRow() > 1) {
    const sessData = sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 4).getValues();
    sessData.forEach(r => {
      const d = new Date(r[1]);
      const dq = quarterOf_(d);
      if (dq.q === targetQ && dq.y === targetY) quarterSessionIds[String(r[0])] = true;
    });
  }
  if (Object.keys(quarterSessionIds).length === 0) return null;

  const nat20s = {}, nat1s = {}, sessions = {};
  const bump = (map, name) => { map[name] = (map[name] || 0) + 1; };

  const rollData = rollsSheet.getRange(2, 1, rollsSheet.getLastRow() - 1, 6).getValues();
  rollData.forEach(r => {
    if (!quarterSessionIds[String(r[0])]) return;
    const name = String(r[1]);
    const present = r[4] === true || r[4] === 'TRUE' || r[4] === 'true';
    if (!present) return;
    const initialRoll = parseInt(r[2]);
    bump(sessions, name);
    if (initialRoll === 20) bump(nat20s, name);
    if (initialRoll === 1) bump(nat1s, name);
  });

  const participated = {}, won = {}, lost = {};
  if (statsSheet && statsSheet.getLastRow() > 1) {
    const statData = statsSheet.getRange(2, 1, statsSheet.getLastRow() - 1, 5).getValues();
    statData.forEach(r => {
      if (!quarterSessionIds[String(r[0])]) return;
      const name = String(r[1]);
      participated[name] = (participated[name] || 0) + (parseInt(r[2]) || 0);
      won[name]          = (won[name]          || 0) + (parseInt(r[3]) || 0);
      lost[name]         = (lost[name]         || 0) + (parseInt(r[4]) || 0);
    });
  }

  // Returns the top scorer(s). Only counts positive values so a stat nobody
  // achieved shows no winner rather than everyone tied at zero.
  function top(map) {
    const vals = Object.keys(map).filter(n => map[n] > 0);
    if (vals.length === 0) return { names: [], value: 0 };
    const max = Math.max.apply(null, vals.map(n => map[n]));
    return { names: vals.filter(n => map[n] === max).sort(), value: max };
  }

  return {
    nat20s:              top(nat20s),
    nat1s:               top(nat1s),
    rolloffParticipated: top(participated),
    rolloffWins:         top(won),
    rolloffLosses:       top(lost),
    sessions:            top(sessions)
  };
}

// ── Trigger setup ─────────────────────────────────────────────
function createArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'archiveQuarterLeaderboard') ScriptApp.deleteTrigger(t);
  });
  [30, 31].forEach(day => {
    ScriptApp.newTrigger('archiveQuarterLeaderboard').timeBased().onMonthDay(day).atHour(23).create();
  });
  Logger.log('Archive triggers created — fires on the 30th and 31st of every month at 11pm.');
}

function setDeploymentUrl(url) {
  PropertiesService.getScriptProperties().setProperty('DEPLOYMENT_URL', url);
  Logger.log('Deployment URL saved: ' + url);
}

// ── Force archive (testing) ───────────────────────────────────
function forceArchiveCurrentQuarter() {
  const qinfo = currentQuarter_();
  const quarterKey = `Q${qinfo.q}_${qinfo.y}`;
  const quarterLabel = `Q${qinfo.q} ${qinfo.y} Final Leaderboard`;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  const lb = getLeaderboard();
  if (!lb || !lb.members) { Logger.log('No leaderboard data.'); return; }
  const archivedAt = new Date().toISOString();
  lb.members.filter(m => m.count > 0).forEach((m, i) => {
    sheet.appendRow([
      quarterKey, quarterLabel, archivedAt, i + 1, m.name,
      m.bayesian !== null ? parseFloat(m.bayesian.toFixed(2)) : '',
      m.avg !== null ? parseFloat(m.avg.toFixed(2)) : '',
      m.count
    ]);
  });
  sendArchiveEmail_(quarterLabel, quarterKey);
  Logger.log(`Force-archived ${quarterLabel}`);
}

// ── Recalculate a past quarter's leaderboard with current scoring ──
// Computes the leaderboard for a specific quarter (not just the current one)
// using the exact same Bayesian + dynamic-C + 33% penalty rules as the live
// board. Used to retroactively apply scoring changes to archived quarters.
function computeLeaderboardForQuarter_(targetQ, targetY) {
  const members = getMembers_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rollsSheet = ss.getSheetByName(SHEET_NAMES.ROLLS);
  const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  if (!rollsSheet || rollsSheet.getLastRow() <= 1) return [];

  const rollData = rollsSheet.getRange(2, 1, rollsSheet.getLastRow() - 1, 6).getValues();
  const sessData = sessionsSheet && sessionsSheet.getLastRow() > 1
    ? sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 4).getValues()
    : [];

  const sessDateMap = {};
  sessData.forEach(r => { sessDateMap[r[0]] = r[1]; });

  const totals = {}, counts = {};
  members.forEach(m => { totals[m.name] = 0; counts[m.name] = 0; });

  rollData.forEach(r => {
    const sessionId = r[0];
    const name = r[1];
    const effectiveRoll = parseFloat(r[5]);
    const present = r[4] === true || r[4] === 'TRUE' || r[4] === 'true';
    if (!present || isNaN(effectiveRoll)) return;
    if (!totals.hasOwnProperty(name)) return;
    const sessDate = sessDateMap[sessionId];
    if (!sessDate) return;
    const dq = quarterOf_(new Date(sessDate));
    if (dq.q !== targetQ || dq.y !== targetY) return;
    totals[name] += effectiveRoll;
    counts[name]++;
  });

  const allRolls = Object.keys(totals).filter(n => counts[n] > 0);
  const globalMean = allRolls.length > 0
    ? allRolls.reduce((s, n) => s + totals[n], 0) / allRolls.reduce((s, n) => s + counts[n], 0)
    : 10.5;

  const maxCount = allRolls.length > 0 ? Math.max(...allRolls.map(n => counts[n])) : BAYESIAN_C;
  const penaltyThreshold = Math.round(maxCount * 0.33);

  return members.map(m => {
    const n = m.name;
    const c = counts[n];
    const rawAvg = c > 0 ? totals[n] / c : null;
    let bayesian = c > 0 ? ((maxCount * globalMean + totals[n]) / (maxCount + c)) : null;
    const penalized = c > 0 && c < penaltyThreshold;
    if (penalized && bayesian !== null) bayesian -= 3;
    return { name: n, avg: rawAvg, bayesian, count: c, penalized };
  }).filter(m => m.count > 0).sort((a, b) => {
    if (a.bayesian === null) return 1;
    if (b.bayesian === null) return -1;
    return b.bayesian - a.bayesian;
  });
}

// DRY RUN — logs what recalculateArchivedLeaderboards would change, writes nothing.
// Run this first, read the execution log, confirm the numbers look right.
function previewRecalculatedLeaderboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log('No archived leaderboards found.'); return; }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const quarterKeys = [];
  data.forEach(r => { const k = String(r[0]); if (quarterKeys.indexOf(k) === -1) quarterKeys.push(k); });

  quarterKeys.forEach(key => {
    const [q, y] = parseQuarterKey_(key);
    const recalced = computeLeaderboardForQuarter_(q, y);
    Logger.log('──────────── ' + key + ' ────────────');
    recalced.forEach((m, i) => {
      const oldRow = data.find(r => String(r[0]) === key && String(r[4]) === m.name);
      const oldBayes = oldRow ? oldRow[5] : '(none)';
      const pen = m.penalized ? '  [PENALTY -3]' : '';
      Logger.log(
        (i + 1) + '. ' + m.name +
        '  old: ' + oldBayes +
        '  →  new: ' + (m.bayesian !== null ? m.bayesian.toFixed(2) : '—') +
        '  (' + m.count + ' rolls)' + pen
      );
    });
  });
  Logger.log('──────────── DRY RUN COMPLETE — nothing was written ────────────');
}

// WRITES the recalculated scores into PastLeaderboards, replacing the archived
// rows for every quarter with freshly computed bayesian/rawAvg/rank values.
// Run previewRecalculatedLeaderboards first and confirm the log looks right.
// Make sure you have a backup copy of the PastLeaderboards tab before running.
function recalculateArchivedLeaderboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log('No archived leaderboards found.'); return; }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  // Preserve the quarterLabel and archivedAt per quarter so we don't lose them
  const quarterMeta = {};
  const quarterOrder = [];
  data.forEach(r => {
    const key = String(r[0]);
    if (!quarterMeta[key]) {
      quarterMeta[key] = { label: String(r[1]), archivedAt: String(r[2]) };
      quarterOrder.push(key);
    }
  });

  // Build the full new set of rows
  const newRows = [];
  quarterOrder.forEach(key => {
    const [q, y] = parseQuarterKey_(key);
    const recalced = computeLeaderboardForQuarter_(q, y);
    const meta = quarterMeta[key];
    recalced.forEach((m, i) => {
      newRows.push([
        key, meta.label, meta.archivedAt, i + 1, m.name,
        m.bayesian !== null ? parseFloat(m.bayesian.toFixed(2)) : '',
        m.avg !== null ? parseFloat(m.avg.toFixed(2)) : '',
        m.count
      ]);
    });
  });

  // Clear existing data rows and write the new ones
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, 8).setValues(newRows);
  }
  Logger.log('Recalculated ' + quarterOrder.length + ' quarter(s), wrote ' + newRows.length + ' rows.');
}

// ── Wipe test data ────────────────────────────────────────────
function wipeSheetsForTesting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_NAMES.SESSIONS, SHEET_NAMES.ROLLS, SHEET_NAMES.SESSION_STATE, SHEET_NAMES.SESSION_ROLLS, SHEET_NAMES.PAST_LEADERBOARDS, SHEET_NAMES.ROLLOFF_STATS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  });
  Logger.log('Test data wiped. Members sheet left intact.');
}
