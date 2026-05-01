// ============================================================
//  INITIATIVE ROLLER — Google Apps Script Backend
//  Stores all data in the bound Google Sheet
// ============================================================

const SHEET_NAMES = {
  CONFIG:   'Config',
  MEMBERS:  'Members',
  SESSIONS: 'Sessions',
  ROLLS:    'Rolls',
  SESSION_STATE: 'SessionState',
  PAST_LEADERBOARDS: 'PastLeaderboards'
};

const BAYESIAN_C = 8; // confidence threshold — sessions needed to "earn" your full average

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
  getOrCreateSheet_(SHEET_NAMES.MEMBERS, ['name', 'defaultPresent']);
  getOrCreateSheet_(SHEET_NAMES.SESSIONS, ['sessionId', 'date', 'status', 'finalOrder']);
  getOrCreateSheet_(SHEET_NAMES.ROLLS, ['sessionId', 'name', 'initialRoll', 'rolloffRoll', 'present', 'effectiveRoll']);
  getOrCreateSheet_(SHEET_NAMES.SESSION_STATE, ['key', 'value']);
  getOrCreateSheet_(SHEET_NAMES.PAST_LEADERBOARDS, ['quarterKey', 'quarterLabel', 'archivedAt', 'rank', 'name', 'bayesian', 'rawAvg', 'count']);

  // Seed default members if Members sheet is empty (no data rows)
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

// ── Session State ────────────────────────────────────────────
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
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

// ── Current Session ──────────────────────────────────────────
function getCurrentSession() {
  const raw = getStateValue_('currentSession');
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function setCurrentSession_(session) {
  setStateValue_('currentSession', JSON.stringify(session));
}

// ── Start a new roll session ─────────────────────────────────
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
      status: 'rolling', // rolling | rolloff | complete
      presentNames,
      rolls: {},           // name -> roll value
      rolloffRounds: [],   // array of { groups: [[name,...]], rolls: {name: value} }
      finalOrder: []
    };
    setCurrentSession_(session);
    return { success: true, session: sanitizeSession_(session) };
  } finally {
    lock.releaseLock();
  }
}

// ── Submit a roll ────────────────────────────────────────────
function submitRoll(name, roll) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session) return { success: false, error: 'No active session' };
    if (session.status !== 'rolling') return { success: false, error: 'Not in rolling phase' };
    if (!session.presentNames.includes(name)) return { success: false, error: 'Not in present list' };
    roll = parseInt(roll);
    if (isNaN(roll) || roll < 1 || roll > 20) return { success: false, error: 'Invalid roll' };

    session.rolls[name] = roll;
    setCurrentSession_(session);
    return { success: true, session: sanitizeSession_(session) };
  } finally {
    lock.releaseLock();
  }
}

// ── Advance after all rolls submitted ────────────────────────
function advanceFromRolling() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session || session.status !== 'rolling') return { success: false, error: 'Not in rolling phase' };
    if (!allPresentsRolled_(session)) return { success: false, error: 'Not everyone has rolled' };

    const tiedGroups = findTiedGroups_(session.presentNames, session.rolls);
    if (tiedGroups.length > 0) {
      session.status = 'rolloff';
      session.rolloffRounds.push({ groups: tiedGroups, rolls: {} });
      setCurrentSession_(session);
      return { success: true, session: sanitizeSession_(session) };
    } else {
      return finalizeSession_(session);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Submit a rolloff roll ────────────────────────────────────
function submitRolloffRoll(name, roll) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = getCurrentSession();
    if (!session || session.status !== 'rolloff') return { success: false, error: 'Not in rolloff phase' };
    const currentRound = session.rolloffRounds[session.rolloffRounds.length - 1];
    const allRolloffNames = currentRound.groups.flat();
    if (!allRolloffNames.includes(name)) return { success: false, error: 'Not in rolloff' };

    roll = parseInt(roll);
    if (isNaN(roll) || roll < 1 || roll > 20) return { success: false, error: 'Invalid roll' };
    currentRound.rolls[name] = roll;
    setCurrentSession_(session);
    return { success: true, session: sanitizeSession_(session) };
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
    const currentRound = session.rolloffRounds[session.rolloffRounds.length - 1];
    const allRolloffNames = currentRound.groups.flat();
    if (!allRolloffNamesRolled_(currentRound)) return { success: false, error: 'Not everyone in rolloff has rolled' };

    const stillTied = findTiedGroups_(allRolloffNames, currentRound.rolls);
    if (stillTied.length > 0) {
      session.rolloffRounds.push({ groups: stillTied, rolls: {} });
      setCurrentSession_(session);
      return { success: true, session: sanitizeSession_(session) };
    } else {
      return finalizeSession_(session);
    }
  } finally {
    lock.releaseLock();
  }
}

// ── Finalize ─────────────────────────────────────────────────
function finalizeSession_(session) {
  const order = computeFinalOrder_(session);
  session.status = 'complete';
  session.finalOrder = order;
  setCurrentSession_(session);

  // Persist to Rolls + Sessions sheets
  persistSession_(session);

  return { success: true, session: sanitizeSession_(session) };
}

function persistSession_(session) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sessionsSheet = ss.getSheetByName(SHEET_NAMES.SESSIONS);
  const rollsSheet = ss.getSheetByName(SHEET_NAMES.ROLLS);

  sessionsSheet.appendRow([
    session.sessionId,
    session.date,
    'complete',
    session.finalOrder.join(', ')
  ]);

  const members = getMembers_();
  members.forEach(m => {
    const present = session.presentNames.includes(m.name);
    if (!present) {
      rollsSheet.appendRow([session.sessionId, m.name, '', '', false, '']);
      return;
    }
    const initialRoll = session.rolls[m.name] || '';
    const rolloffRoll = getLastRolloffRoll_(session, m.name);
    const effective = computeEffectiveRoll_(session, m.name);
    rollsSheet.appendRow([session.sessionId, m.name, initialRoll, rolloffRoll || '', true, effective]);
  });
}

function getLastRolloffRoll_(session, name) {
  for (let i = session.rolloffRounds.length - 1; i >= 0; i--) {
    const round = session.rolloffRounds[i];
    if (round.rolls[name] !== undefined) return round.rolls[name];
  }
  return null;
}

function computeEffectiveRoll_(session, name) {
  const initial = session.rolls[name];
  const rolloff = getLastRolloffRoll_(session, name);
  if (rolloff !== null && rolloff !== undefined) {
    return (initial + rolloff) / 2;
  }
  return initial;
}

// ── Order computation ────────────────────────────────────────
function computeFinalOrder_(session) {
  const resolved = []; // names in order, resolved
  const resolvedSet = new Set();

  // Build a per-name "tiebreaker chain" from rolloff rounds
  function getRolloffScore(name) {
    for (let i = session.rolloffRounds.length - 1; i >= 0; i--) {
      const round = session.rolloffRounds[i];
      if (round.rolls[name] !== undefined) return round.rolls[name];
    }
    return -1;
  }

  // Sort present names: primary = initial roll desc, secondary = rolloff desc
  const sorted = [...session.presentNames].sort((a, b) => {
    const ar = session.rolls[a] || 0;
    const br = session.rolls[b] || 0;
    if (ar !== br) return br - ar;
    return getRolloffScore(b) - getRolloffScore(a);
  });

  return sorted;
}

// ── Helpers ──────────────────────────────────────────────────
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

function allPresentsRolled_(session) {
  return session.presentNames.every(n => session.rolls[n] !== undefined);
}

function allRolloffNamesRolled_(round) {
  const names = round.groups.flat();
  return names.every(n => round.rolls[n] !== undefined);
}

function sanitizeSession_(session) {
  // Returns a safe copy for client — no internal GAS objects
  return JSON.parse(JSON.stringify(session));
}

// ── Leaderboard ──────────────────────────────────────────────
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

  // Build session date map
  const sessDateMap = {};
  sessData.forEach(r => { sessDateMap[r[0]] = r[1]; });

  const qinfo = currentQuarter_();

  // Aggregate per member for current quarter
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
    const d = new Date(sessDate);
    const dq = quarterOf_(d);
    if (dq.q !== qinfo.q || dq.y !== qinfo.y) return;

    totals[name] += effectiveRoll;
    counts[name]++;
  });

  // Compute global mean for Bayesian
  const allRolls = Object.keys(totals).filter(n => counts[n] > 0);
  const globalMean = allRolls.length > 0
    ? allRolls.reduce((s, n) => s + totals[n], 0) / allRolls.reduce((s, n) => s + counts[n], 0)
    : 10.5;

  const results = members.map(m => {
    const n = m.name;
    const c = counts[n];
    const rawAvg = c > 0 ? totals[n] / c : null;
    const bayesian = c > 0 ? ((BAYESIAN_C * globalMean + totals[n]) / (BAYESIAN_C + c)) : null;
    return { name: n, avg: rawAvg, bayesian, count: c, defaultPresent: m.defaultPresent };
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

// ── Quarter helpers ──────────────────────────────────────────
function currentQuarter_() {
  const d = new Date();
  return quarterOf_(d);
}

function quarterOf_(d) {
  const m = d.getMonth();
  const q = Math.floor(m / 3) + 1;
  return { q, y: d.getFullYear() };
}

function currentQuarterLabel_() {
  const { q, y } = currentQuarter_();
  return `Q${q} ${y}`;
}

// ── Poll endpoint (called by client every 3s) ─────────────────
function poll() {
  return {
    session: getCurrentSession(),
    leaderboard: getLeaderboard(),
    lastOrder: getLastOrder(),
    members: getMembers_()
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
    delete session.rolls[name];
    if (session.rolloffRounds && session.rolloffRounds.length > 0) {
      const currentRound = session.rolloffRounds[session.rolloffRounds.length - 1];
      currentRound.groups = currentRound.groups
        .map(g => g.filter(n => n !== name))
        .filter(g => g.length > 1);
      delete currentRound.rolls[name];
    }
    setCurrentSession_(session);
    return { success: true, session: sanitizeSession_(session) };
  } finally {
    lock.releaseLock();
  }
}

// ── Cancel session (manager escape hatch) ────────────────────
function cancelSession() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    clearSessionState_();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ── Quarter-end archive ───────────────────────────────────────
// Called automatically by time-driven trigger on last day of Mar/Jun/Sep/Dec.
// Also callable manually from the script editor if needed.
function archiveQuarterLeaderboard() {
  const now = new Date();
  const month = now.getMonth(); // 0-indexed
  // Only run at end of Q1 (March=2), Q2 (June=5), Q3 (Sep=8), Q4 (Dec=11)
  const quarterEndMonths = [2, 5, 8, 11];
  if (!quarterEndMonths.includes(month)) {
    Logger.log('archiveQuarterLeaderboard: not a quarter-end month, skipping.');
    return;
  }

  const qinfo = currentQuarter_();
  const quarterKey = `Q${qinfo.q}_${qinfo.y}`;
  const quarterLabel = `Q${qinfo.q} ${qinfo.y} Final Leaderboard`;

  // Check if already archived this quarter
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
  const ranked = lb.members.filter(m => m.count > 0);

  ranked.forEach((m, i) => {
    sheet.appendRow([
      quarterKey,
      quarterLabel,
      archivedAt,
      i + 1,
      m.name,
      m.bayesian !== null ? parseFloat(m.bayesian.toFixed(2)) : '',
      m.avg !== null ? parseFloat(m.avg.toFixed(2)) : '',
      m.count
    ]);
  });

  // Send email notification
  sendArchiveEmail_(quarterLabel, quarterKey);

  Logger.log(`Archived ${quarterLabel} with ${ranked.length} entries.`);
}

function sendArchiveEmail_(quarterLabel, quarterKey) {
  try {
    const deploymentUrl = getDeploymentUrl_();
    const historyUrl = deploymentUrl ? `${deploymentUrl}?view=history` : '(open the app and click Past Quarters)';

    const subject = `Initiative Roller — ${quarterLabel} is ready`;
    const body = `Hi,

The ${quarterLabel} has been automatically saved.

View it here:
${historyUrl}

You can embed the history page in Confluence using the iFrame macro with the URL above.

— Initiative Roller`;

    GmailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject, body);
    Logger.log(`Archive email sent for ${quarterLabel}`);
  } catch(e) {
    Logger.log('Failed to send archive email: ' + e.message);
  }
}

function getDeploymentUrl_() {
  // Stored once when the trigger is created
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('DEPLOYMENT_URL') || '';
}

// ── Get past leaderboards (for History page) ──────────────────
function getPastLeaderboards() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();

  // Group rows by quarterKey
  const quarters = {};
  const quarterOrder = [];
  data.forEach(r => {
    const quarterKey   = String(r[0]);
    const quarterLabel = String(r[1]);
    const archivedAt   = String(r[2]);
    const rank         = parseInt(r[3]);
    const name         = String(r[4]);
    const bayesian     = r[5] !== '' ? parseFloat(r[5]) : null;
    const rawAvg       = r[6] !== '' ? parseFloat(r[6]) : null;
    const count        = parseInt(r[7]);

    if (!quarters[quarterKey]) {
      quarters[quarterKey] = { quarterKey, quarterLabel, archivedAt, members: [] };
      quarterOrder.push(quarterKey);
    }
    quarters[quarterKey].members.push({ rank, name, bayesian, rawAvg, count });
  });

  // Return in chronological order (Q1 2025 before Q2 2025, etc.)
  return quarterOrder
    .sort((a, b) => {
      const [qa, ya] = parseQuarterKey_(a);
      const [qb, yb] = parseQuarterKey_(b);
      return ya !== yb ? ya - yb : qa - qb;
    })
    .map(key => quarters[key]);
}

function parseQuarterKey_(key) {
  // Format: Q1_2025
  const m = key.match(/Q(\d)_(\d{4})/);
  return m ? [parseInt(m[1]), parseInt(m[2])] : [0, 0];
}

// ── Trigger setup ─────────────────────────────────────────────
// Run this ONCE from the Apps Script editor after deploying.
// It creates a monthly trigger that fires on the 28th of every month
// (safe for all months). The function itself checks if it's actually
// a quarter-end month before doing anything.
function createArchiveTrigger() {
  // Delete any existing archive triggers first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'archiveQuarterLeaderboard') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Fire on the 28th of every month at 11pm
  ScriptApp.newTrigger('archiveQuarterLeaderboard')
    .timeBased()
    .onMonthDay(28)
    .atHour(23)
    .create();

  Logger.log('Archive trigger created — fires on the 28th of every month at 11pm.');
}

// ── Save deployment URL (run once after deploying) ────────────
function setDeploymentUrl(url) {
  PropertiesService.getScriptProperties().setProperty('DEPLOYMENT_URL', url);
  Logger.log('Deployment URL saved: ' + url);
}

// ── Manual archive (for testing or makeup runs) ───────────────
// Call this from the editor to force-archive the current quarter,
// regardless of what month it is. Useful for testing.
function forceArchiveCurrentQuarter() {
  const qinfo = currentQuarter_();
  const quarterKey = `Q${qinfo.q}_${qinfo.y}`;
  const quarterLabel = `Q${qinfo.q} ${qinfo.y} Final Leaderboard`;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.PAST_LEADERBOARDS);
  const lb = getLeaderboard();
  if (!lb || !lb.members) { Logger.log('No leaderboard data.'); return; }

  const archivedAt = new Date().toISOString();
  const ranked = lb.members.filter(m => m.count > 0);

  ranked.forEach((m, i) => {
    sheet.appendRow([
      quarterKey, quarterLabel, archivedAt,
      i + 1, m.name,
      m.bayesian !== null ? parseFloat(m.bayesian.toFixed(2)) : '',
      m.avg !== null ? parseFloat(m.avg.toFixed(2)) : '',
      m.count
    ]);
  });

  sendArchiveEmail_(quarterLabel, quarterKey);
  Logger.log(`Force-archived ${quarterLabel} with ${ranked.length} entries.`);
}

