/**
 * SOR Attendance System — API (Google Apps Script, bound to a Sheet).
 *
 * The web pages live on Firebase Hosting and call doPost() with a Firebase
 * ID token from Google sign-in. Check-ins also need the rotating code shown
 * on the in-room display, so nobody can check in from home.
 *
 * Sheets:
 *   Roster      ID | Name | Group | Active | Email (you maintain this)
 *   Attendance  ID | Name | Group | <yyyy-MM-dd>… (one column per day)
 *   Log         every check-in / admin change, append-only
 *   Settings    key/value configuration
 */

var SHEET = {
  ROSTER: 'Roster',
  ATTENDANCE: 'Attendance',
  LOG: 'Log',
  SETTINGS: 'Settings'
};
var ROSTER_HEADERS = ['ID', 'Name', 'Group', 'Active', 'Email'];
var ATTENDANCE_HEADERS = ['ID', 'Name', 'Group'];
var LOG_HEADERS = ['Timestamp', 'Date', 'ID', 'Name', 'Status', 'Source', 'Note'];
var FIRST_DATE_COL = ATTENDANCE_HEADERS.length + 1; // 1-based
var DEFAULT_SETTINGS = [
  ['Organization name', 'SOR'],
  ['Late after (HH:mm, blank = never late)', '09:15'],
  ['Self check-in enabled (TRUE/FALSE)', true],
  ['Auto-mark absent time (hour 0-23)', 23],
  ['Admin emails (comma-separated)', '']
];
var STATUS_COLORS = {
  Present: '#d9ead3',
  Late: '#fff2cc',
  Absent: '#f4cccc',
  Excused: '#cfe2f3'
};
var CODE_STEP_SECONDS = 30; // a code is accepted for its window plus the previous one
var CODE_DIGITS = 6;
var MAX_CODE_FAILURES = 5;  // per account per 10 minutes
var TOKEN_CACHE_SECONDS = 300;

// ---------------------------------------------------------------- Menu / setup

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance')
    .addItem('Set up sheets', 'setup')
    .addItem('Set Firebase API key…', 'promptFirebaseApiKey')
    .addSeparator()
    .addItem('Mark unmarked as Absent (today)', 'markAbsentToday')
    .addItem('Install daily auto-absent trigger', 'installDailyTrigger')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();

  var roster = getOrCreateSheet_(ss, SHEET.ROSTER, ROSTER_HEADERS);
  if (roster.getLastRow() < 2) {
    roster.getRange(2, 1, 2, 5).setValues([
      ['S001', 'Example Member', 'Group A', true, 'member1@gmail.com'],
      ['S002', 'Another Member', 'Group B', true, 'member2@gmail.com']
    ]);
  }
  roster.getRange('D2:D').insertCheckboxes();

  var att = getOrCreateSheet_(ss, SHEET.ATTENDANCE, ATTENDANCE_HEADERS);
  applyStatusFormatting_(att);

  getOrCreateSheet_(ss, SHEET.LOG, LOG_HEADERS);

  var settings = getOrCreateSheet_(ss, SHEET.SETTINGS, ['Setting', 'Value']);
  settings.getRange('B:B').setNumberFormat('@'); // keep "09:15" as text
  var existing = settings.getRange(1, 1, settings.getLastRow(), 1).getValues()
    .map(function (r) { return r[0]; });
  DEFAULT_SETTINGS.forEach(function (row) {
    if (existing.indexOf(row[0]) !== -1) return;
    var value = row[1];
    if (row[0].indexOf('Admin emails') === 0) value = Session.getEffectiveUser().getEmail();
    settings.appendRow([row[0], String(value)]);
  });
  getCodeSecret_();

  syncAttendanceRows_(att, getActiveRoster_());
  SpreadsheetApp.getUi().alert(
    'Setup complete.\n\n1. Fill in the Roster sheet (including each Google email).\n' +
    '2. Attendance > Set Firebase API key…\n' +
    '3. Deploy > New deployment > Web app, then put its URL in the site config.');
}

function promptFirebaseApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Firebase API key',
    'Paste the Web API key from Firebase console > Project settings > General:',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = res.getResponseText().trim();
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty('FIREBASE_API_KEY', key);
  ui.alert('Firebase API key saved.');
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'markAbsentToday') ScriptApp.deleteTrigger(t);
  });
  var hour = Number(getSettings_().autoAbsentHour);
  ScriptApp.newTrigger('markAbsentToday')
    .timeBased().everyDays(1).atHour(isNaN(hour) ? 23 : hour).create();
  SpreadsheetApp.getUi().alert('Daily trigger installed (runs around ' +
    (isNaN(hour) ? 23 : hour) + ':00).');
}

// ---------------------------------------------------------------- API

var ACTIONS = {
  me: apiMe_,
  checkIn: apiCheckIn_,
  displayCode: apiDisplayCode_,
  getDay: apiGetDay_,
  saveDay: apiSaveDay_
};

function doGet() {
  return json_({ ok: true, message: 'SOR attendance API is running.' });
}

/** Body: {"action": "...", "idToken": "<Firebase ID token>", ...params} */
function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!ACTIONS.hasOwnProperty(req.action)) throw new Error('Unknown action.');
    var user = verifyIdToken_(req.idToken);
    return json_({ ok: true, data: ACTIONS[req.action](user, req) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function apiMe_(user) {
  var settings = getSettings_();
  var member = AttendanceLogic.findMemberByEmail(getActiveRoster_(), user.email);
  return {
    orgName: settings.orgName,
    today: todayKey_(),
    email: user.email,
    member: member,
    isAdmin: isAdmin_(user, settings),
    selfCheckIn: settings.selfCheckIn,
    statuses: AttendanceLogic.STATUSES
  };
}

/** The signed-in member checks in with the code currently on the display. */
function apiCheckIn_(user, req) {
  var settings = getSettings_();
  if (!settings.selfCheckIn) throw new Error('Check-in is currently closed.');

  var member = AttendanceLogic.findMemberByEmail(getActiveRoster_(), user.email);
  if (!member) {
    throw new Error(user.email + ' is not on the roster. Ask an admin to add this email.');
  }

  var now = new Date();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var dateKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  var cache = CacheService.getScriptCache();
  var failKey = 'codefail_' + user.email;
  var failures = Number(cache.get(failKey) || 0);
  if (failures >= MAX_CODE_FAILURES) {
    throw new Error('Too many wrong codes. Wait 10 minutes and try again.');
  }
  var code = AttendanceLogic.normalizeCode(req.code);
  if (!isCurrentCode_(code, now.getTime())) {
    cache.put(failKey, String(failures + 1), 600);
    appendLog_([[now, dateKey, member.id, member.name, 'Rejected', 'Self check-in',
      'wrong or expired code "' + code + '"']]);
    throw new Error('That code is wrong or has expired. Enter the code on the screen right now.');
  }

  var minutes = AttendanceLogic.parseTime(Utilities.formatDate(now, tz, 'HH:mm'));
  var incoming = AttendanceLogic.statusForCheckIn(minutes, settings.lateCutoff);

  return withLock_(function () {
    var att = SpreadsheetApp.getActive().getSheetByName(SHEET.ATTENDANCE);
    var rowById = syncAttendanceRows_(att, [member]);
    var col = getOrCreateDateColumn_(att, dateKey);
    var cell = att.getRange(rowById[member.id], col);
    var result = AttendanceLogic.mergeCheckIn(String(cell.getValue()), incoming);
    if (result.changed) {
      cell.setValue(result.status);
      appendLog_([[now, dateKey, member.id, member.name, result.status, 'Self check-in',
        user.email]]);
    }
    return {
      name: member.name,
      status: result.status,
      alreadyMarked: !result.changed,
      time: Utilities.formatDate(now, tz, 'HH:mm')
    };
  });
}

/** Admin: the code to show on the in-room display. */
function apiDisplayCode_(user) {
  requireAdmin_(user);
  var nowMs = new Date().getTime();
  var step = AttendanceLogic.timeStep(nowMs, CODE_STEP_SECONDS);
  return {
    code: codeForStep_(step),
    expiresInMs: (step + 1) * CODE_STEP_SECONDS * 1000 - nowMs,
    stepSeconds: CODE_STEP_SECONDS,
    orgName: getSettings_().orgName
  };
}

/** Admin: roster with statuses for a date. */
function apiGetDay_(user, req) {
  requireAdmin_(user);
  var dateKey = req.date;
  assertDateKey_(dateKey);
  var att = SpreadsheetApp.getActive().getSheetByName(SHEET.ATTENDANCE);
  var roster = getActiveRoster_();
  var rowById = syncAttendanceRows_(att, roster);
  var col = findDateColumn_(att, dateKey);
  var values = col > 0 && att.getLastRow() > 1
    ? att.getRange(1, col, att.getLastRow(), 1).getValues()
    : [];
  var members = roster.map(function (m) {
    var v = values.length ? values[rowById[m.id] - 1][0] : '';
    return { id: m.id, name: m.name, group: m.group, status: String(v || '') };
  });
  return {
    date: dateKey,
    members: members,
    summary: AttendanceLogic.summarize(members.map(function (m) { return m.status; }))
  };
}

/** Admin: save statuses for a date. req.records = [{id, status}] */
function apiSaveDay_(user, req) {
  requireAdmin_(user);
  var dateKey = req.date;
  var records = req.records || [];
  assertDateKey_(dateKey);
  records.forEach(function (r) {
    if (!AttendanceLogic.isValidStatus(r.status)) throw new Error('Invalid status: ' + r.status);
  });

  return withLock_(function () {
    var att = SpreadsheetApp.getActive().getSheetByName(SHEET.ATTENDANCE);
    var roster = getActiveRoster_();
    var byId = {};
    roster.forEach(function (m) { byId[m.id] = m; });
    var rowById = syncAttendanceRows_(att, roster);
    var col = getOrCreateDateColumn_(att, dateKey);
    var range = att.getRange(1, col, att.getLastRow(), 1);
    var values = range.getValues();
    var now = new Date();
    var logRows = [];

    records.forEach(function (r) {
      var row = rowById[r.id];
      if (!row || !byId[r.id]) return;
      var before = String(values[row - 1][0] || '');
      if (before === r.status) return;
      values[row - 1][0] = r.status;
      logRows.push([now, dateKey, r.id, byId[r.id].name, r.status || '(cleared)',
        'Admin ' + user.email, before ? 'was ' + before : '']);
    });

    range.setValues(values);
    appendLog_(logRows);
    return { saved: logRows.length };
  });
}

/** Fills every blank cell for today with Absent. Used by the menu and daily trigger. */
function markAbsentToday() {
  var dateKey = todayKey_();
  var count = withLock_(function () {
    var att = SpreadsheetApp.getActive().getSheetByName(SHEET.ATTENDANCE);
    var roster = getActiveRoster_();
    var rowById = syncAttendanceRows_(att, roster);
    var col = getOrCreateDateColumn_(att, dateKey);
    var range = att.getRange(1, col, att.getLastRow(), 1);
    var values = range.getValues();
    var now = new Date();
    var logRows = [];
    roster.forEach(function (m) {
      var row = rowById[m.id];
      if (!values[row - 1][0]) {
        values[row - 1][0] = 'Absent';
        logRows.push([now, dateKey, m.id, m.name, 'Absent', 'Auto', '']);
      }
    });
    range.setValues(values);
    appendLog_(logRows);
    return logRows.length;
  });
  try {
    SpreadsheetApp.getActive().toast(count + ' member(s) marked Absent for ' + dateKey);
  } catch (e) {
    // No UI when run from a trigger.
  }
  return count;
}

// ---------------------------------------------------------------- Helpers

function getOrCreateSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function applyStatusFormatting_(sheet) {
  var range = sheet.getRange(2, FIRST_DATE_COL, sheet.getMaxRows() - 1,
    Math.max(1, sheet.getMaxColumns() - FIRST_DATE_COL + 1));
  var rules = Object.keys(STATUS_COLORS).map(function (status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUS_COLORS[status])
      .setRanges([range])
      .build();
  });
  sheet.setConditionalFormatRules(rules);
  sheet.setFrozenColumns(ATTENDANCE_HEADERS.length);
}

function getSettings_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  var map = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      map[r[0]] = r[1];
    });
  }
  var get = function (i) {
    var key = DEFAULT_SETTINGS[i][0];
    return map.hasOwnProperty(key) ? map[key] : DEFAULT_SETTINGS[i][1];
  };
  var lateRaw = get(1);
  if (lateRaw instanceof Date) lateRaw = Utilities.formatDate(lateRaw,
    SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm');
  return {
    orgName: String(get(0) || 'SOR'),
    lateCutoff: AttendanceLogic.parseTime(lateRaw),
    selfCheckIn: String(get(2)).toUpperCase() !== 'FALSE',
    autoAbsentHour: get(3),
    adminEmails: AttendanceLogic.parseEmailList(get(4))
  };
}

function getActiveRoster_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.ROSTER);
  if (!sheet) throw new Error('Roster sheet missing. Run Attendance > Set up sheets.');
  if (sheet.getLastRow() < 2) return [];
  var seen = {};
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, ROSTER_HEADERS.length).getValues()
    .filter(function (r) {
      var id = String(r[0]).trim();
      var active = r[3] === true || String(r[3]).toUpperCase() === 'TRUE';
      if (!id || !active || seen[id]) return false;
      seen[id] = true;
      return true;
    })
    .map(function (r) {
      return {
        id: String(r[0]).trim(),
        name: String(r[1]).trim(),
        group: String(r[2]).trim(),
        email: String(r[4]).trim().toLowerCase()
      };
    });
}

/** Ensures each member has a row in Attendance; returns {id: rowNumber} for all rows. */
function syncAttendanceRows_(att, members) {
  var lastRow = att.getLastRow();
  var rowById = {};
  if (lastRow > 1) {
    att.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r, i) {
      var id = String(r[0]).trim();
      if (id) rowById[id] = i + 2;
    });
  }
  var newRows = [];
  members.forEach(function (m) {
    if (!rowById[m.id]) {
      newRows.push([m.id, m.name, m.group]);
      rowById[m.id] = lastRow + newRows.length;
    }
  });
  if (newRows.length) {
    att.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }
  return rowById;
}

function readDateHeaders_(att) {
  var lastCol = att.getLastColumn();
  if (lastCol < 1) return [];
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return att.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).trim();
  });
}

/** 1-based column for dateKey, or -1. */
function findDateColumn_(att, dateKey) {
  var headers = readDateHeaders_(att);
  for (var i = FIRST_DATE_COL - 1; i < headers.length; i++) {
    if (headers[i] === dateKey) return i + 1;
  }
  return -1;
}

function getOrCreateDateColumn_(att, dateKey) {
  var existing = findDateColumn_(att, dateKey);
  if (existing > 0) return existing;
  var headers = readDateHeaders_(att);
  var idx = AttendanceLogic.insertionIndexForDate(headers, dateKey, FIRST_DATE_COL - 1);
  var col = idx + 1;
  if (col <= att.getLastColumn()) {
    att.insertColumnBefore(col);
  } else if (col > att.getMaxColumns()) {
    att.insertColumnAfter(att.getMaxColumns());
  }
  att.getRange(1, col).setNumberFormat('@').setValue(dateKey).setFontWeight('bold');
  applyStatusFormatting_(att); // cover newly added columns
  return col;
}

function appendLog_(rows) {
  if (!rows.length) return;
  var log = SpreadsheetApp.getActive().getSheetByName(SHEET.LOG);
  log.getRange(log.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

/**
 * Verifies a Firebase ID token with Google and returns {email, name}.
 * The API key ties the token to your Firebase project.
 */
function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('Please sign in.');
  var apiKey = PropertiesService.getScriptProperties().getProperty('FIREBASE_API_KEY');
  if (!apiKey) throw new Error('Server not configured: run Attendance > Set Firebase API key…');

  var cache = CacheService.getScriptCache();
  var cacheKey = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var res = UrlFetchApp.fetch(
    'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(apiKey),
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ idToken: idToken }),
      muteHttpExceptions: true
    });
  if (res.getResponseCode() !== 200) throw new Error('Your sign-in expired. Please sign in again.');
  var account = (JSON.parse(res.getContentText()).users || [])[0];
  if (!account || !account.email || !account.emailVerified) {
    throw new Error('Please sign in with a verified Google account.');
  }
  var user = { email: account.email.toLowerCase(), name: account.displayName || '' };
  cache.put(cacheKey, JSON.stringify(user), TOKEN_CACHE_SECONDS);
  return user;
}

function isAdmin_(user, settings) {
  return settings.adminEmails.indexOf(user.email) !== -1;
}

function requireAdmin_(user) {
  if (!isAdmin_(user, getSettings_())) throw new Error('Admins only.');
}

function getCodeSecret_() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('CODE_SECRET');
  if (secret) return secret;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    secret = props.getProperty('CODE_SECRET');
    if (!secret) {
      secret = Utilities.getUuid() + Utilities.getUuid();
      props.setProperty('CODE_SECRET', secret);
    }
    return secret;
  } finally {
    lock.releaseLock();
  }
}

function codeForStep_(step) {
  var sig = Utilities.computeHmacSha256Signature(String(step), getCodeSecret_());
  return AttendanceLogic.truncateToCode(sig, CODE_DIGITS);
}

/** Accepts the current window's code and the previous one (for slow typists). */
function isCurrentCode_(code, nowMs) {
  if (code.length !== CODE_DIGITS) return false;
  var step = AttendanceLogic.timeStep(nowMs, CODE_STEP_SECONDS);
  return code === codeForStep_(step) || code === codeForStep_(step - 1);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function assertDateKey_(dateKey) {
  if (!AttendanceLogic.isValidDateKey(dateKey)) throw new Error('Invalid date: ' + dateKey);
}

function todayKey_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    'yyyy-MM-dd');
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The system is busy. Please try again.');
  try {
    var result = fn();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}
