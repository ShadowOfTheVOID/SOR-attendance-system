/**
 * SOR Attendance — check-in API (Google Apps Script bound to the team's
 * "SOR Signups/Attendance" spreadsheet).
 *
 * The web pages live on Firebase Hosting and call doPost() with a Firebase
 * ID token from Google sign-in. A check-in needs the rotating code shown on
 * the display for a location (e.g. "Hangar 391" or "Online"). The time of
 * the check-in decides which of that location's shifts (columns) it counts
 * for; where shifts overlap, the next shift wins.
 *
 * Existing tabs are only read, except for single student cells in the
 * attendance tab. This script adds three tabs of its own:
 *   Check-in Settings   key/value configuration
 *   Check-in Roster     Name (as in column A) | Google email
 *   Check-in Log        every check-in, rejected code and admin edit
 */

// Always admins; can't be removed from the website. The account that set up
// the script is always an admin too. Other admins are managed on /admin.html.
var PERMANENT_ADMINS = ['nhstedd@gmail.com'];

var TAB = {
  SETTINGS: 'Check-in Settings',
  ROSTER: 'Check-in Roster',
  LOG: 'Check-in Log'
};
var ROSTER_HEADERS = ['Name (exactly as in the attendance tab)', 'Google email'];
var LOG_HEADERS = ['Timestamp', 'Shift date', 'Shift', 'Name', 'Status', 'Source', 'Note'];
var SETTING = {
  TAB: 'Attendance tab',
  OPEN: 'Self check-in enabled (TRUE/FALSE)',
  ADMINS: 'Admin emails (comma-separated)',
  PARTIAL: 'Partial after (minutes after shift start, blank = never)',
  ORG: 'Organization name'
};
var DEFAULT_SETTINGS = [
  [SETTING.TAB, 'Offseason 2026'],
  [SETTING.OPEN, 'TRUE'],
  [SETTING.ADMINS, ''],
  [SETTING.PARTIAL, '30'],
  [SETTING.ORG, 'SOR']
];
// Column-A labels in the attendance tab (matched case-insensitively by prefix).
var LABEL = {
  DATE: 'Date',
  SHIFT: 'Shift Number',
  LOCATION: 'Location',
  START: 'Start Time',
  END: 'End Time',
  STUDENTS_AFTER: 'Avg Attendees'
};
var EARLY_CHECK_IN_MINUTES = 30; // check-in opens this long before a shift starts
var CODE_STEP_SECONDS = 30; // a code is accepted for its window plus the previous one
var CODE_DIGITS = 6;
var MAX_CODE_FAILURES = 5;  // per account per 10 minutes
var TOKEN_CACHE_SECONDS = 300;

// ---------------------------------------------------------------- Menu / setup

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Check-in')
    .addItem('Set up check-in tabs', 'setup')
    .addItem('Add student names to Check-in Roster', 'fillRosterNames')
    .addItem('Set Firebase API key…', 'promptFirebaseApiKey')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();

  var settings = getOrCreateSheet_(ss, TAB.SETTINGS, ['Setting', 'Value']);
  settings.getRange('B:B').setNumberFormat('@');
  var existing = settings.getRange(1, 1, settings.getLastRow(), 1).getValues()
    .map(function (r) { return r[0]; });
  DEFAULT_SETTINGS.forEach(function (row) {
    if (existing.indexOf(row[0]) !== -1) return;
    settings.appendRow([row[0], String(row[1])]);
  });

  getOrCreateSheet_(ss, TAB.ROSTER, ROSTER_HEADERS);
  getOrCreateSheet_(ss, TAB.LOG, LOG_HEADERS);
  getCodeSecret_();
  fillRosterNames();
}

/** Appends every student in the attendance tab that is not yet on the Check-in Roster. */
function fillRosterNames() {
  var ss = SpreadsheetApp.getActive();
  var roster = getOrCreateSheet_(ss, TAB.ROSTER, ROSTER_HEADERS);
  var known = getRoster_().map(function (m) { return m.name; });
  var added = readLayout_(getAttendanceSheet_()).students
    .filter(function (s) {
      return !known.some(function (k) { return AttendanceLogic.sameName(k, s.name); });
    })
    .map(function (s) { return [s.name, '']; });
  if (added.length) {
    roster.getRange(roster.getLastRow() + 1, 1, added.length, 2).setValues(added);
  }
  SpreadsheetApp.getUi().alert(added.length + ' name(s) added to "' + TAB.ROSTER +
    '". Fill in each student\'s Google email in column B.');
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

// ---------------------------------------------------------------- API

var ACTIONS = {
  me: apiMe_,
  checkIn: apiCheckIn_,
  shifts: apiShifts_,
  displayCode: apiDisplayCode_,
  getShift: apiGetShift_,
  saveShift: apiSaveShift_,
  listAdmins: apiListAdmins_,
  addAdmin: apiAddAdmin_,
  removeAdmin: apiRemoveAdmin_
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
  var member = AttendanceLogic.findMemberByEmail(getRoster_(), user.email);
  return {
    orgName: settings.orgName,
    today: todayKey_(),
    email: user.email,
    member: member && { name: member.name },
    isAdmin: isAdmin_(user, settings),
    selfCheckIn: settings.selfCheckIn,
    statuses: AttendanceLogic.STATUSES
  };
}

/** The signed-in student checks in with the code currently on a display. */
function apiCheckIn_(user, req) {
  var settings = getSettings_();
  if (!settings.selfCheckIn) throw new Error('Check-in is currently closed.');

  var member = AttendanceLogic.findMemberByEmail(getRoster_(), user.email);
  if (!member) {
    throw new Error(user.email + ' is not on the Check-in Roster. Ask an admin to add it.');
  }

  var now = new Date();
  var cache = CacheService.getScriptCache();
  var failKey = 'codefail_' + user.email;
  var failures = Number(cache.get(failKey) || 0);
  if (failures >= MAX_CODE_FAILURES) {
    throw new Error('Too many wrong codes. Wait 10 minutes and try again.');
  }

  var sheet = getAttendanceSheet_();
  var layout = readLayout_(sheet);
  var today = todayKey_();
  var shifts = readShifts_(sheet, layout, today);
  if (!shifts.length) throw new Error('There is no shift today in "' + sheet.getName() + '".');

  var code = AttendanceLogic.normalizeCode(req.code);
  var location = findLocationForCode_(locationsOf_(shifts), today, code, now.getTime());
  if (location === null) {
    cache.put(failKey, String(failures + 1), 600);
    appendLog_([[now, today, '', member.name, 'Rejected', 'Self check-in',
      user.email + ': wrong or expired code "' + code + '"']]);
    throw new Error('That code is wrong or has expired. Enter the code on the screen right now.');
  }
  var shift = currentShift_(shifts, location);
  if (!shift) {
    throw new Error('No ' + (location || '') + ' shift is running right now. Check-in opens ' +
      EARLY_CHECK_IN_MINUTES + ' minutes before a shift starts.');
  }

  var student = findStudent_(layout, member.name);
  var incoming = AttendanceLogic.statusForCheckIn(nowMinutes_(), shift.start, settings.partialAfter);
  return withLock_(function () {
    var cell = sheet.getRange(student.row, shift.col);
    var result = AttendanceLogic.mergeCheckIn(cell.getValue(), incoming);
    if (result.changed) {
      cell.setValue(result.status);
      appendLog_([[now, shift.date, shift.label, student.name, result.status, 'Self check-in',
        user.email]]);
    }
    return {
      name: student.name,
      status: result.status,
      alreadyMarked: !result.changed,
      shift: shift.label
    };
  });
}

/** Admin: shifts on a date (default today) and which one is running now. */
function apiShifts_(user, req) {
  requireAdmin_(user);
  var date = req.date || todayKey_();
  assertDateKey_(date);
  var sheet = getAttendanceSheet_();
  var shifts = readShifts_(sheet, readLayout_(sheet), date);
  var idx = date === todayKey_() ? AttendanceLogic.pickDefaultShift(shifts, nowMinutes_()) : 0;
  var def = shifts.length ? shifts[Math.max(idx, 0)] : null;
  return {
    date: date,
    tab: sheet.getName(),
    shifts: shifts.map(function (s) { return { key: s.key, label: s.label }; }),
    defaultKey: def ? def.key : null,
    locations: locationsOf_(shifts),
    defaultLocation: def ? def.location : null
  };
}

/**
 * Admin: the code to show on the display for one of today's locations, and
 * the shift a check-in right now would count for.
 */
function apiDisplayCode_(user, req) {
  requireAdmin_(user);
  var sheet = getAttendanceSheet_();
  var today = todayKey_();
  var shifts = readShifts_(sheet, readLayout_(sheet), today);
  var location = String(req.location == null ? '' : req.location);
  if (locationsOf_(shifts).indexOf(location) === -1) {
    throw new Error('No ' + (location || '') + ' shift today. Reload and pick a location.');
  }
  var shift = currentShift_(shifts, location);
  var nowMs = new Date().getTime();
  var step = AttendanceLogic.timeStep(nowMs, CODE_STEP_SECONDS);
  return {
    code: codeFor_(today, location, step),
    expiresInMs: (step + 1) * CODE_STEP_SECONDS * 1000 - nowMs,
    stepSeconds: CODE_STEP_SECONDS,
    location: location,
    shift: shift ? shift.label : null,
    orgName: getSettings_().orgName
  };
}

/** Admin: every student's status for one shift. */
function apiGetShift_(user, req) {
  requireAdmin_(user);
  assertDateKey_(req.date);
  var sheet = getAttendanceSheet_();
  var layout = readLayout_(sheet);
  var shift = findShiftByKey_(readShifts_(sheet, layout, req.date), req.shiftKey);
  var values = readColumn_(sheet, layout, shift.col);
  var members = layout.students.map(function (s) {
    return { name: s.name, status: String(values[s.row - 1] || '') };
  });
  return {
    date: req.date,
    shift: { key: shift.key, label: shift.label },
    members: members,
    summary: AttendanceLogic.summarize(members.map(function (m) { return m.status; }))
  };
}

/** Admin: save statuses for one shift. req.records = [{name, status}] */
function apiSaveShift_(user, req) {
  requireAdmin_(user);
  assertDateKey_(req.date);
  var records = req.records || [];
  records.forEach(function (r) {
    if (!AttendanceLogic.isValidStatus(r.status)) throw new Error('Invalid status: ' + r.status);
  });
  var sheet = getAttendanceSheet_();
  var layout = readLayout_(sheet);
  var shift = findShiftByKey_(readShifts_(sheet, layout, req.date), req.shiftKey);

  return withLock_(function () {
    var values = readColumn_(sheet, layout, shift.col);
    var now = new Date();
    var logRows = [];
    records.forEach(function (r) {
      var student = findStudent_(layout, r.name);
      var before = String(values[student.row - 1] || '');
      if (before === r.status) return;
      // Write single cells so formulas and blank separator rows are never touched.
      sheet.getRange(student.row, shift.col).setValue(r.status);
      logRows.push([now, shift.date, shift.label, student.name, r.status, 'Admin',
        user.email + (before ? ' (was ' + before + ')' : '')]);
    });
    appendLog_(logRows);
    return { saved: logRows.length };
  });
}

/** Admin: everyone who is an admin, and which ones are permanent. */
function apiListAdmins_(user) {
  requireAdmin_(user);
  var permanent = permanentAdmins_();
  var listed = getSettings_().adminEmails.filter(function (e) { return permanent.indexOf(e) === -1; });
  return {
    admins: permanent.map(function (e) { return { email: e, permanent: true }; })
      .concat(listed.map(function (e) { return { email: e, permanent: false }; }))
  };
}

function apiAddAdmin_(user, req) {
  requireAdmin_(user);
  var email = String(req.email == null ? '' : req.email).trim().toLowerCase();
  if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) {
    throw new Error('"' + email + '" is not a valid email address.');
  }
  return withLock_(function () {
    var admins = getSettings_().adminEmails;
    if (admins.indexOf(email) === -1 && permanentAdmins_().indexOf(email) === -1) {
      setSetting_(SETTING.ADMINS, admins.concat([email]).join(', '));
      appendLog_([[new Date(), todayKey_(), '', email, 'Admin added', 'Admin', user.email]]);
    }
    return apiListAdmins_(user);
  });
}

function apiRemoveAdmin_(user, req) {
  requireAdmin_(user);
  var email = String(req.email == null ? '' : req.email).trim().toLowerCase();
  if (permanentAdmins_().indexOf(email) !== -1) {
    throw new Error(email + ' is a permanent admin and can\'t be removed here.');
  }
  if (email === user.email) throw new Error('You can\'t remove yourself. Ask another admin.');
  return withLock_(function () {
    var admins = getSettings_().adminEmails;
    if (admins.indexOf(email) !== -1) {
      setSetting_(SETTING.ADMINS, admins.filter(function (e) { return e !== email; }).join(', '));
      appendLog_([[new Date(), todayKey_(), '', email, 'Admin removed', 'Admin', user.email]]);
    }
    return apiListAdmins_(user);
  });
}

// ---------------------------------------------------------------- Attendance tab

function getAttendanceSheet_() {
  var name = getSettings_().attendanceTab;
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) {
    throw new Error('Attendance tab "' + name + '" not found. Fix it in "' + TAB.SETTINGS + '".');
  }
  return sheet;
}

/** Finds the header rows by their column-A labels, and the student rows below them. */
function readLayout_(sheet) {
  var lastRow = sheet.getLastRow();
  var colA = sheet.getRange(1, 1, lastRow, 1).getValues().map(function (r) { return r[0]; });
  var row = function (label) {
    var i = AttendanceLogic.findLabelRow(colA, label);
    if (i === -1) {
      throw new Error('Could not find a "' + label + '" row in column A of "' + sheet.getName() + '".');
    }
    return i + 1;
  };
  var layout = {
    dateRow: row(LABEL.DATE),
    shiftRow: row(LABEL.SHIFT),
    locationRow: row(LABEL.LOCATION),
    startRow: row(LABEL.START),
    endRow: row(LABEL.END),
    students: []
  };
  for (var r = row(LABEL.STUDENTS_AFTER) + 1; r <= lastRow; r++) {
    var name = String(colA[r - 1] == null ? '' : colA[r - 1]).trim();
    if (name) layout.students.push({ name: name, row: r });
  }
  layout.lastRow = lastRow;
  return layout;
}

/** Shift columns whose date is dateKey, left to right. */
function readShifts_(sheet, layout, dateKey) {
  var lastCol = sheet.getLastColumn();
  var height = Math.max(layout.dateRow, layout.shiftRow, layout.locationRow,
    layout.startRow, layout.endRow);
  var range = sheet.getRange(1, 1, height, lastCol);
  var values = range.getValues();
  var shown = range.getDisplayValues();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var shifts = [];
  var seen = {};
  for (var c = 2; c <= lastCol; c++) {
    var d = values[layout.dateRow - 1][c - 1];
    if (!isDate_(d) || Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== dateKey) continue;
    var cell = function (r) { return String(shown[r - 1][c - 1]).trim(); };
    var s = {
      col: c,
      date: dateKey,
      shiftNumber: cell(layout.shiftRow),
      location: cell(layout.locationRow),
      start: AttendanceLogic.parseTimeLoose(cell(layout.startRow)),
      end: AttendanceLogic.parseTimeLoose(cell(layout.endRow))
    };
    s.label = AttendanceLogic.shiftLabel(s);
    // Keyed by header contents, not column letter, so inserting a column
    // elsewhere doesn't invalidate the code on the display.
    s.key = [dateKey, s.shiftNumber, s.location, cell(layout.startRow)].join('|');
    if (seen[s.key]) s.key += '|' + c;
    seen[s.key] = true;
    shifts.push(s);
  }
  return shifts;
}

function findShiftByKey_(shifts, key) {
  for (var i = 0; i < shifts.length; i++) {
    if (shifts[i].key === key) return shifts[i];
  }
  throw new Error('That shift no longer exists. Reload and pick it again.');
}

function findStudent_(layout, name) {
  for (var i = 0; i < layout.students.length; i++) {
    if (AttendanceLogic.sameName(layout.students[i].name, name)) return layout.students[i];
  }
  throw new Error('"' + name + '" is not a row in the attendance tab. Check the spelling on ' +
    'the Check-in Roster.');
}

function readColumn_(sheet, layout, col) {
  return sheet.getRange(1, col, layout.lastRow, 1).getValues().map(function (r) { return r[0]; });
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

function getSettings_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(TAB.SETTINGS);
  var map = {};
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      map[r[0]] = r[1];
    });
  }
  var get = function (key) {
    if (map.hasOwnProperty(key)) return map[key];
    for (var i = 0; i < DEFAULT_SETTINGS.length; i++) {
      if (DEFAULT_SETTINGS[i][0] === key) return DEFAULT_SETTINGS[i][1];
    }
    return '';
  };
  return {
    attendanceTab: String(get(SETTING.TAB)).trim(),
    selfCheckIn: String(get(SETTING.OPEN)).toUpperCase() !== 'FALSE',
    adminEmails: AttendanceLogic.parseEmailList(get(SETTING.ADMINS)),
    partialAfter: AttendanceLogic.parseMinutes(get(SETTING.PARTIAL)),
    orgName: String(get(SETTING.ORG) || 'SOR')
  };
}

function getRoster_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(TAB.ROSTER);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
    .map(function (r) {
      return { name: String(r[0]).trim(), email: String(r[1]).trim().toLowerCase() };
    })
    .filter(function (m) { return m.name; });
}

function appendLog_(rows) {
  if (!rows.length) return;
  var log = SpreadsheetApp.getActive().getSheetByName(TAB.LOG) ||
    getOrCreateSheet_(SpreadsheetApp.getActive(), TAB.LOG, LOG_HEADERS);
  log.getRange(log.getLastRow() + 1, 1, rows.length, LOG_HEADERS.length).setValues(rows);
}

/**
 * Verifies a Firebase ID token with Google and returns {email, name}.
 * The API key ties the token to your Firebase project.
 */
function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('Please sign in.');
  var apiKey = PropertiesService.getScriptProperties().getProperty('FIREBASE_API_KEY');
  if (!apiKey) throw new Error('Server not configured: run Check-in > Set Firebase API key…');

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

/** PERMANENT_ADMINS plus the account the script runs as (the one that set it up). */
function permanentAdmins_() {
  var out = PERMANENT_ADMINS.map(function (e) { return e.toLowerCase(); });
  var owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (owner && out.indexOf(owner) === -1) out.push(owner);
  return out;
}

function isAdmin_(user, settings) {
  return permanentAdmins_().indexOf(user.email) !== -1 ||
    settings.adminEmails.indexOf(user.email) !== -1;
}

/** Writes a Check-in Settings value, adding the row if it is missing. */
function setSetting_(key, value) {
  var sheet = getOrCreateSheet_(SpreadsheetApp.getActive(), TAB.SETTINGS, ['Setting', 'Value']);
  var last = sheet.getLastRow();
  var keys = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sheet.getRange(last + 1, 1, 1, 2).setValues([[key, value]]);
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

/** Distinct locations among the shifts, in column order. */
function locationsOf_(shifts) {
  var out = [];
  shifts.forEach(function (s) {
    if (out.indexOf(s.location) === -1) out.push(s.location);
  });
  return out;
}

/** The shift at this location that a check-in right now counts for, or null. */
function currentShift_(shifts, location) {
  var here = shifts.filter(function (s) { return s.location === location; });
  var i = AttendanceLogic.pickShiftAt(here, nowMinutes_(), EARLY_CHECK_IN_MINUTES);
  return i === -1 ? null : here[i];
}

/** Each location gets its own code sequence per day, so the code says where the student is. */
function codeFor_(dateKey, location, step) {
  var sig = Utilities.computeHmacSha256Signature(
    step + '|' + dateKey + '|' + location.toLowerCase(), getCodeSecret_());
  return AttendanceLogic.truncateToCode(sig, CODE_DIGITS);
}

/** The location whose current or previous code matches, or null. */
function findLocationForCode_(locations, dateKey, code, nowMs) {
  if (code.length !== CODE_DIGITS) return null;
  var step = AttendanceLogic.timeStep(nowMs, CODE_STEP_SECONDS);
  for (var i = 0; i < locations.length; i++) {
    if (code === codeFor_(dateKey, locations[i], step) ||
      code === codeFor_(dateKey, locations[i], step - 1)) {
      return locations[i];
    }
  }
  return null;
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function assertDateKey_(dateKey) {
  if (!AttendanceLogic.isValidDateKey(dateKey)) throw new Error('Invalid date: ' + dateKey);
}

function todayKey_() {
  return Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    'yyyy-MM-dd');
}

function nowMinutes_() {
  var hhmm = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
    'HH:mm');
  return AttendanceLogic.parseTimeLoose(hhmm);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
