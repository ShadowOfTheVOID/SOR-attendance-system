/**
 * SOR Attendance System — server side (Google Apps Script, bound to a Sheet).
 *
 * Sheets:
 *   Roster      ID | Name | Group | Active        (you maintain this)
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
var ROSTER_HEADERS = ['ID', 'Name', 'Group', 'Active'];
var ATTENDANCE_HEADERS = ['ID', 'Name', 'Group'];
var LOG_HEADERS = ['Timestamp', 'Date', 'ID', 'Name', 'Status', 'Source', 'Note'];
var FIRST_DATE_COL = ATTENDANCE_HEADERS.length + 1; // 1-based
var DEFAULT_SETTINGS = [
  ['Organization name', 'SOR'],
  ['Late after (HH:mm, blank = never late)', '09:15'],
  ['Self check-in enabled (TRUE/FALSE)', true],
  ['Auto-mark absent time (hour 0-23)', 23]
];
var STATUS_COLORS = {
  Present: '#d9ead3',
  Late: '#fff2cc',
  Absent: '#f4cccc',
  Excused: '#cfe2f3'
};
var MAX_PIN_FAILURES = 10;

// ---------------------------------------------------------------- Menu / setup

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Attendance')
    .addItem('Set up sheets', 'setup')
    .addItem('Set admin PIN…', 'promptAdminPin')
    .addSeparator()
    .addItem('Mark unmarked as Absent (today)', 'markAbsentToday')
    .addItem('Install daily auto-absent trigger', 'installDailyTrigger')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();

  var roster = getOrCreateSheet_(ss, SHEET.ROSTER, ROSTER_HEADERS);
  if (roster.getLastRow() < 2) {
    roster.getRange(2, 1, 2, 4).setValues([
      ['S001', 'Example Member', 'Group A', true],
      ['S002', 'Another Member', 'Group B', true]
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
    if (existing.indexOf(row[0]) === -1) settings.appendRow([row[0], String(row[1])]);
  });

  syncAttendanceRows_(att, getActiveRoster_());
  SpreadsheetApp.getUi().alert(
    'Setup complete.\n\n1. Fill in the Roster sheet.\n' +
    '2. Attendance > Set admin PIN…\n' +
    '3. Deploy > New deployment > Web app to get the check-in link.');
}

function promptAdminPin() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('Admin PIN', 'Enter a new admin PIN (at least 4 characters):',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var pin = res.getResponseText().trim();
  if (pin.length < 4) {
    ui.alert('PIN must be at least 4 characters.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  ui.alert('Admin PIN saved.');
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

// ---------------------------------------------------------------- Web app

function doGet() {
  var settings = getSettings_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(settings.orgName + ' Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Public: configuration and names for the check-in page. */
function getAppConfig() {
  var settings = getSettings_();
  return {
    orgName: settings.orgName,
    today: todayKey_(),
    selfCheckIn: settings.selfCheckIn,
    statuses: AttendanceLogic.STATUSES,
    names: getActiveRoster_().map(function (m) { return m.name; })
  };
}

/** Public: a member checks themselves in by ID or full name. */
function checkIn(query) {
  var settings = getSettings_();
  if (!settings.selfCheckIn) throw new Error('Self check-in is currently closed.');

  var member = AttendanceLogic.findMember(getActiveRoster_(), query);
  if (!member) throw new Error('No active member matches "' + query + '". Check your ID or full name.');

  var now = new Date();
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var dateKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
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
      appendLog_([[now, dateKey, member.id, member.name, result.status, 'Self check-in', '']]);
    }
    return {
      name: member.name,
      status: result.status,
      alreadyMarked: !result.changed,
      time: Utilities.formatDate(now, tz, 'HH:mm')
    };
  });
}

/** Admin: roster with statuses for a date. */
function adminGetDay(pin, dateKey) {
  verifyAdmin_(pin);
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

/** Admin: save statuses for a date. records = [{id, status}] */
function adminSaveDay(pin, dateKey, records) {
  verifyAdmin_(pin);
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
      logRows.push([now, dateKey, r.id, byId[r.id].name, r.status || '(cleared)', 'Admin',
        before ? 'was ' + before : '']);
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
    autoAbsentHour: get(3)
  };
}

function getActiveRoster_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(SHEET.ROSTER);
  if (!sheet) throw new Error('Roster sheet missing. Run Attendance > Set up sheets.');
  if (sheet.getLastRow() < 2) return [];
  var seen = {};
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .filter(function (r) {
      var id = String(r[0]).trim();
      var active = r[3] === true || String(r[3]).toUpperCase() === 'TRUE';
      if (!id || !active || seen[id]) return false;
      seen[id] = true;
      return true;
    })
    .map(function (r) {
      return { id: String(r[0]).trim(), name: String(r[1]).trim(), group: String(r[2]).trim() };
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

function verifyAdmin_(pin) {
  var cache = CacheService.getScriptCache();
  var failures = Number(cache.get('pin_failures') || 0);
  if (failures >= MAX_PIN_FAILURES) {
    throw new Error('Too many wrong PIN attempts. Try again in 10 minutes.');
  }
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!expected) throw new Error('Admin PIN not set. In the sheet: Attendance > Set admin PIN…');
  if (String(pin) !== expected) {
    cache.put('pin_failures', String(failures + 1), 600);
    throw new Error('Wrong PIN.');
  }
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
