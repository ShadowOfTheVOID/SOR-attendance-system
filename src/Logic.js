/**
 * Pure attendance logic with no Google Apps Script dependencies.
 * Loaded as a global in Apps Script and via require() in Node tests.
 */
var AttendanceLogic = (function () {
  var STATUSES = ['Present', 'Late', 'Absent', 'Excused'];
  var DATE_KEY_RE = /^\d{4}-\d{2}-\d{2}$/;

  function isValidStatus(status) {
    return status === '' || STATUSES.indexOf(status) !== -1;
  }

  function isValidDateKey(key) {
    if (!DATE_KEY_RE.test(String(key))) return false;
    var parts = key.split('-').map(Number);
    var d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return d.getUTCFullYear() === parts[0] &&
      d.getUTCMonth() === parts[1] - 1 &&
      d.getUTCDate() === parts[2];
  }

  /** "HH:mm" -> minutes after midnight, or null when blank/invalid. */
  function parseTime(value) {
    var m = /^\s*(\d{1,2}):(\d{2})\s*$/.exec(String(value == null ? '' : value));
    if (!m) return null;
    var h = Number(m[1]);
    var min = Number(m[2]);
    if (h > 23 || min > 59) return null;
    return h * 60 + min;
  }

  /** Status for a self check-in at the given time; no cutoff means always Present. */
  function statusForCheckIn(minutesOfDay, lateCutoffMinutes) {
    if (lateCutoffMinutes == null) return 'Present';
    return minutesOfDay > lateCutoffMinutes ? 'Late' : 'Present';
  }

  function normalize(s) {
    return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();
  }

  /** Finds a roster member by Google account email (case-insensitive). */
  function findMemberByEmail(roster, email) {
    var e = normalize(email);
    if (!e) return null;
    for (var i = 0; i < roster.length; i++) {
      if (normalize(roster[i].email) === e) return roster[i];
    }
    return null;
  }

  /** Index of the rotating-code time window containing nowMs. */
  function timeStep(nowMs, stepSeconds) {
    return Math.floor(nowMs / 1000 / stepSeconds);
  }

  /**
   * RFC 4226 dynamic truncation of an HMAC into a numeric code.
   * Accepts signed bytes (as Apps Script returns them) or unsigned.
   */
  function truncateToCode(hmacBytes, digits) {
    var b = hmacBytes.map(function (x) { return x & 0xff; });
    var offset = b[b.length - 1] & 0x0f;
    var bin = ((b[offset] & 0x7f) << 24) | (b[offset + 1] << 16) |
      (b[offset + 2] << 8) | b[offset + 3];
    var code = String(bin % Math.pow(10, digits));
    while (code.length < digits) code = '0' + code;
    return code;
  }

  /** Strips spaces and anything else that is not a digit. */
  function normalizeCode(input) {
    return String(input == null ? '' : input).replace(/\D/g, '');
  }

  function parseEmailList(value) {
    return String(value == null ? '' : value).split(/[\s,;]+/)
      .map(normalize).filter(Boolean);
  }

  /**
   * Decides what a self check-in does to an existing cell. A check-in only
   * fills a blank or Absent cell; it never overrides Present/Late/Excused.
   */
  function mergeCheckIn(existing, incoming) {
    if (!existing || existing === 'Absent') {
      return { status: incoming, changed: true };
    }
    return { status: existing, changed: false };
  }

  /**
   * Index (0-based) at which to insert dateKey into the header so date
   * columns stay sorted. Non-date headers before firstDateIndex are skipped.
   */
  function insertionIndexForDate(dateHeaders, dateKey, firstDateIndex) {
    for (var i = firstDateIndex; i < dateHeaders.length; i++) {
      if (dateHeaders[i] > dateKey) return i;
    }
    return dateHeaders.length;
  }

  function summarize(statuses) {
    var counts = { Present: 0, Late: 0, Absent: 0, Excused: 0, Unmarked: 0 };
    statuses.forEach(function (s) {
      if (counts.hasOwnProperty(s) && s !== 'Unmarked') counts[s]++;
      else counts.Unmarked++;
    });
    return counts;
  }

  return {
    STATUSES: STATUSES,
    isValidStatus: isValidStatus,
    isValidDateKey: isValidDateKey,
    parseTime: parseTime,
    statusForCheckIn: statusForCheckIn,
    findMemberByEmail: findMemberByEmail,
    timeStep: timeStep,
    truncateToCode: truncateToCode,
    normalizeCode: normalizeCode,
    parseEmailList: parseEmailList,
    mergeCheckIn: mergeCheckIn,
    insertionIndexForDate: insertionIndexForDate,
    summarize: summarize
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AttendanceLogic;
}
