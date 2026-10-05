/**
 * Pure attendance logic with no Google Apps Script dependencies.
 * Loaded as a global in Apps Script and via require() in Node tests.
 */
var AttendanceLogic = (function () {
  // Must match the dropdown (data validation) used in the team's attendance tab.
  var STATUSES = ['Not Present', 'Present', 'Partial', 'Unproductive',
    'Absent Excused', 'Absent Unexcused'];
  // Statuses a self check-in may replace with "Present". Partial/Unproductive
  // are admin judgements and are never overwritten by a check-in.
  var CHECK_IN_REPLACES = ['', 'Not Present', 'Absent Excused', 'Absent Unexcused'];
  var DATE_KEY_RE = /^\d{4}-\d{2}-\d{2}$/;
  var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  function normalize(s) {
    return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toLowerCase();
  }

  function isValidStatus(status) {
    return STATUSES.indexOf(status) !== -1;
  }

  function isValidDateKey(key) {
    if (!DATE_KEY_RE.test(String(key))) return false;
    var parts = key.split('-').map(Number);
    var d = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return d.getUTCFullYear() === parts[0] &&
      d.getUTCMonth() === parts[1] - 1 &&
      d.getUTCDate() === parts[2];
  }

  /**
   * Minutes after midnight from a time cell's displayed text, or null.
   * Handles "18:15", "18:15:00", "6:15 PM", and edited cells such as
   * "3:45:00 PM --> 4:45" (the last time wins; AM/PM carries over).
   */
  function parseTimeLoose(text) {
    var re = /(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap]\.?m\.?)?/gi;
    var m;
    var found = [];
    while ((m = re.exec(String(text == null ? '' : text))) !== null) {
      found.push({ h: Number(m[1]), min: Number(m[2]), ap: m[3] ? m[3][0].toLowerCase() : null });
    }
    if (!found.length) return null;
    var t = found[found.length - 1];
    var ap = t.ap;
    for (var i = found.length - 2; !ap && i >= 0; i--) ap = found[i].ap;
    if (t.h > 23 || t.min > 59) return null;
    var h = t.h;
    if (ap === 'p' && h < 12) h += 12;
    if (ap === 'a' && h === 12) h = 0;
    return h * 60 + t.min;
  }

  function formatMinutes(min) {
    if (min == null) return '?';
    var h = Math.floor(min / 60);
    var m = min % 60;
    var ap = h >= 12 ? 'PM' : 'AM';
    var h12 = h % 12 === 0 ? 12 : h % 12;
    return h12 + ':' + (m < 10 ? '0' : '') + m + ' ' + ap;
  }

  function weekday(dateKey) {
    var p = dateKey.split('-').map(Number);
    return DAY_NAMES[new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay()];
  }

  /** Human label for a shift, e.g. "Sat 10/4 · Shift 2 · Hangar 391 · 1:00 PM–4:15 PM". */
  function shiftLabel(s) {
    var p = s.date.split('-').map(Number);
    var parts = [weekday(s.date) + ' ' + p[1] + '/' + p[2]];
    if (s.shiftNumber) parts.push('Shift ' + s.shiftNumber);
    if (s.location) parts.push(s.location);
    parts.push(formatMinutes(s.start) + '–' + formatMinutes(s.end));
    return parts.join(' · ');
  }

  /**
   * Which of today's shifts the display should default to: one running now
   * (or starting within 30 min), preferring the latest start; otherwise the
   * next one; otherwise the last one. Returns an index or -1.
   */
  function pickDefaultShift(shifts, nowMin) {
    if (!shifts.length) return -1;
    var best = -1;
    shifts.forEach(function (s, i) {
      var start = s.start == null ? 0 : s.start;
      var end = s.end == null ? 24 * 60 : s.end;
      if (nowMin >= start - 30 && nowMin <= end) {
        if (best === -1 || start > (shifts[best].start || 0)) best = i;
      }
    });
    if (best !== -1) return best;
    var next = -1;
    shifts.forEach(function (s, i) {
      if (s.start != null && s.start > nowMin &&
        (next === -1 || s.start < shifts[next].start)) next = i;
    });
    return next !== -1 ? next : shifts.length - 1;
  }

  /**
   * The shift a check-in at nowMin belongs to, or -1. A running shift wins;
   * where two overlap, the later-starting (next) one wins. Before any shift
   * starts, check-in opens earlyMin minutes ahead for the next shift.
   * Missing times count as "all day".
   */
  function pickShiftAt(shifts, nowMin, earlyMin) {
    var running = -1;
    var upcoming = -1;
    shifts.forEach(function (s, i) {
      var start = s.start == null ? 0 : s.start;
      var end = s.end == null ? 24 * 60 : s.end;
      if (nowMin >= start && nowMin <= end) {
        if (running === -1 || start >= startOf(shifts[running])) running = i;
      } else if (nowMin < start && nowMin >= start - earlyMin) {
        if (upcoming === -1 || start < startOf(shifts[upcoming])) upcoming = i;
      }
    });
    return running !== -1 ? running : upcoming;
  }

  function startOf(s) {
    return s.start == null ? 0 : s.start;
  }

  /** Row index (0-based) of the first label in column A starting with prefix, or -1. */
  function findLabelRow(colA, prefix) {
    var p = normalize(prefix);
    for (var i = 0; i < colA.length; i++) {
      if (normalize(colA[i]).indexOf(p) === 0) return i;
    }
    return -1;
  }

  /** Finds a roster entry by Google account email (case-insensitive). */
  function findMemberByEmail(roster, email) {
    var e = normalize(email);
    if (!e) return null;
    for (var i = 0; i < roster.length; i++) {
      if (normalize(roster[i].email) === e) return roster[i];
    }
    return null;
  }

  function sameName(a, b) {
    return normalize(a) !== '' && normalize(a) === normalize(b);
  }

  /**
   * "Partial" when checking in more than partialAfterMin minutes after the
   * shift started, otherwise "Present". null partialAfterMin disables it.
   */
  function statusForCheckIn(nowMin, shiftStartMin, partialAfterMin) {
    if (partialAfterMin == null || shiftStartMin == null) return 'Present';
    return nowMin > shiftStartMin + partialAfterMin ? 'Partial' : 'Present';
  }

  /** "parent" when a roster Role cell starts with "parent", otherwise "student". */
  function roleOf(roleCell) {
    return normalize(roleCell).indexOf('parent') === 0 ? 'parent' : 'student';
  }

  function firstName(name) {
    return normalize(name).split(' ')[0] || '';
  }

  /**
   * Whether a parent-slot cell (typed by hand when signing up) refers to this
   * parent: full name, one of their aliases, or just their first name when no
   * other parent on the roster shares it.
   */
  function parentSlotMatches(slotText, parent, allParents) {
    var t = normalize(slotText);
    if (!t) return false;
    if (t === normalize(parent.name)) return true;
    if ((parent.aliases || []).some(function (a) { return normalize(a) === t; })) return true;
    var first = firstName(parent.name);
    if (t !== first) return false;
    return allParents.filter(function (p) { return firstName(p.name) === first; }).length === 1;
  }

  /** "ab***@gmail.com" — enough to recognize your own address, not to read someone else's. */
  function maskEmail(email) {
    var e = String(email == null ? '' : email).trim();
    var at = e.indexOf('@');
    if (at < 1) return '***';
    return e.slice(0, Math.min(2, at)) + '***' + e.slice(at);
  }

  /**
   * Students and parents with emails from the registration form's sheet.
   * header: first row; rows: the data rows. Parent N's email is the first
   * column starting with "Email" after "Parent/Guardian N (First Name)".
   */
  function parseDirectory(header, rows) {
    var h = header.map(function (x) { return normalize(x); });
    var at = function (label) { return h.indexOf(normalize(label)); };
    var sFirst = at('Student Name (First Name)');
    var sLast = at('Student Name (Last Name)');
    var sEmail = at('Student Email');
    var guardians = [];
    for (var n = 1; n <= 6; n++) {
      var f = at('Parent/Guardian ' + n + ' (First Name)');
      if (f === -1) continue;
      var e = -1;
      for (var i = f + 1; i < h.length; i++) {
        if (h[i].indexOf('email') === 0) { e = i; break; }
      }
      guardians.push({ first: f, last: at('Parent/Guardian ' + n + ' (Last Name)'), email: e });
    }
    var clean = function (v) { return String(v == null ? '' : v).trim().replace(/\s+/g, ' '); };
    var fullName = function (r, fi, li) {
      return [fi === -1 ? '' : clean(r[fi]), li === -1 ? '' : clean(r[li])].filter(Boolean).join(' ');
    };
    var out = { students: [], parents: [] };
    rows.forEach(function (r) {
      var student = sFirst === -1 ? '' : fullName(r, sFirst, sLast);
      if (!student) return;
      var se = sEmail === -1 ? '' : clean(r[sEmail]).toLowerCase();
      if (se) out.students.push({ name: student, email: se });
      guardians.forEach(function (g) {
        var name = fullName(r, g.first, g.last);
        var email = g.email === -1 ? '' : clean(r[g.email]).toLowerCase();
        if (name && email) out.parents.push({ name: name, email: email, child: student });
      });
    });
    return out;
  }

  /** What a check-in with status `incoming` does to the student's existing cell. */
  function mergeCheckIn(existing, incoming) {
    var current = String(existing == null ? '' : existing).trim();
    if (CHECK_IN_REPLACES.indexOf(current) !== -1) {
      return { status: incoming || 'Present', changed: true };
    }
    return { status: current, changed: false };
  }

  /** Whole minutes >= 0 from a settings cell, or null when blank/invalid. */
  function parseMinutes(value) {
    var t = String(value == null ? '' : value).trim();
    if (!/^\d+$/.test(t)) return null;
    return Number(t);
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

  function summarize(statuses) {
    var counts = {};
    STATUSES.forEach(function (s) { counts[s] = 0; });
    statuses.forEach(function (s) {
      var k = isValidStatus(s) ? s : 'Not Present';
      counts[k]++;
    });
    return counts;
  }

  return {
    STATUSES: STATUSES,
    isValidStatus: isValidStatus,
    isValidDateKey: isValidDateKey,
    parseTimeLoose: parseTimeLoose,
    formatMinutes: formatMinutes,
    shiftLabel: shiftLabel,
    pickDefaultShift: pickDefaultShift,
    pickShiftAt: pickShiftAt,
    findLabelRow: findLabelRow,
    findMemberByEmail: findMemberByEmail,
    sameName: sameName,
    statusForCheckIn: statusForCheckIn,
    roleOf: roleOf,
    maskEmail: maskEmail,
    parseDirectory: parseDirectory,
    parentSlotMatches: parentSlotMatches,
    mergeCheckIn: mergeCheckIn,
    parseMinutes: parseMinutes,
    timeStep: timeStep,
    truncateToCode: truncateToCode,
    normalizeCode: normalizeCode,
    parseEmailList: parseEmailList,
    summarize: summarize
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = AttendanceLogic;
}
