const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./fakeSheets');

const TOKENS = {
  'tok-ana': { email: 'ana@example.com', emailVerified: true },
  'tok-ben': { email: 'BEN@example.com', emailVerified: true },
  'tok-cara': { email: 'cara@example.com', emailVerified: true },
  'tok-ghost': { email: 'ghost@example.com', emailVerified: true },
  'tok-stranger': { email: 'stranger@example.com', emailVerified: true },
  'tok-unverified': { email: 'ana@example.com', emailVerified: false },
  'tok-owner': { email: 'owner@example.com', emailVerified: true },
  'tok-pat': { email: 'pat@example.com', emailVerified: true },     // parent, signed up as "Pat"
  'tok-quinn': { email: 'quinn@example.com', emailVerified: true }, // parent, not signed up
  'tok-tedd': { email: 'NhsTedd@gmail.com', emailVerified: true }
};

const day = (y, m, d) => new Date(Date.UTC(y, m - 1, d));

/**
 * Builds a tab shaped like the team's "Offseason 2026": labels in column A,
 * shifts from column I, students below "Avg Attendees" with a blank separator.
 */
function buildAttendanceTab(env) {
  const sh = env.app.SpreadsheetApp.getActive().insertSheet('Offseason 2026');
  const cols = [
    // date, shift #, location, start, end
    [day(2026, 10, 3), 1, 'Hangar 391', '3:45:00 PM --> 4:45', '7:15:00 PM -->8:15'],
    [day(2026, 10, 4), 1, 'Hangar 391', '9:45 AM', '1:15 PM'],
    [day(2026, 10, 4), 2, 'Hangar 391', '1:00 PM', '4:15 PM'],
    [day(2026, 10, 4), 1, 'Online', '3:00 PM', '4:00 PM'],
    [day(2026, 10, 6), 1, 'Hangar 391', '5:45 PM', '8:45 PM']
  ];
  const label = (r, text) => sh.set(r, 1, text);
  label(1, 'x5');
  label(2, 'Date');
  label(3, 'Shift Number');
  label(4, 'Location');
  label(5, 'Mentor 1');
  label(12, 'Female Parent');
  label(13, 'Parent 1 - Online is half of a shift');
  label(14, 'Parent 2 - Online is half of a shift');
  [15, 16, 17, 18].forEach((r, i) => label(r, 'Parent ' + (i + 3)));
  label(19, 'Start Time\nShould start 15 min. prior');
  label(20, 'End Time\nShould start 15 min. before');
  label(21, 'Notes / Meeting Objective(s)');
  label(30, '=COUNTIFS(...)');
  label(31, 'Avg Attendees');
  cols.forEach(([d, n, loc, start, end], i) => {
    const c = 9 + i;
    sh.set(2, c, d); sh.set(3, c, n); sh.set(4, c, loc); sh.set(19, c, start); sh.set(20, c, end);
    sh.set(30, c, true); sh.set(31, c, '=COUNTIF(...)');
  });
  const students = [[33, 'Ana Alvarez'], [34, 'Ben Brooks'], [36, 'Cara Chen']]; // row 35 blank separator
  students.forEach(([r, name]) => {
    sh.set(r, 1, name);
    sh.set(r, 2, '=TRUNC(...)');
    cols.forEach((_, i) => sh.set(r, 9 + i, 'Not Present'));
  });
  sh.set(34, 10, 'Partial'); // Ben already judged Partial for 10/4 shift 1
  sh.set(13, 11, 'Pat');     // Pat Parker signed up (first name only) for 10/4 shift 2
  sh.set(12, 11, 'Robin Reyes');

  const map = env.app.SpreadsheetApp.getActive().insertSheet('student-parent-mapping');
  map.getRange(1, 1, 1, 16).setValues([['Student Name (First Name)', 'Student Name (Last Name)',
    'Parent 1 first name', 'Parent 1 last name', 'Parent 1 relation', 'Parent 1 role',
    'Parent 2 first name', 'Parent 2 last name', 'Parent 2 relation', 'Parent 2 role',
    'Parent aliases 1', 'Parent aliases 2', 'Parent aliases 3', 'Parent aliases 4',
    'Parent aliases 5', 'Parent aliases 6']]);
  map.getRange(2, 1, 3, 16).setValues([
    ['Ana ', 'Alvarez', 'Pat', 'Parker ', 'Father', 'parent', 'Quinn', 'Parker', 'Mother', 'parent', 'P. Parker', '', '', '', '', ''],
    ['Cara', 'Chen', 'Robin', 'Reyes', 'Mother', 'parent', '', '', '', '', '', '', '', '', '', ''],
    ['Ana sibling', 'Alvarez', 'Pat', 'Parker', 'Father', 'parent', '', '', '', '', '', '', '', '', '', '']
  ]);
  return sh;
}

function fresh(opts = {}) {
  const env = load({ tokens: TOKENS, now: new Date('2026-10-04T13:10:00Z'), ...opts });
  env.att = buildAttendanceTab(env);
  env.app.setup();
  const roster = env.sheets['Check-in Roster'];
  const emails = { 'Ana Alvarez': 'ana@example.com', 'Ben Brooks': 'ben@example.com', 'Cara Chen': 'cara@example.com',
    'Pat Parker': 'pat@example.com', 'Quinn Parker': 'quinn@example.com' };
  for (let r = 2; r <= roster.getLastRow(); r++) roster.set(r, 2, emails[roster.get(r, 1)] || '');
  roster.getRange(roster.getLastRow() + 1, 1, 1, 2).setValues([['Ghost Person', 'ghost@example.com']]);
  env.roster = roster;
  return env;
}

const ok = (res) => { assert.equal(res.ok, true, res.error); return res.data; };
const fail = (res, re) => { assert.equal(res.ok, false, 'expected failure'); assert.match(res.error, re); };
const shifts = (env, date) => ok(env.call('shifts', 'tok-owner', date ? { date } : {}));
const keyOf = (env, labelPart, date) => shifts(env, date).shifts.find((s) => s.label.includes(labelPart)).key;
const display = (env, location) => ok(env.call('displayCode', 'tok-owner', { location }));
const codeAt = (env, location) => display(env, location).code;
const COL = { oct3: 9, oct4s1: 10, oct4s2: 11, oct4online: 12, oct6: 13 };
const ROW = { ana: 33, ben: 34, cara: 36 };
const logRows = (env) => env.sheets['Check-in Log'].cells.slice(1).filter(Boolean);

test('setup adds only the three check-in tabs and lists students on the roster', () => {
  const env = fresh();
  assert.deepEqual(Object.keys(env.sheets).sort(),
    ['Check-in Log', 'Check-in Roster', 'Check-in Settings', 'Offseason 2026', 'student-parent-mapping']);
  const rows = env.sheets['Check-in Roster'].cells.slice(1).map((r) => [r[0], r[2], r[3]]);
  assert.deepEqual(rows, [
    ['Ana Alvarez', 'Student', ''],
    ['Ben Brooks', 'Student', ''],
    ['Cara Chen', 'Student', ''],
    ['Pat Parker', 'Parent', 'P. Parker'],
    ['Quinn Parker', 'Parent', 'P. Parker'],
    ['Robin Reyes', 'Parent', ''],
    ['Ghost Person', undefined, undefined]
  ], 'students, then parents (deduped across siblings) with family aliases');
  assert.ok(env.props.CODE_SECRET.length > 40);
  assert.equal(ok(env.call('me', 'tok-owner')).isAdmin, true);

  env.app.fillRosterNames(); // re-running adds no duplicates
  assert.equal(env.sheets['Check-in Roster'].getLastRow(), 8);
});

test('shifts lists the day\'s columns, locations, and what is running now', () => {
  const env = fresh(); // 13:10 on 10/4: shift 1 (9:45–1:15) and shift 2 (1:00–4:15) overlap
  const s = shifts(env);
  assert.equal(s.date, '2026-10-04');
  assert.deepEqual(s.shifts.map((x) => x.label), [
    'Sun 10/4 · Shift 1 · Hangar 391 · 9:45 AM–1:15 PM',
    'Sun 10/4 · Shift 2 · Hangar 391 · 1:00 PM–4:15 PM',
    'Sun 10/4 · Shift 1 · Online · 3:00 PM–4:00 PM'
  ]);
  assert.equal(s.defaultKey, s.shifts[1].key);
  assert.deepEqual(s.locations, ['Hangar 391', 'Online']);
  assert.equal(s.defaultLocation, 'Hangar 391');
  const oct3 = shifts(env, '2026-10-03').shifts;
  assert.equal(oct3[0].label, 'Sat 10/3 · Shift 1 · Hangar 391 · 4:45 PM–8:15 PM', 'edited times use the new value');
});

test('the check-in time picks the shift; the overlap counts for the next shift', () => {
  const env = fresh();
  const at = (iso, tok) => { env.setTime(iso); return ok(env.call('checkIn', tok, { code: codeAt(env, 'Hangar 391') })); };

  assert.match(at('2026-10-04T09:20:00Z', 'tok-ana').shift, /Shift 1/, '25 min early: shift 1');
  assert.equal(env.att.get(ROW.ana, COL.oct4s1), 'Present');

  assert.match(at('2026-10-04T13:05:00Z', 'tok-cara').shift, /Shift 2/, 'overlap: next shift');
  assert.equal(env.att.get(ROW.cara, COL.oct4s2), 'Present');
  assert.equal(env.att.get(ROW.cara, COL.oct4s1), 'Not Present');

  assert.match(at('2026-10-04T13:20:00Z', 'tok-ana').shift, /Shift 2/, 'staying for shift 2 checks in again');
  assert.equal(env.att.get(ROW.ana, COL.oct4s2), 'Present');
  assert.equal(env.att.get(ROW.ana, 2), '=TRUNC(...)', 'formula columns untouched');
  assert.equal(env.att.get(35, COL.oct4s2), '', 'separator row untouched');
});

test('checking in more than 30 minutes after the shift starts marks Partial', () => {
  const env = fresh({ now: new Date('2026-10-04T13:30:00Z') }); // shift 2 started 1:00
  assert.equal(ok(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') })).status, 'Present');
  env.setTime('2026-10-04T13:31:00Z');
  const r = ok(env.call('checkIn', 'tok-cara', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(r.status, 'Partial');
  assert.equal(env.att.get(ROW.cara, COL.oct4s2), 'Partial');
  assert.equal(logRows(env).at(-1)[4], 'Partial');
});

test('the Partial threshold is configurable and can be turned off', () => {
  const env = fresh({ now: new Date('2026-10-04T13:20:00Z') });
  const s = env.sheets['Check-in Settings'];
  const row = s.cells.findIndex((r) => r && String(r[0]).startsWith('Partial after')) + 1;
  assert.equal(s.get(row, 2), '30', 'setup writes the default');
  s.set(row, 2, '15');
  assert.equal(ok(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') })).status, 'Partial');
  s.set(row, 2, '');
  env.setTime('2026-10-04T16:00:00Z');
  assert.equal(ok(env.call('checkIn', 'tok-cara', { code: codeAt(env, 'Hangar 391') })).status, 'Present');
});

test('the display shows which shift check-ins are going to', () => {
  const env = fresh({ now: new Date('2026-10-04T13:05:00Z') });
  assert.match(display(env, 'Hangar 391').shift, /Shift 2 · Hangar 391/);
  assert.equal(display(env, 'Online').shift, null, 'online shift not open yet');
  env.setTime('2026-10-04T14:35:00Z');
  assert.match(display(env, 'Online').shift, /Online/);
  fail(env.call('displayCode', 'tok-owner', { location: 'Moon Base' }), /No Moon Base shift today/);
});

test('each location has its own code; a Hangar code never marks the Online column', () => {
  const env = fresh({ now: new Date('2026-10-04T15:10:00Z') });
  const hangar = codeAt(env, 'Hangar 391');
  const online = codeAt(env, 'Online');
  assert.notEqual(hangar, online);
  ok(env.call('checkIn', 'tok-ana', { code: hangar }));
  ok(env.call('checkIn', 'tok-cara', { code: online }));
  assert.equal(env.att.get(ROW.ana, COL.oct4s2), 'Partial', '2h10m into the Hangar shift');
  assert.equal(env.att.get(ROW.ana, COL.oct4online), 'Not Present');
  assert.equal(env.att.get(ROW.cara, COL.oct4online), 'Present', '10 min into the Online shift');
  assert.equal(env.att.get(ROW.cara, COL.oct4s2), 'Not Present');
});

test('a valid code outside any shift time is refused without counting as a wrong code', () => {
  const env = fresh({ now: new Date('2026-10-04T08:00:00Z') });
  const code = codeAt(env, 'Hangar 391');
  for (let i = 0; i < 6; i++) fail(env.call('checkIn', 'tok-ana', { code }), /No Hangar 391 shift is running/);
  env.setTime('2026-10-04T09:30:00Z');
  ok(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') }));
});

test('check-in never overrides Partial/Unproductive, and repeats are no-ops', () => {
  const env = fresh({ now: new Date('2026-10-04T10:00:00Z') });
  const r = ok(env.call('checkIn', 'tok-ben', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(r.alreadyMarked, true);
  assert.equal(env.att.get(ROW.ben, COL.oct4s1), 'Partial');

  ok(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(ok(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') })).alreadyMarked, true);
  assert.equal(logRows(env).filter((r) => r[4] === 'Present').length, 1);
});

test('codes rotate: the previous window is accepted, older codes are not', () => {
  const env = fresh();
  env.setTime('2026-10-04T13:10:05Z');
  const code = codeAt(env, 'Hangar 391');
  env.setTime('2026-10-04T13:10:35Z');
  assert.notEqual(codeAt(env, 'Hangar 391'), code);
  ok(env.call('checkIn', 'tok-ana', { code }));
  env.setTime('2026-10-04T13:11:05Z');
  fail(env.call('checkIn', 'tok-cara', { code }), /wrong or has expired/);
  assert.equal(env.att.get(ROW.cara, COL.oct4s2), 'Not Present');
  assert.equal(logRows(env).at(-1)[4], 'Rejected');
});

test('a code for another day\'s shift does not work today', () => {
  const env = fresh();
  const step = Math.floor(Date.parse('2026-10-04T13:10:00Z') / 30000);
  const forged = env.app.codeFor_('2026-10-06', 'Hangar 391', step);
  fail(env.call('checkIn', 'tok-ana', { code: forged }), /wrong/);
});

test('wrong codes are rate-limited per account', () => {
  const env = fresh();
  for (let i = 0; i < 5; i++) fail(env.call('checkIn', 'tok-ana', { code: '000000' }), /wrong/);
  fail(env.call('checkIn', 'tok-ana', { code: codeAt(env, 'Hangar 391') }), /Too many/);
  ok(env.call('checkIn', 'tok-cara', { code: codeAt(env, 'Hangar 391') }));
});

test('check-in errors: not on roster, roster name missing from tab, no shift today, closed', () => {
  const env = fresh();
  fail(env.call('checkIn', 'tok-stranger', { code: codeAt(env, 'Hangar 391') }), /not on the Check-in Roster/);
  fail(env.call('checkIn', 'tok-ghost', { code: codeAt(env, 'Hangar 391') }), /not a row in the attendance tab/);

  env.setTime('2026-10-05T18:00:00Z');
  fail(env.call('checkIn', 'tok-ana', { code: '123456' }), /no shift today/);

  const s = env.sheets['Check-in Settings'];
  const row = s.cells.findIndex((r) => r && String(r[0]).startsWith('Self check-in')) + 1;
  s.set(row, 2, 'FALSE');
  fail(env.call('checkIn', 'tok-ana', { code: '123456' }), /closed/);
});

test('auth: tokens are required, verified, and cached', () => {
  const env = fresh();
  fail(env.call('me', undefined), /sign in/i);
  fail(env.call('me', 'forged'), /expired/i);
  fail(env.call('me', 'tok-unverified'), /verified/i);
  fail(env.call('nope', 'tok-ana'), /Unknown action/);
  const before = env.fetches.length;
  ok(env.call('me', 'tok-ana'));
  ok(env.call('me', 'tok-ana'));
  assert.equal(env.fetches.length, before + 1);
  assert.deepEqual(ok(env.call('me', 'tok-ana')).member, { name: 'Ana Alvarez', role: 'student' });
});

test('uses the built-in Firebase API key when no script property overrides it', () => {
  const builtIn = 'AIzaSyCDGs1NXDoFwP0jBM5bASz-XfEcXdYOE4M';
  const env = fresh({ apiKey: builtIn, keyInProps: false });
  assert.equal(ok(env.call('me', 'tok-ana')).member.name, 'Ana Alvarez');
  assert.match(env.fetches[0], new RegExp('key=' + builtIn));
});

test('opens the team spreadsheet by ID (standalone script)', () => {
  const env = fresh();
  ok(env.call('me', 'tok-ana'));
  assert.deepEqual([...new Set(env.openedIds)], ['1cNJ4zwLjHkr8MOZk4QYvyJmILBFWjarDGFOgwaHgiJA']);
});

test('admin-only actions reject students', () => {
  const env = fresh();
  fail(env.call('shifts', 'tok-ana', {}), /Admins only/);
  fail(env.call('displayCode', 'tok-ana', { location: 'Hangar 391' }), /Admins only/);
  fail(env.call('getShift', 'tok-ana', { date: '2026-10-04', shiftKey: 'x' }), /Admins only/);
  fail(env.call('saveShift', 'tok-ana', { date: '2026-10-04', shiftKey: 'x', records: [] }), /Admins only/);
  fail(env.call('listAdmins', 'tok-ana'), /Admins only/);
  fail(env.call('addAdmin', 'tok-ana', { email: 'ana@example.com' }), /Admins only/);
  fail(env.call('removeAdmin', 'tok-ana', { email: 'x@example.com' }), /Admins only/);
});

test('admin can read and edit any shift; changes are logged', () => {
  const env = fresh();
  const key = keyOf(env, 'Hangar', '2026-10-06');
  const before = ok(env.call('getShift', 'tok-owner', { date: '2026-10-06', shiftKey: key }));
  assert.deepEqual(before.members.map((m) => [m.name, m.status]),
    [['Ana Alvarez', 'Not Present'], ['Ben Brooks', 'Not Present'], ['Cara Chen', 'Not Present']]);

  const r = ok(env.call('saveShift', 'tok-owner', {
    date: '2026-10-06', shiftKey: key,
    records: [{ name: 'Ana Alvarez', status: 'Absent Excused' }, { name: 'Ben Brooks', status: 'Not Present' }]
  }));
  assert.equal(r.saved, 1, 'unchanged values are skipped');
  assert.equal(env.att.get(ROW.ana, COL.oct6), 'Absent Excused');
  assert.match(logRows(env).at(-1)[6], /owner@example.com \(was Not Present\)/);
});

test('admin input is validated', () => {
  const env = fresh();
  const key = keyOf(env, 'Shift 2');
  fail(env.call('getShift', 'tok-owner', { date: '2026-13-01', shiftKey: key }), /Invalid date/);
  fail(env.call('saveShift', 'tok-owner', { date: '2026-10-04', shiftKey: key,
    records: [{ name: 'Ana Alvarez', status: 'Late' }] }), /Invalid status/);
  fail(env.call('getShift', 'tok-owner', { date: '2026-10-04', shiftKey: 'nope' }), /no longer exists/);
});

test('a wrong attendance tab name gives a clear error', () => {
  const env = fresh();
  const s = env.sheets['Check-in Settings'];
  const row = s.cells.findIndex((r) => r && r[0] === 'Attendance tab') + 1;
  s.set(row, 2, 'Build Season 2027');
  fail(env.call('shifts', 'tok-owner', {}), /"Build Season 2027" not found/);
});

test('the built-in admin and the setup account are permanent admins', () => {
  const env = fresh();
  assert.equal(ok(env.call('me', 'tok-tedd')).isAdmin, true);
  assert.deepEqual(ok(env.call('listAdmins', 'tok-tedd')).admins, [
    { email: 'nhstedd@gmail.com', permanent: true },
    { email: 'owner@example.com', permanent: true }
  ]);
  fail(env.call('removeAdmin', 'tok-owner', { email: 'nhstedd@gmail.com' }), /permanent/);
});

test('admins can add and remove other admins from the website', () => {
  const env = fresh();
  const added = ok(env.call('addAdmin', 'tok-tedd', { email: ' Ana@Example.com ' })).admins;
  assert.deepEqual(added.at(-1), { email: 'ana@example.com', permanent: false });
  assert.equal(ok(env.call('me', 'tok-ana')).isAdmin, true);
  ok(env.call('displayCode', 'tok-ana', { location: 'Hangar 391' }));

  ok(env.call('addAdmin', 'tok-tedd', { email: 'ana@example.com' })); // no duplicate
  assert.equal(ok(env.call('listAdmins', 'tok-tedd')).admins.length, 3);

  fail(env.call('removeAdmin', 'tok-ana', { email: 'ana@example.com' }), /yourself/);
  ok(env.call('addAdmin', 'tok-ana', { email: 'cara@example.com' }));
  ok(env.call('removeAdmin', 'tok-tedd', { email: 'ana@example.com' }));
  assert.equal(ok(env.call('me', 'tok-ana')).isAdmin, false);
  assert.equal(ok(env.call('me', 'tok-cara')).isAdmin, true);

  const s = env.sheets['Check-in Settings'];
  const row = s.cells.findIndex((r) => r && String(r[0]).startsWith('Admin emails')) + 1;
  assert.equal(s.get(row, 2), 'cara@example.com', 'stored in the settings tab');
  assert.deepEqual(logRows(env).filter((r) => /^Admin (added|removed)$/.test(r[4])).map((r) => [r[3], r[4], r[6]]), [
    ['ana@example.com', 'Admin added', 'nhstedd@gmail.com'],
    ['cara@example.com', 'Admin added', 'ana@example.com'],
    ['ana@example.com', 'Admin removed', 'nhstedd@gmail.com']
  ]);
});

test('adding an admin validates the email', () => {
  const env = fresh();
  fail(env.call('addAdmin', 'tok-tedd', { email: 'not-an-email' }), /not a valid email/);
  fail(env.call('addAdmin', 'tok-tedd', { email: 'a@b.com, c@d.com' }), /not a valid email/);
  fail(env.call('addAdmin', 'tok-tedd', {}), /not a valid email/);
});

test('a parent who signed up is confirmed without changing the sheet', () => {
  const env = fresh({ now: new Date('2026-10-04T13:05:00Z') }); // 10/4 shift 2
  const r = ok(env.call('checkIn', 'tok-pat', { code: codeAt(env, 'Hangar 391') }));
  assert.deepEqual([r.role, r.status, r.alreadyMarked], ['parent', 'Checked in', true]);
  assert.match(r.shift, /Shift 2/);
  assert.equal(env.att.get(13, COL.oct4s2), 'Pat', 'their own sign-up text is left as is');
  assert.equal(env.att.get(14, COL.oct4s2), '');
  assert.deepEqual(logRows(env).at(-1).slice(3, 5), ['Pat Parker', 'Parent checked in']);
  assert.deepEqual(ok(env.call('me', 'tok-pat')).member, { name: 'Pat Parker', role: 'parent' });
});

test('a parent who did not sign up is added to the first empty Parent slot', () => {
  const env = fresh({ now: new Date('2026-10-04T13:05:00Z') });
  const r = ok(env.call('checkIn', 'tok-quinn', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(r.alreadyMarked, false);
  assert.equal(env.att.get(14, COL.oct4s2), 'Quinn Parker', 'Parent 2 (Parent 1 was taken)');
  assert.equal(env.att.get(12, COL.oct4s2), 'Robin Reyes', 'Female Parent untouched');
  assert.equal(logRows(env).at(-1)[4], 'Parent added');

  // Checking in again finds the name it wrote.
  assert.equal(ok(env.call('checkIn', 'tok-quinn', { code: codeAt(env, 'Hangar 391') })).alreadyMarked, true);
  assert.equal(env.att.get(15, COL.oct4s2), '');
});

test('a parent check-in never touches student rows and is never Partial', () => {
  const env = fresh({ now: new Date('2026-10-04T15:30:00Z') }); // 2.5h into shift 2
  const r = ok(env.call('checkIn', 'tok-quinn', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(r.status, 'Checked in');
  for (const row of [33, 34, 36]) assert.equal(env.att.get(row, COL.oct4s2), 'Not Present');
});

test('a parent alias on the roster matches a sign-up', () => {
  const env = fresh({ now: new Date('2026-10-06T17:50:00Z') });
  env.att.set(15, COL.oct6, 'p. parker');
  const r = ok(env.call('checkIn', 'tok-quinn', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(r.alreadyMarked, true, 'family alias "P. Parker" from the mapping tab');
});

test('full parent slots give a clear error', () => {
  const env = fresh({ now: new Date('2026-10-06T17:50:00Z') });
  [13, 14, 15, 16, 17, 18].forEach((row) => env.att.set(row, COL.oct6, 'Someone ' + row));
  fail(env.call('checkIn', 'tok-quinn', { code: codeAt(env, 'Hangar 391') }), /parent slots .* are full/);
});

test('an unlinked parent can claim their name and then check in', () => {
  const env = fresh({ now: new Date('2026-10-04T13:05:00Z') });
  // Robin Reyes is on the roster (from the mapping tab) with no email yet.
  const t = { 'tok-robin': { email: 'robin@example.com', emailVerified: true } };
  Object.assign(env.tokens, t);
  const me = ok(env.call('me', 'tok-robin'));
  assert.equal(me.member, null);
  assert.equal(me.canRegisterAsParent, true);

  const opts = ok(env.call('parentOptions', 'tok-robin'));
  assert.deepEqual(opts.parents, ['Robin Reyes']);
  assert.deepEqual(opts.students, ['Ana Alvarez', 'Ben Brooks', 'Cara Chen']);

  const r = ok(env.call('registerParent', 'tok-robin', { name: 'robin  reyes', child: 'Cara Chen' }));
  assert.deepEqual(r, { name: 'Robin Reyes', role: 'parent' });
  const row = env.roster.cells.findIndex((x) => x && x[0] === 'Robin Reyes') + 1;
  assert.equal(env.roster.get(row, 2), 'robin@example.com');
  assert.match(logRows(env).at(-1)[6], /robin@example.com, parent of Cara Chen/);

  const c = ok(env.call('checkIn', 'tok-robin', { code: codeAt(env, 'Hangar 391') }));
  assert.equal(c.role, 'parent');
  assert.equal(c.alreadyMarked, true, 'their Female Parent sign-up is recognized');
});

test('a parent whose name is missing can add it', () => {
  const env = fresh();
  env.tokens['tok-new'] = { email: 'newparent@example.com', emailVerified: true };
  ok(env.call('registerParent', 'tok-new', { name: 'Dana Diaz', child: 'Ben Brooks' }));
  const last = env.roster.cells.at(-1);
  assert.deepEqual(last.slice(0, 3), ['Dana Diaz', 'newparent@example.com', 'Parent']);
  assert.match(logRows(env).at(-1)[6], /parent of Ben Brooks \(new name\)/);
  assert.equal(ok(env.call('me', 'tok-new')).member.role, 'parent');
});

test('parent registration guards against abuse', () => {
  const env = fresh();
  env.tokens['tok-new'] = { email: 'newparent@example.com', emailVerified: true };
  fail(env.call('registerParent', 'tok-ana', { name: 'Robin Reyes', child: 'Cara Chen' }), /already on the Check-in Roster/,
    'a student cannot also become a parent');
  fail(env.call('parentOptions', 'tok-ana'), /already on the Check-in Roster/);
  fail(env.call('registerParent', 'tok-new', { name: 'Pat Parker', child: 'Ana Alvarez' }), /already linked/,
    'cannot take over a linked parent');
  fail(env.call('registerParent', 'tok-new', { name: 'Ana Alvarez', child: 'Ben Brooks' }), /is a student/);
  fail(env.call('registerParent', 'tok-new', { name: 'Dana Diaz', child: 'Nobody' }), /Pick your student/);
  fail(env.call('registerParent', 'tok-new', { name: 'Dana', child: 'Ben Brooks' }), /first and last name/);
  ok(env.call('registerParent', 'tok-new', { name: 'Dana Diaz', child: 'Ben Brooks' }));
  fail(env.call('registerParent', 'tok-new', { name: 'Dana Diaz', child: 'Ben Brooks' }), /already on the Check-in Roster/,
    'only once per account');
});
