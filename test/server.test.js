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
  'tok-owner': { email: 'owner@example.com', emailVerified: true }
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
  return sh;
}

function fresh(opts = {}) {
  const env = load({ tokens: TOKENS, now: new Date('2026-10-04T13:10:00Z'), ...opts });
  env.att = buildAttendanceTab(env);
  env.app.setup();
  const roster = env.sheets['Check-in Roster'];
  const emails = { 'Ana Alvarez': 'ana@example.com', 'Ben Brooks': 'ben@example.com', 'Cara Chen': 'cara@example.com' };
  for (let r = 2; r <= roster.getLastRow(); r++) roster.set(r, 2, emails[roster.get(r, 1)] || '');
  roster.getRange(roster.getLastRow() + 1, 1, 1, 2).setValues([['Ghost Person', 'ghost@example.com']]);
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
    ['Check-in Log', 'Check-in Roster', 'Check-in Settings', 'Offseason 2026']);
  const names = env.sheets['Check-in Roster'].cells.slice(1).map((r) => r[0]);
  assert.deepEqual(names, ['Ana Alvarez', 'Ben Brooks', 'Cara Chen', 'Ghost Person']);
  assert.ok(env.props.CODE_SECRET.length > 40);
  assert.equal(ok(env.call('me', 'tok-owner')).isAdmin, true);

  env.app.fillRosterNames(); // re-running adds no duplicates
  assert.equal(env.sheets['Check-in Roster'].getLastRow(), 5);
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
  assert.deepEqual(ok(env.call('me', 'tok-ana')).member, { name: 'Ana Alvarez' });
});

test('missing Firebase API key is reported', () => {
  const env = fresh({ apiKey: null });
  fail(env.call('me', 'tok-ana'), /Firebase API key/);
});

test('admin-only actions reject students', () => {
  const env = fresh();
  fail(env.call('shifts', 'tok-ana', {}), /Admins only/);
  fail(env.call('displayCode', 'tok-ana', { location: 'Hangar 391' }), /Admins only/);
  fail(env.call('getShift', 'tok-ana', { date: '2026-10-04', shiftKey: 'x' }), /Admins only/);
  fail(env.call('saveShift', 'tok-ana', { date: '2026-10-04', shiftKey: 'x', records: [] }), /Admins only/);
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
