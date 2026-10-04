const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./fakeSheets');

const TOKENS = {
  'tok-jane': { email: 'Jane@Example.com', emailVerified: true },
  'tok-john': { email: 'john@example.com', emailVerified: true },
  'tok-old': { email: 'old@example.com', emailVerified: true },
  'tok-stranger': { email: 'stranger@example.com', emailVerified: true },
  'tok-unverified': { email: 'jane@example.com', emailVerified: false },
  'tok-owner': { email: 'owner@example.com', emailVerified: true }
};

function fresh(opts = {}) {
  const env = load({ tokens: TOKENS, ...opts });
  env.app.setup();
  env.sheets.Roster.getRange(2, 1, 3, 5).setValues([
    ['S001', 'Jane Doe', 'A', true, 'jane@example.com'],
    ['S002', 'John Smith', 'B', true, 'JOHN@example.com'],
    ['S003', 'Old Member', 'A', false, 'old@example.com']
  ]);
  return env;
}

const ok = (res) => { assert.equal(res.ok, true, res.error); return res.data; };
const fail = (res, re) => { assert.equal(res.ok, false); assert.match(res.error, re); };
const displayCode = (env) => ok(env.call('displayCode', 'tok-owner')).code;
const checkIn = (env, tok, code = displayCode(env)) => env.call('checkIn', tok, { code });

const header = (env) => env.sheets.Attendance.getRange(1, 1, 1, env.sheets.Attendance.getLastColumn()).getValues()[0];
const rowOf = (env, id) => env.sheets.Attendance.cells.findIndex((r) => r && r[0] === id) + 1;
const cell = (env, id, date) => env.sheets.Attendance.get(rowOf(env, id), header(env).indexOf(date) + 1);
const logStatuses = (env) => env.sheets.Log.cells.slice(1).map((r) => r[4]);

test('setup creates sheets, makes the owner admin, and creates a code secret', () => {
  const env = fresh();
  for (const n of ['Roster', 'Attendance', 'Log', 'Settings']) assert.ok(env.sheets[n], n);
  assert.deepEqual(header(env), ['ID', 'Name', 'Group']);
  assert.ok(env.props.CODE_SECRET.length > 40);
  assert.equal(ok(env.call('me', 'tok-owner')).isAdmin, true);
});

test('me identifies the roster member from their Google email', () => {
  const env = fresh();
  const me = ok(env.call('me', 'tok-jane'));
  assert.equal(me.email, 'jane@example.com');
  assert.equal(me.member.id, 'S001');
  assert.equal(me.isAdmin, false);
  assert.equal(ok(env.call('me', 'tok-stranger')).member, null);
});

test('requests without a valid verified token are rejected', () => {
  const env = fresh();
  fail(env.call('me', undefined), /sign in/i);
  fail(env.call('me', 'forged'), /expired/i);
  fail(env.call('me', 'tok-unverified'), /verified/i);
  fail(env.call('nope', 'tok-jane'), /Unknown action/);
});

test('tokens are rejected when the Firebase API key is not configured', () => {
  const env = fresh({ apiKey: null });
  fail(env.call('me', 'tok-jane'), /Firebase API key/);
});

test('verified tokens are cached instead of re-checked every request', () => {
  const env = fresh();
  ok(env.call('me', 'tok-jane'));
  ok(env.call('me', 'tok-jane'));
  assert.equal(env.fetches.length, 1);
});

test('check-in with the current code marks Present, then Late after cutoff', () => {
  const env = fresh();
  env.setTime('2026-10-05T09:00:00Z');
  const r = ok(checkIn(env, 'tok-jane'));
  assert.equal(r.status, 'Present');
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present');

  env.setTime('2026-10-05T09:30:00Z');
  assert.equal(ok(checkIn(env, 'tok-john')).status, 'Late');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Late');
});

test('codes rotate: previous window accepted, older codes rejected', () => {
  const env = fresh();
  env.setTime('2026-10-05T09:00:05Z');
  const code = displayCode(env);
  env.setTime('2026-10-05T09:00:35Z');
  assert.notEqual(displayCode(env), code, 'code changed after 30s');
  ok(checkIn(env, 'tok-jane', code)); // one window late is still fine

  env.setTime('2026-10-05T09:01:05Z');
  fail(checkIn(env, 'tok-john', code), /wrong or has expired/);
  assert.equal(cell(env, 'S002', '2026-10-05'), '');
  assert.ok(logStatuses(env).includes('Rejected'));
});

test('wrong codes are rejected and rate-limited per account', () => {
  const env = fresh();
  fail(checkIn(env, 'tok-jane', ''), /wrong/);
  for (let i = 0; i < 4; i++) fail(checkIn(env, 'tok-jane', '000000'), /wrong/);
  fail(checkIn(env, 'tok-jane'), /Too many/);
  ok(checkIn(env, 'tok-john')); // other accounts unaffected
});

test('only active roster emails can check in', () => {
  const env = fresh();
  fail(checkIn(env, 'tok-stranger'), /not on the roster/);
  fail(checkIn(env, 'tok-old'), /not on the roster/);
});

test('repeat check-in does not change the mark or log again', () => {
  const env = fresh();
  ok(checkIn(env, 'tok-jane'));
  env.setTime('2026-10-05T11:00:00Z');
  const r = ok(checkIn(env, 'tok-jane'));
  assert.equal(r.alreadyMarked, true);
  assert.equal(r.status, 'Present');
  assert.deepEqual(logStatuses(env), ['Present']);
});

test('self check-in can be disabled in Settings', () => {
  const env = fresh();
  const s = env.sheets.Settings;
  const row = s.cells.findIndex((r) => r && String(r[0]).startsWith('Self check-in')) + 1;
  s.set(row, 2, 'FALSE');
  fail(checkIn(env, 'tok-jane'), /closed/);
});

test('admin-only actions reject non-admins', () => {
  const env = fresh();
  fail(env.call('displayCode', 'tok-jane'), /Admins only/);
  fail(env.call('getDay', 'tok-jane', { date: '2026-10-05' }), /Admins only/);
  fail(env.call('saveDay', 'tok-jane', { date: '2026-10-05', records: [] }), /Admins only/);
});

test('admin save writes statuses, inserts dates in order, and logs changes', () => {
  const env = fresh();
  ok(checkIn(env, 'tok-jane')); // 2026-10-05
  ok(env.call('saveDay', 'tok-owner', { date: '2026-10-07', records: [{ id: 'S001', status: 'Excused' }] }));
  ok(env.call('saveDay', 'tok-owner', { date: '2026-10-06', records: [{ id: 'S002', status: 'Absent' }] }));
  assert.deepEqual(header(env).slice(3), ['2026-10-05', '2026-10-06', '2026-10-07']);
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present', 'existing column survived insert');
  assert.equal(cell(env, 'S001', '2026-10-07'), 'Excused');
  assert.equal(cell(env, 'S002', '2026-10-06'), 'Absent');

  const day = ok(env.call('getDay', 'tok-owner', { date: '2026-10-05' }));
  assert.deepEqual(day.members.map((m) => [m.id, m.status]), [['S001', 'Present'], ['S002', '']]);
  assert.equal(day.summary.Unmarked, 1);

  const r = ok(env.call('saveDay', 'tok-owner', { date: '2026-10-05', records: [{ id: 'S001', status: 'Present' }] }));
  assert.equal(r.saved, 0, 'unchanged values are not logged');
});

test('admin input is validated', () => {
  const env = fresh();
  fail(env.call('getDay', 'tok-owner', { date: '2026-13-01' }), /Invalid date/);
  fail(env.call('saveDay', 'tok-owner', { date: '2026-10-05', records: [{ id: 'S001', status: 'Here' }] }), /Invalid status/);
});

test('markAbsentToday fills only blank cells', () => {
  const env = fresh();
  ok(checkIn(env, 'tok-jane'));
  assert.equal(env.app.markAbsentToday(), 1);
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Absent');
  assert.equal(env.app.markAbsentToday(), 0);
});

test('a late check-in replaces an automatic Absent', () => {
  const env = fresh();
  env.app.markAbsentToday();
  env.setTime('2026-10-05T10:00:00Z');
  assert.equal(ok(checkIn(env, 'tok-john')).status, 'Late');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Late');
});

test('many date columns grow the sheet past its initial width', () => {
  const env = fresh();
  for (let d = 1; d <= 30; d++) {
    ok(env.call('saveDay', 'tok-owner', {
      date: `2026-11-${String(d).padStart(2, '0')}`, records: [{ id: 'S001', status: 'Present' }]
    }));
  }
  assert.equal(header(env).length, 33);
  assert.equal(cell(env, 'S001', '2026-11-30'), 'Present');
});
