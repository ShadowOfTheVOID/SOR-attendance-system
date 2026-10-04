const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./fakeSheets');

function fresh(opts) {
  const env = load(opts);
  env.app.setup();
  const roster = env.sheets.Roster;
  roster.getRange(2, 1, 3, 4).setValues([
    ['S001', 'Jane Doe', 'A', true],
    ['S002', 'John Smith', 'B', true],
    ['S003', 'Old Member', 'A', false]
  ]);
  return env;
}

const header = (env) => env.sheets.Attendance.getRange(1, 1, 1, env.sheets.Attendance.getLastColumn()).getValues()[0];
const rowOf = (env, id) => env.sheets.Attendance.cells.findIndex((r) => r && r[0] === id) + 1;
const cell = (env, id, date) => env.sheets.Attendance.get(rowOf(env, id), header(env).indexOf(date) + 1);

test('setup creates all sheets with headers', () => {
  const env = fresh();
  for (const n of ['Roster', 'Attendance', 'Log', 'Settings']) assert.ok(env.sheets[n], n);
  assert.deepEqual(header(env), ['ID', 'Name', 'Group']);
});

test('check-in marks Present before cutoff and Late after', () => {
  const env = fresh();
  env.setTime('2026-10-05T09:00:00Z');
  let r = env.app.checkIn('s001');
  assert.equal(r.status, 'Present');
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present');

  env.setTime('2026-10-05T09:30:00Z');
  r = env.app.checkIn('John Smith');
  assert.equal(r.status, 'Late');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Late');
  assert.equal(env.sheets.Log.getLastRow(), 3);
});

test('repeat check-in does not change the mark or log again', () => {
  const env = fresh();
  env.app.checkIn('S001');
  env.setTime('2026-10-05T11:00:00Z');
  const r = env.app.checkIn('S001');
  assert.equal(r.alreadyMarked, true);
  assert.equal(r.status, 'Present');
  assert.equal(env.sheets.Log.getLastRow(), 2);
});

test('inactive or unknown members are rejected', () => {
  const env = fresh();
  assert.throws(() => env.app.checkIn('S003'), /No active member/);
  assert.throws(() => env.app.checkIn('nobody'), /No active member/);
});

test('self check-in can be disabled in Settings', () => {
  const env = fresh();
  const s = env.sheets.Settings;
  const row = s.cells.findIndex((r) => r && String(r[0]).startsWith('Self check-in')) + 1;
  s.set(row, 2, 'FALSE');
  assert.throws(() => env.app.checkIn('S001'), /closed/);
});

test('admin save writes statuses, inserts dates in order, and logs changes', () => {
  const env = fresh();
  env.app.checkIn('S001'); // 2026-10-05
  env.app.adminSaveDay('1234', '2026-10-07', [{ id: 'S001', status: 'Excused' }]);
  env.app.adminSaveDay('1234', '2026-10-06', [{ id: 'S002', status: 'Absent' }]);
  assert.deepEqual(header(env).slice(3), ['2026-10-05', '2026-10-06', '2026-10-07']);
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present', 'existing column survived insert');
  assert.equal(cell(env, 'S001', '2026-10-07'), 'Excused');
  assert.equal(cell(env, 'S002', '2026-10-06'), 'Absent');

  const day = env.app.adminGetDay('1234', '2026-10-05');
  assert.deepEqual(day.members.map((m) => [m.id, m.status]), [['S001', 'Present'], ['S002', '']]);
  assert.equal(day.summary.Unmarked, 1);

  const r = env.app.adminSaveDay('1234', '2026-10-05', [{ id: 'S001', status: 'Present' }]);
  assert.equal(r.saved, 0, 'unchanged values are not logged');
});

test('admin endpoints require the PIN and valid input', () => {
  const env = fresh();
  assert.throws(() => env.app.adminGetDay('nope', '2026-10-05'), /Wrong PIN/);
  assert.throws(() => env.app.adminGetDay('1234', '2026-13-01'), /Invalid date/);
  assert.throws(() => env.app.adminSaveDay('1234', '2026-10-05', [{ id: 'S001', status: 'Here' }]), /Invalid status/);
});

test('repeated wrong PINs lock admin out', () => {
  const env = fresh();
  for (let i = 0; i < 10; i++) assert.throws(() => env.app.adminGetDay('x', '2026-10-05'));
  assert.throws(() => env.app.adminGetDay('1234', '2026-10-05'), /Too many/);
});

test('markAbsentToday fills only blank cells', () => {
  const env = fresh();
  env.app.checkIn('S001');
  assert.equal(env.app.markAbsentToday(), 1);
  assert.equal(cell(env, 'S001', '2026-10-05'), 'Present');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Absent');
  assert.equal(env.app.markAbsentToday(), 0);
});

test('a late check-in replaces an automatic Absent', () => {
  const env = fresh();
  env.app.markAbsentToday();
  env.setTime('2026-10-05T10:00:00Z');
  assert.equal(env.app.checkIn('S002').status, 'Late');
  assert.equal(cell(env, 'S002', '2026-10-05'), 'Late');
});

test('many date columns grow the sheet past its initial width', () => {
  const env = fresh();
  for (let d = 1; d <= 30; d++) {
    env.app.adminSaveDay('1234', `2026-11-${String(d).padStart(2, '0')}`, [{ id: 'S001', status: 'Present' }]);
  }
  assert.equal(header(env).length, 33);
  assert.equal(cell(env, 'S001', '2026-11-30'), 'Present');
});
