const test = require('node:test');
const assert = require('node:assert/strict');
const L = require('../src/Logic.js');

test('parseTime', () => {
  assert.equal(L.parseTime('09:15'), 555);
  assert.equal(L.parseTime(' 0:00 '), 0);
  assert.equal(L.parseTime('23:59'), 1439);
  assert.equal(L.parseTime('24:00'), null);
  assert.equal(L.parseTime(''), null);
  assert.equal(L.parseTime(null), null);
  assert.equal(L.parseTime('9am'), null);
});

test('statusForCheckIn', () => {
  assert.equal(L.statusForCheckIn(555, 555), 'Present');
  assert.equal(L.statusForCheckIn(556, 555), 'Late');
  assert.equal(L.statusForCheckIn(1000, null), 'Present');
});

test('findMemberByEmail is case-insensitive and exact', () => {
  const roster = [
    { id: 'S001', email: 'Jane@Example.com' },
    { id: 'S002', email: '' }
  ];
  assert.equal(L.findMemberByEmail(roster, ' jane@example.com ').id, 'S001');
  assert.equal(L.findMemberByEmail(roster, 'jane@example'), null);
  assert.equal(L.findMemberByEmail(roster, ''), null, 'blank email never matches blank roster cell');
});

test('truncateToCode matches the RFC 4226 example', () => {
  const hex = '1f8698690e02ca16618550ef7f19da8e945b555a';
  const bytes = hex.match(/../g).map((h) => parseInt(h, 16));
  assert.equal(L.truncateToCode(bytes, 6), '872921');
  const signed = bytes.map((x) => (x > 127 ? x - 256 : x));
  assert.equal(L.truncateToCode(signed, 6), '872921');
});

test('timeStep, normalizeCode, parseEmailList', () => {
  assert.equal(L.timeStep(59999, 30), 1);
  assert.equal(L.timeStep(60000, 30), 2);
  assert.equal(L.normalizeCode(' 123 456 '), '123456');
  assert.deepEqual(L.parseEmailList('A@x.com, b@y.com;\nc@z.com'), ['a@x.com', 'b@y.com', 'c@z.com']);
  assert.deepEqual(L.parseEmailList(''), []);
});

test('mergeCheckIn never overrides an existing mark except Absent', () => {
  assert.deepEqual(L.mergeCheckIn('', 'Present'), { status: 'Present', changed: true });
  assert.deepEqual(L.mergeCheckIn('Absent', 'Late'), { status: 'Late', changed: true });
  assert.deepEqual(L.mergeCheckIn('Present', 'Late'), { status: 'Present', changed: false });
  assert.deepEqual(L.mergeCheckIn('Excused', 'Present'), { status: 'Excused', changed: false });
});

test('insertionIndexForDate keeps date columns sorted', () => {
  const h = ['ID', 'Name', 'Group', '2026-10-01', '2026-10-03'];
  assert.equal(L.insertionIndexForDate(h, '2026-10-02', 3), 4);
  assert.equal(L.insertionIndexForDate(h, '2026-09-30', 3), 3);
  assert.equal(L.insertionIndexForDate(h, '2026-10-04', 3), 5);
  assert.equal(L.insertionIndexForDate(['ID', 'Name', 'Group'], '2026-10-04', 3), 3);
});

test('date and status validation', () => {
  assert.ok(L.isValidDateKey('2026-02-28'));
  assert.ok(!L.isValidDateKey('2026-02-30'));
  assert.ok(!L.isValidDateKey('2026-2-3'));
  assert.ok(L.isValidStatus('Late'));
  assert.ok(L.isValidStatus(''));
  assert.ok(!L.isValidStatus('present'));
});

test('summarize', () => {
  assert.deepEqual(L.summarize(['Present', 'Late', '', 'Absent', 'Present', 'Excused']),
    { Present: 2, Late: 1, Absent: 1, Excused: 1, Unmarked: 1 });
});
