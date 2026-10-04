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

test('findMember by id or unique name, case-insensitive', () => {
  const roster = [
    { id: 'S001', name: 'Jane Doe' },
    { id: 'S002', name: 'John  Smith' },
    { id: 'S003', name: 'Sam Lee' },
    { id: 'S004', name: 'sam lee' }
  ];
  assert.equal(L.findMember(roster, ' s001 ').id, 'S001');
  assert.equal(L.findMember(roster, 'JANE DOE').id, 'S001');
  assert.equal(L.findMember(roster, 'john smith').id, 'S002');
  assert.equal(L.findMember(roster, 'Sam Lee'), null, 'ambiguous names do not match');
  assert.equal(L.findMember(roster, 'Jane'), null, 'partial names do not match');
  assert.equal(L.findMember(roster, ''), null);
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
