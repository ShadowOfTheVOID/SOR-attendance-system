const test = require('node:test');
const assert = require('node:assert/strict');
const L = require('../src/Logic.js');

test('parseTimeLoose handles the formats found in the sheet', () => {
  assert.equal(L.parseTimeLoose('18:15'), 18 * 60 + 15);
  assert.equal(L.parseTimeLoose('18:15:00'), 18 * 60 + 15);
  assert.equal(L.parseTimeLoose('6:15 PM'), 18 * 60 + 15);
  assert.equal(L.parseTimeLoose('12:45 PM'), 12 * 60 + 45);
  assert.equal(L.parseTimeLoose('12:10 am'), 10);
  assert.equal(L.parseTimeLoose('3:45:00 PM --> 4:45'), 16 * 60 + 45, 'edited time: last wins, PM carries');
  assert.equal(L.parseTimeLoose('7:15:00 PM -->8:15'), 20 * 60 + 15);
  assert.equal(L.parseTimeLoose(''), null);
  assert.equal(L.parseTimeLoose('TBD'), null);
  assert.equal(L.parseTimeLoose('25:00'), null);
});

test('shiftLabel and formatMinutes', () => {
  assert.equal(L.formatMinutes(0), '12:00 AM');
  assert.equal(L.formatMinutes(13 * 60 + 5), '1:05 PM');
  assert.equal(L.shiftLabel({ date: '2026-10-04', shiftNumber: '2', location: 'Hangar 391', start: 780, end: 975 }),
    'Sun 10/4 · Shift 2 · Hangar 391 · 1:00 PM–4:15 PM');
});

test('pickDefaultShift prefers running shift with latest start, then next, then last', () => {
  const shifts = [
    { start: 9 * 60 + 45, end: 13 * 60 + 15 },
    { start: 13 * 60, end: 16 * 60 + 15 },
    { start: 15 * 60, end: 16 * 60 }
  ];
  assert.equal(L.pickDefaultShift(shifts, 10 * 60), 0);
  assert.equal(L.pickDefaultShift(shifts, 13 * 60 + 5), 1, 'overlap: later start wins');
  assert.equal(L.pickDefaultShift(shifts, 12 * 60 + 45), 1, 'starts within 30 min');
  assert.equal(L.pickDefaultShift(shifts, 8 * 60), 0, 'next upcoming');
  assert.equal(L.pickDefaultShift(shifts, 22 * 60), 2, 'after all: last');
  assert.equal(L.pickDefaultShift([], 600), -1);
});

test('pickShiftAt: running shift, overlap goes to the next shift, early window', () => {
  const s = [
    { start: 9 * 60 + 45, end: 13 * 60 + 15 }, // 9:45–1:15
    { start: 13 * 60, end: 16 * 60 + 15 }      // 1:00–4:15
  ];
  assert.equal(L.pickShiftAt(s, 10 * 60, 30), 0);
  assert.equal(L.pickShiftAt(s, 12 * 60 + 59, 30), 0, 'not stolen by next shift\'s early window');
  assert.equal(L.pickShiftAt(s, 13 * 60, 30), 1, 'overlap starts: next shift');
  assert.equal(L.pickShiftAt(s, 13 * 60 + 15, 30), 1, 'still overlap: next shift');
  assert.equal(L.pickShiftAt(s, 9 * 60 + 15, 30), 0, '30 min early');
  assert.equal(L.pickShiftAt(s, 9 * 60 + 14, 30), -1, 'too early');
  assert.equal(L.pickShiftAt(s, 16 * 60 + 16, 30), -1, 'after the last shift');
  assert.equal(L.pickShiftAt([{ start: null, end: null }], 600, 30), 0, 'no times = all day');
  // 3 chained shifts with 15-min overlaps (like 10/3)
  const chain = [{ start: 510, end: 630 }, { start: 615, end: 735 }, { start: 720, end: 840 }];
  assert.deepEqual([600, 620, 700, 725, 800].map((t) => L.pickShiftAt(chain, t, 30)), [0, 1, 1, 2, 2]);
});

test('findLabelRow matches the start of column A labels', () => {
  const colA = ['x5', 'Date', 'Shift Number', 'Start Time\nShould start 15 min. prior', 'Avg Attendees'];
  assert.equal(L.findLabelRow(colA, 'date'), 1);
  assert.equal(L.findLabelRow(colA, 'start time'), 3);
  assert.equal(L.findLabelRow(colA, 'end time'), -1);
});

test('findMemberByEmail and sameName', () => {
  const roster = [{ name: 'Jane Doe', email: 'Jane@Example.com' }, { name: 'X', email: '' }];
  assert.equal(L.findMemberByEmail(roster, ' jane@example.com ').name, 'Jane Doe');
  assert.equal(L.findMemberByEmail(roster, ''), null);
  assert.ok(L.sameName('Jane  Doe ', 'jane doe'));
  assert.ok(!L.sameName('', ''));
});

test('statusForCheckIn and parseMinutes', () => {
  assert.equal(L.statusForCheckIn(13 * 60 + 30, 13 * 60, 30), 'Present', 'exactly at the limit');
  assert.equal(L.statusForCheckIn(13 * 60 + 31, 13 * 60, 30), 'Partial');
  assert.equal(L.statusForCheckIn(12 * 60 + 40, 13 * 60, 30), 'Present', 'early');
  assert.equal(L.statusForCheckIn(23 * 60, 13 * 60, null), 'Present', 'disabled');
  assert.equal(L.statusForCheckIn(23 * 60, null, 30), 'Present', 'no start time');
  assert.equal(L.parseMinutes(' 30 '), 30);
  assert.equal(L.parseMinutes(0), 0);
  assert.equal(L.parseMinutes(''), null);
  assert.equal(L.parseMinutes('soon'), null);
});

test('mergeCheckIn only replaces absent/blank cells', () => {
  assert.deepEqual(L.mergeCheckIn('Not Present', 'Partial'), { status: 'Partial', changed: true });
  for (const s of ['', 'Not Present', 'Absent Excused', 'Absent Unexcused', null]) {
    assert.deepEqual(L.mergeCheckIn(s), { status: 'Present', changed: true }, String(s));
  }
  for (const s of ['Present', 'Partial', 'Unproductive']) {
    assert.deepEqual(L.mergeCheckIn(s), { status: s, changed: false }, s);
  }
});

test('truncateToCode matches the RFC 4226 example', () => {
  const bytes = '1f8698690e02ca16618550ef7f19da8e945b555a'.match(/../g).map((h) => parseInt(h, 16));
  assert.equal(L.truncateToCode(bytes, 6), '872921');
  assert.equal(L.truncateToCode(bytes.map((x) => (x > 127 ? x - 256 : x)), 6), '872921');
});

test('timeStep, normalizeCode, parseEmailList', () => {
  assert.equal(L.timeStep(59999, 30), 1);
  assert.equal(L.timeStep(60000, 30), 2);
  assert.equal(L.normalizeCode(' 123 456 '), '123456');
  assert.deepEqual(L.parseEmailList('A@x.com, b@y.com;\nc@z.com'), ['a@x.com', 'b@y.com', 'c@z.com']);
});

test('validation and summarize', () => {
  assert.ok(L.isValidDateKey('2026-02-28'));
  assert.ok(!L.isValidDateKey('2026-02-30'));
  assert.ok(L.isValidStatus('Absent Excused'));
  assert.ok(!L.isValidStatus('Late'));
  const c = L.summarize(['Present', 'Present', '', 'Partial']);
  assert.equal(c.Present, 2);
  assert.equal(c['Not Present'], 1);
  assert.equal(c.Partial, 1);
});

test('roleOf and parentSlotMatches', () => {
  assert.equal(L.roleOf('Parent'), 'parent');
  assert.equal(L.roleOf(' parent volunteer'), 'parent');
  assert.equal(L.roleOf(''), 'student');
  assert.equal(L.roleOf('Student'), 'student');

  const simi = { name: 'Simi Raj', aliases: ['S. Raj'] };
  const sam = { name: 'Samir Vapiwala', aliases: [] };
  const sam2 = { name: 'Samir Gupta', aliases: [] };
  const all = [simi, sam, sam2];
  assert.ok(L.parentSlotMatches('Simi Raj ', simi, all));
  assert.ok(L.parentSlotMatches('simi', simi, all), 'unique first name');
  assert.ok(L.parentSlotMatches('S. Raj', simi, all), 'alias');
  assert.ok(!L.parentSlotMatches('Samir', sam, all), 'ambiguous first name');
  assert.ok(L.parentSlotMatches('samir vapiwala', sam, all));
  assert.ok(!L.parentSlotMatches('', simi, all));
  assert.ok(!L.parentSlotMatches('Simi Rajan', simi, all));
});

test('maskEmail', () => {
  assert.equal(L.maskEmail('jane.doe@gmail.com'), 'ja***@gmail.com');
  assert.equal(L.maskEmail('j@x.org'), 'j***@x.org');
  assert.equal(L.maskEmail('nope'), '***');
});

test('parseDirectory reads students and up to three guardians from the form sheet', () => {
  const header = ['Timestamp', 'Student Name (First Name)', 'Student Name (Last Name)', 'Student Email',
    'Parent/Guardian 1 (First Name)', 'Parent/Guardian 1 (Last Name)', 'Relationship to Youth', 'Email', 'Phone number',
    'Parent/Guardian 2 (First Name)', 'Parent/Guardian 2 (Last Name)', 'Relationship to Youth 2', 'Email 2',
    'Parent/Guardian 3 (First Name)', 'Parent/Guardian 3 (Last Name)', 'Email 3'];
  const rows = [
    ['t', 'Ana ', ' Alvarez', 'Ana@Gmail.com ', 'Pat', 'Parker', 'Father', 'PAT@x.com', '1', 'Quinn', 'Parker', 'Mother', '', '', '', ''],
    ['t', 'Ben', 'Brooks', '', 'Sam', 'Brooks', 'Mom', 'sam@x.com', '2', '', '', '', '', 'Lee', 'Brooks', 'lee@x.com'],
    ['t', '', '', 'orphan@x.com', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  assert.deepEqual(L.parseDirectory(header, rows), {
    students: [{ name: 'Ana Alvarez', email: 'ana@gmail.com' }],
    parents: [
      { name: 'Pat Parker', email: 'pat@x.com', child: 'Ana Alvarez' },
      { name: 'Sam Brooks', email: 'sam@x.com', child: 'Ben Brooks' },
      { name: 'Lee Brooks', email: 'lee@x.com', child: 'Ben Brooks' }
    ]
  });
});
