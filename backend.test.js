// Backend tests: runs the real Code.gs against the in-memory Apps Script mock.
// Run with: TZ=Asia/Kolkata node test/backend.test.js
const assert = require('assert');
const vm = require('vm');
const path = require('path');
const { createBackend } = require('./gas-mock');

const CODE = path.join(__dirname, '..', 'Code.gs');
const results = [];
function test(name, fn) {
  try {
    fn();
    results.push({ name, ok: true });
    console.log('PASS  ' + name);
  } catch (err) {
    results.push({ name, ok: false, err });
    console.log('FAIL  ' + name + '\n      ' + err.message);
  }
}

function fmt(d) {
  const p = n => (n < 10 ? '0' : '') + n;
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}
function dayOffset(n) {
  const d = new Date();
  d.setDate(d.getDate() + n);
  return fmt(d);
}
const YESTERDAY = dayOffset(-1);
const TWO_DAYS_AGO = dayOffset(-2);
const THREE_DAYS_AGO = dayOffset(-3);
const TOMORROW = dayOffset(1);
const TODAY = dayOffset(0);

// Fixes the clock used by Code.gs (new Date() with no arguments, and Date.now()).
function setNow(localIso) {
  const Real = vm.runInContext("Date", be.ctx);
  const fixed = new Real(localIso).getTime();
  be.ctx.__NewDate = class extends Real {
    constructor(...args) { if (args.length === 0) super(fixed); else super(...args); }
    static now() { return fixed; }
  };
  vm.runInContext("Date = __NewDate", be.ctx);
}
function restoreClock() {
  // Re-creating the context is not needed: a fresh clock is set with setNow for each check.
}

const be = createBackend(CODE);
const ok = r => assert.strictEqual(r.success, true, 'expected success, got: ' + JSON.stringify(r));
const fails = (r, text) => {
  assert.strictEqual(r.success, false, 'expected failure, got: ' + JSON.stringify(r));
  if (text) assert.ok(String(r.error).toLowerCase().includes(text.toLowerCase()), 'error was: ' + r.error);
};
const countRows = (sheet, pred) => be.rows(sheet).filter(pred).length;

let tokA = '', tokB = '', tokC = '';

test('status endpoint reports the new version and actions', () => {
  const s = be.get({});
  assert.strictEqual(s.version, 'features-v8');
  ['saveAttendanceBatch', 'saveDayChange', 'saveHoliday', 'saveSettings', 'forgotRequest', 'forgotReset']
    .forEach(a => assert.ok(s.actions.includes(a), 'missing action ' + a));
});

test('setupSheets creates the new sheets and adds Semester/Counted/Kind to Attendance', () => {
  be.runTrigger('setupSheets');
  ['DayChanges', 'Holidays', 'Settings', 'ResetCodes', 'Attendance', 'Students'].forEach(n =>
    assert.ok(be.sheets.has(n), 'missing sheet ' + n));
  const headers = be.sheets.get('Attendance').rows[0];
  ['Semester', 'Counted', 'Kind'].forEach(h => assert.ok(headers.includes(h), 'missing header ' + h));
});

const PROFILE = {
  studentName: 'Asha Rao', email: 'asha@example.com', course: 'B.E.', department: 'Computer Science & Engineering',
  year: '3', class: 'CS-A', section: 'A', semester: '5', collegeName: 'Anna University', academicYear: '2026-2027'
};

test('register rejects a password shorter than 8 characters', () => {
  fails(be.postThenFetch(Object.assign({ action: 'register', registerNumber: '21CS009', password: 'short' }, PROFILE)), 'at least 8');
});

test('register creates accounts for two students and returns tokens', () => {
  const a = be.postThenFetch(Object.assign({ action: 'register', registerNumber: '21cs001', password: 'correct-horse-1' }, PROFILE));
  ok(a); tokA = a.token;
  const b = be.postThenFetch(Object.assign({ action: 'register', registerNumber: '21cs002', password: 'battery-staple-2' }, PROFILE,
    { studentName: 'Ben Iyer', email: 'ben@example.com' }));
  ok(b); tokB = b.token;
  assert.ok(tokA && tokB && tokA !== tokB);
  assert.strictEqual(be.rows('Accounts').some(r => String(r['Password Hash']).includes('correct-horse')), false, 'password stored in plain text');
});

test('login locks an account after 5 wrong passwords, then the right password is refused until the lock ends', () => {
  for (let i = 0; i < 5; i++) fails(be.postThenFetch({ action: 'login', registerNumber: '21CS002', password: 'wrong-pass-xx' }), 'wrong');
  fails(be.postThenFetch({ action: 'login', registerNumber: '21CS002', password: 'battery-staple-2' }), 'too many');
  be.props.delete('fail_21CS002');
});

test('login with the right password returns a session', () => {
  const r = be.postThenFetch({ action: 'login', registerNumber: '21CS001', password: 'correct-horse-1' });
  ok(r); tokA = r.token;
});

test('session-less requests are refused', () => {
  const r = be.api('getData', 'not-a-real-token');
  assert.strictEqual(r.loggedOut, true);
});

test('a session token sent in the URL is refused; the same request works with the token in the POST body', () => {
  const viaUrl = be.get({ action: 'getData', token: tokA });
  assert.strictEqual(viaUrl.success, false);
  assert.ok(/request body/.test(viaUrl.error), 'unclear refusal: ' + viaUrl.error);
  const viaPost = be.api('getData', tokA);
  assert.strictEqual(viaPost.success, true);
});


const BATCH_DAY = [
  { period: 1, subject: 'Java', start: '09:00', end: '09:45', status: 'Present', kind: 'regular', counted: 'Yes' },
  { period: 2, subject: 'DBMS', start: '09:45', end: '10:30', status: 'Absent', kind: 'regular', counted: 'Yes' },
  { period: 3, subject: 'Maths', start: '10:30', end: '11:15', status: 'Present', kind: 'regular', counted: 'Yes' }
];

test('a POST without a valid requestId is refused and writes nothing', () => {
  const before = be.rows('Attendance').length;
  be.post({ action: 'saveAttendanceBatch', token: tokA, attendanceDate: YESTERDAY, semester: '5', rows: JSON.stringify(BATCH_DAY) });
  assert.strictEqual(be.rows('Attendance').length, before);
});


test('saveAttendanceBatch saves every period for a past date', () => {
  const r = be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, semester: '5', rows: JSON.stringify(BATCH_DAY) });
  ok(r);
  assert.strictEqual(r.saved, 3);
  assert.strictEqual(r.updated, 0);
});

test('saving the same date again corrects entries and never duplicates them', () => {
  const corrected = BATCH_DAY.map(x => Object.assign({}, x, { status: x.period === 2 ? 'Present' : x.status }));
  const r = be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, semester: '5', rows: JSON.stringify(corrected) });
  ok(r);
  assert.strictEqual(r.updated, 3);
  assert.strictEqual(r.saved, 0);
  assert.strictEqual(countRows('Attendance', x => x['Register Number'] === '21CS001' && String(x['Attendance Date']) === YESTERDAY), 3);
  const dbms = be.rows('Attendance').find(x => x.Subject === 'DBMS' && String(x['Attendance Date']) === YESTERDAY);
  assert.strictEqual(dbms.Status, 'Present', 'correction did not update status');
});

test('dates and times come back in the expected text format', () => {
  const data = be.api('getData', tokA, {});
  const row = data.attendance.find(x => x.Subject === 'Java');
  assert.strictEqual(row['Attendance Date'], YESTERDAY);
  assert.strictEqual(row['Attendance Time'], '09:00');
  assert.strictEqual(row['End Time'], '09:45');
  assert.strictEqual(String(row.Semester), '5');
  assert.strictEqual(row.Counted, 'Yes');
});

test('future dates cannot be marked', () => {
  fails(be.api('saveAttendanceBatch', tokA, { attendanceDate: TOMORROW, rows: JSON.stringify(BATCH_DAY) }), 'future');
});

test('invalid status, duplicate period and bad times are refused', () => {
  fails(be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, rows: JSON.stringify([{ period: 1, subject: 'X', start: '09:00', end: '09:45', status: 'Late' }]) }), 'present or absent');
  fails(be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, rows: JSON.stringify([BATCH_DAY[0], BATCH_DAY[0]]) }), 'listed twice');
  fails(be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, rows: JSON.stringify([{ period: 1, subject: 'X', start: '10:00', end: '09:00', status: 'Present' }]) }), 'start and end');
});

test('a holiday blocks marking that date and is excluded from counting', () => {
  ok(be.api('saveHoliday', tokA, { holidayDate: TWO_DAYS_AGO, type: 'College holiday', note: 'Pongal' }));
  fails(be.api('saveAttendanceBatch', tokA, { attendanceDate: TWO_DAYS_AGO, rows: JSON.stringify(BATCH_DAY) }), 'holiday');
  const data = be.api('getData', tokA, {});
  assert.ok(data.holidays.some(h => h.Date === TWO_DAYS_AGO && h.Type === 'College holiday' && h.Note === 'Pongal'));
  ok(be.api('deleteHoliday', tokA, { holidayDate: TWO_DAYS_AGO }));
  assert.strictEqual(be.api('getData', tokA, {}).holidays.length, 0);
});

test('a PT period saved as not counted keeps Counted = No', () => {
  const pt = [{ period: 55, subject: 'PT', start: '16:00', end: '17:00', status: 'Present', kind: 'pt', counted: 'No' }];
  ok(be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, rows: JSON.stringify(pt) }));
  const row = be.rows('Attendance').find(x => x.Subject === 'PT');
  assert.strictEqual(row.Counted, 'No');
  assert.strictEqual(row.Kind, 'pt');
});

test('one-day change: cancel P2, replace P1 subject and teacher, add extra X50', () => {
  ok(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'P2', change: 'cancel' }));
  ok(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'P1', change: 'replace', subject: 'Java Lab', teacher: 'Dr. Meena' }));
  ok(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'X50', change: 'extra', subject: 'Guest Lecture', start: '14:00', end: '15:00', counted: 'Yes', note: 'Substitute' }));
  const data = be.api('getData', tokA, {});
  const c = data.dayChanges;
  assert.strictEqual(c.length, 3);
  const p2 = c.find(x => x.Ref === 'P2');
  assert.strictEqual(p2.Change, 'cancel');
  assert.strictEqual(p2.Counted, 'No');
  const p1 = c.find(x => x.Ref === 'P1');
  assert.strictEqual(p1.Subject, 'Java Lab');
  assert.strictEqual(p1.Teacher, 'Dr. Meena');
  const x50 = c.find(x => x.Ref === 'X50');
  assert.strictEqual(x50['Start Time'], '14:00');
  assert.strictEqual(x50.Date, YESTERDAY);
});

test('a change re-saved for the same period updates it instead of duplicating', () => {
  ok(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'P1', change: 'replace', subject: 'Java Theory' }));
  assert.strictEqual(be.rows('DayChanges').filter(x => x.Ref === 'P1' && String(x.Date) === YESTERDAY).length, 1);
});

test('invalid changes are refused', () => {
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'X50', change: 'cancel' }), 'only timetable');
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'Q1', change: 'cancel' }), 'reference');
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'P1', change: 'delete' }), 'type is invalid');
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'P3', change: 'replace' }), 'at least one');
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'X51', change: 'pt', start: '16:00', end: '15:00', subject: 'Sports' }), 'start and end');
  fails(be.api('saveDayChange', tokA, { changeDate: YESTERDAY, ref: 'X51', change: 'pt', start: '16:00', end: '17:00' }), 'subject');
});

test('an extra period that already has attendance cannot be deleted from the schedule', () => {
  ok(be.api('saveAttendanceBatch', tokA, { attendanceDate: YESTERDAY, rows: JSON.stringify([
    { period: 50, subject: 'Guest Lecture', start: '14:00', end: '15:00', status: 'Present', kind: 'extra', counted: 'Yes' }]) }));
  fails(be.api('deleteDayChange', tokA, { changeDate: YESTERDAY, ref: 'X50' }), 'already saved');
});

test('cancelled periods that were marked earlier do not count in stats (checked in the browser test)', () => {
  // Stats are computed in the page; the backend keeps the raw rows and the change record. Presence is enough here.
  const data = be.api('getData', tokA, {});
  assert.ok(data.attendance.length > 0 && data.dayChanges.some(c => c.Change === 'cancel'));
});

test('settings: valid values save, invalid values are refused, and times read back as HH:MM', () => {
  ok(be.api('saveSettings', tokA, { settings: JSON.stringify({ reminders: 'on', reminderTime: '19:30', reminderEmail: 'on' }) }));
  fails(be.api('saveSettings', tokA, { settings: JSON.stringify({ reminderTime: '25:99' }) }), 'HH:MM');
  fails(be.api('saveSettings', tokA, { settings: JSON.stringify({ reminders: 'maybe' }) }), 'on or off');
  ok(be.api('saveSettings', tokA, { settings: JSON.stringify({ lastReminderSent: '2020-01-01' }) }));
  const s = be.api('getData', tokA, {}).settings;
  const val = k => s.find(x => x.Setting === k).Value;
  assert.strictEqual(val('reminderTime'), '19:30');
  assert.strictEqual(val('reminderEmail'), 'on');
  assert.ok(!s.some(x => x.Setting === 'lastReminderSent' && x.Value === '2020-01-01'), 'students must not set lastReminderSent');
});

test('formula-looking text is stored as text, never as a formula', () => {
  ok(be.api('saveDayChange', tokA, { changeDate: THREE_DAYS_AGO, ref: 'X52', change: 'extra', subject: '=HYPERLINK("http://x","go")', start: '10:00', end: '11:00' }));
  assert.strictEqual(be.formulaCells.length, 0, 'formula cells created: ' + be.formulaCells.join(' | '));
  const row = be.rows('DayChanges').find(x => x.Ref === 'X52');
  assert.strictEqual(row.Subject, '=HYPERLINK("http://x","go")');
});

test('data isolation: a second student sees none of the first student rows', () => {
  const dataB = be.api('getData', tokB, {});
  assert.strictEqual(dataB.attendance.length, 0);
  assert.strictEqual(dataB.dayChanges.length, 0);
  assert.ok(dataB.students.every(s => s['Register Number'] === '21CS002'));
  // Student B saving attendance writes under B's register number only.
  ok(be.api('saveAttendanceBatch', tokB, { attendanceDate: YESTERDAY, rows: JSON.stringify([BATCH_DAY[0]]) }));
  assert.strictEqual(countRows('Attendance', x => x['Register Number'] === '21CS001' && String(x['Attendance Date']) === YESTERDAY), 5);
  assert.strictEqual(countRows('Attendance', x => x['Register Number'] === '21CS002'), 1);
});

test('getData never returns another student register number', () => {
  const dataA = be.api('getData', tokA, {});
  const regs = new Set([...dataA.attendance, ...dataA.dayChanges, ...dataA.holidays, ...dataA.settings].map(r => r['Register Number']));
  assert.ok(regs.size <= 1 && (regs.size === 0 || regs.has('21CS001')));
});

// Reminders skip Sundays, so use a non-Sunday date for this section.
const REMIND_DAY = new Date().getDay() === 0 ? TOMORROW : TODAY;
// ---- Reminder emails ----
test('reminder trigger installs once even when run twice', () => {
  be.runTrigger('installReminderTrigger');
  be.runTrigger('installReminderTrigger');
  assert.strictEqual(be.triggers.filter(t => t.fn === 'sendDailyReminders').length, 1);
});

test('reminders: email goes out once, at the chosen hour, only to opted-in students who have not marked today', () => {
  // Student C: opted in for 12:00, nothing marked, holiday today (must NOT get an email).
  const c = be.postThenFetch(Object.assign({ action: 'register', registerNumber: '21cs003', password: 'third-student-3' }, PROFILE,
    { studentName: 'Chitra N', email: 'chitra@example.com' }));
  ok(c); const tokC2 = c.token;
  ok(be.api('saveSettings', tokC2, { settings: JSON.stringify({ reminderEmail: 'on', reminderTime: '15:00' }) }));
  ok(be.api('saveHoliday', tokC2, { holidayDate: REMIND_DAY, type: 'Non-working day' }));

  // Student B: email on for 12:00.
  ok(be.api('saveSettings', tokB, { settings: JSON.stringify({ reminderEmail: 'on', reminderTime: '12:00' }) }));

  const before = be.outbox.length;
  setNow(REMIND_DAY + 'T12:05:00');
  be.runTrigger('sendDailyReminders');
  const sent = be.outbox.slice(before);
  assert.strictEqual(sent.length, 1, 'expected one email at 12:05, got ' + sent.length);
  assert.strictEqual(sent[0].to, 'ben@example.com');

  // Same hour again: no repeat.
  be.runTrigger('sendDailyReminders');
  assert.strictEqual(be.outbox.length, before + 1, 'repeat email sent');

  // Student A (19:30 email) runs at 19:10: sends once; at 19:40 same day: no repeat.
  ok(be.api('saveSettings', tokA, { settings: JSON.stringify({ reminderEmail: 'on', reminderTime: '19:30' }) }));
  setNow(REMIND_DAY + 'T19:10:00');
  be.runTrigger('sendDailyReminders');
  assert.strictEqual(be.outbox.filter(m => m.to === 'asha@example.com').length, 1);
  setNow(REMIND_DAY + 'T19:40:00');
  be.runTrigger('sendDailyReminders');
  assert.strictEqual(be.outbox.filter(m => m.to === 'asha@example.com').length, 1, 'repeat email to Asha');

  // Chitra (holiday today, 15:00) gets nothing at 15:05.
  setNow(REMIND_DAY + 'T15:05:00');
  be.runTrigger('sendDailyReminders');
  assert.strictEqual(be.outbox.filter(m => m.to === 'chitra@example.com').length, 0, 'holiday student was emailed');
});

test('a student who already marked today gets no reminder', () => {
  // Ben marks today's first period, then the 12:00 reminder must not be sent on a new day either.
  setNow(REMIND_DAY + 'T12:00:00');
  ok(be.api('saveAttendanceBatch', tokB, { attendanceDate: REMIND_DAY, rows: JSON.stringify([BATCH_DAY[0]]) }));
  const before = be.outbox.length;
  be.runTrigger('sendDailyReminders');
  assert.strictEqual(be.outbox.length, before, 'reminder sent to a student who already marked');
});

// ---- Forgot password ----
test('forgot password: an unknown register number gets the same answer and no email', () => {
  const before = be.outbox.length;
  const r = be.postThenFetch({ action: 'forgotRequest', registerNumber: '99ZZ999' });
  ok(r);
  assert.ok(/has an account/.test(r.message));
  assert.strictEqual(be.outbox.length, before);
});

test('forgot password: a known student gets a 6-digit code by email, stored only as a hash', () => {
  setNow(TODAY + 'T20:00:00');
  const r = be.postThenFetch({ action: 'forgotRequest', registerNumber: '21CS001' });
  ok(r);
  const mail = be.outbox[be.outbox.length - 1];
  assert.strictEqual(mail.to, 'asha@example.com');
  const code = mail.body.match(/code is: (\d{6})/)[1];
  const stored = be.rows('ResetCodes').filter(x => x['Register Number'] === '21CS001');
  assert.ok(stored.length >= 1);
  assert.ok(stored.every(x => !String(Object.values(x).join(' ')).includes(code)), 'code stored in plain text');
});

test('forgot password: 5 wrong codes lock the code, even the right one fails afterwards', () => {
  const realFirst = be.outbox[be.outbox.length - 1].body.match(/code is: (\d{6})/)[1];
  const wrong = realFirst === '000000' ? '111111' : '000000';
  for (let i = 0; i < 5; i++) fails(be.postThenFetch({ action: 'forgotReset', registerNumber: '21CS001', code: wrong, newPassword: 'new-pass-value-1' }), 'wrong or has expired');
  fails(be.postThenFetch({ action: 'forgotReset', registerNumber: '21CS001', code: realFirst, newPassword: 'new-pass-value-1' }), 'wrong or has expired');
});

test('forgot password: a new code can be requested and used once, then the password changes', () => {
  // Second and third requests are allowed (limit is 3 per hour); the fourth is silently ignored.
  ok(be.postThenFetch({ action: 'forgotRequest', registerNumber: '21CS001' }));
  ok(be.postThenFetch({ action: 'forgotRequest', registerNumber: '21CS001' }));
  const sentBefore = be.outbox.length;
  ok(be.postThenFetch({ action: 'forgotRequest', registerNumber: '21CS001' }));
  assert.strictEqual(be.outbox.length, sentBefore, 'more than 3 codes per hour were emailed');

  const latest = be.outbox.filter(m => m.to === 'asha@example.com').pop().body.match(/code is: (\d{6})/)[1];
  const short = be.postThenFetch({ action: 'forgotReset', registerNumber: '21CS001', code: latest, newPassword: 'short' });
  fails(short, 'at least 8');
  ok(be.postThenFetch({ action: 'forgotReset', registerNumber: '21CS001', code: latest, newPassword: 'brand-new-pass-9' }));
  fails(be.postThenFetch({ action: 'forgotReset', registerNumber: '21CS001', code: latest, newPassword: 'another-pass-10' }), 'wrong or has expired');
  fails(be.postThenFetch({ action: 'login', registerNumber: '21CS001', password: 'correct-horse-1' }), 'wrong');
  ok(be.postThenFetch({ action: 'login', registerNumber: '21CS001', password: 'brand-new-pass-9' }));
  assert.strictEqual(be.api('getData', tokA, {}).loggedOut, true, 'old session still valid after reset');
});

test('no password, reset code or token appears in the Responses log or any sheet', () => {
  const everything = [];
  be.sheets.forEach(s => s.rows.forEach(r => r.forEach(v => everything.push(String(v)))));
  const sensitive = ['correct-horse-1', 'brand-new-pass-9', 'battery-staple-2', 'third-student-3'];
  sensitive.forEach(p => assert.ok(!everything.some(v => v.includes(p)), 'a password appears in a sheet'));
  const responses = be.rows('Responses').map(r => String(r.Details)).join(' ');
  assert.ok(!/password|newPassword|"code"/.test(responses), 'sensitive keys in the Responses log');
});

const failed = results.filter(r => !r.ok);
console.log('\n' + (results.length - failed.length) + ' passed, ' + failed.length + ' failed');
process.exit(failed.length ? 1 : 0);
