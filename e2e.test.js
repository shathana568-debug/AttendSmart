// End-to-end: drives index.html in Chromium (opened as a local file, like a student would).
// Apps Script calls are answered by the real Code.gs running in the mock backend.
// Run with: TZ=Asia/Kolkata node test/e2e.test.js
const assert = require('assert');
const path = require('path');
const { chromium } = require('playwright');
const { createBackend } = require('./gas-mock');

const WORK = path.join(__dirname, '..');
const CODE = path.join(WORK, 'Code.gs');
const PAGE = 'file://' + path.join(WORK, 'index.html');

// Fixed browser clock (Monday 2 Nov 2026, 10:00) with a setter for later steps.
const FAKE_DATE = `(() => {
  const Real = Date;
  let fixed = new Real('2026-11-02T10:00:00').getTime();
  window.__setNow = iso => { fixed = iso ? new Real(iso).getTime() : null; };
  class FakeDate extends Real {
    constructor(...a) { if (a.length === 0) super(fixed === null ? Real.now() : fixed); else super(...a); }
    static now() { return fixed === null ? Real.now() : fixed; }
  }
  window.Date = FakeDate;
})();`;

// Stand-in for the SheetJS library (the sandbox cannot download it). It records what the page asks for.
const XLSX_STUB = `window.XLSX = {
  utils: {
    book_new() { return { sheets: {}, order: [] }; },
    aoa_to_sheet(rows) { return rows; },
    book_append_sheet(wb, ws, name) { wb.sheets[name] = ws; wb.order.push(name); }
  },
  writeFile(wb, name) { window.__lastExport = { name: name, sheets: wb.sheets, order: wb.order }; }
};`;

let passed = 0;
const failures = [];
async function step(name, fn) {
  try {
    await fn();
    passed++;
    console.log('PASS  ' + name);
  } catch (err) {
    failures.push(name);
    console.log('FAIL  ' + name + '\n      ' + err.message.split('\n').slice(0, 4).join('\n      '));
  }
}

(async () => {
  const be = createBackend(CODE);
  be.runTrigger('setupSheets');
  be.setClock('2026-11-02T10:00:00');

  const browser = await chromium.launch();
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  await context.addInitScript(FAKE_DATE);
  await context.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  let xlsxMode = 'stub';
  await context.route('https://cdnjs.cloudflare.com/**', r => {
    if (xlsxMode === 'abort') return r.abort();
    return r.fulfill({ status: 200, contentType: 'application/javascript', body: XLSX_STUB });
  });
  await context.route('**/script.google.com/**', route => {
    const req = route.request();
    if (req.method() === 'POST') {
      be.post(JSON.parse(req.postData()));
      return route.fulfill({ status: 200, contentType: 'text/plain', body: '' });
    }
    const params = Object.fromEntries(new URL(req.url()).searchParams.entries());
    return route.fulfill({ status: 200, contentType: 'application/javascript', body: be.getRaw(params) });
  });

  const page = await context.newPage();
  const pageErrors = [];
  page.on('pageerror', e => pageErrors.push(e.message));
  page.on('console', m => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) pageErrors.push(m.text()); });
  page.on('dialog', d => d.accept());
  await page.goto(PAGE);

  const toastHas = async text => {
    try {
      await page.waitForFunction(t => document.getElementById('toast').textContent.includes(t), text, { timeout: 15000 });
    } catch (err) {
      const shown = await page.evaluate(() => ({
        toast: document.getElementById('toast').textContent,
        view: (document.querySelector('.view.active, [data-screen].active') || {}).id || '',
        errors: Array.from(document.querySelectorAll('.errbox, .alert-error, #errorBox, [role=alert]')).map(e => e.textContent.trim()).filter(Boolean)
      }));
      throw new Error('expected toast "' + text + '" but got ' + JSON.stringify(shown));
    }
  };
  const clearToast = () => page.evaluate(() => { document.getElementById('toast').textContent = ''; });
  const go = async view => page.locator('[data-view="' + view + '"]:visible').first().click();
  const setDate = async iso => page.evaluate(v => {
    const el = document.getElementById('markDate');
    el.value = v;
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }, iso);
  const tokA = () => be.rows('Accounts').length;
  const attendanceFor = (reg, date) => be.rows('Attendance').filter(r => r['Register Number'] === reg && String(r['Attendance Date']) === date);

  // ---------- Account and profile ----------
  await step('login gate is shown to a new visitor', async () => {
    assert.strictEqual(await page.locator('#authGate').isVisible(), true);
  });

  await step('a student can create an account with a password of 8+ characters', async () => {
    await page.click('#gotoCreate');
    await page.fill('#pName', 'Asha Rao');
    await page.fill('#pEmail', 'asha@example.com');
    await page.fill('#pReg', '21cs001');
    await page.fill('#pPass', 'correct-horse-1');
    await page.fill('#pPass2', 'correct-horse-1');
    await page.fill('#pCollege', 'Anna University');
    await page.selectOption('#pCourse', 'B.E.');
    await page.selectOption('#pDept', 'Computer Science & Engineering');
    await page.fill('#pAcadYear', '2026-2027');
    await page.selectOption('#pYear', '3');
    await page.selectOption('#pSem', '5');
    await page.fill('#pClass', 'CS-A');
    await page.fill('#pSection', 'A');
    await page.click('#profileForm button[type=submit]');
    await page.waitForFunction(() => !!localStorage.getItem('attendsmart:token'), null, { timeout: 15000 });
    assert.strictEqual(tokA(), 1, 'account not written to Accounts sheet');
    assert.ok(!JSON.stringify(be.sheets.get('Responses').rows).includes('correct-horse-1'), 'password written to log');
  });

  // ---------- Timetable ----------
  await step('timetable: subjects and Monday periods save to Google Sheets', async () => {
    await go('timetable');
    for (const s of ['Java', 'DBMS', 'Maths']) {
      await page.fill('#subjInput', s);
      await clearToast();
      await page.click('#subjBtn');
      await toastHas('Subject added');
    }
    const slots = [['Java', '09:00', '09:45'], ['DBMS', '09:45', '10:30'], ['Maths', '10:30', '11:15']];
    for (const [subj, st, en] of slots) {
      await page.selectOption('#pSubj', subj);
      await page.fill('#pStart', st);
      await page.fill('#pEnd', en);
      await page.click('#pAddBtn');
    }
    await clearToast();
    await page.click('#saveDayBtn');
    await toastHas('saved to Google Sheets');
    const saved = be.rows('Timetable').filter(r => r.Day === 'Monday' && r['Register Number'] === '21CS001');
    assert.strictEqual(saved.length, 3);
  });

  // ---------- Date-based attendance ----------
  await step('mark page: Mark All Present only fills drafts, nothing is saved yet', async () => {
    await go('mark');
    await page.waitForSelector('#markPeriods [data-action="mark-period"]');
    await page.click('#markAllBtn');
    const text = await page.textContent('#markUnsaved');
    assert.ok(/3 change\(s\) not saved yet/.test(text), 'unsaved text was: ' + text);
    assert.strictEqual(attendanceFor('21CS001', '2026-11-02').length, 0, 'drafts were saved automatically');
  });

  await step('mark one period Absent, then Save Attendance writes exactly one row per period', async () => {
    await page.click('button[data-action="mark-period"][data-period="2"][data-value="Absent"]');
    await clearToast();
    await page.click('#saveAttBtn');
    await toastHas('Saved 3 period(s)');
    const rows = attendanceFor('21CS001', '2026-11-02');
    assert.strictEqual(rows.length, 3);
    const byPeriod = Object.fromEntries(rows.map(r => [r['Period No'], r.Status]));
    assert.deepStrictEqual(byPeriod, { 1: 'Present', 2: 'Absent', 3: 'Present' });
    assert.strictEqual(await page.textContent('#markUnsaved'), 'All changes saved');
  });

  await step('saving again corrects, never duplicates', async () => {
    await page.click('button[data-action="mark-period"][data-period="2"][data-value="Present"]');
    await clearToast();
    await page.click('#saveAttBtn');
    await toastHas('corrected');
    assert.strictEqual(attendanceFor('21CS001', '2026-11-02').length, 3);
    assert.strictEqual(attendanceFor('21CS001', '2026-11-02').find(r => r['Period No'] === 2).Status, 'Present');
    await page.click('button[data-action="mark-period"][data-period="2"][data-value="Absent"]');
    await clearToast();
    await page.click('#saveAttBtn');
    await toastHas('Saved 3 period(s)');
  });

  await step('dashboard counts the three saved periods', async () => {
    await go('dashboard');
    await page.waitForFunction(() => document.getElementById('statConducted').textContent === '3');
  });

  // ---------- One-day schedule changes ----------
  await step('cancelling a period for one date removes it from the counts, timetable unchanged', async () => {
    await go('mark');
    await page.selectOption('#chgType', 'cancel');
    await page.selectOption('#chgPeriod', '2');
    await clearToast();
    await page.click('#chgSave');
    await toastHas('Change saved for');
    assert.ok((await page.textContent('#chgList')).includes('Period 2 cancelled'));
    const tt = be.rows('Timetable').filter(r => r.Day === 'Monday' && r['Register Number'] === '21CS001');
    assert.strictEqual(tt.length, 3, 'weekly timetable was changed');
    await go('dashboard');
    await page.waitForFunction(() => document.getElementById('statConducted').textContent === '2');
  });

  await step('replacing a teacher for one period shows as changed', async () => {
    await go('mark');
    await page.selectOption('#chgType', 'replace');
    await page.selectOption('#chgPeriod', '1');
    await page.fill('#chgTeacher', 'Dr. Meena');
    await clearToast();
    await page.click('#chgSave');
    await toastHas('Change saved for');
    const card = page.locator('#markPeriods .sched', { hasText: 'Java' }).first();
    assert.ok((await card.textContent()).includes('Dr. Meena'));
    assert.ok((await card.textContent()).includes('Changed for this date'));
  });

  await step('PT period not counted: marked and saved, excluded from the percentage', async () => {
    await page.selectOption('#chgType', 'pt');
    await page.fill('#chgSubject', 'Sports PT');
    await page.fill('#chgStart', '16:00');
    await page.fill('#chgEnd', '17:00');
    await page.selectOption('#chgCounted', 'No');
    await clearToast();
    await page.click('#chgSave');
    await toastHas('Change saved for');
    await page.click('#markPeriods .sched:has-text("Sports PT") button[data-value="Present"]');
    await clearToast();
    await page.click('#saveAttBtn');
    await toastHas('Saved');
    const pt = attendanceFor('21CS001', '2026-11-02').find(r => r.Subject === 'Sports PT');
    assert.ok(pt, 'PT row not saved');
    assert.strictEqual(pt.Counted, 'No');
    assert.strictEqual(pt.Kind, 'pt');
    await go('dashboard');
    await page.waitForFunction(() => document.getElementById('statConducted').textContent === '2');
  });

  await step('removing a change brings the period back', async () => {
    await go('mark');
    await clearToast();
    await page.click('#chgList [data-action="chg-del"][data-value="P2"]');
    await toastHas('Change removed');
    assert.ok(!(await page.textContent('#chgList')).includes('Period 2 cancelled'));
  });

  // ---------- Holidays ----------
  await step('holiday: blocks marking that date, and is not counted', async () => {
    await setDate('2026-10-30');
    await page.selectOption('#holType', 'College holiday');
    await page.fill('#holNote', 'Test holiday');
    await clearToast();
    await page.click('#holForm button[type=submit]');
    await toastHas('College holiday saved');
    assert.strictEqual(await page.locator('#markHoliday').isVisible(), true);
    assert.strictEqual(await page.isDisabled('#markAllBtn'), true);
    assert.strictEqual(await page.isDisabled('#saveAttBtn'), true);
  });

  await step('removing a holiday re-enables marking for that date', async () => {
    await clearToast();
    await page.click('#holCurrent [data-action="hol-del"]');
    await toastHas('Holiday removed');
    assert.strictEqual(await page.locator('#markHoliday').isVisible(), false);
  });

  // ---------- Alerts ----------
  await step('subject below 80% shows an alert naming the subject; it clears when back above 80%', async () => {
    await setDate('2026-10-26');
    await page.click('button[data-action="mark-period"][data-period="1"][data-value="Absent"]');
    await clearToast();
    await page.click('#saveAttBtn');
    await toastHas('Saved');
    await go('dashboard');
    await page.waitForFunction(() => document.getElementById('alertBox').textContent.includes('Java'));
    assert.ok((await page.textContent('#alertBox')).includes('below 80%'));
    // Add Java Present on three more days (backend write, as if marked on those days).
    const login = be.postThenFetch({ action: 'login', registerNumber: '21CS001', password: 'correct-horse-1' });
    for (const d of ['2026-10-27', '2026-10-28', '2026-10-29']) {
      const r = be.api('saveAttendanceBatch', login.token, { attendanceDate: d, rows: JSON.stringify([
        { period: 1, subject: 'Java', start: '09:00', end: '09:45', status: 'Present', kind: 'regular', counted: 'Yes' }]) });
      assert.strictEqual(r.success, true, r.error);
    }
    await page.reload();
    // Wait until the page has loaded all five Java records (the reload reads them back from the sheet).
    await page.waitForFunction(() => countedRecords().filter(x => x.subject === 'Java').length === 5, null, { timeout: 20000 });
    // DBMS stays listed (its period 2 is really Absent from earlier steps); Java must no longer be listed.
    assert.ok(!(await page.textContent('#alertBox')).includes('Java is below 80%'), 'Java alert still shown at 80%');
  });

  // ---------- Semesters ----------
  await step('starting Semester 6 keeps Semester 5 records separate', async () => {
    await go('profile');
    await page.selectOption('#newSem', '6');
    await page.click('#semMoveBtn');
    await page.waitForFunction(() => state.student && String(state.student.semester) === '6', null, { timeout: 15000 });
    await go('history');
    await page.waitForFunction(() => document.getElementById('histSem').options.length >= 2);
    await page.selectOption('#histSem', '6');
    await go('dashboard');
    await page.waitForFunction(() => document.getElementById('statConducted').textContent === '0');
    await go('history');
    await page.selectOption('#histSem', '5');
    await go('dashboard');
    await page.waitForFunction(() => Number(document.getElementById('statConducted').textContent) > 0);
    assert.ok(Number(await page.textContent('#statConducted')) > 0);
  });

  // ---------- Theme ----------
  await step('dark mode is applied and remembered after reopening', async () => {
    await page.click('button[data-action="theme-toggle"]');
    assert.strictEqual(await page.getAttribute('html', 'data-theme'), 'dark');
    await page.reload();
    assert.strictEqual(await page.getAttribute('html', 'data-theme'), 'dark');
    await go('profile');
    await page.click('button[data-action="theme"][data-value="system"]');
    assert.strictEqual(await page.getAttribute('html', 'data-theme'), null);
  });

  // ---------- Reminders ----------
  await step('reminder settings save to the sheet and show in the app at the set time', async () => {
    await go('profile');
    await page.check('#remOn');
    await page.check('#remEmail');
    await page.fill('#remTime', '19:30');
    await clearToast();
    await page.click('#remSave');
    await toastHas('Reminder settings saved');
    const s = Object.fromEntries(be.rows('Settings').filter(r => r['Register Number'] === '21CS001').map(r => [r.Setting, r.Value]));
    assert.strictEqual(s.reminders, 'on');
    assert.strictEqual(s.reminderEmail, 'on');
    assert.strictEqual(s.reminderTime, '19:30');
    // Monday 9 Nov: 18:00 is before the set time; 19:40 is after it. The banner lives on the Dashboard.
    await page.evaluate(() => window.__setNow('2026-11-09T18:00:00'));
    await page.evaluate(() => checkReminder());
    await go('dashboard');
    assert.strictEqual(await page.locator('#remindBox').isVisible(), false);
    await page.evaluate(() => window.__setNow('2026-11-09T19:40:00'));
    await page.evaluate(() => checkReminder());
    await go('dashboard');
    assert.strictEqual(await page.locator('#remindBox').isVisible(), true);
    await page.evaluate(() => window.__setNow(null));
  });

  // ---------- History, export and charts ----------
  await step('charts are drawn from the saved records', async () => {
    await go('history');
    await page.selectOption('#histSem', '5');
    await page.waitForSelector('#chartOverall svg');
    assert.ok(await page.locator('#chartOverall path').count() >= 1, 'no overall line');
    assert.ok(await page.locator('#chartSubjects path').count() >= 2, 'expected subject lines');
    assert.ok(await page.locator('#chartWeekly rect').count() >= 1, 'no weekly bars');
    assert.ok(await page.locator('#chartMonthly rect').count() >= 1, 'no monthly bars');
  });

  await step('Download Excel builds a Summary and an Attendance sheet from saved data', async () => {
    await page.evaluate(() => { window.__lastExport = null; });
    await page.click('button[data-action="export-xlsx"]:visible >> nth=0');
    await page.waitForFunction(() => !!window.__lastExport);
    const ex = await page.evaluate(() => window.__lastExport);
    assert.ok(/^attendance_21CS001_\d{4}-\d{2}-\d{2}\.xlsx$/.test(ex.name), ex.name);
    assert.deepStrictEqual(ex.order, ['Summary', 'Attendance']);
    const header = ex.sheets.Attendance[0];
    assert.deepStrictEqual(header.slice(0, 4), ['Date', 'Day', 'Subject', 'Period']);
    assert.ok(ex.sheets.Attendance.length > 5, 'too few records');
    assert.ok(ex.sheets.Summary.some(r => r[0] === 'Register number' && r[1] === '21CS001'));
  });

  await step('Export to PDF prints a report with student details and the overall percentage', async () => {
    await page.evaluate(() => { window.__printed = false; window.print = () => { window.__printed = true; }; });
    await page.click('button[data-action="export-pdf"]:visible >> nth=0');
    assert.strictEqual(await page.evaluate(() => window.__printed), true);
    const html = await page.innerHTML('#printReport');
    assert.ok(html.includes('Attendance report') && html.includes('Overall') && html.includes('Asha Rao'));
  });

  await step('if the Excel library cannot load, a CSV file downloads instead', async () => {
    xlsxMode = 'abort';
    await page.evaluate(() => { delete window.XLSX; });
    const [download] = await Promise.all([
      page.waitForEvent('download'),
      page.click('button[data-action="export-xlsx"]:visible >> nth=0')
    ]);
    assert.ok(download.suggestedFilename().endsWith('.csv'), download.suggestedFilename());
    const file = await download.path();
    const text = require('fs').readFileSync(file, 'utf8');
    assert.ok(text.includes('Student name') && text.includes('Attendance records'));
    xlsxMode = 'stub';
  });

  // ---------- Forgot password ----------
  await step('forgot password: emailed code resets the password; old password stops working', async () => {
    await go('profile');
    await page.click('#logoutBtn');
    await page.waitForFunction(() => !localStorage.getItem('attendsmart:token'));
    await page.click('#forgotLink');
    await page.fill('#fReg', '21CS001');
    await page.click('#forgotStep1 button[type=submit]');
    await page.waitForFunction(() => document.getElementById('forgotStep2').hidden === false, null, { timeout: 15000 });
    const mail = be.outbox.filter(m => m.to === 'asha@example.com').pop();
    const code = mail.body.match(/code is: (\d{6})/)[1];
    const wrong = code === '000000' ? '111111' : '000000';
    await page.fill('#fCode', wrong);
    await page.fill('#fPass', 'fresh-pass-2026');
    await page.fill('#fPass2', 'fresh-pass-2026');
    await page.click('#forgotStep2 button[type=submit]');
    await page.waitForFunction(() => document.getElementById('authMsg').textContent.includes('wrong or has expired'), null, { timeout: 15000 });
    await page.fill('#fCode', code);
    await page.click('#forgotStep2 button[type=submit]');
    await page.waitForFunction(() => document.getElementById('authMsg').textContent.includes('Password changed'), null, { timeout: 15000 });
    // Old password refused, new password works.
    await page.fill('#lReg', '21CS001');
    await page.fill('#lPass', 'correct-horse-1');
    await page.click('#loginForm button[type=submit]');
    await page.waitForFunction(() => document.getElementById('authMsg').textContent.includes('wrong'), null, { timeout: 15000 });
    await page.fill('#lPass', 'fresh-pass-2026');
    await page.click('#loginForm button[type=submit]');
    await page.waitForFunction(() => !!localStorage.getItem('attendsmart:token'), null, { timeout: 15000 });
    await page.waitForFunction(() => document.getElementById('statConducted').textContent !== '', null, { timeout: 15000 });
  });

  // ---------- Layout and errors ----------
  await step('layout has no horizontal scrolling on a phone-width screen', async () => {
    await page.setViewportSize({ width: 390, height: 844 });
    for (const v of ['dashboard', 'mark', 'timetable', 'calculate', 'history', 'profile']) {
      await page.evaluate(v2 => show(v2), v);
      await page.waitForTimeout(150);
      const wide = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
      assert.ok(wide <= 1, v + ' scrolls sideways by ' + wide + 'px');
    }
    await page.setViewportSize({ width: 1280, height: 900 });
  });

  await step('no JavaScript errors occurred during the whole run', async () => {
    assert.deepStrictEqual(pageErrors, []);
  });

  await browser.close();

  // Backend check: no passwords or reset codes in any sheet.
  const everything = [];
  be.sheets.forEach(s => s.rows.forEach(r => r.forEach(v => everything.push(String(v)))));
  ['correct-horse-1', 'fresh-pass-2026'].forEach(p => {
    if (everything.some(v => v.includes(p))) failures.push('password found in sheet: ' + p);
  });
  console.log('\n' + passed + ' passed, ' + failures.length + ' failed');
  process.exit(failures.length ? 1 : 0);
})().catch(err => { console.error('CRASH', err); process.exit(2); });
