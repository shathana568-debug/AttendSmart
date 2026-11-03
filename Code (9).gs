/**
 * Smart Attendance Tracker - Google Apps Script backend (Code.gs)
 *
 * Sheets used (created automatically; existing data is never deleted):
 *   Students   - student profiles (one row per registration)
 *   Attendance - one row per student, date and period
 *   Subjects   - subjects each student added (keyed by register number)
 *   Timetable  - each student's periods per day (keyed by register number)
 *   Responses  - log of every request received
 *
 * Actions (sent as the "action" parameter):
 *   registerStudent    -> saves a student profile
 *   saveAttendance     -> saves, or corrects, one period's Present/Absent status
 *   saveSubjects       -> replaces the student's subject list
 *   saveTimetableDay   -> replaces the student's periods for one day
 *   getData            -> returns the student's profile, attendance, subjects and timetable
 *                         (requires "registerNumber")
 */

// Leave empty if this script is opened from the Google Sheet (Extensions > Apps Script).
// Otherwise paste the spreadsheet ID from the sheet's URL.
const SPREADSHEET_ID = '';

// Shown by the API status check so you can confirm which version is deployed.
const CODE_VERSION = 'features-v8';
const SUPPORTED_ACTIONS = ['register', 'login', 'forgotRequest', 'forgotReset', 'registerStudent', 'saveAttendance', 'saveAttendanceBatch', 'saveSubjects', 'saveTimetableDay', 'saveDayChange', 'deleteDayChange', 'saveHoliday', 'deleteHoliday', 'saveSettings', 'getData', 'extractTimetable', 'uploadTimetable', 'fetchResult'];

const MAX_TEXT_LENGTH = 200;
const MAX_PERIODS = 20;
const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const SHEETS = {
  students: {
    name: 'Students',
    headers: ['Timestamp', 'Student Name', 'Register Number', 'Email', 'Course',
      'Department', 'Year', 'Class', 'Section', 'Semester', 'College Name', 'Academic Year']
  },
  attendance: {
    name: 'Attendance',
    headers: ['Timestamp', 'Register Number', 'Student Name', 'Subject', 'Attendance Date',
      'Attendance Time', 'Status', 'Period No', 'End Time', 'Day', 'Semester', 'Counted', 'Kind']
  },
  subjects: {
    name: 'Subjects',
    headers: ['Timestamp', 'Register Number', 'Subject']
  },
  timetable: {
    name: 'Timetable',
    headers: ['Timestamp', 'Register Number', 'Day', 'Period No', 'Type', 'Start Time', 'End Time', 'Subject']
  },
  sessions: {
    name: 'Sessions',
    headers: ['Timestamp', 'Token Hash', 'Register Number', 'Expires']
  },
  accounts: {
    name: 'Accounts',
    headers: ['Timestamp', 'Register Number', 'Salt', 'Password Hash']
  },
  dayChanges: {
    name: 'DayChanges',
    headers: ['Timestamp', 'Register Number', 'Date', 'Ref', 'Change', 'Subject', 'Teacher',
      'Start Time', 'End Time', 'Counted', 'Note']
  },
  holidays: {
    name: 'Holidays',
    headers: ['Timestamp', 'Register Number', 'Date', 'Type', 'Note']
  },
  settings: {
    name: 'Settings',
    headers: ['Timestamp', 'Register Number', 'Setting', 'Value']
  },
  resetCodes: {
    name: 'ResetCodes',
    headers: ['Timestamp', 'Register Number', 'Salt', 'Code Hash', 'Expires', 'Attempts', 'Used']
  },
  responses: {
    name: 'Responses',
    headers: ['Timestamp', 'Action', 'Register Number', 'Details', 'Result']
  }
};

const STUDENT_REQUIRED = ['studentName', 'registerNumber', 'email', 'course', 'department',
  'year', 'class', 'section', 'semester', 'collegeName', 'academicYear'];
const ATTENDANCE_REQUIRED = ['registerNumber', 'subject', 'attendanceDate', 'attendanceTime',
  'endTime', 'periodNo', 'status'];

/* ---------- Entry points ---------- */

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!params.action) {
    return respond({ success: true, message: 'Smart Attendance Tracker API is running.', version: CODE_VERSION, actions: SUPPORTED_ACTIONS }, params);
  }
  return handleRequest(params);
}

function doPost(e) {
  let body = {};
  try {
    const raw = e && e.postData && e.postData.contents ? String(e.postData.contents).trim() : '';
    if (raw.charAt(0) === '{') body = JSON.parse(raw);
  } catch (err) {
    return respond({ success: false, error: 'Request body is not valid JSON.' }, {});
  }
  const bodyAction = clean(body.action);
  if (bodyAction === 'uploadTimetable') {
    try {
      return respond(withAuth(body, saveUpload), {});
    } catch (err) {
      return respond(fail('Upload failed: ' + err.message), {});
    }
  }
  // Every other request is a POST with a requestId; the result is collected through fetchResult.
  return handlePost(bodyAction, body);
}

/* ---------- Routing ---------- */

function handleRequest(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return respond({ success: false, error: 'Server is busy. Please try again in a moment.' }, params);
  }

  try {
    // The session token must never travel in a URL (URLs are kept in browser history and server logs).
    if (params.token) return respond(fail('Send the session token in the request body, not the URL.'), params);
    const action = clean(params.action);
    const result = runAction(params);
    if (action !== 'getData' && action !== 'fetchResult') {
      logResponse(action, clean(params.registerNumber).toUpperCase(), params, result);
    }
    return respond(result, params);
  } catch (err) {
    return respond({ success: false, error: 'Server error: ' + err.message }, params);
  } finally {
    lock.releaseLock();
  }
}

// Runs one action and returns its result object. Used by the GET route and by the POST route.
function runAction(params) {
  const action = clean(params.action);
  let result;
  switch (action) {
      case 'register':
      case 'login': result = fail('Please log in from the app.'); break;
      case 'registerStudent': result = withAuth(params, saveStudent); break;
      case 'saveAttendance': result = withAuth(params, saveAttendance); break;
      case 'saveSubjects': result = withAuth(params, saveSubjects); break;
      case 'saveTimetableDay': result = withAuth(params, saveTimetableDay); break;
      case 'getData': result = withAuth(params, getData); break;
      case 'fetchResult': result = fetchResult(params); break;
      case 'extractTimetable': result = withAuth(params, extractTimetable); break;
      case 'saveAttendanceBatch': result = withAuth(params, saveAttendanceBatch); break;
      case 'saveDayChange': result = withAuth(params, saveDayChange); break;
      case 'deleteDayChange': result = withAuth(params, deleteDayChange); break;
      case 'saveHoliday': result = withAuth(params, saveHoliday); break;
      case 'deleteHoliday': result = withAuth(params, deleteHoliday); break;
      case 'saveSettings': result = withAuth(params, saveSettings); break;
      default: result = { success: false, error: 'Unknown action: ' + action, version: CODE_VERSION, actions: SUPPORTED_ACTIONS };
  }
  return result;
}

// Returns JSONP when a valid "callback" name is given (works from file://), otherwise plain JSON.
function respond(obj, params) {
  const text = JSON.stringify(obj);
  const cb = clean(params && params.callback);
  if (/^[A-Za-z_$][\w$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + text + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function fail(message) {
  return { success: false, error: message };
}

/* ---------- Sheet helpers ---------- */

function getSpreadsheet() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No spreadsheet found. Set SPREADSHEET_ID in Code.gs.');
  return ss;
}

// Creates the sheet if needed and adds any missing header columns at the end.
function getSheet(def) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(def.name);
  if (!sheet) sheet = ss.insertSheet(def.name);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(clean);
    def.headers.forEach(function (h) {
      if (current.indexOf(h) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h).setFontWeight('bold');
        current.push(h);
      }
    });
  }
  return sheet;
}

function headerRow(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(clean);
}

// Returns each data row as an object keyed by its header name (plus _row = sheet row number).
function readRecords(def) {
  const sheet = getSheet(def);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(clean);
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const obj = { _row: i + 1 };
    headers.forEach(function (h, c) {
      if (h) obj[h] = formatCell(values[i][c], h);
    });
    out.push(obj);
  }
  return out;
}

// Appends a row, placing each value under the matching header.
function appendObject(def, obj) {
  const sheet = getSheet(def);
  const row = headerRow(sheet).map(function (h) {
    if (h === 'Timestamp') return new Date();
    return obj[h] === undefined ? '' : obj[h];
  });
  sheet.appendRow(row);
}

// Updates selected columns of an existing row.
function setRowValues(def, rowNumber, obj) {
  const sheet = getSheet(def);
  const headers = headerRow(sheet);
  Object.keys(obj).forEach(function (key) {
    const col = headers.indexOf(key);
    if (col !== -1) sheet.getRange(rowNumber, col + 1).setValue(obj[key]);
  });
}

// Deletes every row for which predicate(rowObject) is true (bottom-up so row numbers stay valid).
function deleteRowsWhere(def, predicate) {
  const sheet = getSheet(def);
  const rows = readRecords(def);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (predicate(rows[i])) sheet.deleteRow(rows[i]._row);
  }
}

function formatCell(value, header) {
  if (!(value instanceof Date)) return value;
  const tz = Session.getScriptTimeZone();
  if (header === 'Attendance Date' || header === 'Date') return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  if (header === 'Attendance Time' || header === 'Start Time' || header === 'End Time' || header === 'Value') {
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  return Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm:ss');
}

/* ---------- Validation helpers ---------- */

function clean(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

function missingFields(params, required) {
  return required.filter(function (field) { return clean(params[field]) === ''; });
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidDate(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!m) return false;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const date = new Date(y, mo - 1, d);
  return date.getFullYear() === y && date.getMonth() === mo - 1 && date.getDate() === d;
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

// "9:05" or "09:05" -> "09:05"; anything else -> ""
function normalizeTime(value) {
  const m = /^(\d{1,2}):(\d{2})$/.exec(clean(value));
  if (!m) return '';
  const h = Number(m[1]), mi = Number(m[2]);
  if (h > 23 || mi > 59) return '';
  return pad2(h) + ':' + pad2(mi);
}

function parseList(value) {
  const text = clean(value);
  if (!text) return [];
  try {
    const list = JSON.parse(text);
    return Array.isArray(list) ? list : null;
  } catch (err) {
    return null;
  }
}

function findStudent(regNo) {
  const rows = readRecords(SHEETS.students).filter(function (r) {
    return clean(r['Register Number']).toUpperCase() === regNo;
  });
  return rows.length ? rows[rows.length - 1] : null;
}

/* ---------- Actions ---------- */

function saveStudent(p) {
  const missing = missingFields(p, STUDENT_REQUIRED);
  if (missing.length) return fail('Missing required fields: ' + missing.join(', '));

  const tooLong = STUDENT_REQUIRED.filter(function (f) { return clean(p[f]).length > MAX_TEXT_LENGTH; });
  if (tooLong.length) return fail('Values too long: ' + tooLong.join(', '));

  const email = clean(p.email);
  if (!isValidEmail(email)) return fail('Invalid email address.');

  const regNo = clean(p.registerNumber).toUpperCase();
  appendObject(SHEETS.students, {
    'Student Name': clean(p.studentName),
    'Register Number': regNo,
    'Email': email,
    'Course': clean(p.course),
    'Department': clean(p.department),
    'Year': clean(p.year),
    'Class': clean(p.class),
    'Section': clean(p.section),
    'Semester': clean(p.semester),
    'College Name': clean(p.collegeName),
    'Academic Year': clean(p.academicYear)
  });
  return { success: true, message: 'Student profile saved.', registerNumber: regNo };
}

// Saves a period's status. If the same student, date and period already has a row,
// that row is corrected instead of creating a duplicate.
function saveAttendance(p) {
  const missing = missingFields(p, ATTENDANCE_REQUIRED);
  if (missing.length) return fail('Missing required fields: ' + missing.join(', '));

  const rawStatus = clean(p.status).toLowerCase();
  const status = rawStatus === 'present' ? 'Present' : rawStatus === 'absent' ? 'Absent' : '';
  if (!status) return fail('Status must be Present or Absent.');

  const date = clean(p.attendanceDate);
  if (!isValidDate(date)) return fail('Attendance date must be in YYYY-MM-DD format.');

  const start = normalizeTime(p.attendanceTime);
  const end = normalizeTime(p.endTime);
  if (!start || !end) return fail('Start and end time must be HH:MM (24-hour).');
  if (start >= end) return fail('End time must be after the start time.');

  const periodNo = clean(p.periodNo);
  if (!/^\d{1,2}$/.test(periodNo)) return fail('Period number is invalid.');

  const subject = clean(p.subject);
  if (subject.length > MAX_TEXT_LENGTH) return fail('Subject name is too long.');

  const regNo = clean(p.registerNumber).toUpperCase();
  const student = findStudent(regNo);
  if (!student) return fail('Register number is not registered: ' + regNo);
  const name = clean(p.studentName) || clean(student['Student Name']);

  const existing = readRecords(SHEETS.attendance).filter(function (r) {
    return clean(r['Register Number']).toUpperCase() === regNo &&
      String(r['Attendance Date']) === date &&
      String(r['Period No']) === periodNo;
  });

  if (existing.length) {
    setRowValues(SHEETS.attendance, existing[0]._row, {
      'Timestamp': new Date(),
      'Student Name': name,
      'Subject': subject,
      'Attendance Time': start,
      'End Time': end,
      'Status': status
    });
    return { success: true, updated: true, message: 'Attendance corrected.', status: status };
  }

  appendObject(SHEETS.attendance, {
    'Register Number': regNo,
    'Student Name': name,
    'Subject': subject,
    'Attendance Date': date,
    'Attendance Time': start,
    'End Time': end,
    'Period No': Number(periodNo),
    'Status': status,
    'Day': weekdayName(date),
    'Semester': clean(p.semester) || clean(student['Semester']),
    'Counted': 'Yes',
    'Kind': 'regular'
  });
  return { success: true, updated: false, message: 'Attendance saved.', status: status };
}

// Replaces the student's subject list (duplicates ignored, case-insensitive).
function saveSubjects(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!regNo) return fail('Register number is required.');
  if (!findStudent(regNo)) return fail('Register number is not registered. Save your profile first.');

  const list = parseList(p.subjects);
  if (list === null) return fail('Subjects must be a list.');

  const seen = {};
  const names = [];
  list.forEach(function (item) {
    const n = clean(item);
    if (!n || n.length > 60) return;
    const key = n.toLowerCase();
    if (!seen[key]) { seen[key] = true; names.push(n); }
  });
  if (names.length > 60) return fail('Too many subjects (maximum 60).');

  deleteRowsWhere(SHEETS.subjects, function (r) {
    return clean(r['Register Number']).toUpperCase() === regNo;
  });
  names.forEach(function (n) {
    appendObject(SHEETS.subjects, { 'Register Number': regNo, 'Subject': n });
  });
  return { success: true, message: 'Subjects saved.', count: names.length };
}

// Replaces the student's periods for one day. Class periods are numbered 1, 2, 3...
// in time order; breaks have no period number.
function saveTimetableDay(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!regNo) return fail('Register number is required.');
  if (!findStudent(regNo)) return fail('Register number is not registered. Save your profile first.');

  const day = clean(p.day);
  if (DAYS.indexOf(day) === -1) return fail('Day must be Monday to Saturday.');

  const list = parseList(p.rows);
  if (list === null) return fail('Periods data is not valid.');
  if (list.length > MAX_PERIODS) return fail('A day can have at most ' + MAX_PERIODS + ' periods.');

  const rows = [];
  for (let i = 0; i < list.length; i++) {
    const r = list[i];
    const label = 'Period ' + (i + 1) + ': ';
    if (!Array.isArray(r)) return fail(label + 'data is not valid.');
    const type = clean(r[0]).toLowerCase();
    const start = normalizeTime(r[1]);
    const end = normalizeTime(r[2]);
    const subject = clean(r[3]);
    if (type !== 'class' && type !== 'break') return fail(label + 'type must be class or break.');
    if (!start || !end) return fail(label + 'enter a valid start and end time.');
    if (start >= end) return fail(label + 'end time must be after the start time.');
    if (type === 'class' && !subject) return fail(label + 'subject name is required.');
    if (subject.length > MAX_TEXT_LENGTH) return fail(label + 'subject name is too long.');
    rows.push({ type: type, start: start, end: end, subject: type === 'class' ? subject : '' });
  }

  rows.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].start < rows[i - 1].end) {
      return fail('Periods overlap: ' + rows[i - 1].start + ' - ' + rows[i - 1].end +
        ' and ' + rows[i].start + ' - ' + rows[i].end + '.');
    }
  }

  deleteRowsWhere(SHEETS.timetable, function (r) {
    return clean(r['Register Number']).toUpperCase() === regNo && r['Day'] === day;
  });

  let periodNo = 0;
  rows.forEach(function (r) {
    if (r.type === 'class') periodNo++;
    appendObject(SHEETS.timetable, {
      'Register Number': regNo,
      'Day': day,
      'Period No': r.type === 'class' ? periodNo : '',
      'Type': r.type,
      'Start Time': r.start,
      'End Time': r.end,
      'Subject': r.subject
    });
  });

  // Keep every subject used in a timetable in the Subjects list so it can be reused.
  const known = readRecords(SHEETS.subjects)
    .filter(function (r) { return clean(r['Register Number']).toUpperCase() === regNo; })
    .map(function (r) { return clean(r['Subject']).toLowerCase(); });
  rows.forEach(function (r) {
    if (r.type !== 'class') return;
    const key = r.subject.toLowerCase();
    if (known.indexOf(key) === -1) {
      known.push(key);
      appendObject(SHEETS.subjects, { 'Register Number': regNo, 'Subject': r.subject });
    }
  });

  return { success: true, message: day + ' timetable saved.', count: rows.length };
}

function getData(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!regNo) return fail('registerNumber is required.');

  const mine = function (r) { return clean(r['Register Number']).toUpperCase() === regNo; };
  return {
    success: true,
    students: readRecords(SHEETS.students).filter(mine),
    attendance: readRecords(SHEETS.attendance).filter(mine),
    subjects: readRecords(SHEETS.subjects).filter(mine),
    timetable: readRecords(SHEETS.timetable).filter(mine),
    dayChanges: readRecords(SHEETS.dayChanges).filter(mine),
    holidays: readRecords(SHEETS.holidays).filter(mine),
    settings: readRecords(SHEETS.settings).filter(mine)
  };
}

/* ---------- Date-based attendance, schedule changes, holidays and settings ---------- */

const CHANGE_KINDS = ['cancel', 'replace', 'extra', 'pt'];
const HOLIDAY_TYPES = ['College holiday', 'Non-working day'];
const SETTING_KEYS = ['reminders', 'reminderTime', 'reminderEmail'];
const RESET_MINUTES = 15;
const RESET_MAX_ATTEMPTS = 5;
const RESET_MAX_PER_HOUR = 3;
const RESET_BAD_CODE = 'The code is wrong or has expired. Request a new code and try again.';
const AUTH_ACTIONS = { login: loginAccount, register: registerAccount, forgotRequest: forgotRequest, forgotReset: forgotReset };

// Text starting with = + - or @ would be read by Sheets as a formula, so it is stored as plain text.
function safeText(v) {
  const s = clean(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function todayISO() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function sameStudent(row, regNo) {
  return clean(row['Register Number']).toUpperCase() === regNo;
}

function isHolidayFor(regNo, date) {
  return readRecords(SHEETS.holidays).some(function (h) {
    return sameStudent(h, regNo) && String(h['Date']) === date;
  });
}

// Updates the one row that matches the predicate, or appends a new row. Returns true when updated.
function upsertRow(def, predicate, obj) {
  const rows = readRecords(def).filter(predicate);
  if (rows.length) {
    setRowValues(def, rows[rows.length - 1]._row, obj);
    return true;
  }
  appendObject(def, obj);
  return false;
}

// Saves Present/Absent for every period in one request. Existing periods for the date are corrected,
// never duplicated. Holidays and future dates are refused.
function saveAttendanceBatch(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  const student = findStudent(regNo);
  if (!student) return fail('Register number is not registered. Save your profile first.');

  const date = clean(p.attendanceDate);
  if (!isValidDate(date)) return fail('Attendance date must be in YYYY-MM-DD format.');
  if (date > todayISO()) return fail('You cannot mark attendance for a future date.');
  if (isHolidayFor(regNo, date)) return fail('This date is a holiday, so attendance cannot be marked.');

  const semester = clean(p.semester) || clean(student['Semester']);
  if (semester.length > 20) return fail('Semester value is too long.');

  const list = parseList(p.rows);
  if (list === null) return fail('Attendance entries are not valid.');
  if (!list.length) return fail('Nothing to save. Mark at least one period first.');
  if (list.length > 40) return fail('Too many entries in one save (maximum 40).');

  const rows = [];
  const seen = {};
  for (let i = 0; i < list.length; i++) {
    const r = list[i] || {};
    const label = 'Entry ' + (i + 1) + ': ';
    const period = Number(r.period);
    if (!Number.isInteger(period) || period < 1 || period > 99) return fail(label + 'period number is invalid.');
    if (seen[period]) return fail(label + 'period ' + period + ' is listed twice.');
    seen[period] = true;
    const raw = clean(r.status).toLowerCase();
    const status = raw === 'present' ? 'Present' : raw === 'absent' ? 'Absent' : '';
    if (!status) return fail(label + 'status must be Present or Absent.');
    const start = normalizeTime(r.start);
    const end = normalizeTime(r.end);
    if (!start || !end || start >= end) return fail(label + 'start and end times are invalid.');
    const subject = clean(r.subject);
    if (!subject || subject.length > MAX_TEXT_LENGTH) return fail(label + 'subject name is missing or too long.');
    const kind = ['regular', 'extra', 'pt'].indexOf(clean(r.kind)) !== -1 ? clean(r.kind) : 'regular';
    const counted = clean(r.counted) === 'No' ? 'No' : 'Yes';
    rows.push({ period: period, status: status, start: start, end: end, subject: subject, kind: kind, counted: counted });
  }

  const byPeriod = {};
  readRecords(SHEETS.attendance)
    .filter(function (r) { return sameStudent(r, regNo) && String(r['Attendance Date']) === date; })
    .forEach(function (r) { byPeriod[String(r['Period No'])] = r; });

  const name = clean(student['Student Name']);
  let saved = 0;
  let updated = 0;
  rows.forEach(function (r) {
    const fields = {
      'Timestamp': new Date(),
      'Student Name': name,
      'Subject': safeText(r.subject),
      'Attendance Time': "'" + r.start,
      'End Time': "'" + r.end,
      'Status': r.status,
      'Semester': semester,
      'Counted': r.counted,
      'Kind': r.kind
    };
    const prev = byPeriod[String(r.period)];
    if (prev) {
      setRowValues(SHEETS.attendance, prev._row, fields);
      updated++;
    } else {
      fields['Register Number'] = regNo;
      fields['Attendance Date'] = "'" + date;
      fields['Period No'] = r.period;
      fields['Day'] = weekdayName(date);
      appendObject(SHEETS.attendance, fields);
      saved++;
    }
  });
  return { success: true, message: 'Attendance saved.', saved: saved, updated: updated, date: date };
}

// One-day schedule change. ref = "P<n>" for a timetable period (cancel / replace)
// or "X<n>" (50-99) for an extra class or PT/activity period.
function saveDayChange(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!findStudent(regNo)) return fail('Register number is not registered. Save your profile first.');

  const date = clean(p.changeDate);
  if (!isValidDate(date)) return fail('Date must be in YYYY-MM-DD format.');

  const ref = clean(p.ref).toUpperCase();
  if (!/^(P\d{1,2}|X\d{2})$/.test(ref)) return fail('Period reference is invalid.');
  const refNo = Number(ref.slice(1));
  const isExtra = ref.charAt(0) === 'X';
  if (isExtra && (refNo < 50 || refNo > 99)) return fail('Extra period numbers must be 50 to 99.');
  if (!isExtra && (refNo < 1 || refNo > MAX_PERIODS)) return fail('Period number is invalid.');

  const change = clean(p.change).toLowerCase();
  if (CHANGE_KINDS.indexOf(change) === -1) return fail('Change type is invalid.');
  if (isExtra && (change === 'cancel' || change === 'replace')) return fail('Only timetable periods can be cancelled or replaced.');
  if (!isExtra && (change === 'extra' || change === 'pt')) return fail('Extra and PT periods need an X reference.');

  const subject = clean(p.subject);
  const teacher = clean(p.teacher);
  const note = clean(p.note);
  if (subject.length > MAX_TEXT_LENGTH || teacher.length > MAX_TEXT_LENGTH || note.length > MAX_TEXT_LENGTH) {
    return fail('Subject, teacher and note must be ' + MAX_TEXT_LENGTH + ' characters or fewer.');
  }
  if ((change === 'extra' || change === 'pt') && !subject) return fail('Enter a subject or activity name.');

  const timingGiven = clean(p.start) !== '' || clean(p.end) !== '';
  if (change === 'replace' && !subject && !teacher && !timingGiven) {
    return fail('Change at least one of subject, teacher or timing.');
  }

  let start = '';
  let end = '';
  if (change === 'extra' || change === 'pt' || (change === 'replace' && timingGiven)) {
    start = normalizeTime(p.start);
    end = normalizeTime(p.end);
    if (!start || !end || start >= end) return fail('Enter a valid start and end time (end must be after start).');
  }

  const counted = change === 'cancel' ? 'No' : (clean(p.counted) === 'No' ? 'No' : 'Yes');
  upsertRow(SHEETS.dayChanges,
    function (r) { return sameStudent(r, regNo) && String(r['Date']) === date && clean(r['Ref']).toUpperCase() === ref; },
    {
      'Timestamp': new Date(),
      'Register Number': regNo,
      'Date': "'" + date,
      'Ref': ref,
      'Change': change,
      'Subject': safeText(subject),
      'Teacher': safeText(teacher),
      'Start Time': start ? "'" + start : '',
      'End Time': end ? "'" + end : '',
      'Counted': counted,
      'Note': safeText(note)
    });
  return { success: true, message: 'Schedule change saved for ' + date + '.' };
}

function deleteDayChange(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!findStudent(regNo)) return fail('Register number is not registered.');
  const date = clean(p.changeDate);
  if (!isValidDate(date)) return fail('Date must be in YYYY-MM-DD format.');
  const ref = clean(p.ref).toUpperCase();
  if (!/^(P\d{1,2}|X\d{2})$/.test(ref)) return fail('Period reference is invalid.');

  if (ref.charAt(0) === 'X') {
    const used = readRecords(SHEETS.attendance).some(function (r) {
      return sameStudent(r, regNo) && String(r['Attendance Date']) === date && String(r['Period No']) === ref.slice(1);
    });
    if (used) return fail('Attendance is already saved for this extra period. Correct it in Mark Attendance first.');
  }

  deleteRowsWhere(SHEETS.dayChanges, function (r) {
    return sameStudent(r, regNo) && String(r['Date']) === date && clean(r['Ref']).toUpperCase() === ref;
  });
  return { success: true, message: 'Schedule change removed.' };
}

function saveHoliday(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!findStudent(regNo)) return fail('Register number is not registered. Save your profile first.');
  const date = clean(p.holidayDate);
  if (!isValidDate(date)) return fail('Date must be in YYYY-MM-DD format.');
  const type = clean(p.type);
  if (HOLIDAY_TYPES.indexOf(type) === -1) return fail('Choose College holiday or Non-working day.');
  const note = clean(p.note);
  if (note.length > MAX_TEXT_LENGTH) return fail('Note must be ' + MAX_TEXT_LENGTH + ' characters or fewer.');
  upsertRow(SHEETS.holidays,
    function (r) { return sameStudent(r, regNo) && String(r['Date']) === date; },
    {
      'Timestamp': new Date(),
      'Register Number': regNo,
      'Date': "'" + date,
      'Type': type,
      'Note': safeText(note)
    });
  return { success: true, message: type + ' saved for ' + date + '.' };
}

function deleteHoliday(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!findStudent(regNo)) return fail('Register number is not registered.');
  const date = clean(p.holidayDate);
  if (!isValidDate(date)) return fail('Date must be in YYYY-MM-DD format.');
  deleteRowsWhere(SHEETS.holidays, function (r) { return sameStudent(r, regNo) && String(r['Date']) === date; });
  return { success: true, message: 'Holiday removed.' };
}

// Only the three reminder keys can be set from the app. lastReminderSent is written by the trigger only.
function saveSettings(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!findStudent(regNo)) return fail('Register number is not registered. Save your profile first.');
  let input;
  try {
    input = JSON.parse(clean(p.settings) || '{}');
  } catch (err) {
    return fail('Settings are not valid.');
  }
  if (!input || typeof input !== 'object' || Array.isArray(input)) return fail('Settings are not valid.');

  const values = {};
  if (input.reminders !== undefined) {
    if (['on', 'off'].indexOf(input.reminders) === -1) return fail('Reminders must be on or off.');
    values.reminders = input.reminders;
  }
  if (input.reminderEmail !== undefined) {
    if (['on', 'off'].indexOf(input.reminderEmail) === -1) return fail('Email reminders must be on or off.');
    values.reminderEmail = input.reminderEmail;
  }
  if (input.reminderTime !== undefined) {
    const t = normalizeTime(input.reminderTime);
    if (!t) return fail('Reminder time must be HH:MM (24-hour).');
    values.reminderTime = t;
  }
  SETTING_KEYS.forEach(function (key) {
    if (values[key] === undefined) return;
    upsertRow(SHEETS.settings,
      function (r) { return sameStudent(r, regNo) && clean(r['Setting']) === key; },
      {
        'Timestamp': new Date(),
        'Register Number': regNo,
        'Setting': key,
        'Value': values[key] === values.reminderTime ? "'" + values[key] : values[key]
      });
  });
  return { success: true, message: 'Reminder settings saved.' };
}

/* ---------- Reminder emails (time-driven trigger) ---------- */

// Runs every hour. Emails a student only when: email reminders are on, the current hour matches
// their reminder time, they have not been emailed today, they have marked nothing today and it is not
// a holiday or Sunday. Each student gets at most one email per day.
function sendDailyReminders() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const hour = Utilities.formatDate(now, tz, 'HH');
  if (weekdayName(today) === 'Sunday') return;

  const byReg = {};
  readRecords(SHEETS.settings).forEach(function (s) {
    const reg = clean(s['Register Number']).toUpperCase();
    byReg[reg] = byReg[reg] || {};
    byReg[reg][clean(s['Setting'])] = clean(s['Value']);
  });
  const marked = {};
  readRecords(SHEETS.attendance).forEach(function (r) {
    if (String(r['Attendance Date']) === today) marked[clean(r['Register Number']).toUpperCase()] = true;
  });
  const onHoliday = {};
  readRecords(SHEETS.holidays).forEach(function (h) {
    if (String(h['Date']) === today) onHoliday[clean(h['Register Number']).toUpperCase()] = true;
  });

  Object.keys(byReg).forEach(function (reg) {
    try {
      const s = byReg[reg];
      if (s.reminderEmail !== 'on') return;
      if (normalizeTime(s.reminderTime).slice(0, 2) !== hour) return;
      if (s.lastReminderSent === today) return;
      if (marked[reg] || onHoliday[reg]) return;
      const student = findStudent(reg);
      const email = student ? clean(student['Email']) : '';
      if (!student || !isValidEmail(email)) return;

      MailApp.sendEmail({
        to: email,
        subject: 'AttendSmart: mark today\'s attendance',
        body: 'Hi ' + clean(student['Student Name']) + ',\n\n' +
          'You have not marked attendance for ' + today + ' yet. Open AttendSmart to mark your periods.\n\n' +
          'To stop these emails, turn off email reminders in My Profile.'
      });
      upsertRow(SHEETS.settings,
        function (r) { return sameStudent(r, reg) && clean(r['Setting']) === 'lastReminderSent'; },
        { 'Timestamp': new Date(), 'Register Number': reg, 'Setting': 'lastReminderSent', 'Value': "'" + today });
    } catch (err) {
      Logger.log('Reminder failed for ' + reg + ': ' + err.message);
    }
  });
}

// Run once from the Apps Script editor (Run > installReminderTrigger) and approve the permissions.
function installReminderTrigger() {
  removeReminderTrigger();
  ScriptApp.newTrigger('sendDailyReminders').timeBased().everyHours(1).create();
  Logger.log('Reminder trigger installed. It checks every hour.');
}

function removeReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyReminders') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- Forgot password: emailed 6-digit code ---------- */

// Always answers the same way, so the response never shows whether a register number or email exists.
function forgotRequest(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!regNo) return fail('Register number is required.');
  const generic = {
    success: true,
    message: 'If this register number has an account with an email on file, a 6-digit code has been sent. It expires in ' + RESET_MINUTES + ' minutes.'
  };
  const acc = findAccount(regNo);
  const student = findStudent(regNo);
  const email = student ? clean(student['Email']) : '';
  if (!acc || !isValidEmail(email)) return generic;

  const hourAgo = Date.now() - 60 * 60000;
  const codes = readRecords(SHEETS.resetCodes).filter(function (r) { return sameStudent(r, regNo); });
  const recent = codes.filter(function (r) { return Number(r['Expires']) - RESET_MINUTES * 60000 > hourAgo; });
  if (recent.length >= RESET_MAX_PER_HOUR) return generic;

  codes.filter(function (r) { return clean(r['Used']) !== 'Yes'; })
    .forEach(function (r) { setRowValues(SHEETS.resetCodes, r._row, { 'Used': 'Yes' }); });

  const code = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16) % 1000000)).slice(-6);
  const salt = newToken().slice(0, 16);
  appendObject(SHEETS.resetCodes, {
    'Register Number': regNo,
    'Salt': salt,
    'Code Hash': hashPassword(salt, code),
    'Expires': Date.now() + RESET_MINUTES * 60000,
    'Attempts': 0,
    'Used': 'No'
  });
  MailApp.sendEmail({
    to: email,
    subject: 'AttendSmart password reset code',
    body: 'Your AttendSmart password reset code is: ' + code + '\n\n' +
      'It expires in ' + RESET_MINUTES + ' minutes and can be used once.\n' +
      'If you did not ask for this, ignore this email. Never share this code.'
  });
  return generic;
}

// Checks the newest unused code. Wrong codes count as attempts; after RESET_MAX_ATTEMPTS the code is dead.
function forgotReset(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  const code = clean(p.code);
  const newPassword = String(p.newPassword || '');
  if (!regNo || !/^\d{6}$/.test(code)) return fail(RESET_BAD_CODE);
  if (newPassword.length < MIN_PASSWORD) return fail('Password must be at least ' + MIN_PASSWORD + ' characters.');

  const open = readRecords(SHEETS.resetCodes).filter(function (r) {
    return sameStudent(r, regNo) && clean(r['Used']) !== 'Yes';
  });
  const latest = open[open.length - 1];
  if (!latest || Number(latest['Expires']) < Date.now() || Number(latest['Attempts']) >= RESET_MAX_ATTEMPTS) {
    return fail(RESET_BAD_CODE);
  }
  if (hashPassword(clean(latest['Salt']), code) !== clean(latest['Code Hash'])) {
    setRowValues(SHEETS.resetCodes, latest._row, { 'Attempts': Number(latest['Attempts']) + 1 });
    return fail(RESET_BAD_CODE);
  }

  setRowValues(SHEETS.resetCodes, latest._row, { 'Used': 'Yes' });
  const salt = newToken().slice(0, 16);
  appendObject(SHEETS.accounts, {
    'Register Number': regNo,
    'Salt': salt,
    'Password Hash': hashPassword(salt, newPassword)
  });
  deleteRowsWhere(SHEETS.sessions, function (r) { return sameStudent(r, regNo); }); // log out everywhere
  PropertiesService.getScriptProperties().deleteProperty('fail_' + regNo);
  return { success: true, message: 'Password changed. Log in with your new password.' };
}

/* ---------- Responses log ---------- */

function logResponse(action, regNo, params, result) {
  try {
    const copy = Object.assign({}, params);
    delete copy.callback;
    delete copy.password;
    delete copy.newPassword;
    delete copy.code;
    delete copy.token;
    delete copy.requestId;
    const details = JSON.stringify(copy).slice(0, 500);
    const outcome = result && result.success ? 'OK' : (result && result.error) || 'Failed';
    appendObject(SHEETS.responses, {
      'Action': action,
      'Register Number': regNo,
      'Details': details,
      'Result': outcome
    });
  } catch (err) {
    // Logging must never block the request.
  }
}

/* ---------- Setup (run once from the Apps Script editor) ---------- */

function setupSheets() {
  Object.keys(SHEETS).forEach(function (key) { getSheet(SHEETS[key]); });
  Logger.log('All sheets are ready.');
}

/* ---------- Timetable upload and automatic reading ---------- */

const UPLOAD_FOLDER_NAME = 'AttendSmart Timetable Uploads';
const UPLOAD_TYPES = {
  'image/jpeg': '.jpg',
  'image/png': '.png',
  'application/pdf': '.pdf',
  'application/msword': '.doc',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx'
};
const WORD_TYPES = ['application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'];
// Gemini model used to read timetables. Change it here if Google retires it.
const GEMINI_MODEL = 'gemini-2.5-flash';
const MAX_UPLOAD_BASE64 = 7000000; // about 5 MB file
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function weekdayName(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return '';
  return WEEKDAY_NAMES[new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getDay()];
}

function validUploadId(id) {
  return /^[A-Za-z0-9_-]{6,60}$/.test(id);
}

function getUploadFolder() {
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER_NAME);
}

// Called from doPost. The browser sends the file in the request body and it is stored in Drive.
function saveUpload(body) {
  const reg = clean(body.registerNumber).toUpperCase();
  if (!reg || !findStudent(reg)) return fail('Register number is not registered. Save your profile first.');
  const id = clean(body.uploadId);
  if (!validUploadId(id)) return fail('Upload ID is invalid.');
  const mime = clean(body.mimeType);
  const ext = UPLOAD_TYPES[mime];
  if (!ext) return fail('Unsupported file type. Use JPG, PNG, PDF, DOC or DOCX.');
  const data = String(body.data || '');
  if (!data) return fail('The file is empty.');
  if (data.length > MAX_UPLOAD_BASE64) return fail('The file is larger than 5 MB.');
  const bytes = Utilities.base64Decode(data);
  getUploadFolder().createFile(Utilities.newBlob(bytes, mime, 'TT_' + reg + '_' + id + ext));
  return { success: true, message: 'Upload received.' };
}

function findUploadFile(reg, id) {
  const prefix = 'TT_' + reg + '_' + id + '.';
  const it = getUploadFolder().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(prefix) === 0) return f;
  }
  return null;
}

// Word files are converted to PDF through Google Drive (needs the Drive advanced service).
function wordToPdf(blob) {
  const temp = Drive.Files.create({ name: 'tmp-timetable', mimeType: 'application/vnd.google-apps.document' }, blob);
  try {
    return DriveApp.getFileById(temp.id).getAs('application/pdf');
  } finally {
    DriveApp.getFileById(temp.id).setTrashed(true);
  }
}

function callGemini(key, blob) {
  const prompt = 'You read a college class timetable from the attached file. Return ONLY JSON in this exact shape: ' +
    '{"days":[{"day":"Monday","periods":[{"type":"class","start":"09:30","end":"10:30","subject":"Data Structures"},' +
    '{"type":"break","start":"11:00","end":"11:15","subject":""}]}]}. ' +
    'Include only Monday to Saturday. Use 24-hour HH:MM times. Use type "break" for tea, lunch or other breaks. ' +
    'Copy subject names as printed. Do not invent periods, subjects or times; leave out anything you cannot read clearly.';
  const body = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } }
      ]
    }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0 }
  };
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('The reading service returned error ' + res.getResponseCode() + '. Check GEMINI_API_KEY and GEMINI_MODEL.');
  }
  const json = JSON.parse(res.getContentText());
  const cand = json.candidates && json.candidates[0];
  const part = cand && cand.content && cand.content.parts && cand.content.parts[0];
  if (!part || !part.text) throw new Error('The reading service returned no timetable data.');
  return JSON.parse(part.text);
}

// Keeps only valid Monday-Saturday periods, sorted, with no overlaps.
function normalizeExtraction(obj) {
  const out = {};
  const list = obj && Array.isArray(obj.days) ? obj.days : [];
  list.forEach(function (d) {
    const day = DAYS.find(x => x.toLowerCase() === clean(d && d.day).toLowerCase());
    if (!day || !d || !Array.isArray(d.periods)) return;
    const rows = [];
    d.periods.forEach(function (r) {
      if (!r) return;
      const start = normalizeTime(r.start);
      const end = normalizeTime(r.end);
      if (!start || !end || start >= end) return;
      const type = clean(r.type).toLowerCase() === 'break' ? 'break' : 'class';
      const subject = type === 'class' ? clean(r.subject).slice(0, MAX_TEXT_LENGTH) : '';
      if (type === 'class' && !subject) return;
      rows.push({ type: type, start: start, end: end, subject: subject });
    });
    rows.sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; });
    const kept = [];
    rows.forEach(function (r) {
      const last = kept[kept.length - 1];
      if (!last || r.start >= last.end) kept.push(r);
    });
    if (kept.length) out[day] = kept.slice(0, MAX_PERIODS);
  });
  return out;
}

// Reads the uploaded file. Called by the page, which retries while the upload is still arriving.
function extractTimetable(p) {
  const reg = clean(p.registerNumber).toUpperCase();
  if (!reg || !findStudent(reg)) return fail('Register number is not registered.');
  const id = clean(p.uploadId);
  if (!validUploadId(id)) return fail('Upload ID is invalid.');
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    return fail('Automatic reading is not set up. Add GEMINI_API_KEY under Project Settings > Script Properties, or use manual setup.');
  }

  const file = findUploadFile(reg, id);
  if (!file) return { success: false, pending: true, error: 'The upload has not arrived yet.' };

  let blob = file.getBlob();
  if (WORD_TYPES.indexOf(blob.getContentType()) !== -1) blob = wordToPdf(blob);

  const days = normalizeExtraction(callGemini(key, blob));
  if (!Object.keys(days).length) {
    return fail('No timetable could be read from the file. Try a clearer image or PDF, or use manual setup.');
  }
  file.setTrashed(true);
  return { success: true, days: days, message: 'Timetable read. Check every period before saving.' };
}

/* ---------- Accounts, passwords and sessions ---------- */
const SESSION_DAYS = 30;
const MIN_PASSWORD = 8;
const LOGIN_MAX_FAILS = 5;
const LOGIN_LOCK_MINUTES = 15;

function hashPassword(salt, password) {
  let digest = salt + ':' + password;
  for (let i = 0; i < 2000; i++) {
    digest = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digest));
  }
  return digest;
}

function findAccount(regNo) {
  const rows = readRecords(SHEETS.accounts).filter(function (r) {
    return clean(r['Register Number']).toUpperCase() === regNo;
  });
  return rows.length ? rows[rows.length - 1] : null;
}

function newToken() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase();
}

function tokenHash(token) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
}

// Sessions are kept in the Sessions sheet (not a short-lived cache), so students stay logged in for 30 days.
function issueToken(regNo) {
  const token = newToken();
  appendObject(SHEETS.sessions, {
    'Token Hash': tokenHash(token),
    'Register Number': regNo,
    'Expires': Date.now() + SESSION_DAYS * 86400000
  });
  return token;
}

// Returns the register number for a valid, unexpired session token, otherwise null.
function sessionRegNo(token) {
  const t = clean(token).toLowerCase();
  if (!/^[a-f0-9]{64}$/.test(t)) return null;
  const h = tokenHash(t);
  const rows = readRecords(SHEETS.sessions).filter(function (r) { return clean(r['Token Hash']) === h; });
  if (!rows.length) return null;
  const row = rows[rows.length - 1];
  if (Number(row['Expires']) < Date.now()) return null;
  return clean(row['Register Number']).toUpperCase() || null;
}

function loggedOut() {
  return { success: false, loggedOut: true, error: 'Your session has ended. Please log in again.' };
}

// Runs an action only for the logged-in student. The register number always comes from the session,
// so a student can never read or change another student's rows.
function withAuth(params, fn) {
  const regNo = sessionRegNo(params.token);
  if (!regNo) return loggedOut();
  params.registerNumber = regNo;
  return fn(params);
}

function registerAccount(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  const password = String(p.password || '');
  if (!regNo) return fail('Register number is required.');
  if (password.length < MIN_PASSWORD) return fail('Password must be at least ' + MIN_PASSWORD + ' characters.');
  if (findAccount(regNo)) return fail('This register number already has an account. Please log in.');

  const saved = saveStudent(p); // checks the profile fields first
  if (!saved.success) return saved;

  const salt = newToken().slice(0, 16);
  appendObject(SHEETS.accounts, {
    'Register Number': regNo,
    'Salt': salt,
    'Password Hash': hashPassword(salt, password)
  });
  return { success: true, message: 'Account created.', registerNumber: regNo, token: issueToken(regNo) };
}

function failState(regNo) {
  const raw = PropertiesService.getScriptProperties().getProperty('fail_' + regNo);
  return raw ? JSON.parse(raw) : { n: 0, until: 0 };
}

function loginAccount(p) {
  const regNo = clean(p.registerNumber).toUpperCase();
  if (!regNo) return fail('Register number or password is wrong.');
  const props = PropertiesService.getScriptProperties();
  const st = failState(regNo);
  if (st.until > Date.now()) {
    return fail('Too many wrong attempts. Please wait ' + LOGIN_LOCK_MINUTES + ' minutes and try again.');
  }
  const acc = findAccount(regNo);
  const password = String(p.password || '');
  const ok = !!acc && hashPassword(clean(acc['Salt']), password) === clean(acc['Password Hash']);
  if (!ok) {
    st.n += 1;
    if (st.n >= LOGIN_MAX_FAILS) { st.n = 0; st.until = Date.now() + LOGIN_LOCK_MINUTES * 60000; }
    props.setProperty('fail_' + regNo, JSON.stringify(st));
    return fail('Register number or password is wrong.');
  }
  props.deleteProperty('fail_' + regNo);
  return { success: true, message: 'Logged in.', registerNumber: regNo, token: issueToken(regNo) };
}

// Login and register arrive as a POST (password not in the web address). The result is stored
// briefly and collected by the page through fetchResult.
function handlePost(action, body) {
  const id = clean(body.requestId);
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) return respond(fail('Request ID is invalid.'), {});
  if (action === 'fetchResult') return respond(fail('Use a GET request to collect a result.'), {});
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return respond(fail('Server is busy. Please try again.'), {});
  }
  let result;
  try {
    result = Object.prototype.hasOwnProperty.call(AUTH_ACTIONS, action) ? AUTH_ACTIONS[action](body) : runAction(body);
  } catch (err) {
    result = fail('Server error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
  if (action !== 'getData') logResponse(action, clean(body.registerNumber).toUpperCase(), body, result);
  storeResult(id, result);
  return respond({ success: true, stored: true }, {});
}

function storeResult(id, result) {
  let text = JSON.stringify({ at: Date.now(), result: result });
  if (text.length > 95000) {
    text = JSON.stringify({ at: Date.now(), result: fail('The result is too large to return. Please reload the page.') });
  }
  CacheService.getScriptCache().put('res_' + id, text, 300);
}

function fetchResult(p) {
  const id = clean(p.requestId);
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) return fail('Request ID is invalid.');
  const cache = CacheService.getScriptCache();
  const key = 'res_' + id;
  const raw = cache.get(key);
  if (!raw) return { success: false, pending: true, error: 'Still working.' };
  cache.remove(key);
  const stored = JSON.parse(raw);
  if (Date.now() - stored.at > 300000) return fail('This request expired. Please try again.');
  return stored.result;
}
