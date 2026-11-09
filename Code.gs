const SHEETS = {
  settings: 'Settings',
  students: 'Students',
  attendance: 'Attendance',
  calls: 'CallLogs',
  reports: 'DailyReports'
};
const SESSIONS = ['S1', 'S2', 'S3'];
const SESSION_TTL_SECONDS = 21600;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FEA Attendance & Retention Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Creates missing application sheets and initial settings without overwriting existing data. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ensureSheet_(ss, SHEETS.settings, ['Key', 'Value']);
  ensureSheet_(ss, SHEETS.students, ['ID', 'Name', 'Phone', 'Active', 'Expected return', 'Session', 'Historical Absences']);
  ensureSheet_(ss, SHEETS.attendance, ['Date', 'Student ID', 'Name', 'Session', 'Status']);
  ensureSheet_(ss, SHEETS.calls, ['Timestamp', 'Student ID', 'Name', 'Reason', 'Action', 'Outcome', 'FollowUp', 'CardStyle', 'CopiedText']);
  ensureSheet_(ss, SHEETS.reports, ['Date', 'S1_Lesson', 'S2_Lesson', 'S3_Lesson', 'ReportText']);

  const existing = settings.getLastRow() > 1
    ? settings.getRange(2, 1, settings.getLastRow() - 1, 2).getValues()
    : [];
  const keys = new Set(existing.map(row => String(row[0])));
  const defaults = [
    ['ADMIN_CODE', 'FEA-ADM-' + randomCode_()],
    ['STAFF_CODE', 'FEA-STF-' + randomCode_()],
    ['VIEW_CODE', 'FEA-VIW-' + randomCode_()],
    ['TOTAL_CLASSES_HELD', '0']
  ];
  defaults.forEach(row => {
    if (!keys.has(row[0])) settings.appendRow(row);
  });
  Logger.log('Setup complete. Access codes are stored in the Settings sheet; keep that sheet restricted to administrators.');
  return { success: true, message: 'Setup complete. Check the Settings sheet for access codes.' };
}

function verifyAccessCode(code) {
  const input = String(code || '').trim();
  if (!input) return { success: false };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.settings);
  if (!sheet) throw new Error('Run setup() before signing in.');
  const values = sheet.getDataRange().getValues();
  const codeRoles = { ADMIN_CODE: 'admin', STAFF_CODE: 'staff', VIEW_CODE: 'view' };
  let role = null;
  values.slice(1).forEach(row => {
    const expectedRole = codeRoles[String(row[0])];
    if (expectedRole && String(row[1]).trim() === input) role = expectedRole;
  });
  if (!role) return { success: false };
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('attendance_session_' + token, role, SESSION_TTL_SECONDS);
  return { success: true, role: role, token: token, expiresInSeconds: SESSION_TTL_SECONDS };
}

function getAppData(token, selectedDate, selectedSession) {
  requireRole_(token, ['admin', 'staff', 'view']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentsSheet = requireSheet_(ss, SHEETS.students);
  const attendanceSheet = requireSheet_(ss, SHEETS.attendance);
  const settingsSheet = requireSheet_(ss, SHEETS.settings);
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const date = isDateString_(selectedDate) ? selectedDate : today;
  const session = SESSIONS.includes(String(selectedSession).toUpperCase()) ? String(selectedSession).toUpperCase() : 'S1';
  const settingRows = getRows_(settingsSheet);
  const settings = {};
  settingRows.forEach(row => settings[String(row[0])] = row[1]);
  const attendanceRows = getRows_(attendanceSheet);

  const selectedAttendance = {};
  attendanceRows.forEach(row => {
    const rowDate = normalizeDate_(row[0], tz);
    if (rowDate === date && String(row[3]).toUpperCase() === session) {
      selectedAttendance[String(row[1])] = String(row[4]);
    }
  });

  const rawStudents = getRows_(studentsSheet);
  const students = rawStudents.filter(row => row[0] !== '' && row[0] != null).map(row => {
    const id = String(row[0]);
    const studentSession = String(row[5] || '').trim().toUpperCase();
    const history = attendanceRows
      .filter(record => String(record[1]) === id && String(record[3]).toUpperCase() === studentSession && normalizeDate_(record[0], tz))
      .sort((a, b) => normalizeDate_(b[0], tz).localeCompare(normalizeDate_(a[0], tz)));
    let consecutiveAbsences = 0;
    for (const record of history) {
      if (String(record[4]).toLowerCase() === 'absent') consecutiveAbsences++;
      else break;
    }
    const expectedReturn = normalizeDate_(row[4], tz);
    const active = String(row[3]).trim().toLowerCase() === 'yes';
    return {
      id: id,
      name: String(row[1] || ''),
      phone: String(row[2] || ''),
      active: active,
      expectedReturn: expectedReturn,
      isHiddenByReturn: Boolean(expectedReturn && expectedReturn > today),
      session: studentSession,
      historicalAbsences: Math.max(0, Number(row[6]) || 0),
      totalAbsences: Math.max(0, Number(row[6]) || 0) + history.filter(record => String(record[4]).toLowerCase() === 'absent').length,
      consecutiveAbsences: consecutiveAbsences,
      attendanceStatus: selectedAttendance[id] || ''
    };
  });

  return {
    students: students,
    today: today,
    selectedDate: date,
    selectedSession: session,
    totalClassesHeld: Math.max(0, Number(settings.TOTAL_CLASSES_HELD) || 0)
  };
}

function saveBulkAttendance(token, dateStr, sessionVal, absentIds) {
  requireRole_(token, ['admin', 'staff']);
  if (!isDateString_(dateStr)) throw new Error('Choose a valid attendance date.');
  const session = String(sessionVal || '').toUpperCase();
  if (!SESSIONS.includes(session)) throw new Error('Invalid session.');
  if (!Array.isArray(absentIds)) throw new Error('Invalid absent-student list.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attendanceSheet = requireSheet_(ss, SHEETS.attendance);
  const studentRows = getRows_(requireSheet_(ss, SHEETS.students));
  const sessionStudents = studentRows.filter(row =>
    String(row[3]).trim().toLowerCase() === 'yes' && String(row[5]).trim().toUpperCase() === session
  );
  const validIds = new Set(sessionStudents.map(row => String(row[0])));
  const absentSet = new Set(absentIds.map(String));
  absentSet.forEach(id => { if (!validIds.has(id)) throw new Error('An absent student is not in the selected active session.'); });

  const data = attendanceSheet.getDataRange().getValues();
  const header = data.length ? data[0] : ['Date', 'Student ID', 'Name', 'Session', 'Status'];
  const kept = [header];
  data.slice(1).forEach(row => {
    if (normalizeDate_(row[0], Session.getScriptTimeZone()) !== dateStr || String(row[3]).toUpperCase() !== session) kept.push(row);
  });
  const attendanceDate = Utilities.parseDate(dateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sessionStudents.forEach(student => {
    const id = String(student[0]);
    kept.push([attendanceDate, student[0], student[1], session, absentSet.has(id) ? 'Absent' : 'Present']);
  });
  attendanceSheet.clearContents();
  if (kept.length) attendanceSheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);
  return { success: true, count: sessionStudents.length };
}

function updateExpectedReturn(token, studentId, returnDateStr) {
  requireRole_(token, ['admin', 'staff']);
  if (returnDateStr && !isDateString_(returnDateStr)) throw new Error('Choose a valid return date.');
  const sheet = requireSheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEETS.students);
  const rows = getRows_(sheet);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(studentId)) {
      sheet.getRange(i + 2, 5).setValue(returnDateStr ? Utilities.parseDate(returnDateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '');
      return { success: true };
    }
  }
  throw new Error('Student not found.');
}

function saveCallRecord(token, payload) {
  requireRole_(token, ['admin', 'staff']);
  if (!payload || !payload.studentId) throw new Error('Invalid call record.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const student = findStudent_(ss, payload.studentId);
  requireSheet_(ss, SHEETS.calls).appendRow([
    new Date(), student.id, student.name,
    cleanText_(payload.reason), cleanText_(payload.action), cleanText_(payload.outcome),
    cleanText_(payload.followUp), cleanText_(payload.cardStyle), cleanText_(payload.copiedText)
  ]);
  if (payload.expectedReturn) updateExpectedReturn(token, student.id, payload.expectedReturn);
  return { success: true };
}

function generateLessonReport(token, lessons) {
  requireRole_(token, ['admin', 'staff', 'view']);
  lessons = lessons || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const students = getRows_(requireSheet_(ss, SHEETS.students)).filter(row => String(row[3]).trim().toLowerCase() === 'yes');
  const attendance = getRows_(requireSheet_(ss, SHEETS.attendance));
  const absentBySession = {};
  SESSIONS.forEach(session => absentBySession[session] = new Set());
  attendance.forEach(row => {
    if (normalizeDate_(row[0], tz) === today && String(row[4]).toLowerCase() === 'absent') {
      const session = String(row[3]).toUpperCase();
      if (absentBySession[session]) absentBySession[session].add(String(row[1]));
    }
  });
  let text = '📊 *DAILY BRANCH REPORT - FEA*\n📅 Date: ' + today + '\n━━━━━━━━━━━━━━━━━━━━━\n';
  SESSIONS.forEach(session => {
    const list = students.filter(row => String(row[5]).trim().toUpperCase() === session);
    const absent = list.filter(row => absentBySession[session].has(String(row[0])));
    const present = list.length - absent.length;
    const lesson = cleanText_(lessons[session]) || 'N/A';
    text += '\n*' + session + ' (Lesson: ' + lesson + ')*\nAttendance: ' + present + '/' + list.length + '\n';
    text += absent.length ? 'Absent (' + absent.length + '): ' + absent.map(row => String(row[1])).join(', ') + '\n' : 'Absent: None (100%)\n';
  });
  requireSheet_(ss, SHEETS.reports).appendRow([new Date(), cleanText_(lessons.S1), cleanText_(lessons.S2), cleanText_(lessons.S3), text]);
  return { reportText: text };
}

function saveOpeningBalance(token, totalClasses, studentAbsences) {
  requireRole_(token, ['admin']);
  const classes = Number(totalClasses);
  if (!Number.isInteger(classes) || classes < 0) throw new Error('Total classes must be a non-negative whole number.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = requireSheet_(ss, SHEETS.settings);
  const rows = settings.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === 'TOTAL_CLASSES_HELD') {
      settings.getRange(i + 1, 2).setValue(classes);
      found = true;
      break;
    }
  }
  if (!found) settings.appendRow(['TOTAL_CLASSES_HELD', classes]);

  const absences = studentAbsences || {};
  const sheet = requireSheet_(ss, SHEETS.students);
  const students = getRows_(sheet);
  students.forEach((row, i) => {
    const key = String(row[0]);
    if (Object.prototype.hasOwnProperty.call(absences, key)) {
      const value = Number(absences[key]);
      if (!Number.isInteger(value) || value < 0) throw new Error('Historical absences must be non-negative whole numbers.');
      sheet.getRange(i + 2, 7).setValue(value);
    }
  });
  return { success: true };
}

function makeBackupCopy(token) {
  requireRole_(token, ['admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const name = ss.getName() + '_Backup_' + timestamp;
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(name);
  return { success: true, fileName: name, url: copy.getUrl() };
}

function requireRole_(token, allowedRoles) {
  const role = token && CacheService.getScriptCache().get('attendance_session_' + String(token));
  if (!role || allowedRoles.indexOf(role) === -1) throw new Error('Your session expired or you do not have permission. Please sign in again.');
  return role;
}
function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}
function requireSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setup() first.');
  return sheet;
}
function getRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  return values.length > 1 ? values.slice(1) : [];
}
function normalizeDate_(value, timezone) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? '' : Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd');
}
function isDateString_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ''))) return false;
  const parsed = new Date(String(value) + 'T00:00:00Z');
  return !isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}
function randomCode_() { return Math.floor(1000 + Math.random() * 9000); }
function cleanText_(value) { return String(value == null ? '' : value).trim().slice(0, 5000); }
function findStudent_(ss, id) {
  const row = getRows_(requireSheet_(ss, SHEETS.students)).find(item => String(item[0]) === String(id));
  if (!row) throw new Error('Student not found.');
  return { id: String(row[0]), name: String(row[1] || '') };
}
