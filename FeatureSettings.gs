const LMS_FEATURES = [
  { key: 'FEATURE_ATTENDANCE', label: 'Attendance' },
  { key: 'FEATURE_REMINDERS', label: 'Absence reminders and follow-ups' },
  { key: 'FEATURE_PERFORMANCE', label: 'Performance tracking' },
  { key: 'FEATURE_STUDENT_IMPORT', label: 'Student add / import / paste' },
  { key: 'FEATURE_REPORTS', label: 'Daily summary reports' },
  { key: 'FEATURE_INSIGHTS', label: 'Insights and analytics' },
  { key: 'FEATURE_BACKUPS', label: 'Google Drive backups' }
];

/** Adds a spreadsheet menu so an administrator can manage optional modules. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LMS Controls')
    .addItem('Feature settings', 'showFeatureSettings')
    .addToUi();
}

function showFeatureSettings() {
  const html = HtmlService.createHtmlOutputFromFile('FeatureSettings')
    .setTitle('LMS feature settings');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Returns toggle states only after an administrator authenticates with the app access code. */
function getFeatureSettings(accessCode) {
  const auth = verifyAccessCode(accessCode);
  if (!auth.success || auth.role !== 'admin') throw new Error('Admin access code required.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) throw new Error('Run setup() before configuring features.');
  const values = sheet.getDataRange().getValues();
  const stored = {};
  values.slice(1).forEach(row => stored[String(row[0])] = String(row[1]).toLowerCase());
  return LMS_FEATURES.map(feature => ({
    key: feature.key,
    label: feature.label,
    enabled: stored[feature.key] !== 'false'
  }));
}

/** Saves module toggle states; omitted keys are left unchanged. */
function saveFeatureSettings(accessCode, updates) {
  const auth = verifyAccessCode(accessCode);
  if (!auth.success || auth.role !== 'admin') throw new Error('Admin access code required.');
  if (!updates || typeof updates !== 'object' || Array.isArray(updates)) throw new Error('Invalid settings payload.');
  const allowed = new Set(LMS_FEATURES.map(feature => feature.key));
  Object.keys(updates).forEach(key => {
    if (!allowed.has(key) || typeof updates[key] !== 'boolean') throw new Error('Invalid feature setting: ' + key);
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) throw new Error('Run setup() before configuring features.');
  const rows = sheet.getDataRange().getValues();
  const rowByKey = new Map();
  rows.slice(1).forEach((row, index) => rowByKey.set(String(row[0]), index + 2));
  Object.keys(updates).forEach(key => {
    if (rowByKey.has(key)) sheet.getRange(rowByKey.get(key), 2).setValue(updates[key] ? 'TRUE' : 'FALSE');
    else sheet.appendRow([key, updates[key] ? 'TRUE' : 'FALSE']);
  });
  SpreadsheetApp.flush();
  return { success: true, savedAt: new Date().toISOString() };
}

/** Internal helper for future modules to enforce a feature switch on the server. */
function isLmsFeatureEnabled_(featureKey) {
  const allowed = new Set(LMS_FEATURES.map(feature => feature.key));
  if (!allowed.has(featureKey)) throw new Error('Unknown feature key.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sheet) return true;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === featureKey) return String(rows[i][1]).toLowerCase() !== 'false';
  }
  return true;
}
