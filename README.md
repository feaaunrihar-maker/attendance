# attendance

आपके दिए गए ऑपरेशनल प्लान और आर्किटेक्चर के विश्लेषण के आधार पर, यह नया ब्लूप्रिंट आपके पिछले सिस्टम से काफी अधिक उन्नत और व्यवस्थित है।

### योजना का मुख्य विश्लेषण (Key Analysis):

1. **डेटाबेस स्ट्रक्चर**:
* `Students` शीट में हेडर: `ID` | `Name` | `Phone` | `Active` (Yes/No) | `Expected return` | `Session` (S1/S2/S3).
* मूल `Attendance` शीट सुरक्षित रखी जाएगी।
* रोल-बेस्ड एक्सेस कोड्स (Admin, Staff, View-only) के लिए `Settings` शीट का उपयोग।


2. **नई कार्यप्रणाली (New Core Features to Integrate)**:
* **Role-Based Login / Access Codes**: Admin (फुल कंट्रोल), Staff (अटेंडेंस + कॉलिंग), View-only (सिर्फ व्यू)।
* **स्मार्ट री-अपीयरेंस (Expected Return Date)**: यदि कोई छात्र वापसी की तारीख (Expected return) देता है, तो वह उस तारीख तक रिमाइंडर्स में हाइड रहेगा।
* **4-स्टाइल कॉल कार्ड्स और चिप्स**: Reason, Action, Outcome, और Next Follow-up चिप्स के साथ 4 अलग-अलग कार्ड लेआउट और वन-क्लिक कॉपी।
* **लेसन नंबर (L-) के साथ दैनिक रिपोर्ट**: प्रत्येक सेशन (S1, S2, S3) के लिए Lesson Number इनपुट और WhatsApp एक्सपोर्ट।
* **ओपनिंग बैलेंस और ट्रेंड्स (Insights)**: पुराने दिनों के कुल क्लासेज और छात्र-वार ऐतिहासिक अनुपस्थिति को जोड़कर सटीक कुल अटेंडेंस प्रतिशत निकालना।
* **ड्राइव बैकअप यूटिलिटी**: एक क्लिक में पूरे स्प्रेडशीट का टाइमस्टैम्प्ड स्नैपशॉट Google Drive में सेव करना।


---

यहाँ पूरा एकीकृत समाधान दिया गया है। इसे सिंगल-फाइल आर्किटेक्चर (`Code.gs` और `Index.html`) के रूप में तैयार किया गया है।

---

### भाग 1: Google Apps Script बैकएंड (`Code.gs`)

Google Apps Script एडिटर में `Code.gs` का सारा पुराना कोड हटाकर यह पूरा कोड पेस्ट करें:

```javascript
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('FEA Retention & Operations Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. इनिशियलाइजेशन और ऑटो-सेटअप (पहली बार रन करने के लिए)
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Settings शीट सेटअप
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Key', 'Value']);
    settingsSheet.appendRow(['ADMIN_CODE', 'FEA-ADM-' + Math.floor(1000 + Math.random() * 9000)]);
    settingsSheet.appendRow(['STAFF_CODE', 'FEA-STF-' + Math.floor(1000 + Math.random() * 9000)]);
    settingsSheet.appendRow(['VIEW_CODE', 'FEA-VIW-' + Math.floor(1000 + Math.random() * 9000)]);
    settingsSheet.appendRow(['TOTAL_CLASSES_HELD', '0']);
  }

  // Students शीट
  let stuSheet = ss.getSheetByName('Students');
  if (!stuSheet) {
    stuSheet = ss.insertSheet('Students');
    stuSheet.appendRow(['ID', 'Name', 'Phone', 'Active', 'Expected return', 'Session', 'Historical Absences']);
  }

  // Attendance शीट
  let attSheet = ss.getSheetByName('Attendance');
  if (!attSheet) {
    attSheet = ss.insertSheet('Attendance');
    attSheet.appendRow(['Date', 'Student ID', 'Name', 'Session', 'Status']);
  }

  // CallLogs शीट
  let callSheet = ss.getSheetByName('CallLogs');
  if (!callSheet) {
    callSheet = ss.insertSheet('CallLogs');
    callSheet.appendRow(['Timestamp', 'Student ID', 'Name', 'Reason', 'Action', 'Outcome', 'FollowUp', 'CardStyle', 'CopiedText']);
  }

  // DailyReports टैब
  let repSheet = ss.getSheetByName('DailyReports');
  if (!repSheet) {
    repSheet = ss.insertSheet('DailyReports');
    repSheet.appendRow(['Date', 'S1_Lesson', 'S2_Lesson', 'S3_Lesson', 'ReportText']);
  }

  const keys = settingsSheet.getDataRange().getValues();
  Logger.log("=== ACCESS CODES INITIALIZED ===");
  keys.forEach(r => Logger.log(r[0] + ": " + r[1]));
}

// 2. लॉगिन और रोल सत्यापन
function verifyAccessCode(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  
  let role = null;
  data.forEach(r => {
    if (r[0] === 'ADMIN_CODE' && String(r[1]).trim() === code.trim()) role = 'admin';
    if (r[0] === 'STAFF_CODE' && String(r[1]).trim() === code.trim()) role = 'staff';
    if (r[0] === 'VIEW_CODE' && String(r[1]).trim() === code.trim()) role = 'view';
  });
  
  return { success: role !== null, role: role };
}

// 3. मुख्य डेटा लोड करना (डैशबोर्ड, रिमाइंडर्स, और एनालिटिक्स)
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  const setSheet = ss.getSheetByName('Settings');
  
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayDateObj = new Date(todayStr);

  // Settings
  const setData = setSheet.getDataRange().getValues();
  let totalClassesHeld = 0;
  setData.forEach(r => {
    if (r[0] === 'TOTAL_CLASSES_HELD') totalClassesHeld = Number(r[1]) || 0;
  });

  // Students
  const stuRows = stuSheet.getDataRange().getValues();
  stuRows.shift();
  
  // Attendance
  const attRows = attSheet.getDataRange().getValues();
  attRows.shift();

  // आज अनुपस्थित छात्रों का सेट
  const todayAbsentMap = new Map();
  attRows.forEach(r => {
    const rDate = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rDate === todayStr) {
      todayAbsentMap.set(String(r[1]), r[4]); // ID -> Status
    }
  });

  const students = stuRows.map(r => {
    const sId = String(r[0]);
    const name = String(r[1]);
    const phone = String(r[2]);
    const active = String(r[3]).trim().toLowerCase() === 'yes';
    
    // Expected return date handling
    let expectedReturnStr = '';
    let isHiddenByReturn = false;
    if (r[4]) {
      expectedReturnStr = Utilities.formatDate(new Date(r[4]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (new Date(expectedReturnStr) > todayDateObj) {
        isHiddenByReturn = true;
      }
    }

    const session = String(r[5]);
    const histAbsences = Number(r[6]) || 0;

    // Consecutive absence calculation
    const history = attRows
      .filter(row => String(row[1]) === sId)
      .sort((a, b) => new Date(b[0]) - new Date(a[0]));

    let consecutiveAbsences = 0;
    for (let record of history) {
      if (record[4] === 'Absent') consecutiveAbsences++;
      else break;
    }

    const appAbsences = history.filter(row => row[4] === 'Absent').length;
    const totalAbsences = histAbsences + appAbsences;

    return {
      id: sId,
      name: name,
      phone: phone,
      active: active,
      expectedReturn: expectedReturnStr,
      isHiddenByReturn: isHiddenByReturn,
      session: session,
      consecutiveAbsences: consecutiveAbsences,
      totalAbsences: totalAbsences,
      isAbsentToday: todayAbsentMap.get(sId) === 'Absent'
    };
  });

  return {
    students: students,
    today: todayStr,
    totalClassesHeld: totalClassesHeld
  };
}

// 4. अटेंडेंस सेव करना (पुराने रिकॉर्ड्स को ओवरराइट या अपडेट करता है)
function saveBulkAttendance(dateStr, sessionVal, absentIds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attSheet = ss.getSheetByName('Attendance');
  const stuSheet = ss.getSheetByName('Students');
  
  const absentSet = new Set(absentIds.map(String));
  const stuRows = stuSheet.getDataRange().getValues();
  stuRows.shift();
  
  const sessionStudents = stuRows.filter(r => 
    String(r[3]).trim().toLowerCase() === 'yes' && 
    String(r[5]).toUpperCase() === sessionVal.toUpperCase()
  );

  const attData = attSheet.getDataRange().getValues();
  const rowsToKeep = [];
  rowsToKeep.push(attData[0]); // Header

  // उस तारीख और सेशन की पुरानी एंट्रीज को हटाएं ताकि ओवरराइट हो सके
  for (let i = 1; i < attData.length; i++) {
    const row = attData[i];
    const rDate = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const rSess = String(row[3]).toUpperCase();
    if (rDate === dateStr && rSess === sessionVal.toUpperCase()) {
      continue; // Drop existing
    }
    rowsToKeep.push(row);
  }

  // नई एंट्रीज जोड़ें
  sessionStudents.forEach(stu => {
    const sId = String(stu[0]);
    const status = absentSet.has(sId) ? 'Absent' : 'Present';
    rowsToKeep.push([new Date(dateStr), stu[0], stu[1], sessionVal, status]);
  });

  attSheet.clearContents();
  attSheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

  return { success: true };
}

// 5. एक्सपेक्टेड रिटर्न डेट अपडेट करना
function updateExpectedReturn(studentId, returnDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(studentId)) {
      sheet.getRange(i + 1, 5).setValue(returnDateStr ? new Date(returnDateStr) : '');
      break;
    }
  }
  return { success: true };
}

// 6. कॉल लॉग दर्ज करना (चिप्स और 4 कार्ड स्टाइल्स के साथ)
function saveCallRecord(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CallLogs');
  sheet.appendRow([
    new Date(),
    payload.studentId,
    payload.name,
    payload.reason,
    payload.action,
    payload.outcome,
    payload.followUp,
    payload.cardStyle,
    payload.copiedText
  ]);

  if (payload.expectedReturn) {
    updateExpectedReturn(payload.studentId, payload.expectedReturn);
  }
  return { success: true };
}

// 7. दैनिक रिपोर्ट तैयार करना (Lesson L- नंबर्स के साथ)
function generateLessonReport(lessons) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stuSheet = ss.getSheetByName('Students');
  const attSheet = ss.getSheetByName('Attendance');
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const stuRows = stuSheet.getDataRange().getValues();
  stuRows.shift();
  const activeStudents = stuRows.filter(r => String(r[3]).toLowerCase() === 'yes');

  const attRows = attSheet.getDataRange().getValues();
  attRows.shift();
  
  const todayAbsentSet = new Set();
  attRows.forEach(r => {
    const rDate = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rDate === todayStr && r[4] === 'Absent') {
      todayAbsentSet.add(String(r[1]));
    }
  });

  let text = `📊 *DAILY BRANCH REPORT - FEA*\n📅 Date: ${todayStr}\n━━━━━━━━━━━━━━━━━━━━━\n`;

  ['S1', 'S2', 'S3'].forEach(sess => {
    const list = activeStudents.filter(s => String(s[5]).toUpperCase() === sess);
    const absent = list.filter(s => todayAbsentSet.has(String(s[0])));
    const presentCount = list.length - absent.length;
    const lNum = lessons[sess] || 'N/A';

    text += `\n*${sess} (Lesson: ${lNum})*\n`;
    text += `Attendance: ${presentCount}/${list.length}\n`;
    if (absent.length > 0) {
      text += `Absent (${absent.length}): ` + absent.map(s => s[1]).join(', ') + `\n`;
    } else {
      text += `Absent: None (100%)\n`;
    }
  });

  return { reportText: text };
}

// 8. ओपनिंग बैलेंस सेट करना
function saveOpeningBalance(totalClasses, studentAbsences) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName('Settings');
  const stuSheet = ss.getSheetByName('Students');

  // Update Settings
  const setData = setSheet.getDataRange().getValues();
  for (let i = 1; i < setData.length; i++) {
    if (setData[i][0] === 'TOTAL_CLASSES_HELD') {
      setSheet.getRange(i + 1, 2).setValue(totalClasses);
      break;
    }
  }

  // Update Students Historical Absences
  const stuData = stuSheet.getDataRange().getValues();
  for (let i = 1; i < stuData.length; i++) {
    const id = String(stuData[i][0]);
    if (studentAbsences[id] !== undefined) {
      stuSheet.getRange(i + 1, 7).setValue(studentAbsences[id]);
    }
  }

  return { success: true };
}

// 9. Google Drive बैकअप स्नैपशॉट
function makeBackupCopy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const copyName = `${ss.getName()}_Backup_${timeStamp}`;
  
  const copy = file.makeCopy(copyName);
  return { success: true, fileName: copyName, url: copy.getUrl() };
}
```

---

### भाग 2: संपूर्ण फ्रंटएंड यूजर इंटरफेस (`Index.html`)

Apps Script में `Index.html` खोलें और यह संपूर्ण डार्क-थीम इंटरफेस पेस्ट करें:

```html
<!DOCTYPE html>
<html lang="hi" data-bs-theme="dark">
<head>
  <meta charset="UTF-8">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
  <style>
    body { background-color: #0e1117; color: #e6edf3; font-family: system-ui, -apple-system, sans-serif; }
    .card { background-color: #161b22; border: 1px solid #30363d; border-radius: 8px; }
    .chip { cursor: pointer; user-select: none; transition: 0.2s; }
    .chip.active { background-color: #238636 !important; color: white !important; }
    .student-pill { cursor: pointer; user-select: none; transition: all 0.2s; }
    .student-pill.absent { background-color: #da3633 !important; color: white !important; border-color: #f85149 !important; }
    .nav-tabs .nav-link { color: #8b949e; border: none; font-weight: 500; }
    .nav-tabs .nav-link.active { background-color: transparent; color: #58a6ff; border-bottom: 2px solid #58a6ff; }
  </style>
</head>
<body class="p-2 p-md-3">

  <!-- लॉगिन स्क्रीन (Login View) -->
  <div id="loginScreen" class="d-flex align-items-center justify-content-center" style="min-height: 80vh;">
    <div class="card p-4" style="max-width: 400px; width: 100%;">
      <div class="text-center mb-3">
        <i class="bi bi-shield-lock-fill text-primary display-5"></i>
        <h4 class="mt-2 fw-bold">FEA Portal Login</h4>
        <p class="text-secondary small">एक्सेस कोड दर्ज करें (Admin, Staff या View-only)</p>
      </div>
      <div class="mb-3">
        <input type="password" id="accessCodeInput" class="form-control text-center fs-5" placeholder="FEA-XXX-0000">
      </div>
      <button class="btn btn-primary w-100 fw-bold" onclick="handleLogin()">लॉगिन करें</button>
      <div id="loginError" class="text-danger small text-center mt-2" style="display:none;">अमान्य एक्सेस कोड!</div>
    </div>
  </div>

  <!-- मुख्य एप्लीकेशन (Main App) -->
  <div id="mainApp" style="display: none;">
    <!-- हेडर -->
    <div class="d-flex flex-wrap justify-content-between align-items-center mb-3 pb-2 border-bottom border-secondary">
      <div>
        <h5 class="m-0 fw-bold"><i class="bi bi-mortarboard-fill text-primary me-2"></i>FEA Attendance & Retention</h5>
        <small class="text-secondary" id="userRoleBadge"></small>
      </div>
      <div class="d-flex gap-2 align-items-center">
        <span class="badge bg-dark border border-secondary" id="todayLabel"></span>
        <button class="btn btn-outline-danger btn-sm" onclick="logout()"><i class="bi bi-box-arrow-right"></i></button>
      </div>
    </div>

    <!-- नेविगेशन टैब्स -->
    <ul class="nav nav-tabs mb-3" id="appTabs">
      <li class="nav-item"><a class="nav-link active" onclick="switchTab('attTab')"><i class="bi bi-check2-square me-1"></i> Attendance</a></li>
      <li class="nav-item"><a class="nav-link" onclick="switchTab('remindersTab')"><i class="bi bi-bell me-1"></i> Reminders (2+ Days)</a></li>
      <li class="nav-item"><a class="nav-link" onclick="switchTab('reportTab')"><i class="bi bi-file-text me-1"></i> Report (L-)</a></li>
      <li class="nav-item"><a class="nav-link" onclick="switchTab('insightsTab')"><i class="bi bi-graph-up me-1"></i> Insights</a></li>
      <li class="nav-item" id="adminSettingsNav" style="display:none;"><a class="nav-link" onclick="switchTab('settingsTab')"><i class="bi bi-gear me-1"></i> Settings & Backup</a></li>
    </ul>

    <!-- 1. ATTENDANCE TAB -->
    <div id="attTab" class="tab-content-area">
      <div class="card p-3 mb-3">
        <div class="row g-2 align-items-center">
          <div class="col-md-3">
            <label class="small text-secondary">तारीख चुनें</label>
            <input type="date" id="attDate" class="form-control form-control-sm">
          </div>
          <div class="col-md-3">
            <label class="small text-secondary">सेशन चुनें</label>
            <select id="attSession" class="form-select form-select-sm" onchange="renderAttendancePills()">
              <option value="S1">Session 1 (S1)</option>
              <option value="S2">Session 2 (S2)</option>
              <option value="S3">Session 3 (S3)</option>
            </select>
          </div>
          <div class="col-md-6 text-md-end mt-3">
            <button class="btn btn-success btn-sm px-3" onclick="submitAttendance()">
              <i class="bi bi-save me-1"></i> Save Attendance
            </button>
          </div>
        </div>
      </div>

      <div class="card p-3">
        <div class="small text-secondary mb-2">केवल अनुपस्थित (Absent) छात्रों पर टैप करें (वे लाल हो जाएंगे):</div>
        <div class="d-flex flex-wrap gap-2" id="attendancePillsBox"></div>
      </div>
    </div>

    <!-- 2. REMINDERS TAB (2+ DAYS ABSENT) -->
    <div id="remindersTab" class="tab-content-area" style="display:none;">
      <div class="card p-3">
        <h6 class="fw-bold text-danger mb-3"><i class="bi bi-exclamation-triangle-fill me-1"></i> लगातार अनुपस्थित छात्र (2 या अधिक दिन)</h6>
        <div class="table-responsive">
          <table class="table table-dark table-hover align-middle mb-0">
            <thead>
              <tr class="text-secondary small">
                <th>नाम व ID</th>
                <th>सेशन</th>
                <th>फोन</th>
                <th>गायब दिन</th>
                <th>वापसी तिथि (Expected Return)</th>
                <th>एक्शन</th>
              </tr>
            </thead>
            <tbody id="remindersTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- 3. REPORT TAB (LESSON L-) -->
    <div id="reportTab" class="tab-content-area" style="display:none;">
      <div class="card p-3 mb-3">
        <h6 class="fw-bold mb-3"><i class="bi bi-journal-text me-1"></i> लेसन नंबर (L-) दर्ज करें</h6>
        <div class="row g-3">
          <div class="col-md-4">
            <label class="small text-secondary">S1 Lesson</label>
            <input type="text" id="lS1" class="form-control form-control-sm" placeholder="e.g. L-124">
          </div>
          <div class="col-md-4">
            <label class="small text-secondary">S2 Lesson</label>
            <input type="text" id="lS2" class="form-control form-control-sm" placeholder="e.g. L-124">
          </div>
          <div class="col-md-4">
            <label class="small text-secondary">S3 Lesson</label>
            <input type="text" id="lS3" class="form-control form-control-sm" placeholder="e.g. L-124">
          </div>
        </div>
        <button class="btn btn-primary btn-sm mt-3" onclick="generateReport()">
          <i class="bi bi-file-earmark-text me-1"></i> रिपोर्ट तैयार करें
        </button>
      </div>

      <div class="card p-3" id="reportResultBox" style="display:none;">
        <div class="d-flex justify-content-between align-items-center mb-2">
          <span class="small fw-bold">WhatsApp तैयार फॉर्मेट:</span>
          <button class="btn btn-success btn-sm" onclick="copyReport()"><i class="bi bi-clipboard-check me-1"></i> Copy Report</button>
        </div>
        <textarea id="reportTextPreview" class="form-control bg-black text-light font-monospace small" rows="10" readonly></textarea>
      </div>
    </div>

    <!-- 4. INSIGHTS & OPENING BALANCE -->
    <div id="insightsTab" class="tab-content-area" style="display:none;">
      <div class="card p-3 mb-3">
        <h6 class="fw-bold"><i class="bi bi-bar-chart-line me-1"></i> ब्रांच ट्रेंड्स</h6>
        <div class="row g-2 text-center mt-2">
          <div class="col-4">
            <div class="p-2 border border-secondary rounded bg-black">
              <div class="small text-secondary">कुल एनरोल्ड</div>
              <h4 class="m-0 text-primary" id="insightTotalStudents">0</h4>
            </div>
          </div>
          <div class="col-4">
            <div class="p-2 border border-secondary rounded bg-black">
              <div class="small text-secondary">कुल आयोजित क्लासेज</div>
              <h4 class="m-0 text-info" id="insightTotalClasses">0</h4>
            </div>
          </div>
          <div class="col-4">
            <div class="p-2 border border-secondary rounded bg-black">
              <div class="small text-secondary">औसत उपस्थिति</div>
              <h4 class="m-0 text-success" id="insightAvgRate">0%</h4>
            </div>
          </div>
        </div>
      </div>

      <!-- ओपनिंग बैलेंस कार्ड (Admin Only) -->
      <div class="card p-3" id="openingBalanceCard" style="display:none;">
        <h6 class="fw-bold mb-2"><i class="bi bi-calculator me-1"></i> Opening Balance (ऐतिहासिक डेटा)</h6>
        <p class="text-secondary small">ऐप शुरू करने से पहले आयोजित कुल क्लासेस और छात्रों की पिछली अनुपस्थिति यहाँ दर्ज करें।</p>
        <div class="row g-2 align-items-center mb-3">
          <div class="col-md-4">
            <label class="small text-secondary">कुल पिछली क्लासेस</label>
            <input type="number" id="histClasses" class="form-control form-control-sm">
          </div>
          <div class="col-md-4 mt-md-4">
            <button class="btn btn-outline-primary btn-sm" onclick="saveHistoricalBalance()">सेव ओपनिंग बैलेंस</button>
          </div>
        </div>
      </div>
    </div>

    <!-- 5. SETTINGS & BACKUP (Admin Only) -->
    <div id="settingsTab" class="tab-content-area" style="display:none;">
      <div class="card p-3 mb-3">
        <h6 class="fw-bold"><i class="bi bi-cloud-arrow-up me-1"></i> Google Drive स्नैपशॉट बैकअप</h6>
        <p class="text-secondary small">साप्ताहिक बैकअप के लिए नीचे क्लिक करें। यह Google Drive में टाइमस्टैम्प के साथ कॉपी बना देगा।</p>
        <button class="btn btn-outline-warning btn-sm" onclick="triggerBackup()">
          <i class="bi bi-copy me-1"></i> Make a backup copy
        </button>
        <div id="backupMsg" class="mt-2 small text-success"></div>
      </div>
    </div>
  </div>

  <!-- कॉल लॉग मोडल (चिप्स और 4 कार्ड स्टाइल्स के साथ) -->
  <div class="modal fade" id="callModal" tabindex="-1">
    <div class="modal-dialog modal-lg">
      <div class="modal-content bg-dark border-secondary">
        <div class="modal-header border-secondary">
          <h5 class="modal-title"><i class="bi bi-telephone-outbound text-warning me-2"></i>कॉल फॉलो-अप लॉग</h5>
          <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
        </div>
        <div class="modal-body">
          <input type="hidden" id="callStuId">
          <div class="d-flex justify-content-between mb-3 border-bottom border-secondary pb-2">
            <div>
              <h6 class="m-0 fw-bold" id="callStuName"></h6>
              <small class="text-secondary" id="callStuPhone"></small>
            </div>
            <div class="col-md-4">
              <label class="small text-secondary">Expected Return Date</label>
              <input type="date" id="callExpectedReturn" class="form-control form-control-sm">
            </div>
          </div>

          <!-- चिप्स सेलेक्शन -->
          <div class="mb-2">
            <label class="small text-secondary d-block">1. Reason (कारण):</label>
            <div class="d-flex flex-wrap gap-1" id="chipsReason">
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'reason')">तबीयत खराब</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'reason')">पारिवारिक कार्य / शादी</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'reason')">कॉलेज / परीक्षा</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'reason')">कॉल नहीं उठा</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'reason')">अन्य</span>
            </div>
          </div>

          <div class="mb-2">
            <label class="small text-secondary d-block">2. Action (कार्रवाई):</label>
            <div class="d-flex flex-wrap gap-1" id="chipsAction">
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'action')">छात्र से बात हुई</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'action')">अभिभावक से बात हुई</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'action')">WhatsApp संदेश भेजा</span>
            </div>
          </div>

          <div class="mb-2">
            <label class="small text-secondary d-block">3. Outcome (परिणाम):</label>
            <div class="d-flex flex-wrap gap-1" id="chipsOutcome">
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'outcome')">कल उपस्थित होगा</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'outcome')">निश्चित तिथि पर आएगा</span>
              <span class="badge bg-secondary chip" onclick="selectChip(this, 'outcome')">सत्र बदलने का अनुरोध</span>
            </div>
          </div>

          <!-- कार्ड स्टाइल्स -->
          <div class="mb-2 mt-3">
            <label class="small text-secondary d-block fw-bold">4. कार्ड स्टाइल चुनें (WhatsApp कॉपी के लिए):</label>
            <div class="btn-group btn-group-sm w-100" role="group">
              <input type="radio" class="btn-check" name="cardStyle" id="cs1" checked onclick="updatePreviewCard()">
              <label class="btn btn-outline-secondary" for="cs1">Style 1 (Formal)</label>
              <input type="radio" class="btn-check" name="cardStyle" id="cs2" onclick="updatePreviewCard()">
              <label class="btn btn-outline-secondary" for="cs2">Style 2 (Bullet)</label>
              <input type="radio" class="btn-check" name="cardStyle" id="cs3" onclick="updatePreviewCard()">
              <label class="btn btn-outline-secondary" for="cs3">Style 3 (Compact)</label>
              <input type="radio" class="btn-check" name="cardStyle" id="cs4" onclick="updatePreviewCard()">
              <label class="btn btn-outline-secondary" for="cs4">Style 4 (Urgent Alert)</label>
            </div>
          </div>

          <!-- प्रीव्यू -->
          <textarea id="cardPreviewText" class="form-control bg-black text-light font-monospace small mt-2" rows="5" readonly></textarea>
        </div>
        <div class="modal-footer border-secondary justify-content-between">
          <button class="btn btn-outline-success btn-sm" onclick="copyCardText()"><i class="bi bi-clipboard me-1"></i> Copy Selected Text</button>
          <button class="btn btn-primary btn-sm" onclick="saveCallLog()"><i class="bi bi-check2-circle me-1"></i> Save Call Log</button>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
  <script>
    let currentUserRole = null;
    let appData = null;
    let absentSelectedSet = new Set();
    let callModal = null;

    document.addEventListener("DOMContentLoaded", () => {
      callModal = new bootstrap.Modal(document.getElementById('callModal'));
      const today = new Date().toISOString().split('T')[0];
      document.getElementById('attDate').value = today;
      document.getElementById('todayLabel').innerText = "तारीख: " + today;
    });

    function handleLogin() {
      const code = document.getElementById('accessCodeInput').value;
      google.script.run.withSuccessHandler(res => {
        if (res.success) {
          currentUserRole = res.role;
          document.getElementById('loginScreen').style.display = 'none';
          document.getElementById('mainApp').style.display = 'block';
          document.getElementById('userRoleBadge').innerText = "Role: " + res.role.toUpperCase();
          if (res.role === 'admin') {
            document.getElementById('adminSettingsNav').style.display = 'block';
            document.getElementById('openingBalanceCard').style.display = 'block';
          }
          fetchData();
        } else {
          document.getElementById('loginError').style.display = 'block';
        }
      }).verifyAccessCode(code);
    }

    function logout() {
      currentUserRole = null;
      document.getElementById('mainApp').style.display = 'none';
      document.getElementById('loginScreen').style.display = 'flex';
      document.getElementById('accessCodeInput').value = '';
    }

    function switchTab(tabId) {
      document.querySelectorAll('.tab-content-area').forEach(el => el.style.display = 'none');
      document.querySelectorAll('#appTabs .nav-link').forEach(el => el.classList.remove('active'));
      document.getElementById(tabId).style.display = 'block';
      event.target.classList.add('active');
    }

    function fetchData() {
      google.script.run.withSuccessHandler(data => {
        appData = data;
        renderAttendancePills();
        renderReminders();
        renderInsights();
      }).getAppData();
    }

    // 1. Attendance Pills Logic
    function renderAttendancePills() {
      const sess = document.getElementById('attSession').value;
      const box = document.getElementById('attendancePillsBox');
      box.innerHTML = '';
      absentSelectedSet.clear();

      const students = appData.students.filter(s => s.active && s.session === sess);
      students.forEach(s => {
        const span = document.createElement('span');
        span.className = 'badge border border-secondary p-2 student-pill ' + (s.isAbsentToday ? 'absent' : 'bg-dark');
        span.innerText = `${s.name} (${s.id.slice(-4)})`;
        if (s.isAbsentToday) absentSelectedSet.add(s.id);

        span.onclick = () => {
          if (currentUserRole === 'view') return;
          if (absentSelectedSet.has(s.id)) {
            absentSelectedSet.delete(s.id);
            span.classList.remove('absent');
            span.classList.add('bg-dark');
          } else {
            absentSelectedSet.add(s.id);
            span.classList.add('absent');
            span.classList.remove('bg-dark');
          }
        };
        box.appendChild(span);
      });
    }

    function submitAttendance() {
      if (currentUserRole === 'view') return alert("View-only को अनुमति नहीं है");
      const d = document.getElementById('attDate').value;
      const sess = document.getElementById('attSession').value;
      google.script.run.withSuccessHandler(() => {
        alert("अटेंडेंस सेव हो गई!");
        fetchData();
      }).saveBulkAttendance(d, sess, Array.from(absentSelectedSet));
    }

    // 2. Reminders Tab (Consecutive >= 2)
    function renderReminders() {
      const tbody = document.getElementById('remindersTableBody');
      tbody.innerHTML = '';

      const alerts = appData.students.filter(s => s.active && s.consecutiveAbsences >= 2 && !s.isHiddenByReturn);
      if (alerts.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="text-center py-3 text-secondary">कोई 2+ दिन अनुपस्थित छात्र नहीं हैं।</td></tr>`;
        return;
      }

      alerts.forEach(s => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><span class="fw-bold">${s.name}</span><br><small class="text-secondary">${s.id}</small></td>
          <td><span class="badge bg-secondary">${s.session}</span></td>
          <td><a href="tel:${s.phone}" class="text-info">${s.phone}</a></td>
          <td><span class="badge bg-danger">${s.consecutiveAbsences} Days</span></td>
          <td><input type="date" class="form-control form-control-sm" style="width:140px;" onchange="saveReturnDate('${s.id}', this.value)" value="${s.expectedReturn || ''}"></td>
          <td>
            <button class="btn btn-outline-warning btn-sm" onclick="openCallModal('${s.id}')">
              <i class="bi bi-telephone-forward"></i> Call Done
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function saveReturnDate(stuId, val) {
      google.script.run.withSuccessHandler(() => {
        fetchData();
      }).updateExpectedReturn(stuId, val);
    }

    // 3. Call Log Modal & Chip logic
    let selectedChips = { reason: '', action: '', outcome: '' };
    let currentCallStu = null;

    function openCallModal(stuId) {
      currentCallStu = appData.students.find(s => s.id === stuId);
      document.getElementById('callStuId').value = currentCallStu.id;
      document.getElementById('callStuName').innerText = currentCallStu.name;
      document.getElementById('callStuPhone').innerText = `ID: ${currentCallStu.id} | Phone: ${currentCallStu.phone} | Session: ${currentCallStu.session}`;
      document.getElementById('callExpectedReturn').value = currentCallStu.expectedReturn || '';
      
      // Reset chips
      selectedChips = { reason: '', action: '', outcome: '' };
      document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      updatePreviewCard();
      callModal.show();
    }

    function selectChip(el, type) {
      el.parentElement.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      el.classList.add('active');
      selectedChips[type] = el.innerText;
      updatePreviewCard();
    }

    function updatePreviewCard() {
      if (!currentCallStu) return;
      const s1 = document.getElementById('cs1').checked;
      const s2 = document.getElementById('cs2').checked;
      const s3 = document.getElementById('cs3').checked;
      const s4 = document.getElementById('cs4').checked;
      const retDate = document.getElementById('callExpectedReturn').value || 'N/A';

      let text = '';
      if (s1) {
        text = `📞 *FEA Retention Follow-Up*\nStudent: ${currentCallStu.name} (${currentCallStu.id})\nReason: ${selectedChips.reason || 'N/A'}\nAction: ${selectedChips.action || 'N/A'}\nOutcome: ${selectedChips.outcome || 'N/A'}\nExpected Return: ${retDate}`;
      } else if (s2) {
        text = `• *Call Update*: ${currentCallStu.name} [${currentCallStu.session}]\n• *Status*: ${selectedChips.reason}\n• *Action Taken*: ${selectedChips.action}\n• *Next Steps*: Return by ${retDate}`;
      } else if (s3) {
        text = `[${currentCallStu.name} - ${currentCallStu.session}] ${selectedChips.reason} -> ${selectedChips.action} -> Return: ${retDate}`;
      } else if (s4) {
        text = `🚨 *URGENT ATTENDANCE ALERT*\n${currentCallStu.name} has been away for ${currentCallStu.consecutiveAbsences} days.\nUpdate: ${selectedChips.reason} | Follow-up: ${retDate}`;
      }
      document.getElementById('cardPreviewText').value = text;
    }

    function copyCardText() {
      navigator.clipboard.writeText(document.getElementById('cardPreviewText').value).then(() => alert("कार्ड टेक्स्ट कॉपी हो गया!"));
    }

    function saveCallLog() {
      const payload = {
        studentId: currentCallStu.id,
        name: currentCallStu.name,
        reason: selectedChips.reason,
        action: selectedChips.action,
        outcome: selectedChips.outcome,
        followUp: document.getElementById('callExpectedReturn').value,
        expectedReturn: document.getElementById('callExpectedReturn').value,
        cardStyle: document.querySelector('input[name="cardStyle"]:checked').id,
        copiedText: document.getElementById('cardPreviewText').value
      };
      google.script.run.withSuccessHandler(() => {
        callModal.hide();
        fetchData();
      }).saveCallRecord(payload);
    }

    // 4. Report L- Generation
    function generateReport() {
      const lessons = {
        S1: document.getElementById('lS1').value,
        S2: document.getElementById('lS2').value,
        S3: document.getElementById('lS3').value
      };
      google.script.run.withSuccessHandler(res => {
        document.getElementById('reportResultBox').style.display = 'block';
        document.getElementById('reportTextPreview').value = res.reportText;
      }).generateLessonReport(lessons);
    }

    function copyReport() {
      navigator.clipboard.writeText(document.getElementById('reportTextPreview').value).then(() => alert("रिपोर्ट कॉपी हो गई!"));
    }

    // 5. Insights & Opening Balance
    function renderInsights() {
      const active = appData.students.filter(s => s.active);
      document.getElementById('insightTotalStudents').innerText = active.length;
      document.getElementById('insightTotalClasses').innerText = appData.totalClassesHeld;
      document.getElementById('histClasses').value = appData.totalClassesHeld;

      let totalPossible = active.length * (appData.totalClassesHeld || 1);
      let totalAbs = active.reduce((acc, s) => acc + s.totalAbsences, 0);
      let rate = totalPossible > 0 ? Math.round(((totalPossible - totalAbs) / totalPossible) * 100) : 100;
      document.getElementById('insightAvgRate').innerText = `${Math.max(0, rate)}%`;
    }

    function saveHistoricalBalance() {
      const val = Number(document.getElementById('histClasses').value);
      google.script.run.withSuccessHandler(() => {
        alert("ओपनिंग बैलेंस सेव हो गया!");
        fetchData();
      }).saveOpeningBalance(val, {});
    }

    // 6. Settings Backup
    function triggerBackup() {
      document.getElementById('backupMsg').innerText = "बैकअप बन रहा है...";
      google.script.run.withSuccessHandler(res => {
        document.getElementById('backupMsg').innerHTML = `✅ बैकअप तैयार: <a href="${res.url}" target="_blank" class="text-info">${res.fileName}</a>`;
      }).makeBackupCopy();
    }
  </script>
</body>
</html>
```

---

### कार्यान्वयन के अगले 3 त्वरित कदम (Execution Steps):

1. **`Code.gs` और `Index.html` अपडेट करें**:
* Google Sheet में **Extensions > Apps Script** खोलें।
* `Code.gs` में भाग 1 का कोड पेस्ट करें।
* `Index.html` में भाग 2 का कोड पेस्ट करें और **Ctrl + S** दबाकर सेव करें।


2. **`setup` फ़ंक्शन रन करें**:
* ऊपर टूलबार में फ़ंक्शन ड्रॉपडाउन से **`setup`** चुनें और **Run** दबाएं।
* अगर अनुमति (Permissions) माँगे तो Allow करें।
* नीचे Execution log में आपको **Admin**, **Staff**, और **View-only** एक्सेस कोड दिखाई देंगे (जैसे `FEA-ADM-1234`)।


3. **New Version Deploy करें**:
* ऊपर दाएँ **Deploy > Manage deployments** पर जाएँ।
* पेंसिल (Edit) पर क्लिक करें, Version में **New version** चुनें, और **Deploy** कर दें।