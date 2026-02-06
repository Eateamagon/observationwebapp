/**
 * Debug.gs - Comprehensive debugging suite for Classroom Observation System
 * 
 * Add this file to your Apps Script project for troubleshooting.
 * Run functions from the dropdown menu and check View > Logs for output.
 */

// ============================================================================
// MAIN DEBUG DASHBOARD
// ============================================================================

/**
 * Run all diagnostic tests and output a summary
 */
function runFullDiagnostics() {
  Logger.log('='.repeat(60));
  Logger.log('CLASSROOM OBSERVATION SYSTEM - DIAGNOSTIC REPORT');
  Logger.log('Run at: ' + new Date().toISOString());
  Logger.log('='.repeat(60));
  
  testSpreadsheetAccess();
  testAllSheets();
  testCurrentUser();
  testTeacherData();
  testBellSchedules();
  testObservations();
  testSubRequests();
  testEmailCapability();
  testCalendarCapability();
  
  Logger.log('='.repeat(60));
  Logger.log('DIAGNOSTIC COMPLETE');
  Logger.log('='.repeat(60));
}

// ============================================================================
// SPREADSHEET & SHEET TESTS
// ============================================================================

function testSpreadsheetAccess() {
  Logger.log('\n--- SPREADSHEET ACCESS TEST ---');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✓ Spreadsheet Name: ' + ss.getName());
    Logger.log('✓ Spreadsheet ID: ' + ss.getId());
    Logger.log('✓ URL: ' + ss.getUrl());
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testAllSheets() {
  Logger.log('\n--- SHEET EXISTENCE TEST ---');
  const requiredSheets = [
    'Teachers', 'Rooms', 'BellSchedules', 'PrepPeriods', 
    'LunchPeriods', 'Observations', 'SubstituteRequests', 
    'Admins', 'ReadOnlyUsers', 'AuditLog'
  ];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  requiredSheets.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const rowCount = sheet.getLastRow();
      const colCount = sheet.getLastColumn();
      Logger.log('✓ ' + sheetName + ': ' + rowCount + ' rows, ' + colCount + ' columns');
      
      // Show headers
      if (rowCount > 0 && colCount > 0) {
        const headers = sheet.getRange(1, 1, 1, colCount).getValues()[0];
        Logger.log('  Headers: ' + JSON.stringify(headers));
      }
    } else {
      Logger.log('✗ MISSING: ' + sheetName);
    }
  });
}

// ============================================================================
// USER & AUTH TESTS
// ============================================================================

function testCurrentUser() {
  Logger.log('\n--- CURRENT USER TEST ---');
  try {
    const email = Session.getActiveUser().getEmail();
    Logger.log('✓ Current Email: ' + email);
    
    const user = getCurrentUser();
    if (user) {
      Logger.log('✓ User Object: ' + JSON.stringify(user, null, 2));
    } else {
      Logger.log('✗ getCurrentUser() returned null - user may not be in Teachers/Admins sheet');
    }
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

// ============================================================================
// DATA TESTS
// ============================================================================

function testTeacherData() {
  Logger.log('\n--- TEACHER DATA TEST ---');
  try {
    const teachers = Data.getAllTeachers();
    Logger.log('✓ Total active teachers: ' + teachers.length);
    
    if (teachers.length > 0) {
      Logger.log('✓ Sample teacher: ' + JSON.stringify(teachers[0], null, 2));
      
      // Test grade filtering
      [6, 7, 8].forEach(function(grade) {
        const byGrade = Data.getTeachersByGrade(grade);
        Logger.log('  Grade ' + grade + ': ' + byGrade.length + ' teachers');
      });
    }
    
    // Test multi-grade teachers
    const multiGrade = teachers.filter(function(t) { return t.grades && t.grades.length > 1; });
    Logger.log('✓ Multi-grade teachers: ' + multiGrade.length);
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
    Logger.log('  Stack: ' + e.stack);
  }
}

function testBellSchedules() {
  Logger.log('\n--- BELL SCHEDULE TEST ---');
  try {
    [6, 7, 8].forEach(function(grade) {
      const schedule = Data.getBellSchedule(grade);
      Logger.log('Grade ' + grade + ': ' + schedule.length + ' periods');
      if (schedule.length > 0) {
        Logger.log('  First period: ' + JSON.stringify(schedule[0]));
        Logger.log('  Last period: ' + JSON.stringify(schedule[schedule.length - 1]));
      }
    });
    
    // Check lunch periods
    const lunchPeriods = Data.getAllRows('LunchPeriods');
    Logger.log('✓ Lunch periods configured: ' + lunchPeriods.length);
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testObservations() {
  Logger.log('\n--- OBSERVATIONS TEST ---');
  try {
    const allObs = Data.getAllObservations({});
    Logger.log('✓ Total observations: ' + allObs.length);
    
    const byStatus = {};
    allObs.forEach(function(o) {
      byStatus[o.status] = (byStatus[o.status] || 0) + 1;
    });
    Logger.log('✓ By status: ' + JSON.stringify(byStatus));
    
    if (allObs.length > 0) {
      Logger.log('✓ Latest observation: ' + JSON.stringify(allObs[0], null, 2));
    }
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testSubRequests() {
  Logger.log('\n--- SUBSTITUTE REQUESTS TEST ---');
  try {
    const pending = Data.getSubRequestsByStatus('pending');
    Logger.log('✓ Pending requests: ' + pending.length);
    
    const approved = Data.getSubRequestsByStatus('approved');
    Logger.log('✓ Approved requests: ' + approved.length);
    
    const denied = Data.getSubRequestsByStatus('denied');
    Logger.log('✓ Denied requests: ' + denied.length);
    
    if (pending.length > 0) {
      Logger.log('✓ Sample pending request: ' + JSON.stringify(pending[0], null, 2));
    }
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

// ============================================================================
// EMAIL & CALENDAR TESTS
// ============================================================================

function testEmailCapability() {
  Logger.log('\n--- EMAIL CAPABILITY TEST ---');
  try {
    const quota = MailApp.getRemainingDailyQuota();
    Logger.log('✓ Remaining email quota: ' + quota);
    
    if (quota < 10) {
      Logger.log('⚠ WARNING: Low email quota!');
    }
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testCalendarCapability() {
  Logger.log('\n--- CALENDAR CAPABILITY TEST ---');
  try {
    const calendars = CalendarApp.getAllCalendars();
    Logger.log('✓ Accessible calendars: ' + calendars.length);
    
    const defaultCal = CalendarApp.getDefaultCalendar();
    Logger.log('✓ Default calendar: ' + defaultCal.getName());
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

// ============================================================================
// FUNCTION-SPECIFIC TESTS
// ============================================================================

function testGetAvailableSlots() {
  Logger.log('\n--- GET AVAILABLE SLOTS TEST ---');
  try {
    const teachers = Data.getAllTeachers();
    if (teachers.length === 0) {
      Logger.log('✗ No teachers to test with');
      return;
    }
    
    const teacher = teachers[0];
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    // Skip to Monday if weekend
    while (tomorrow.getDay() === 0 || tomorrow.getDay() === 6) {
      tomorrow.setDate(tomorrow.getDate() + 1);
    }
    const dateStr = Utilities.formatDate(tomorrow, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    
    Logger.log('Testing with teacher: ' + teacher.name + ' (' + teacher.id + ')');
    Logger.log('Testing date: ' + dateStr);
    
    const result = getAvailableSlots(teacher.id, dateStr);
    Logger.log('Result: ' + JSON.stringify(result, null, 2));
    
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
    Logger.log('  Stack: ' + e.stack);
  }
}

function testGetDashboardData() {
  Logger.log('\n--- GET DASHBOARD DATA TEST ---');
  try {
    const data = getDashboardData();
    Logger.log('Result: ' + JSON.stringify(data, null, 2));
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
    Logger.log('  Stack: ' + e.stack);
  }
}

function testGetMyObservationStatus() {
  Logger.log('\n--- GET MY OBSERVATION STATUS TEST ---');
  try {
    const status = getMyObservationStatus();
    Logger.log('Result: ' + JSON.stringify(status, null, 2));
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testGetMyObservations() {
  Logger.log('\n--- GET MY OBSERVATIONS TEST ---');
  try {
    const obs = getMyObservations();
    Logger.log('Result: ' + JSON.stringify(obs, null, 2));
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

function testGetPendingSubRequests() {
  Logger.log('\n--- GET PENDING SUB REQUESTS TEST ---');
  try {
    const requests = getPendingSubRequests();
    Logger.log('Result: ' + JSON.stringify(requests, null, 2));
  } catch (e) {
    Logger.log('✗ FAILED: ' + e.message);
  }
}

// ============================================================================
// DATA VALIDATION
// ============================================================================

function validateTeacherData() {
  Logger.log('\n--- TEACHER DATA VALIDATION ---');
  const teachers = Data.getAllRows('Teachers');
  let issues = [];
  
  teachers.forEach(function(t, index) {
    const row = index + 2; // Account for header
    
    if (!t.id) issues.push('Row ' + row + ': Missing ID');
    if (!t.email) issues.push('Row ' + row + ': Missing email');
    if (!t.name) issues.push('Row ' + row + ': Missing name');
    if (!t.room) issues.push('Row ' + row + ': Missing room');
    
    // Check active value
    const active = String(t.active).toUpperCase();
    if (!['TRUE', 'FALSE', 'ACTIVE', 'INACTIVE', 'YES', 'NO', '1', '0'].includes(active)) {
      issues.push('Row ' + row + ': Invalid active value "' + t.active + '"');
    }
    
    // Check email format
    if (t.email && !t.email.includes('@')) {
      issues.push('Row ' + row + ': Invalid email format "' + t.email + '"');
    }
  });
  
  if (issues.length === 0) {
    Logger.log('✓ All teacher data is valid');
  } else {
    Logger.log('✗ Found ' + issues.length + ' issues:');
    issues.forEach(function(issue) { Logger.log('  - ' + issue); });
  }
}

function validateBellScheduleData() {
  Logger.log('\n--- BELL SCHEDULE DATA VALIDATION ---');
  const schedules = Data.getAllRows('BellSchedules');
  let issues = [];
  
  schedules.forEach(function(s, index) {
    const row = index + 2;
    
    if (!s.grade) issues.push('Row ' + row + ': Missing grade');
    if (!s.period) issues.push('Row ' + row + ': Missing period');
    if (!s.startTime) issues.push('Row ' + row + ': Missing startTime');
    if (!s.endTime) issues.push('Row ' + row + ': Missing endTime');
  });
  
  // Check for complete schedules per grade
  [6, 7, 8].forEach(function(grade) {
    const gradeSchedule = schedules.filter(function(s) { return Number(s.grade) === grade; });
    if (gradeSchedule.length === 0) {
      issues.push('Grade ' + grade + ': No bell schedule defined');
    } else if (gradeSchedule.length < 6) {
      issues.push('Grade ' + grade + ': Only ' + gradeSchedule.length + ' periods defined (expected 6-8)');
    }
  });
  
  if (issues.length === 0) {
    Logger.log('✓ All bell schedule data is valid');
  } else {
    Logger.log('✗ Found ' + issues.length + ' issues:');
    issues.forEach(function(issue) { Logger.log('  - ' + issue); });
  }
}

// ============================================================================
// CLEANUP & MAINTENANCE
// ============================================================================

function listAllObservationIds() {
  Logger.log('\n--- ALL OBSERVATION IDS ---');
  const obs = Data.getAllObservations({});
  obs.forEach(function(o) {
    Logger.log(o.id + ' | ' + o.date + ' | ' + o.observerName + ' -> ' + o.teacherName + ' | ' + o.status);
  });
}

function listAllSubRequestIds() {
  Logger.log('\n--- ALL SUB REQUEST IDS ---');
  const requests = Data.getAllRows('SubstituteRequests');
  requests.forEach(function(r) {
    Logger.log(r.id + ' | ' + r.date + ' | ' + r.requesterName + ' | ' + r.status);
  });
}

// ============================================================================
// QUICK FIXES
// ============================================================================

/**
 * Fix any observations with null periods
 */
function fixNullPeriods() {
  Logger.log('\n--- FIXING NULL PERIODS ---');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Observations');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const periodsCol = headers.indexOf('periods');
  
  if (periodsCol === -1) {
    Logger.log('✗ periods column not found');
    return;
  }
  
  let fixed = 0;
  for (let i = 1; i < data.length; i++) {
    const periods = data[i][periodsCol];
    if (!periods || periods === '' || periods === 'null') {
      sheet.getRange(i + 1, periodsCol + 1).setValue('[]');
      fixed++;
    }
  }
  
  Logger.log('✓ Fixed ' + fixed + ' rows');
}

/**
 * Reset the AuditLog sheet (keeps headers)
 */
function clearAuditLog() {
  Logger.log('\n--- CLEARING AUDIT LOG ---');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AuditLog');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
    Logger.log('✓ Audit log cleared');
  } else {
    Logger.log('✓ Audit log already empty');
  }
}
