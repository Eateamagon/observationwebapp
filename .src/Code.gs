/**
 * Classroom Observation Management System
 * Main Entry Point and Routing
 * 
 * This file handles:
 * - Web app entry point (doGet)
 * - User authentication
 * - Role-based routing
 * - Core server functions called from frontend
 */

// ============================================================================
// CONFIGURATION - SCHOOL-SPECIFIC SETTINGS
// ============================================================================
// 
// *** IMPORTANT: Update these values for each school deployment ***
//
const CONFIG = {
  // ==========================================================================
  // SCHOOL-SPECIFIC SETTINGS (MUST UPDATE FOR EACH SCHOOL)
  // ==========================================================================
  
  // Google Spreadsheet ID for this school's data
  // Get this from the spreadsheet URL: docs.google.com/spreadsheets/d/{THIS_ID}/edit
  SPREADSHEET_ID: '1QrHoixG5JTduc7jMD3p1K8UZATeWT67Yc_XLfVCvOaA',
  
  // School name (appears in emails and notifications)
  SCHOOL_NAME: 'Kate Collins Middle School',
  
  // School abbreviation (for branding - up to 4 characters)
  SCHOOL_ABBR: 'KCMS',
  
  // Substitute coordinator email for this school
  SUB_COORDINATOR_EMAIL: 'dfolks@waynesboro.k12.va.us',
  
  // Grade levels at this school
  GRADES: [6, 7, 8, 9],
  
  // Allowed email domain (leave empty to allow any)
  ALLOWED_EMAIL_DOMAIN: 'waynesboro.k12.va.us',
  
  // ==========================================================================
  // DIVISION-WIDE SETTINGS (Same for all schools)
  // ==========================================================================
  
  // Observation deadline day/month (year auto-calculated)
  OBSERVATION_DEADLINE_MONTH: 4,  // April
  OBSERVATION_DEADLINE_DAY: 19,   // 19th
  
  // Timezone
  TIMEZONE: 'America/New_York',
  
  // Minimum days advance notice for substitute requests
  MIN_ADVANCE_DAYS_FOR_SUB: 2,
  
  // Lock timeout for concurrent access prevention
  LOCK_TIMEOUT_MS: 30000,
  
  // Observation form template ID (Google Doc) - shared across division
  OBSERVATION_FORM_ID: '1VAJlJ7SZpRo4wyuautRXpoGFR9RHV-UIz4LcbTwBfyU',
};

// ============================================================================
// AUTO-CALCULATED VALUES (No manual updates needed!)
// ============================================================================

/**
 * Get current school year string (e.g., "2025-2026")
 * Auto-rolls on August 1st
 */
function getSchoolYear() {
  const now = new Date();
  const year = now.getMonth() >= 7 ? now.getFullYear() : now.getFullYear() - 1;
  return `${year}-${year + 1}`;
}

/**
 * Get observation deadline for current school year
 * Returns Date object for April 19 (or configured date) of spring semester
 */
function getObservationDeadline() {
  const now = new Date();
  const deadlineYear = now.getMonth() >= 7 ? now.getFullYear() + 1 : now.getFullYear();
  return new Date(deadlineYear, CONFIG.OBSERVATION_DEADLINE_MONTH - 1, CONFIG.OBSERVATION_DEADLINE_DAY, 23, 59, 59);
}

/**
 * Get school year start date (August 1)
 */
function getSchoolYearStart() {
  const now = new Date();
  const startYear = now.getMonth() >= 7 ? now.getFullYear() : now.getFullYear() - 1;
  return new Date(startYear, 7, 1); // August 1
}

// ============================================================================
// WEB APP ENTRY POINT
// ============================================================================

/**
 * Main entry point for the web app
 * @param {Object} e - Event object from Google
 * @return {HtmlOutput} The rendered HTML page
 */
function doGet(e) {
  const user = getCurrentUser();
  
  if (!user) {
    // Check if they have a valid domain but aren't registered
    const email = Session.getActiveUser().getEmail();
    if (email && CONFIG.ALLOWED_EMAIL_DOMAIN && email.endsWith('@' + CONFIG.ALLOWED_EMAIL_DOMAIN)) {
      // Show access request form
      return HtmlService.createHtmlOutput(getAccessRequestPage(email))
        .setTitle('Request Access - ' + CONFIG.SCHOOL_NAME);
    }
    return HtmlService.createHtmlOutput('<h1>Access Denied</h1><p>You must be signed in with a school Google account.</p>')
      .setTitle('Access Denied');
  }
  
  const template = HtmlService.createTemplateFromFile('Index');
  template.user = user;
  template.config = {
    schoolYear: getSchoolYear(),
    grades: CONFIG.GRADES,
    minAdvanceDays: CONFIG.MIN_ADVANCE_DAYS_FOR_SUB,
    schoolName: CONFIG.SCHOOL_NAME,
    schoolAbbr: getCurrentBranding() // Dynamic branding from Settings sheet
  };
  
  return template.evaluate()
    .setTitle(CONFIG.SCHOOL_NAME + ' - Observation System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include external HTML/CSS/JS files
 * @param {string} filename - Name of file to include
 * @return {string} File contents
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ============================================================================
// AUTHENTICATION & AUTHORIZATION
// ============================================================================

/**
 * Get current authenticated user with role information
 * @return {Object|null} User object with email, name, role, grade, etc.
 */
function getCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return null;
    
    const role = Permissions.getUserRole(email);
    const teacherInfo = role === 'teacher' || role === 'admin' 
      ? Data.getTeacherByEmail(email) 
      : null;
    
    return {
      email: email,
      role: role,
      name: teacherInfo ? teacherInfo.name : email.split('@')[0],
      teacherId: teacherInfo ? teacherInfo.id : null,
      grade: teacherInfo ? teacherInfo.grade : null,
      room: teacherInfo ? teacherInfo.room : null
    };
  } catch (error) {
    Logger.log('Error getting current user: ' + error.message);
    return null;
  }
}

/**
 * Server-side role check (called before sensitive operations)
 * @param {string} requiredRole - Minimum role required ('readonly', 'teacher', 'admin')
 * @return {boolean} Whether user has required role
 */
function checkRole(requiredRole) {
  const user = getCurrentUser();
  if (!user) return false;
  
  const roleHierarchy = { 'readonly': 1, 'teacher': 2, 'admin': 3 };
  return roleHierarchy[user.role] >= roleHierarchy[requiredRole];
}

// ============================================================================
// OBSERVATION MANAGEMENT (Server Functions)
// ============================================================================

/**
 * Get teacher's observation status for the school year
 * Requires 1 observation between now and the deadline
 * @return {Object} Status object with count, hasMetRequirement, observations
 */
function getMyObservationStatus() {
  const user = getCurrentUser();
  if (!user || !user.teacherId) {
    return { error: 'Not authenticated as teacher' };
  }
  
  const now = new Date();
  const schoolYearStart = getSchoolYearStart();
  const deadline = getObservationDeadline();
  
  // Get all observations for this school year
  const observations = Data.getObservationsByObserver(user.teacherId, schoolYearStart, deadline);
  
  // Calculate days remaining until deadline
  const daysRemaining = Math.ceil((deadline - now) / (1000 * 60 * 60 * 24));
  
  return {
    count: observations.length,
    hasMetRequirement: observations.length >= 1,
    observations: observations,
    deadline: Utilities.formatDate(deadline, CONFIG.TIMEZONE, 'MMMM d, yyyy'),
    daysRemaining: daysRemaining > 0 ? daysRemaining : 0,
    isPastDeadline: now > deadline,
    schoolYear: getSchoolYear()
  };
}

/**
 * Get available rooms/teachers for a specific grade
 * @param {number} grade - Grade level (6, 7, or 8)
 * @return {Object[]} Array of available teacher/room objects
 */
function getTeachersByGrade(grade) {
  if (!checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  if (!CONFIG.GRADES.includes(Number(grade))) {
    return { error: 'Invalid grade' };
  }
  
  return Data.getTeachersByGrade(Number(grade));
}

/**
 * Get all teachers (for cross-grade observation)
 * @return {Object[]} Array of all teacher objects
 */
function getAllTeachers() {
  if (!checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  return Data.getAllTeachers();
}

/**
 * Get bell schedule for a specific grade
 * @param {number} grade - Grade level
 * @return {Object[]} Array of period objects with times
 */
function getBellSchedule(grade) {
  if (!checkRole('readonly')) {
    return { error: 'Unauthorized' };
  }
  
  // 6th grade has own schedule, 7th/8th/Electives/SPED/ELL share
  const scheduleGrade = Number(grade) === 6 ? 6 : 7;
  return Data.getBellSchedule(scheduleGrade);
}

/**
 * Get prep periods for a grade
 * @param {number} grade - Grade level
 * @return {Object[]} Array of prep period info
 */
function getPrepPeriods(grade) {
  if (!checkRole('readonly')) {
    return { error: 'Unauthorized' };
  }
  
  return Data.getPrepPeriods(Number(grade));
}

/**
 * Get available time slots for a specific teacher on a date
 * @param {string} teacherId - Teacher to be observed
 * @param {string} dateStr - Date string (YYYY-MM-DD)
 * @return {Object} Available slots with conflict info
 */
function getAvailableSlots(teacherId, dateStr) {
  const user = getCurrentUser();
  if (!user || !checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  const date = new Date(dateStr + 'T00:00:00');
  const teacher = Data.getTeacherById(teacherId);
  
  if (!teacher) {
    return { error: 'Teacher not found' };
  }
  
  // Get bell schedule - use primary grade, or default to 7 for support staff
  const primaryGrade = Array.isArray(teacher.grades) && teacher.grades.length > 0 
    ? teacher.grades[0] 
    : (teacher.grade || 7);
  const scheduleGrade = primaryGrade === 6 ? 6 : 7;
  const bellSchedule = Data.getBellSchedule(scheduleGrade);
  
  // Get unavailable periods for the TEACHER being observed
  // Support staff have no fixed restrictions
  const isSupport = teacher.type === 'support';
  
  // Get teacher's unavailable periods
  let unavailablePeriodNumbers = [];
  if (!isSupport) {
    // First check new unavailablePeriods field
    if (teacher.unavailablePeriods) {
      let teacherUnavailable = teacher.unavailablePeriods;
      if (typeof teacherUnavailable === 'string') {
        try {
          teacherUnavailable = JSON.parse(teacherUnavailable);
        } catch (e) {
          teacherUnavailable = teacherUnavailable.split(',').map(p => Number(p.trim())).filter(p => p);
        }
      }
      if (Array.isArray(teacherUnavailable)) {
        teacherUnavailable.forEach(p => {
          const pNum = Number(p);
          if (pNum && !unavailablePeriodNumbers.includes(pNum)) {
            unavailablePeriodNumbers.push(pNum);
          }
        });
      }
    }
    
    // Backward compatibility: check old lunchPeriod field
    if (teacher.lunchPeriod) {
      const teacherLunch = Number(teacher.lunchPeriod);
      if (teacherLunch && !unavailablePeriodNumbers.includes(teacherLunch)) {
        unavailablePeriodNumbers.push(teacherLunch);
      }
    }
    
    // Fall back to grade-based lunch periods if no teacher-specific settings
    if (unavailablePeriodNumbers.length === 0) {
      const teacherGrades = Array.isArray(teacher.grades) ? teacher.grades : [teacher.grade];
      teacherGrades.forEach(g => {
        const gradeLunches = Data.getLunchPeriods(Number(g));
        gradeLunches.forEach(l => {
          if (!unavailablePeriodNumbers.includes(Number(l.period))) {
            unavailablePeriodNumbers.push(Number(l.period));
          }
        });
      });
    }
  }
  
  // Get existing observations for this date/teacher
  const existingObs = Data.getObservationsForTeacherOnDate(teacherId, date);
  
  // Get observer's existing observations on this date
  const observerObs = Data.getObservationsByObserverOnDate(user.teacherId, date);
  
  // Get observer info for their schedule constraints
  const observer = Data.getTeacherById(user.teacherId);
  const observerIsSupport = observer && observer.type === 'support';
  
  // Build slot availability
  const slots = bellSchedule.map(period => {
    const periodNum = Number(period.period);
    const isUnavailable = unavailablePeriodNumbers.includes(periodNum);
    const isBooked = existingObs.some(o => {
      const obsPeriods = Array.isArray(o.periods) ? o.periods : [];
      return obsPeriods.includes(periodNum) || obsPeriods.includes(String(periodNum));
    });
    const observerBusy = observerObs.some(o => {
      const obsPeriods = Array.isArray(o.periods) ? o.periods : [];
      return obsPeriods.includes(periodNum) || obsPeriods.includes(String(periodNum));
    });
    
    // Check if observer is being observed during this period
    const observerBeingObserved = Data.getObservationsForTeacherOnDate(user.teacherId, date)
      .some(o => {
        const obsPeriods = Array.isArray(o.periods) ? o.periods : [];
        return obsPeriods.includes(periodNum) || obsPeriods.includes(String(periodNum));
      });
    
    // Determine availability
    const available = !isUnavailable && !isBooked && !observerBusy && !observerBeingObserved;
    
    // Determine reason for unavailability
    let reason = null;
    if (isUnavailable) {
      reason = 'Teacher unavailable';
    } else if (isBooked) {
      reason = 'Already has observer';
    } else if (observerBusy) {
      reason = 'You have another observation';
    } else if (observerBeingObserved) {
      reason = 'You are being observed';
    }
    
    return {
      period: period.period,
      startTime: period.startTime,
      endTime: period.endTime,
      available: available,
      isUnavailable: isUnavailable,
      isBooked: isBooked,
      observerBusy: observerBusy,
      observerBeingObserved: observerBeingObserved,
      reason: reason
    };
  });
  
  return {
    teacher: teacher,
    date: dateStr,
    slots: slots
  };
}

/**
 * Create a new observation
 * @param {Object} observationData - Observation details
 * @return {Object} Result with success/error
 */
function createObservation(observationData) {
  const user = getCurrentUser();
  if (!user || !checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  // Server-side validation
  const validation = validateObservation(observationData, user);
  if (!validation.valid) {
    return { error: validation.message };
  }
  
  // Acquire lock to prevent race conditions
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    // Re-validate after acquiring lock (double-check)
    const revalidation = validateObservation(observationData, user);
    if (!revalidation.valid) {
      return { error: revalidation.message };
    }
    
    // Check for requirement status
    const status = getMyObservationStatus();
    const alreadyMet = status.hasMetRequirement;
    
    // Create the observation
    const observation = {
      id: Utilities.getUuid(),
      observerId: user.teacherId,
      observerEmail: user.email,
      observerName: user.name,
      teacherId: observationData.teacherId,
      teacherName: observationData.teacherName,
      room: observationData.room,
      grade: observationData.grade,
      date: observationData.date,
      periods: observationData.periods,
      needsSub: observationData.needsSub || false,
      subStatus: observationData.needsSub ? 'pending' : 'not_needed',
      status: observationData.needsSub ? 'pending_sub' : 'confirmed',
      createdAt: new Date().toISOString(),
      createdBy: user.email
    };
    
    // Save observation
    const result = Data.createObservation(observation);
    
    if (result.success) {
      // Log the action
      Data.logAudit({
        action: 'CREATE_OBSERVATION',
        userId: user.email,
        details: JSON.stringify(observation),
        timestamp: new Date().toISOString()
      });
      
      // If sub needed, create sub request and notify coordinator
      if (observation.needsSub) {
        createSubstituteRequest(observation);
      }
      
      // Return success with bonus info
      return {
        success: true,
        observationId: observation.id,
        alreadyMetRequirement: alreadyMet,
        message: alreadyMet 
          ? 'Observation scheduled! You\'ve already met this month\'s requirement - PD points awarded!'
          : 'Observation scheduled successfully!'
      };
    } else {
      return { error: result.error || 'Failed to create observation' };
    }
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cancel an observation
 * @param {string} observationId - ID of observation to cancel
 * @return {Object} Result with success/error
 */
function cancelObservation(observationId) {
  const user = getCurrentUser();
  if (!user || !checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  const observation = Data.getObservationById(observationId);
  
  if (!observation) {
    return { error: 'Observation not found' };
  }
  
  // Check permission: own observation or admin
  if (observation.observerId !== user.teacherId && user.role !== 'admin') {
    return { error: 'You can only cancel your own observations' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    // Delete calendar events for both participants
    deleteObservationCalendarEvents(observation);
    
    const result = Data.updateObservation(observationId, { 
      status: 'canceled',
      canceledAt: new Date().toISOString(),
      canceledBy: user.email
    });
    
    if (result.success) {
      // Cancel any pending sub request
      if (observation.needsSub) {
        Data.updateSubRequest(observationId, { status: 'canceled' });
      }
      
      Data.logAudit({
        action: 'CANCEL_OBSERVATION',
        userId: user.email,
        details: JSON.stringify({ observationId, reason: 'User canceled', calendarEventsDeleted: true }),
        timestamp: new Date().toISOString()
      });
      
      return { success: true, message: 'Observation canceled and calendar events removed' };
    } else {
      return { error: result.error || 'Failed to cancel observation' };
    }
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reschedule an observation
 * @param {string} observationId - ID of observation to reschedule
 * @param {Object} newData - New date/periods
 * @return {Object} Result with success/error
 */
function rescheduleObservation(observationId, newData) {
  const user = getCurrentUser();
  if (!user || !checkRole('teacher')) {
    return { error: 'Unauthorized' };
  }
  
  const observation = Data.getObservationById(observationId);
  
  if (!observation) {
    return { error: 'Observation not found' };
  }
  
  if (observation.observerId !== user.teacherId && user.role !== 'admin') {
    return { error: 'You can only reschedule your own observations' };
  }
  
  // Validate new slot
  const validation = validateObservation({
    ...observation,
    date: newData.date,
    periods: newData.periods,
    needsSub: newData.needsSub
  }, user, observationId);
  
  if (!validation.valid) {
    return { error: validation.message };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    const result = Data.updateObservation(observationId, {
      date: newData.date,
      periods: newData.periods,
      needsSub: newData.needsSub,
      subStatus: newData.needsSub ? 'pending' : 'not_needed',
      status: newData.needsSub ? 'pending_sub' : 'confirmed',
      rescheduledAt: new Date().toISOString(),
      rescheduledBy: user.email
    });
    
    if (result.success) {
      // Handle sub request changes
      if (newData.needsSub && !observation.needsSub) {
        createSubstituteRequest({ ...observation, ...newData, id: observationId });
      } else if (!newData.needsSub && observation.needsSub) {
        Data.updateSubRequest(observationId, { status: 'canceled' });
      }
      
      Data.logAudit({
        action: 'RESCHEDULE_OBSERVATION',
        userId: user.email,
        details: JSON.stringify({ observationId, oldDate: observation.date, newDate: newData.date }),
        timestamp: new Date().toISOString()
      });
      
      return { success: true, message: 'Observation rescheduled' };
    } else {
      return { error: result.error || 'Failed to reschedule' };
    }
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get teacher's observations
 * @return {Object[]} Array of observation objects
 */
function getMyObservations() {
  const user = getCurrentUser();
  if (!user || !user.teacherId) {
    return { error: 'Not authenticated as teacher' };
  }
  
  return Data.getObservationsByObserver(user.teacherId);
}

// ============================================================================
// SUBSTITUTE COVERAGE
// ============================================================================

/**
 * Create a substitute coverage request
 * @param {Object} observation - Observation requiring coverage
 */
function createSubstituteRequest(observation) {
  const request = {
    id: Utilities.getUuid(),
    observationId: observation.id,
    requesterId: observation.observerId,
    requesterEmail: observation.observerEmail,
    requesterName: observation.observerName,
    date: observation.date,
    periods: observation.periods,
    status: 'pending',
    createdAt: new Date().toISOString()
  };
  
  Data.createSubRequest(request);
  
  // Send email to sub coordinator
  try {
    const dateFormatted = Utilities.formatDate(
      new Date(observation.date + 'T00:00:00'), 
      CONFIG.TIMEZONE, 
      'EEEE, MMMM d, yyyy'
    );
    
    MailApp.sendEmail({
      to: CONFIG.SUB_COORDINATOR_EMAIL,
      subject: `Substitute Coverage Request - ${observation.observerName}`,
      body: `A substitute coverage request has been submitted.\n\n` +
            `Teacher: ${observation.observerName}\n` +
            `Date: ${dateFormatted}\n` +
            `Periods: ${observation.periods.join(', ')}\n\n` +
            `Please review and approve in the Observation System.`
    });
  } catch (e) {
    Logger.log('Failed to send sub request email: ' + e.message);
  }
  
  Data.logAudit({
    action: 'CREATE_SUB_REQUEST',
    userId: observation.observerEmail,
    details: JSON.stringify(request),
    timestamp: new Date().toISOString()
  });
}

// ============================================================================
// VALIDATION
// ============================================================================

/**
 * Validate observation data
 * @param {Object} data - Observation data to validate
 * @param {Object} user - Current user
 * @param {string} excludeId - Observation ID to exclude (for rescheduling)
 * @return {Object} Validation result
 */
function validateObservation(data, user, excludeId = null) {
  // Required fields
  if (!data.teacherId || !data.date || !data.periods || data.periods.length === 0) {
    return { valid: false, message: 'Missing required fields' };
  }
  
  // Can't observe yourself
  if (data.teacherId === user.teacherId) {
    return { valid: false, message: 'You cannot observe yourself' };
  }
  
  // Valid date (not in past, not weekend)
  const obsDate = new Date(data.date + 'T00:00:00');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  if (obsDate < today) {
    return { valid: false, message: 'Cannot schedule observations in the past' };
  }
  
  const dayOfWeek = obsDate.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return { valid: false, message: 'Cannot schedule observations on weekends' };
  }
  
  // Note: Sub requests are allowed regardless of advance notice
  // The sub coordinator will handle short-notice requests as needed
  
  // Check teacher exists
  const teacher = Data.getTeacherById(data.teacherId);
  if (!teacher) {
    return { valid: false, message: 'Selected teacher not found' };
  }
  
  // Check periods are valid (not lunch) - only for classroom teachers
  if (teacher.type !== 'support' && teacher.grade) {
    const lunchPeriods = Data.getLunchPeriods(teacher.grade);
    const hasLunch = data.periods.some(p => lunchPeriods.some(l => l.period === p));
    if (hasLunch) {
      return { valid: false, message: 'Lunch periods cannot be observed' };
    }
  }
  
  // Check slot availability
  const existingObs = Data.getObservationsForTeacherOnDate(data.teacherId, obsDate)
    .filter(o => o.id !== excludeId && o.status !== 'canceled');
  
  for (const period of data.periods) {
    if (existingObs.some(o => o.periods.includes(period))) {
      return { valid: false, message: `Period ${period} already has an observer scheduled` };
    }
  }
  
  // Check observer isn't already observing or being observed
  const observerObs = Data.getObservationsByObserverOnDate(user.teacherId, obsDate)
    .filter(o => o.id !== excludeId && o.status !== 'canceled');
  
  const observerBeingObserved = Data.getObservationsForTeacherOnDate(user.teacherId, obsDate)
    .filter(o => o.id !== excludeId && o.status !== 'canceled');
  
  for (const period of data.periods) {
    if (observerObs.some(o => o.periods.includes(period))) {
      return { valid: false, message: `You already have an observation during period ${period}` };
    }
    if (observerBeingObserved.some(o => o.periods.includes(period))) {
      return { valid: false, message: `You are being observed during period ${period}` };
    }
  }
  
  return { valid: true };
}

// ============================================================================
// ADMIN FUNCTIONS
// ============================================================================

/**
 * Get all observations (admin view)
 * @param {Object} filters - Optional filters (date range, status, etc.)
 * @return {Object[]} Array of observations
 */
function getAllObservations(filters = {}) {
  if (!checkRole('readonly')) {
    return { error: 'Unauthorized' };
  }
  
  return Data.getAllObservations(filters);
}

/**
 * Get pending substitute requests (admin)
 * @return {Object[]} Array of pending requests
 */
function getPendingSubRequests() {
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  return Data.getSubRequestsByStatus('pending');
}

/**
 * Approve a substitute request (admin)
 * @param {string} requestId - Sub request ID
 * @return {Object} Result
 */
function approveSubRequest(requestId) {
  const user = getCurrentUser();
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    const request = Data.getSubRequestById(requestId);
    if (!request) {
      return { error: 'Request not found' };
    }
    
    // Get observation details for the email
    const observation = Data.getObservationById(request.observationId);
    const teacher = observation ? Data.getTeacherById(observation.teacherId) : null;
    
    // Get bell schedule for times
    const scheduleGrade = teacher && teacher.grade ? teacher.grade : 7;
    const bellSchedule = Data.getBellSchedule(scheduleGrade === 6 ? 6 : 7);
    
    // Get period times
    const periods = request.periods || (observation ? observation.periods : []);
    let periodDetails = '';
    let startTime = '';
    let endTime = '';
    
    if (bellSchedule.length > 0 && periods.length > 0) {
      const firstPeriod = bellSchedule.find(b => Number(b.period) === Number(periods[0]));
      const lastPeriod = bellSchedule.find(b => Number(b.period) === Number(periods[periods.length - 1]));
      if (firstPeriod) startTime = firstPeriod.startTime;
      if (lastPeriod) endTime = lastPeriod.endTime;
      
      periodDetails = periods.map(p => {
        const slot = bellSchedule.find(b => Number(b.period) === Number(p));
        return slot ? `Period ${p}: ${slot.startTime} - ${slot.endTime}` : `Period ${p}`;
      }).join('\n           ');
    }
    
    // Format date nicely
    const obsDate = new Date(request.date);
    const dateStr = Utilities.formatDate(obsDate, CONFIG.TIMEZONE, 'EEEE, MMMM d, yyyy');
    
    Data.updateSubRequest(requestId, { 
      status: 'approved',
      approvedBy: user.email,
      approvedAt: new Date().toISOString()
    });
    
    Data.updateObservation(request.observationId, {
      subStatus: 'approved',
      status: 'confirmed'
    });
    
    // Build nice email
    const teacherName = teacher ? teacher.name : (observation ? observation.teacherName : 'Unknown');
    const room = teacher ? teacher.room : (observation ? observation.room : 'TBD');
    
    const emailBody = `
Hello ${request.requesterName},

Great news! Your substitute coverage request has been APPROVED! ✓

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
OBSERVATION DETAILS
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

Date:        ${dateStr}
Teacher:     ${teacherName}
Room:        ${room}
Coverage:    ${periodDetails || `Period(s) ${periods.join(', ')}`}
Time:        ${startTime && endTime ? `${startTime} - ${endTime}` : 'See schedule'}

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

Your observation is now confirmed. A substitute will cover your class during this time.

Approved by: ${user.name || user.email}

Thank you for participating in our peer observation program!

- ${CONFIG.SCHOOL_NAME} Administration
    `.trim();
    
    const emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%); color: white; padding: 24px; text-align: center; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0;">✓ Coverage Approved!</h2>
        </div>
        <div style="padding: 24px; background: #f8f9fa; border-radius: 0 0 8px 8px;">
          <p>Hello ${request.requesterName},</p>
          <p>Great news! Your substitute coverage request has been <strong style="color: #27ae60;">APPROVED</strong>!</p>
          
          <div style="background: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #27ae60;">
            <h3 style="color: #333; margin-top: 0; margin-bottom: 16px;">📋 Observation Details</h3>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 10px 0; color: #666; width: 100px;"><strong>Date:</strong></td>
                <td style="padding: 10px 0; font-weight: 600;">${dateStr}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;"><strong>Teacher:</strong></td>
                <td style="padding: 10px 0;">${teacherName}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;"><strong>Room:</strong></td>
                <td style="padding: 10px 0;">${room}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666; vertical-align: top;"><strong>Coverage:</strong></td>
                <td style="padding: 10px 0;">
                  ${periods.map(p => {
                    const slot = bellSchedule.find(b => Number(b.period) === Number(p));
                    return `<div style="margin-bottom: 4px;">Period ${p}${slot ? `: ${slot.startTime} - ${slot.endTime}` : ''}</div>`;
                  }).join('')}
                </td>
              </tr>
              ${startTime && endTime ? `
              <tr>
                <td style="padding: 10px 0; color: #666;"><strong>Total Time:</strong></td>
                <td style="padding: 10px 0; font-weight: 600; color: #27ae60;">${startTime} - ${endTime}</td>
              </tr>
              ` : ''}
            </table>
          </div>
          
          <p style="background: #e8f5e9; padding: 12px; border-radius: 6px; color: #2e7d32;">
            ✓ Your observation is now <strong>confirmed</strong>. A substitute will cover your class during this time.
          </p>
          
          <p style="color: #666; font-size: 14px;">Approved by: ${user.name || user.email}</p>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
          
          <p style="color: #666; font-size: 12px; text-align: center;">
            ${CONFIG.SCHOOL_NAME} • Classroom Observation System
          </p>
        </div>
      </div>
    `;
    
    // Notify requester
    try {
      MailApp.sendEmail({
        to: request.requesterEmail,
        subject: `✓ Coverage Approved: ${teacherName} on ${dateStr}`,
        body: emailBody,
        htmlBody: emailHtml
      });
    } catch (e) {
      Logger.log('Failed to send approval email: ' + e.message);
    }
    
    Data.logAudit({
      action: 'APPROVE_SUB_REQUEST',
      userId: user.email,
      details: JSON.stringify({ requestId }),
      timestamp: new Date().toISOString()
    });
    
    return { success: true, message: 'Request approved' };
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deny a substitute request (admin)
 * @param {string} requestId - Sub request ID
 * @param {string} reason - Denial reason
 * @return {Object} Result
 */
function denySubRequest(requestId, reason) {
  const user = getCurrentUser();
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    const request = Data.getSubRequestById(requestId);
    if (!request) {
      return { error: 'Request not found' };
    }
    
    // Get observation details for the email
    const observation = Data.getObservationById(request.observationId);
    const teacher = observation ? Data.getTeacherById(observation.teacherId) : null;
    
    // Format date nicely
    const obsDate = new Date(request.date);
    const dateStr = Utilities.formatDate(obsDate, CONFIG.TIMEZONE, 'EEEE, MMMM d, yyyy');
    const teacherName = teacher ? teacher.name : (observation ? observation.teacherName : 'Unknown');
    const periods = request.periods || (observation ? observation.periods : []);
    
    Data.updateSubRequest(requestId, { 
      status: 'denied',
      deniedBy: user.email,
      deniedAt: new Date().toISOString(),
      denyReason: reason
    });
    
    // Cancel the associated observation
    Data.updateObservation(request.observationId, {
      subStatus: 'denied',
      status: 'canceled',
      cancelReason: 'Substitute coverage denied'
    });
    
    // Delete calendar events if they exist
    if (observation) {
      deleteObservationCalendarEvents(observation);
    }
    
    const emailBody = `
Hello ${request.requesterName},

Unfortunately, your substitute coverage request has been denied.

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
REQUEST DETAILS
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

Date:        ${dateStr}
Teacher:     ${teacherName}
Period(s):   ${periods.join(', ')}

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

REASON FOR DENIAL:
${reason || 'No reason provided'}

â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

Your observation has been automatically canceled. You may schedule a new observation that doesn't require substitute coverage, or try again on a different date.

If you have questions, please contact the substitute coordinator.

- ${CONFIG.SCHOOL_NAME} Administration
    `.trim();
    
    const emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%); color: white; padding: 24px; text-align: center; border-radius: 8px 8px 0 0;">
          <h2 style="margin: 0;">Coverage Request Denied</h2>
        </div>
        <div style="padding: 24px; background: #f8f9fa; border-radius: 0 0 8px 8px;">
          <p>Hello ${request.requesterName},</p>
          <p>Unfortunately, your substitute coverage request has been <strong style="color: #e74c3c;">denied</strong>.</p>
          
          <div style="background: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #e74c3c;">
            <h3 style="color: #333; margin-top: 0; margin-bottom: 16px;">📋 Request Details</h3>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 10px 0; color: #666; width: 100px;"><strong>Date:</strong></td>
                <td style="padding: 10px 0;">${dateStr}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;"><strong>Teacher:</strong></td>
                <td style="padding: 10px 0;">${teacherName}</td>
              </tr>
              <tr>
                <td style="padding: 10px 0; color: #666;"><strong>Period(s):</strong></td>
                <td style="padding: 10px 0;">${periods.join(', ')}</td>
              </tr>
            </table>
          </div>
          
          <div style="background: #ffebee; padding: 16px; border-radius: 8px; margin: 20px 0;">
            <strong style="color: #c62828;">Reason for Denial:</strong>
            <p style="margin: 8px 0 0 0; color: #333;">${reason || 'No reason provided'}</p>
          </div>
          
          <p style="color: #666;">Your observation has been automatically canceled. You may schedule a new observation that doesn't require substitute coverage, or try again on a different date.</p>
          
          <p style="color: #666; font-size: 14px;">If you have questions, please contact the substitute coordinator.</p>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
          
          <p style="color: #666; font-size: 12px; text-align: center;">
            ${CONFIG.SCHOOL_NAME} • Classroom Observation System
          </p>
        </div>
      </div>
    `;
    
    // Notify requester
    try {
      MailApp.sendEmail({
        to: request.requesterEmail,
        subject: `Coverage Denied: ${teacherName} on ${dateStr}`,
        body: emailBody,
        htmlBody: emailHtml
      });
    } catch (e) {
      Logger.log('Failed to send denial email: ' + e.message);
    }
    
    Data.logAudit({
      action: 'DENY_SUB_REQUEST',
      userId: user.email,
      details: JSON.stringify({ requestId, reason }),
      timestamp: new Date().toISOString()
    });
    
    return { success: true, message: 'Request denied and observation canceled' };
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Admin edit observation
 * @param {string} observationId - Observation ID
 * @param {Object} updates - Fields to update
 * @return {Object} Result
 */
function adminEditObservation(observationId, updates) {
  const user = getCurrentUser();
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    const result = Data.updateObservation(observationId, {
      ...updates,
      modifiedAt: new Date().toISOString(),
      modifiedBy: user.email
    });
    
    if (result.success) {
      Data.logAudit({
        action: 'ADMIN_EDIT_OBSERVATION',
        userId: user.email,
        details: JSON.stringify({ observationId, updates }),
        timestamp: new Date().toISOString()
      });
      
      return { success: true };
    } else {
      return { error: result.error || 'Failed to update' };
    }
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Admin delete observation
 * @param {string} observationId - Observation ID
 * @return {Object} Result
 */
function adminDeleteObservation(observationId) {
  const user = getCurrentUser();
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return { error: 'System busy. Please try again.' };
    }
    
    const result = Data.deleteObservation(observationId);
    
    if (result.success) {
      Data.logAudit({
        action: 'ADMIN_DELETE_OBSERVATION',
        userId: user.email,
        details: JSON.stringify({ observationId }),
        timestamp: new Date().toISOString()
      });
      
      return { success: true };
    } else {
      return { error: result.error || 'Failed to delete' };
    }
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Get audit log (admin)
 * @param {number} limit - Number of entries to return
 * @return {Object[]} Audit log entries
 */
function getAuditLog(limit = 100) {
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  return Data.getAuditLog(limit);
}

// ============================================================================
// ACCESS REQUEST SYSTEM
// ============================================================================

/**
 * Generate access request page HTML
 * @param {string} email - User's email
 * @return {string} HTML content
 */
function getAccessRequestPage(email) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
               background: linear-gradient(135deg, #F5EEF8 0%, #E8DAEF 100%); 
               min-height: 100vh; display: flex; align-items: center; justify-content: center; margin: 0; }
        .container { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 10px 40px rgba(0,0,0,0.1); 
                     max-width: 450px; width: 90%; text-align: center; }
        h1 { color: #5B2C6F; margin-bottom: 8px; }
        .subtitle { color: #666; margin-bottom: 24px; }
        .email { background: #F5EEF8; padding: 12px; border-radius: 6px; font-weight: 500; margin-bottom: 24px; }
        input, select { width: 100%; padding: 12px; margin-bottom: 16px; border: 2px solid #E8DAEF; 
                        border-radius: 8px; font-size: 1rem; box-sizing: border-box; }
        input:focus, select:focus { outline: none; border-color: #5B2C6F; }
        button { background: linear-gradient(135deg, #5B2C6F 0%, #4A235A 100%); color: white; 
                 padding: 14px 28px; border: none; border-radius: 8px; font-size: 1rem; 
                 cursor: pointer; width: 100%; font-weight: 500; }
        button:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(91,44,111,0.3); }
        .success { background: #dcfce7; color: #166534; padding: 16px; border-radius: 8px; display: none; }
        .error { background: #fee2e2; color: #dc2626; padding: 16px; border-radius: 8px; display: none; margin-bottom: 16px; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>Request Access</h1>
        <p class="subtitle">${CONFIG.SCHOOL_NAME} Observation System</p>
        <div class="email">${email}</div>
        <div class="error" id="error"></div>
        <form id="requestForm">
          <input type="text" id="name" placeholder="Your Full Name" required>
          <input type="text" id="room" placeholder="Room Number (optional)">
          <select id="grade">
            <option value="">Select Primary Grade Level</option>
            ${CONFIG.GRADES.map(g => `<option value="${g}">Grade ${g}</option>`).join('')}
          </select>
          <select id="role">
            <option value="teacher">Teacher</option>
            <option value="readonly">Read-Only Access</option>
          </select>
          <button type="submit">Submit Request</button>
        </form>
        <div class="success" id="success">
          <strong>Request Submitted!</strong><br>
          An administrator will review your request and grant access soon.
        </div>
      </div>
      <script>
        document.getElementById('requestForm').onsubmit = function(e) {
          e.preventDefault();
          const data = {
            email: '${email}',
            name: document.getElementById('name').value,
            room: document.getElementById('room').value,
            grade: document.getElementById('grade').value,
            role: document.getElementById('role').value
          };
          google.script.run
            .withSuccessHandler(function(result) {
              if (result.success) {
                document.getElementById('requestForm').style.display = 'none';
                document.getElementById('success').style.display = 'block';
              } else {
                document.getElementById('error').textContent = result.error || 'Failed to submit request';
                document.getElementById('error').style.display = 'block';
              }
            })
            .withFailureHandler(function(err) {
              document.getElementById('error').textContent = 'Error: ' + err.message;
              document.getElementById('error').style.display = 'block';
            })
            .submitAccessRequest(data);
        };
      </script>
    </body>
    </html>
  `;
}

/**
 * Submit an access request
 * @param {Object} data - Request data {email, name, room, grade, role}
 * @return {Object} Result
 */
function submitAccessRequest(data) {
  if (!data.email || !data.name) {
    return { error: 'Name and email are required' };
  }
  
  // Validate email domain if configured
  if (CONFIG.ALLOWED_EMAIL_DOMAIN && !data.email.endsWith('@' + CONFIG.ALLOWED_EMAIL_DOMAIN)) {
    return { error: 'Invalid email domain' };
  }
  
  // Check if already registered
  const existingUser = Data.findUserByEmail(data.email);
  if (existingUser) {
    return { error: 'You already have access. Try refreshing the page.' };
  }
  
  // Check if request already exists
  const existingRequest = Data.getAccessRequest(data.email);
  if (existingRequest && existingRequest.status === 'pending') {
    return { error: 'You already have a pending request' };
  }
  
  // Create access request
  const request = {
    id: 'req-' + Date.now(),
    email: data.email,
    name: data.name,
    room: data.room || '',
    grade: data.grade || '',
    requestedRole: data.role || 'teacher',
    status: 'pending',
    createdAt: new Date().toISOString()
  };
  
  Data.createAccessRequest(request);
  
  // Notify admins
  notifyAdminsOfAccessRequest(request);
  
  return { success: true };
}

/**
 * Get pending access requests (admin only)
 * @return {Object[]} Pending requests
 */
function getPendingAccessRequests() {
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  return Data.getAccessRequestsByStatus('pending');
}

/**
 * Approve an access request (admin only)
 * @param {string} requestId - Request ID
 * @return {Object} Result
 */
function approveAccessRequest(requestId) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const request = Data.getAccessRequestById(requestId);
  if (!request) {
    return { error: 'Request not found' };
  }
  
  // Add user to appropriate sheet
  if (request.requestedRole === 'teacher') {
    const teacherData = {
      id: 't-' + Date.now(),
      email: request.email,
      name: request.name,
      room: request.room,
      grade: request.grade,
      grades: request.grade ? JSON.stringify([Number(request.grade)]) : '[]',
      type: 'classroom',
      active: 'TRUE',
      prepPeriods: '',
      lunchPeriod: ''
    };
    Data.createTeacher(teacherData);
  } else if (request.requestedRole === 'readonly') {
    Data.addReadOnlyUser({ id: 'ro-' + Date.now(), email: request.email, name: request.name });
  }
  
  // Update request status
  Data.updateAccessRequest(requestId, { status: 'approved', approvedBy: user.email, approvedAt: new Date().toISOString() });
  
  // Log audit
  Data.logAudit({
    action: 'APPROVE_ACCESS_REQUEST',
    userId: user.email,
    details: JSON.stringify({ requestId, email: request.email, role: request.requestedRole }),
    timestamp: new Date().toISOString()
  });
  
  return { success: true, message: `Access granted to ${request.name}` };
}

/**
 * Deny an access request (admin only)
 * @param {string} requestId - Request ID
 * @param {string} reason - Denial reason
 * @return {Object} Result
 */
function denyAccessRequest(requestId, reason) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const request = Data.getAccessRequestById(requestId);
  if (!request) {
    return { error: 'Request not found' };
  }
  
  Data.updateAccessRequest(requestId, { 
    status: 'denied', 
    deniedBy: user.email, 
    deniedAt: new Date().toISOString(),
    denyReason: reason || ''
  });
  
  Data.logAudit({
    action: 'DENY_ACCESS_REQUEST',
    userId: user.email,
    details: JSON.stringify({ requestId, email: request.email, reason }),
    timestamp: new Date().toISOString()
  });
  
  return { success: true };
}

/**
 * Notify admins of new access request
 * @param {Object} request - The access request
 */
function notifyAdminsOfAccessRequest(request) {
  try {
    const admins = Data.getAllAdmins();
    if (!admins || admins.length === 0) return;
    
    const subject = `[${CONFIG.SCHOOL_ABBR}] New Access Request: ${request.name}`;
    const body = `
      <div style="font-family: Arial, sans-serif; max-width: 500px;">
        <h2 style="color: #5B2C6F;">New Access Request</h2>
        <p><strong>Name:</strong> ${request.name}</p>
        <p><strong>Email:</strong> ${request.email}</p>
        <p><strong>Room:</strong> ${request.room || 'Not specified'}</p>
        <p><strong>Grade:</strong> ${request.grade || 'Not specified'}</p>
        <p><strong>Requested Role:</strong> ${request.requestedRole}</p>
        <p style="margin-top: 20px;">Log in to the Observation System to approve or deny this request.</p>
      </div>
    `;
    
    admins.forEach(admin => {
      if (admin.email) {
        MailApp.sendEmail({
          to: admin.email,
          subject: subject,
          htmlBody: body
        });
      }
    });
  } catch (e) {
    Logger.log('Failed to notify admins: ' + e.message);
  }
}

// ============================================================================
// ARCHIVE & RESET SYSTEM
// ============================================================================

/**
 * Archive current year's data and reset for new year (admin only)
 * @param {boolean} confirmReset - Must be true to proceed
 * @return {Object} Result
 */
function archiveAndResetYear(confirmReset) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  if (confirmReset !== true) {
    return { error: 'Please confirm the reset by passing true' };
  }
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const schoolYear = getSchoolYear();
  
  try {
    // 1. Archive Observations sheet
    const obsSheet = ss.getSheetByName('Observations');
    if (obsSheet && obsSheet.getLastRow() > 1) {
      const archiveName = `Archive_Observations_${schoolYear}`;
      
      // Check if archive already exists
      let archiveSheet = ss.getSheetByName(archiveName);
      if (!archiveSheet) {
        archiveSheet = obsSheet.copyTo(ss);
        archiveSheet.setName(archiveName);
      }
      
      // Clear observations (keep headers)
      if (obsSheet.getLastRow() > 1) {
        obsSheet.deleteRows(2, obsSheet.getLastRow() - 1);
      }
    }
    
    // 2. Archive SubstituteRequests sheet
    const subSheet = ss.getSheetByName('SubstituteRequests');
    if (subSheet && subSheet.getLastRow() > 1) {
      const archiveName = `Archive_SubRequests_${schoolYear}`;
      
      let archiveSheet = ss.getSheetByName(archiveName);
      if (!archiveSheet) {
        archiveSheet = subSheet.copyTo(ss);
        archiveSheet.setName(archiveName);
      }
      
      if (subSheet.getLastRow() > 1) {
        subSheet.deleteRows(2, subSheet.getLastRow() - 1);
      }
    }
    
    // 3. Archive and clear AccessRequests if it exists
    const accessSheet = ss.getSheetByName('AccessRequests');
    if (accessSheet && accessSheet.getLastRow() > 1) {
      const archiveName = `Archive_AccessReq_${schoolYear}`;
      
      let archiveSheet = ss.getSheetByName(archiveName);
      if (!archiveSheet) {
        archiveSheet = accessSheet.copyTo(ss);
        archiveSheet.setName(archiveName);
      }
      
      if (accessSheet.getLastRow() > 1) {
        accessSheet.deleteRows(2, accessSheet.getLastRow() - 1);
      }
    }
    
    // 4. Log the archive action
    Data.logAudit({
      action: 'ARCHIVE_AND_RESET_YEAR',
      userId: user.email,
      details: JSON.stringify({ schoolYear, archivedAt: new Date().toISOString() }),
      timestamp: new Date().toISOString()
    });
    
    return { 
      success: true, 
      message: `Successfully archived ${schoolYear} data! The system is ready for the new school year.`
    };
    
  } catch (e) {
    Logger.log('Archive error: ' + e.message);
    return { error: 'Archive failed: ' + e.message };
  }
}

/**
 * Get archive status and list of archived years (admin only)
 * @return {Object} Archive info
 */
function getArchiveStatus() {
  if (!checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheets = ss.getSheets();
  
  const archives = sheets
    .filter(s => s.getName().startsWith('Archive_'))
    .map(s => ({
      name: s.getName(),
      rows: s.getLastRow() - 1
    }));
  
  const obsSheet = ss.getSheetByName('Observations');
  const currentObservations = obsSheet ? Math.max(0, obsSheet.getLastRow() - 1) : 0;
  
  return {
    currentSchoolYear: getSchoolYear(),
    currentObservations: currentObservations,
    archives: archives,
    canArchive: currentObservations > 0
  };
}

// ============================================================================
// BULK TEACHER IMPORT
// ============================================================================

/**
 * Bulk import teachers from parsed data (admin only)
 * @param {Object[]} teachers - Array of teacher objects
 * @return {Object} Result with counts
 */
function bulkImportTeachers(teachers) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  if (!Array.isArray(teachers) || teachers.length === 0) {
    return { error: 'No teachers provided' };
  }
  
  let added = 0;
  let skipped = 0;
  let errors = [];
  
  teachers.forEach((t, index) => {
    try {
      // Validate required fields
      if (!t.email || !t.name) {
        errors.push(`Row ${index + 1}: Missing email or name`);
        skipped++;
        return;
      }
      
      // Check if already exists
      const existing = Data.findTeacherByEmail(t.email);
      if (existing) {
        skipped++;
        return;
      }
      
      // Validate email domain if configured
      if (CONFIG.ALLOWED_EMAIL_DOMAIN && !t.email.endsWith('@' + CONFIG.ALLOWED_EMAIL_DOMAIN)) {
        errors.push(`Row ${index + 1}: Invalid email domain`);
        skipped++;
        return;
      }
      
      // Create teacher
      const teacherData = {
        id: 't-' + Date.now() + '-' + index,
        email: t.email.trim().toLowerCase(),
        name: t.name.trim(),
        room: t.room || '',
        grade: t.grade || '',
        grades: t.grade ? JSON.stringify([Number(t.grade)]) : '[]',
        type: t.type || 'classroom',
        active: 'TRUE',
        prepPeriods: '',
        lunchPeriod: ''
      };
      
      Data.createTeacher(teacherData);
      added++;
      
    } catch (e) {
      errors.push(`Row ${index + 1}: ${e.message}`);
      skipped++;
    }
  });
  
  Data.logAudit({
    action: 'BULK_IMPORT_TEACHERS',
    userId: user.email,
    details: JSON.stringify({ added, skipped, errors: errors.length }),
    timestamp: new Date().toISOString()
  });
  
  return {
    success: true,
    added: added,
    skipped: skipped,
    errors: errors.slice(0, 10), // Return first 10 errors
    message: `Added ${added} teacher(s), skipped ${skipped}`
  };
}

/**
 * Parse CSV/TSV text into teacher objects
 * @param {string} text - Raw text (CSV or TSV)
 * @return {Object[]} Parsed teachers
 */
function parseTeacherImportText(text) {
  if (!text || typeof text !== 'string') {
    return { error: 'No text provided' };
  }
  
  const lines = text.trim().split('\n');
  if (lines.length < 2) {
    return { error: 'Need at least a header row and one data row' };
  }
  
  // Detect delimiter
  const delimiter = lines[0].includes('\t') ? '\t' : ',';
  
  // Parse header
  const headers = lines[0].split(delimiter).map(h => h.trim().toLowerCase());
  
  // Map common header variations
  const headerMap = {
    'email': ['email', 'e-mail', 'email address'],
    'name': ['name', 'full name', 'teacher name', 'teacher'],
    'room': ['room', 'room number', 'room #', 'classroom'],
    'grade': ['grade', 'grade level', 'grade_level']
  };
  
  const columnIndex = {};
  for (const [field, variations] of Object.entries(headerMap)) {
    const idx = headers.findIndex(h => variations.includes(h));
    if (idx !== -1) columnIndex[field] = idx;
  }
  
  if (columnIndex.email === undefined || columnIndex.name === undefined) {
    return { error: 'Could not find "email" and "name" columns in header' };
  }
  
  // Parse data rows
  const teachers = [];
  for (let i = 1; i < lines.length; i++) {
    const values = lines[i].split(delimiter).map(v => v.trim().replace(/^["']|["']$/g, ''));
    
    if (values.length < 2 || !values[columnIndex.email]) continue;
    
    teachers.push({
      email: values[columnIndex.email],
      name: values[columnIndex.name],
      room: columnIndex.room !== undefined ? values[columnIndex.room] : '',
      grade: columnIndex.grade !== undefined ? values[columnIndex.grade] : ''
    });
  }
  
  return { teachers, count: teachers.length };
}

// ============================================================================
// BRANDING SETTINGS
// ============================================================================

/**
 * Update school branding (admin only)
 * This updates the CONFIG.SCHOOL_ABBR in the script
 * @param {string} newBranding - New branding text
 * @return {Object} Result
 */
function updateBranding(newBranding) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized' };
  }
  
  if (!newBranding || typeof newBranding !== 'string') {
    return { error: 'Invalid branding text' };
  }
  
  // Sanitize input
  newBranding = newBranding.trim();
  
  if (newBranding.length === 0) {
    return { error: 'Branding text cannot be empty' };
  }
  
  if (newBranding.length > 20) {
    return { error: 'Branding text must be 20 characters or less' };
  }
  
  try {
    // Store branding in a Settings sheet (create if doesn't exist)
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('Settings');
      settingsSheet.appendRow(['key', 'value', 'updatedAt', 'updatedBy']);
    }
    
    // Find or create branding setting
    const data = settingsSheet.getDataRange().getValues();
    let brandingRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'SCHOOL_ABBR') {
        brandingRow = i + 1;
        break;
      }
    }
    
    const timestamp = new Date().toISOString();
    
    if (brandingRow > 0) {
      // Update existing row
      settingsSheet.getRange(brandingRow, 2, 1, 3).setValues([[newBranding, timestamp, user.email]]);
    } else {
      // Add new row
      settingsSheet.appendRow(['SCHOOL_ABBR', newBranding, timestamp, user.email]);
    }
    
    // Log the change
    Data.logAudit({
      action: 'UPDATE_BRANDING',
      userId: user.email,
      details: JSON.stringify({ newBranding, previousBranding: CONFIG.SCHOOL_ABBR }),
      timestamp: timestamp
    });
    
    return { 
      success: true, 
      newBranding: newBranding,
      message: 'Branding updated! Create a new deployment for changes to take effect for all users.'
    };
    
  } catch (e) {
    Logger.log('Branding update error: ' + e.message);
    return { error: 'Failed to update branding: ' + e.message };
  }
}

/**
 * Get current branding from Settings sheet (or CONFIG default)
 * @return {string} Current branding
 */
function getCurrentBranding() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'SCHOOL_ABBR' && data[i][1]) {
          return data[i][1];
        }
      }
    }
  } catch (e) {
    Logger.log('Error getting branding: ' + e.message);
  }
  
  return CONFIG.SCHOOL_ABBR;
}

// ============================================================================
// DASHBOARD DATA
// ============================================================================

/**
 * Get dashboard statistics
 * @return {Object} Dashboard data
 */
function getDashboardData() {
  if (!checkRole('readonly')) {
    return { error: 'Unauthorized' };
  }
  
  const now = new Date();
  const schoolYearStart = getSchoolYearStart();
  const deadline = getObservationDeadline();
  
  const startOfWeek = new Date(now);
  startOfWeek.setDate(now.getDate() - now.getDay());
  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6);
  
  const confirmedObs = Data.getAllObservations({ status: 'confirmed' });
  const pendingSubObs = Data.getAllObservations({ status: 'pending_sub' });
  const allObs = confirmedObs.concat(pendingSubObs);

  // School year observations (for requirement tracking)
  const yearObs = allObs.filter(o => {
    const d = new Date(o.date);
    return d >= schoolYearStart && d <= deadline;
  });
  const weekObs = allObs.filter(o => {
    const d = new Date(o.date);
    return d >= startOfWeek && d <= endOfWeek;
  });
  const todayObs = allObs.filter(o => {
    const d = new Date(o.date);
    const today = new Date();
    return d.toDateString() === today.toDateString();
  });
  
  const pendingSubs = Data.getSubRequestsByStatus('pending');
  
  // Teachers who haven't observed this school year
  const teachers = Data.getAllTeachers();
  const teachersWhoObserved = new Set(yearObs.map(o => o.observerId));
  const teachersNotObserved = teachers.filter(t => !teachersWhoObserved.has(t.id));
  
  return {
    totalObservations: allObs.length,
    monthlyObservations: yearObs.length,  // Now represents school year total
    weeklyObservations: weekObs.length,
    todayObservations: todayObs.length,
    pendingSubRequests: pendingSubs.length,
    teachersNotObservedThisMonth: teachersNotObserved.length,
    teachersNotObservedList: teachersNotObserved.map(t => ({ 
      id: t.id, 
      name: t.name, 
      grade: t.grade,
      grades: t.grades || [t.grade]
    })),
    todaySchedule: todayObs.map(o => ({
      time: Array.isArray(o.periods) ? o.periods.join(', ') : o.periods,
      observer: o.observerName,
      teacher: o.teacherName,
      room: o.room
    })),
    deadline: Utilities.formatDate(deadline, CONFIG.TIMEZONE, 'MMMM d, yyyy')
  };
}

// ============================================================================
// INITIALIZATION
// ============================================================================

/**
 * Initialize spreadsheet with required sheets
 * Run this once to set up the database
 */
function initializeSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheets = {
    'Teachers': ['id', 'email', 'name', 'grade', 'room', 'department', 'type', 'active', 'unavailablePeriods'],
    'Rooms': ['id', 'roomNumber', 'grade', 'capacity', 'type'],
    'BellSchedules': ['id', 'grade', 'period', 'startTime', 'endTime', 'type'],
    'Observations': ['id', 'observerId', 'observerEmail', 'observerName', 'teacherId', 'teacherName', 'room', 'grade', 'date', 'periods', 'needsSub', 'subStatus', 'status', 'createdAt', 'createdBy', 'modifiedAt', 'modifiedBy', 'canceledAt', 'canceledBy', 'cancelReason', 'rescheduledAt', 'rescheduledBy', 'observerCalendarEventId', 'teacherCalendarEventId'],
    'SubstituteRequests': ['id', 'observationId', 'requesterId', 'requesterEmail', 'requesterName', 'date', 'periods', 'status', 'createdAt', 'approvedBy', 'approvedAt', 'deniedBy', 'deniedAt', 'denyReason'],
    'Admins': ['email', 'name', 'addedAt'],
    'ReadOnlyUsers': ['email', 'name', 'addedAt'],
    'AuditLog': ['id', 'action', 'userId', 'details', 'timestamp']
  };
  
  for (const [sheetName, headers] of Object.entries(sheets)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  
  Logger.log('Spreadsheet initialized with all required sheets');
}

/**
 * Add sample data for testing
 */
function addSampleData() {
  // Sample bell schedules
  const bell6 = [
    { grade: 6, period: 1, startTime: '7:45', endTime: '8:35', type: 'class' },
    { grade: 6, period: 2, startTime: '8:39', endTime: '9:29', type: 'class' },
    { grade: 6, period: 3, startTime: '9:33', endTime: '10:23', type: 'class' },
    { grade: 6, period: 4, startTime: '10:27', endTime: '11:17', type: 'lunch' },
    { grade: 6, period: 5, startTime: '11:21', endTime: '12:11', type: 'class' },
    { grade: 6, period: 6, startTime: '12:15', endTime: '1:05', type: 'class' },
    { grade: 6, period: 7, startTime: '1:09', endTime: '1:59', type: 'class' },
    { grade: 6, period: 8, startTime: '2:03', endTime: '2:50', type: 'class' }
  ];
  
  const bell78 = [
    { grade: 7, period: 1, startTime: '7:45', endTime: '8:35', type: 'class' },
    { grade: 7, period: 2, startTime: '8:39', endTime: '9:29', type: 'class' },
    { grade: 7, period: 3, startTime: '9:33', endTime: '10:23', type: 'class' },
    { grade: 7, period: 4, startTime: '10:27', endTime: '11:17', type: 'class' },
    { grade: 7, period: 5, startTime: '11:21', endTime: '12:11', type: 'lunch' },
    { grade: 7, period: 6, startTime: '12:15', endTime: '1:05', type: 'class' },
    { grade: 7, period: 7, startTime: '1:09', endTime: '1:59', type: 'class' },
    { grade: 7, period: 8, startTime: '2:03', endTime: '2:50', type: 'class' }
  ];
  
  // Add to sheets
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bellSheet = ss.getSheetByName('BellSchedules');
  
  bell6.forEach(b => {
    bellSheet.appendRow([Utilities.getUuid(), b.grade, b.period, b.startTime, b.endTime, b.type]);
  });
  
  bell78.forEach(b => {
    bellSheet.appendRow([Utilities.getUuid(), b.grade, b.period, b.startTime, b.endTime, b.type]);
  });
  
  // Add lunch periods
  const lunchSheet = ss.getSheetByName('LunchPeriods');
  lunchSheet.appendRow([Utilities.getUuid(), 6, 4]);
  lunchSheet.appendRow([Utilities.getUuid(), 7, 5]);
  lunchSheet.appendRow([Utilities.getUuid(), 8, 5]);
  
  // Add prep periods (example)
  const prepSheet = ss.getSheetByName('PrepPeriods');
  prepSheet.appendRow([Utilities.getUuid(), 6, 7]);
  prepSheet.appendRow([Utilities.getUuid(), 7, 8]);
  prepSheet.appendRow([Utilities.getUuid(), 8, 8]);
  
  Logger.log('Sample data added');
}

// ============================================================================
// OBSERVATION FORM & NOTIFICATION SYSTEM
// ============================================================================

/**
 * Generate a force-copy link for the observation form
 * @return {string} Force copy URL
 */
function getObservationFormCopyLink() {
  return 'https://docs.google.com/document/d/' + CONFIG.OBSERVATION_FORM_ID + '/copy';
}

/**
 * Send observation confirmation with form link, email, and calendar invite
 * @param {string} observationId - The observation ID
 * @return {Object} Result with success/error
 */
function sendObservationConfirmation(observationId) {
  const user = getCurrentUser();
  if (!user) {
    return { error: 'Unauthorized' };
  }
  
  try {
    // Get observation details
    const observation = Data.getObservationById(observationId);
    if (!observation) {
      return { error: 'Observation not found' };
    }
    
    // Get teacher being observed
    const teacher = Data.getTeacherById(observation.teacherId);
    if (!teacher) {
      return { error: 'Teacher not found' };
    }
    
    // Get observer info
    const observer = Data.getTeacherById(observation.observerId);
    if (!observer) {
      return { error: 'Observer not found' };
    }
    
    // Generate form copy link
    const formLink = getObservationFormCopyLink();
    
    // Get bell schedule for timing
    const scheduleGrade = teacher.grade || 7;
    const bellSchedule = Data.getBellSchedule(scheduleGrade === 6 ? 6 : 7);
    
    // Find start and end times based on periods
    const periods = observation.periods || [];
    let startTime = '8:15 AM';
    let endTime = '9:00 AM';
    
    if (bellSchedule.length > 0 && periods.length > 0) {
      const firstPeriod = bellSchedule.find(b => b.period === periods[0]);
      const lastPeriod = bellSchedule.find(b => b.period === periods[periods.length - 1]);
      if (firstPeriod) startTime = firstPeriod.startTime;
      if (lastPeriod) endTime = lastPeriod.endTime;
    }
    
    // Format the date
    const obsDate = new Date(observation.date);
    const dateStr = Utilities.formatDate(obsDate, CONFIG.TIMEZONE, 'EEEE, MMMM d, yyyy');
    
    // Build email body
    const emailBody = `
Hello ${observer.name},

Your classroom observation has been confirmed!

OBSERVATION DETAILS:
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”
Teacher:     ${teacher.name}
Room:        ${teacher.room}
Date:        ${dateStr}
Period(s):   ${periods.join(', ')}
Time:        ${startTime} - ${endTime}
â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”â”

OBSERVATION FORM:
Click the link below to get your copy of the observation form:
${formLink}

A calendar invite has been sent to your email.

Thank you for participating in our peer observation program!

- ${CONFIG.SCHOOL_NAME} Administration
    `.trim();
    
    // Build HTML email
    const emailHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: #5B2C6F; color: white; padding: 20px; text-align: center;">
          <h2 style="margin: 0;">Observation Confirmed!</h2>
        </div>
        <div style="padding: 20px; background: #f8f9fa;">
          <p>Hello ${observer.name},</p>
          <p>Your classroom observation has been confirmed!</p>
          
          <div style="background: white; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #D4AC0D;">
            <h3 style="color: #5B2C6F; margin-top: 0;">Observation Details</h3>
            <table style="width: 100%;">
              <tr><td style="padding: 8px 0; color: #666;"><strong>Teacher:</strong></td><td>${teacher.name}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;"><strong>Room:</strong></td><td>${teacher.room}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;"><strong>Date:</strong></td><td>${dateStr}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;"><strong>Period(s):</strong></td><td>${periods.join(', ')}</td></tr>
              <tr><td style="padding: 8px 0; color: #666;"><strong>Time:</strong></td><td>${startTime} - ${endTime}</td></tr>
            </table>
          </div>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${formLink}" style="display: inline-block; background: #D4AC0D; color: #000; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold;">
              📋 Get Your Observation Form
            </a>
          </div>
          
          <p style="color: #666; font-size: 14px;">A calendar invite has been added to your calendar.</p>
          
          <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
          
          <p style="color: #666; font-size: 12px; text-align: center;">
            ${CONFIG.SCHOOL_NAME} • Classroom Observation System
          </p>
        </div>
      </div>
    `;
    
    // Send email
    MailApp.sendEmail({
      to: observer.email,
      subject: `Observation Confirmed: ${teacher.name} on ${dateStr}`,
      body: emailBody,
      htmlBody: emailHtml
    });
    
    // Create calendar event
    createObservationCalendarEvent(observation, observer, teacher, obsDate, startTime, endTime, formLink);
    
    // Log the action
    Data.appendRow('AuditLog', {
      id: Utilities.getUuid(),
      timestamp: new Date().toISOString(),
      userId: user.email,
      action: 'SENT_CONFIRMATION: ' + observationId,
      details: 'Email and calendar invite sent to ' + observer.email
    });
    
    return { 
      success: true, 
      message: 'Confirmation sent!',
      formLink: formLink
    };
    
  } catch (e) {
    Logger.log('Error sending confirmation: ' + e.message);
    return { error: 'Failed to send confirmation: ' + e.message };
  }
}

/**
 * Create calendar events for both observer and teacher being observed
 * @param {Object} observation - The observation record
 * @param {Object} observer - The observer's teacher record
 * @param {Object} teacher - The teacher being observed
 * @param {Date} obsDate - The observation date
 * @param {string} startTime - Start time string
 * @param {string} endTime - End time string
 * @param {string} formLink - Link to observation form
 */
function createObservationCalendarEvent(observation, observer, teacher, obsDate, startTime, endTime, formLink) {
  try {
    // Parse times
    const startDateTime = parseTimeToDate(obsDate, startTime);
    const endDateTime = parseTimeToDate(obsDate, endTime);
    
    const periodStr = (observation.periods || []).join(', ');
    let observerEventId = null;
    let teacherEventId = null;
    
    // ========================================
    // Create event for OBSERVER
    // ========================================
    try {
      const observerCalendar = CalendarApp.getDefaultCalendar();
      
      const observerEvent = observerCalendar.createEvent(
        '👀 Observing: ' + teacher.name + ' (Room ' + teacher.room + ')',
        startDateTime,
        endDateTime,
        {
          description: 'Classroom Observation\n\n' +
            '📍 You are OBSERVING this class\n\n' +
            'Teacher: ' + teacher.name + '\n' +
            'Room: ' + teacher.room + '\n' +
            'Period(s): ' + periodStr + '\n\n' +
            '📋 Observation Form:\n' + formLink + '\n\n' +
            '---\n' +
            'Created by Classroom Observation System',
          location: 'Room ' + teacher.room,
          guests: teacher.email,
          sendInvites: false // We'll send a separate invite to teacher
        }
      );
      
      // Set reminders for observer
      observerEvent.addPopupReminder(15); // 15 minutes before
      observerEvent.addPopupReminder(60); // 1 hour before
      
      observerEventId = observerEvent.getId();
      Logger.log('Observer calendar event created: ' + observerEventId);
      
    } catch (e) {
      Logger.log('Error creating observer calendar event: ' + e.message);
    }
    
    // ========================================
    // Create event for TEACHER being observed
    // ========================================
    try {
      // Try to create event on teacher's calendar by sending them an invite
      // We create it on the current user's calendar and invite the teacher
      const calendar = CalendarApp.getDefaultCalendar();
      
      const teacherEvent = calendar.createEvent(
        '📋 Being Observed by: ' + observer.name,
        startDateTime,
        endDateTime,
        {
          description: 'Classroom Observation\n\n' +
            '👀 You are BEING OBSERVED during this period\n\n' +
            'Observer: ' + observer.name + '\n' +
            'Your Room: ' + teacher.room + '\n' +
            'Period(s): ' + periodStr + '\n\n' +
            '---\n' +
            'Created by Classroom Observation System',
          location: 'Room ' + teacher.room,
          guests: teacher.email,
          sendInvites: true // Send invite to teacher
        }
      );
      
      teacherEventId = teacherEvent.getId();
      Logger.log('Teacher calendar event created: ' + teacherEventId);
      
    } catch (e) {
      Logger.log('Error creating teacher calendar event: ' + e.message);
    }
    
    // ========================================
    // Store event IDs in observation record
    // ========================================
    if (observerEventId || teacherEventId) {
      Data.updateObservation(observation.id, {
        observerCalendarEventId: observerEventId || '',
        teacherCalendarEventId: teacherEventId || ''
      });
      Logger.log('Stored calendar event IDs for observation: ' + observation.id);
    }
    
  } catch (e) {
    Logger.log('Error in createObservationCalendarEvent: ' + e.message);
    // Don't fail the whole process if calendar fails
  }
}

/**
 * Delete calendar events for an observation (called when canceling)
 * @param {Object} observation - The observation record with calendar event IDs
 */
function deleteObservationCalendarEvents(observation) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Delete observer's event
    if (observation.observerCalendarEventId) {
      try {
        const observerEvent = calendar.getEventById(observation.observerCalendarEventId);
        if (observerEvent) {
          observerEvent.deleteEvent();
          Logger.log('Deleted observer calendar event: ' + observation.observerCalendarEventId);
        }
      } catch (e) {
        Logger.log('Could not delete observer event: ' + e.message);
      }
    }
    
    // Delete teacher's event
    if (observation.teacherCalendarEventId) {
      try {
        const teacherEvent = calendar.getEventById(observation.teacherCalendarEventId);
        if (teacherEvent) {
          teacherEvent.deleteEvent();
          Logger.log('Deleted teacher calendar event: ' + observation.teacherCalendarEventId);
        }
      } catch (e) {
        Logger.log('Could not delete teacher event: ' + e.message);
      }
    }
    
  } catch (e) {
    Logger.log('Error deleting calendar events: ' + e.message);
    // Don't fail the cancellation if calendar deletion fails
  }
}

/**
 * Parse a time string and date into a Date object
 */
function parseTimeToDate(baseDate, timeStr) {
  const date = new Date(baseDate);
  
  // Handle various time formats: "8:15", "8:15 AM", "14:30"
  let hours = 8;
  let minutes = 0;
  
  if (timeStr) {
    const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (timeMatch) {
      hours = parseInt(timeMatch[1]);
      minutes = parseInt(timeMatch[2]);
      
      const ampm = timeMatch[3];
      if (ampm) {
        if (ampm.toUpperCase() === 'PM' && hours !== 12) {
          hours += 12;
        } else if (ampm.toUpperCase() === 'AM' && hours === 12) {
          hours = 0;
        }
      } else {
        // No AM/PM - assume school hours
        // If hour is less than 7, it's probably PM
        if (hours < 7) {
          hours += 12;
        }
      }
    }
  }
  
  date.setHours(hours, minutes, 0, 0);
  return date;
}

/**
 * Get observation form link (called from frontend)
 * @return {Object} Object with form link
 */
function getObservationFormLink() {
  return {
    link: getObservationFormCopyLink(),
    formId: CONFIG.OBSERVATION_FORM_ID
  };
}

// ============================================================================
// TEACHER SCHEDULE MANAGEMENT (ADMIN)
// ============================================================================

/**
 * Get all teachers with their schedule info for management (admin only)
 * @return {Object[]} Array of teachers with full schedule info
 */
function getTeachersForManagement() {
  if (!checkRole('admin')) {
    return { error: 'Unauthorized - Admin only' };
  }
  
  const teachers = Data.getAllTeachers();
  return teachers.map(t => {
    // Parse unavailablePeriods - the new unified field
    let unavailablePeriods = [];
    if (t.unavailablePeriods) {
      if (typeof t.unavailablePeriods === 'string') {
        try {
          unavailablePeriods = JSON.parse(t.unavailablePeriods);
        } catch (e) {
          unavailablePeriods = t.unavailablePeriods.split(',').map(p => Number(p.trim())).filter(p => p);
        }
      } else if (Array.isArray(t.unavailablePeriods)) {
        unavailablePeriods = t.unavailablePeriods;
      }
    }
    
    // Parse grades
    let grades = [];
    if (t.grades) {
      if (typeof t.grades === 'string') {
        try {
          grades = JSON.parse(t.grades);
        } catch (e) {
          grades = t.grades.split(',').map(g => Number(g.trim())).filter(g => g);
        }
      } else if (Array.isArray(t.grades)) {
        grades = t.grades;
      }
    } else if (t.grade) {
      grades = [Number(t.grade)];
    }
    
    return {
      id: t.id,
      name: t.name,
      email: t.email,
      grade: t.grade,
      grades: grades,
      room: t.room,
      unavailablePeriods: unavailablePeriods,
      lunchPeriod: t.lunchPeriod ? Number(t.lunchPeriod) : null,
      type: t.type || 'classroom'
    };
  });
}

// Keep old function name for backwards compatibility
function getTeachersWithPrepPeriods() {
  return getTeachersForManagement();
}

/**
 * Bulk update teacher schedules (grades, prep periods, lunch) - admin only
 * @param {Object[]} updates - Array of {teacherId, grades, prepPeriods, lunchPeriod}
 * @return {Object} Result
 */
function bulkUpdateTeacherSchedules(updates) {
  const user = getCurrentUser();
  if (!user || !checkRole('admin')) {
    return { error: 'Unauthorized - Admin only' };
  }
  
  if (!Array.isArray(updates)) {
    return { error: 'Invalid input' };
  }
  
  let successCount = 0;
  let errorCount = 0;
  
  updates.forEach(u => {
    const updateData = {};
    
    // Handle unavailable periods - store as JSON string
    if (u.unavailablePeriods !== undefined) {
      updateData.unavailablePeriods = JSON.stringify(u.unavailablePeriods);
      // Clear old lunchPeriod field since we're using unavailablePeriods now
      updateData.lunchPeriod = '';
    }
    
    // Backward compatibility: handle old lunchPeriod updates
    if (u.lunchPeriod !== undefined && u.unavailablePeriods === undefined) {
      updateData.lunchPeriod = u.lunchPeriod || '';
    }
    
    const result = Data.updateRow('Teachers', u.teacherId, updateData);
    if (result.success) {
      successCount++;
    } else {
      errorCount++;
      Logger.log('Failed to update teacher ' + u.teacherId + ': ' + result.error);
    }
  });
  
  Data.logAudit({
    action: 'BULK_UPDATE_TEACHER_UNAVAILABLE_TIMES',
    userId: user.email,
    details: JSON.stringify({ total: updates.length, success: successCount, errors: errorCount }),
    timestamp: new Date().toISOString()
  });
  
  return {
    success: true,
    message: 'Updated ' + successCount + ' teacher(s)' + (errorCount > 0 ? ', ' + errorCount + ' error(s)' : '')
  };
}

/**
 * Update a single teacher's unavailable periods (admin only)
 * @param {string} teacherId - Teacher ID
 * @param {Object} scheduleData - {unavailablePeriods}
 * @return {Object} Result
 */
function updateTeacherSchedule(teacherId, scheduleData) {
  return bulkUpdateTeacherSchedules([{
    teacherId: teacherId,
    ...scheduleData
  }]);
}
