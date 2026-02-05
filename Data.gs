/**
 * Data Access Layer
 * All CRUD operations for Google Sheets database
 * 
 * This file handles:
 * - Reading from and writing to sheets
 * - Data transformations
 * - Query operations
 */

const Data = {
  
  // ============================================================================
  // HELPER FUNCTIONS
  // ============================================================================
  
  /**
   * Get spreadsheet instance (cached)
   */
  getSpreadsheet: function() {
    if (!this._ss) {
      this._ss = CONFIG.SPREADSHEET_ID 
        ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        : SpreadsheetApp.getActiveSpreadsheet();
    }
    return this._ss;
  },
  
  /**
   * Get a sheet by name
   * @param {string} sheetName - Name of the sheet
   * @return {Sheet} The sheet object
   */
  getSheet: function(sheetName) {
    return this.getSpreadsheet().getSheetByName(sheetName);
  },
  
  /**
   * Get all data from a sheet as objects
   * @param {string} sheetName - Name of the sheet
   * @return {Object[]} Array of row objects
   */
  getAllRows: function(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const headers = data[0];
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        // Convert Date objects to ISO strings for browser compatibility
        if (row[i] instanceof Date) {
          obj[h] = row[i].toISOString();
        } else {
          obj[h] = row[i];
        }
      });
      return obj;
    });
  },
  
  /**
   * Find row index by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - ID to find
   * @param {string} idColumn - Column name for ID (default: 'id')
   * @return {number} Row index (1-based) or -1 if not found
   */
  findRowById: function(sheetName, id, idColumn = 'id') {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return -1;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.indexOf(idColumn);
    
    if (idColIndex === -1) return -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === id) {
        return i + 1; // 1-based row number
      }
    }
    return -1;
  },
  
  /**
   * Update a row by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - ID of row to update
   * @param {Object} updates - Object with column:value pairs
   * @return {Object} Result object
   */
  updateRow: function(sheetName, id, updates) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const rowIndex = this.findRowById(sheetName, id);
    if (rowIndex === -1) return { success: false, error: 'Row not found' };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (const [column, value] of Object.entries(updates)) {
      const colIndex = headers.indexOf(column);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(value);
      }
    }
    
    return { success: true };
  },
  
  /**
   * Append a row to a sheet
   * @param {string} sheetName - Name of the sheet
   * @param {Object} data - Object with column:value pairs
   * @return {Object} Result object
   */
  appendRow: function(sheetName, data) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => data[h] !== undefined ? data[h] : '');
    
    sheet.appendRow(row);
    return { success: true };
  },
  
  /**
   * Delete a row by ID
   * @param {string} sheetName - Name of the sheet
   * @param {string} id - ID of row to delete
   * @return {Object} Result object
   */
  deleteRow: function(sheetName, id) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };
    
    const rowIndex = this.findRowById(sheetName, id);
    if (rowIndex === -1) return { success: false, error: 'Row not found' };
    
    sheet.deleteRow(rowIndex);
    return { success: true };
  },
  
  // ============================================================================
  // TEACHERS
  // ============================================================================
  
  /**
   * Get all active teachers
   * @return {Object[]} Array of teacher objects
   */
  getAllTeachers: function() {
    return this.getAllRows('Teachers')
      .filter(t => {
        const active = String(t.active).toUpperCase();
        return active === 'TRUE' || active === 'ACTIVE' || active === 'YES' || active === '1';
      })
      .map(t => this.parseTeacherGrades(t));
  },
  
  /**
   * Parse teacher grades from comma-separated string to array
   * @param {Object} teacher - Teacher object
   * @return {Object} Teacher with grades as array
   */
  parseTeacherGrades: function(teacher) {
    if (!teacher.grade) {
      teacher.grades = [];
      teacher.grade = null;
    } else if (typeof teacher.grade === 'string' && teacher.grade.includes(',')) {
      teacher.grades = teacher.grade.split(',').map(g => Number(g.trim()));
      teacher.grade = teacher.grades[0]; // Primary grade for schedule purposes
    } else {
      teacher.grades = [Number(teacher.grade)];
      teacher.grade = Number(teacher.grade);
    }
    // Default type to 'classroom' if not specified
    teacher.type = teacher.type || 'classroom';
    return teacher;
  },
  
  /**
   * Get teachers by grade
   * @param {number} grade - Grade level
   * @return {Object[]} Array of teacher objects
   */
  getTeachersByGrade: function(grade) {
    return this.getAllTeachers().filter(t => 
      t.grades && t.grades.includes(Number(grade))
    );
  },
  
  /**
   * Get teacher by ID
   * @param {string} id - Teacher ID
   * @return {Object|null} Teacher object or null
   */
  getTeacherById: function(id) {
    const teacher = this.getAllRows('Teachers')
      .filter(t => {
        const active = String(t.active).toUpperCase();
        return active === 'TRUE' || active === 'ACTIVE' || active === 'YES' || active === '1';
      })
      .find(t => t.id === id);
    return teacher ? this.parseTeacherGrades(teacher) : null;
  },
  
  /**
   * Get teacher by email
   * @param {string} email - Teacher email
   * @return {Object|null} Teacher object or null
   */
  getTeacherByEmail: function(email) {
    const teacher = this.getAllRows('Teachers')
      .filter(t => {
        const active = String(t.active).toUpperCase();
        return active === 'TRUE' || active === 'ACTIVE' || active === 'YES' || active === '1';
      })
      .find(t => t.email.toLowerCase() === email.toLowerCase());
    return teacher ? this.parseTeacherGrades(teacher) : null;
  },
  
  // ============================================================================
  // ROOMS
  // ============================================================================
  
  /**
   * Get all rooms
   * @return {Object[]} Array of room objects
   */
  getAllRooms: function() {
    return this.getAllRows('Rooms');
  },
  
  /**
   * Get rooms by grade
   * @param {number} grade - Grade level
   * @return {Object[]} Array of room objects
   */
  getRoomsByGrade: function(grade) {
    return this.getAllRooms().filter(r => Number(r.grade) === grade);
  },
  
  // ============================================================================
  // BELL SCHEDULES
  // ============================================================================
  
  /**
   * Get bell schedule for a grade
   * @param {number} grade - Grade level (6 or 7 for 7/8)
   * @return {Object[]} Array of period objects sorted by period number
   */
  getBellSchedule: function(grade) {
    return this.getAllRows('BellSchedules')
      .filter(b => Number(b.grade) === grade)
      .map(b => {
        // Format times - handle both Date objects and strings
        b.startTime = this.formatTimeValue(b.startTime);
        b.endTime = this.formatTimeValue(b.endTime);
        return b;
      })
      .sort((a, b) => Number(a.period) - Number(b.period));
  },
  
  /**
   * Format a time value to HH:MM string
   * @param {Date|string} value - Time value from sheet
   * @return {string} Formatted time string
   */
  formatTimeValue: function(value) {
    if (!value) return '';
    
    // If it's already a simple time string, return it
    if (typeof value === 'string' && !value.includes('T')) {
      return value;
    }
    
    // If it's a Date object or ISO string, extract the time
    try {
      const date = new Date(value);
      let hours = date.getUTCHours();
      const minutes = date.getUTCMinutes();
      
      // Convert to 12-hour format
      const ampm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12;
      if (hours === 0) hours = 12;
      
      return hours + ':' + (minutes < 10 ? '0' : '') + minutes;
    } catch (e) {
      return String(value);
    }
  },
  
  // ============================================================================
  // PREP & LUNCH PERIODS
  // ============================================================================
  
  /**
   * Get prep periods for a grade
   * @param {number} grade - Grade level
   * @return {Object[]} Array of prep period objects
   */
  getPrepPeriods: function(grade) {
    return this.getAllRows('PrepPeriods').filter(p => Number(p.grade) === grade);
  },
  
  /**
   * Get lunch periods for a grade
   * @param {number} grade - Grade level
   * @return {Object[]} Array of lunch period objects
   */
  getLunchPeriods: function(grade) {
    return this.getAllRows('LunchPeriods').filter(l => Number(l.grade) === grade);
  },
  
  // ============================================================================
  // OBSERVATIONS
  // ============================================================================
  
  /**
   * Create a new observation
   * @param {Object} observation - Observation data
   * @return {Object} Result object
   */
  createObservation: function(observation) {
    // Convert periods array to string for storage
    const storageData = {
      ...observation,
      periods: JSON.stringify(observation.periods)
    };
    return this.appendRow('Observations', storageData);
  },
  
  /**
   * Get observation by ID
   * @param {string} id - Observation ID
   * @return {Object|null} Observation object or null
   */
  getObservationById: function(id) {
    const obs = this.getAllRows('Observations').find(o => o.id === id);
    if (obs && typeof obs.periods === 'string') {
      try {
        obs.periods = JSON.parse(obs.periods);
      } catch (e) {
        obs.periods = [];
      }
    }
    return obs || null;
  },
  
  /**
   * Update an observation
   * @param {string} id - Observation ID
   * @param {Object} updates - Fields to update
   * @return {Object} Result object
   */
  updateObservation: function(id, updates) {
    if (updates.periods && Array.isArray(updates.periods)) {
      updates.periods = JSON.stringify(updates.periods);
    }
    return this.updateRow('Observations', id, updates);
  },
  
  /**
   * Delete an observation
   * @param {string} id - Observation ID
   * @return {Object} Result object
   */
  deleteObservation: function(id) {
    return this.deleteRow('Observations', id);
  },
  
  /**
   * Get all observations with optional filters
   * @param {Object} filters - Optional filters (status, startDate, endDate)
   * @return {Object[]} Array of observation objects
   */
  getAllObservations: function(filters = {}) {
    let observations = this.getAllRows('Observations').map(o => {
      // Parse periods
      if (typeof o.periods === 'string') {
        try {
          o.periods = JSON.parse(o.periods);
        } catch (e) {
          o.periods = [];
        }
      }
      // Ensure date is a string (YYYY-MM-DD)
      if (o.date instanceof Date) {
        o.date = Utilities.formatDate(o.date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      } else if (typeof o.date === 'string' && o.date.includes('T')) {
        o.date = o.date.split('T')[0];
      }
      return o;
    });
    
    if (filters.status) {
      observations = observations.filter(o => o.status === filters.status);
    }
    
    if (filters.startDate) {
      const start = new Date(filters.startDate);
      observations = observations.filter(o => new Date(o.date) >= start);
    }
    
    if (filters.endDate) {
      const end = new Date(filters.endDate);
      observations = observations.filter(o => new Date(o.date) <= end);
    }
    
    return observations.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  },
  
  /**
   * Get observations by observer (teacher doing the observing)
   * @param {string} observerId - Observer's teacher ID
   * @param {Date} startDate - Optional start date
   * @param {Date} endDate - Optional end date
   * @return {Object[]} Array of observation objects
   */
  getObservationsByObserver: function(observerId, startDate = null, endDate = null) {
    let observations = this.getAllObservations()
      .filter(o => o.observerId === observerId && o.status !== 'canceled');
    
    if (startDate) {
      observations = observations.filter(o => new Date(o.date) >= startDate);
    }
    if (endDate) {
      observations = observations.filter(o => new Date(o.date) <= endDate);
    }
    
    return observations;
  },
  
  /**
   * Get observations by observer on a specific date
   * @param {string} observerId - Observer's teacher ID
   * @param {Date} date - The date
   * @return {Object[]} Array of observation objects
   */
  getObservationsByObserverOnDate: function(observerId, date) {
    const dateStr = Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    return this.getAllObservations()
      .filter(o => o.observerId === observerId && o.date === dateStr);
  },
  
  /**
   * Get observations for a teacher (being observed) on a specific date
   * @param {string} teacherId - Teacher ID being observed
   * @param {Date} date - The date
   * @return {Object[]} Array of observation objects
   */
  getObservationsForTeacherOnDate: function(teacherId, date) {
    const dateStr = Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    return this.getAllObservations()
      .filter(o => o.teacherId === teacherId && o.date === dateStr && o.status !== 'canceled');
  },
  
  // ============================================================================
  // SUBSTITUTE REQUESTS
  // ============================================================================
  
  /**
   * Create a substitute request
   * @param {Object} request - Request data
   * @return {Object} Result object
   */
  createSubRequest: function(request) {
    const storageData = {
      ...request,
      periods: JSON.stringify(request.periods)
    };
    return this.appendRow('SubstituteRequests', storageData);
  },
  
  /**
   * Get substitute request by ID
   * @param {string} id - Request ID
   * @return {Object|null} Request object or null
   */
  getSubRequestById: function(id) {
    const req = this.getAllRows('SubstituteRequests').find(r => r.id === id);
    if (req && typeof req.periods === 'string') {
      try {
        req.periods = JSON.parse(req.periods);
      } catch (e) {
        req.periods = [];
      }
    }
    return req || null;
  },
  
  /**
   * Get substitute request by observation ID
   * @param {string} observationId - Observation ID
   * @return {Object|null} Request object or null
   */
  getSubRequestByObservationId: function(observationId) {
    const req = this.getAllRows('SubstituteRequests').find(r => r.observationId === observationId);
    if (req && typeof req.periods === 'string') {
      try {
        req.periods = JSON.parse(req.periods);
      } catch (e) {
        req.periods = [];
      }
    }
    return req || null;
  },
  
  /**
   * Update a substitute request
   * @param {string} id - Request ID (or observation ID)
   * @param {Object} updates - Fields to update
   * @return {Object} Result object
   */
  updateSubRequest: function(id, updates) {
    // Try to find by request ID first, then by observation ID
    let req = this.getSubRequestById(id);
    if (!req) {
      req = this.getSubRequestByObservationId(id);
    }
    if (!req) return { success: false, error: 'Request not found' };
    
    return this.updateRow('SubstituteRequests', req.id, updates);
  },
  
  /**
   * Get substitute requests by status
   * @param {string} status - Status to filter by
   * @return {Object[]} Array of request objects
   */
  getSubRequestsByStatus: function(status) {
    return this.getAllRows('SubstituteRequests')
      .filter(r => r.status === status)
      .map(r => {
        // Ensure periods is an array
        if (typeof r.periods === 'string') {
          try {
            r.periods = JSON.parse(r.periods);
          } catch (e) {
            r.periods = [];
          }
        }
        // Ensure date is a string
        if (r.date instanceof Date) {
          r.date = Utilities.formatDate(r.date, CONFIG.TIMEZONE, 'yyyy-MM-dd');
        } else if (typeof r.date === 'string' && r.date.includes('T')) {
          r.date = r.date.split('T')[0];
        }
        return r;
      })
      .sort((a, b) => new Date(a.date) - new Date(b.date));
  },
  
  // ============================================================================
  // ADMINS & READ-ONLY USERS
  // ============================================================================
  
  /**
   * Get all admin emails
   * @return {string[]} Array of admin emails (lowercase)
   */
  getAdminEmails: function() {
    return this.getAllRows('Admins').map(a => a.email.toLowerCase());
  },
  
  /**
   * Get all read-only user emails
   * @return {string[]} Array of read-only emails (lowercase)
   */
  getReadOnlyEmails: function() {
    return this.getAllRows('ReadOnlyUsers').map(r => r.email.toLowerCase());
  },
  
  /**
   * Check if email is admin
   * @param {string} email - Email to check
   * @return {boolean} True if admin
   */
  isAdmin: function(email) {
    return this.getAdminEmails().includes(email.toLowerCase());
  },
  
  /**
   * Check if email is read-only user
   * @param {string} email - Email to check
   * @return {boolean} True if read-only
   */
  isReadOnly: function(email) {
    return this.getReadOnlyEmails().includes(email.toLowerCase());
  },
  
  // ============================================================================
  // AUDIT LOG
  // ============================================================================
  
  /**
   * Log an audit entry
   * @param {Object} entry - Audit entry data
   * @return {Object} Result object
   */
  logAudit: function(entry) {
    entry.id = Utilities.getUuid();
    return this.appendRow('AuditLog', entry);
  },
  
  /**
   * Get audit log entries
   * @param {number} limit - Maximum entries to return
   * @return {Object[]} Array of audit entries
   */
  getAuditLog: function(limit = 100) {
    return this.getAllRows('AuditLog')
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
      .slice(0, limit);
  },
  
  // ============================================================================
  // ACCESS REQUESTS
  // ============================================================================
  
  /**
   * Create AccessRequests sheet if it doesn't exist
   */
  ensureAccessRequestsSheet: function() {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName('AccessRequests');
    if (!sheet) {
      sheet = ss.insertSheet('AccessRequests');
      sheet.appendRow(['id', 'email', 'name', 'room', 'grade', 'requestedRole', 'status', 'createdAt', 'approvedBy', 'approvedAt', 'deniedBy', 'deniedAt', 'denyReason']);
    }
    return sheet;
  },
  
  /**
   * Create an access request
   * @param {Object} request - Request data
   * @return {Object} Result
   */
  createAccessRequest: function(request) {
    this.ensureAccessRequestsSheet();
    return this.appendRow('AccessRequests', request);
  },
  
  /**
   * Get access request by email
   * @param {string} email - Email to find
   * @return {Object|null} Request or null
   */
  getAccessRequest: function(email) {
    this.ensureAccessRequestsSheet();
    const requests = this.getAllRows('AccessRequests');
    return requests.find(r => r.email && r.email.toLowerCase() === email.toLowerCase());
  },
  
  /**
   * Get access request by ID
   * @param {string} id - Request ID
   * @return {Object|null} Request or null
   */
  getAccessRequestById: function(id) {
    this.ensureAccessRequestsSheet();
    const requests = this.getAllRows('AccessRequests');
    return requests.find(r => r.id === id);
  },
  
  /**
   * Get access requests by status
   * @param {string} status - Status filter
   * @return {Object[]} Matching requests
   */
  getAccessRequestsByStatus: function(status) {
    this.ensureAccessRequestsSheet();
    return this.getAllRows('AccessRequests').filter(r => r.status === status);
  },
  
  /**
   * Update an access request
   * @param {string} id - Request ID
   * @param {Object} updates - Fields to update
   * @return {Object} Result
   */
  updateAccessRequest: function(id, updates) {
    return this.updateRow('AccessRequests', id, updates);
  },
  
  /**
   * Find any user by email (teacher, admin, or readonly)
   * @param {string} email - Email to find
   * @return {Object|null} User or null
   */
  findUserByEmail: function(email) {
    const lowerEmail = email.toLowerCase();
    
    // Check teachers
    const teacher = this.getAllTeachers().find(t => t.email && t.email.toLowerCase() === lowerEmail);
    if (teacher) return { ...teacher, role: 'teacher' };
    
    // Check admins
    const admin = this.getAllAdmins().find(a => a.email && a.email.toLowerCase() === lowerEmail);
    if (admin) return { ...admin, role: 'admin' };
    
    // Check readonly
    const readonly = this.getAllRows('ReadOnlyUsers').find(r => r.email && r.email.toLowerCase() === lowerEmail);
    if (readonly) return { ...readonly, role: 'readonly' };
    
    return null;
  },
  
  /**
   * Find teacher by email
   * @param {string} email - Email to find
   * @return {Object|null} Teacher or null
   */
  findTeacherByEmail: function(email) {
    const lowerEmail = email.toLowerCase();
    return this.getAllTeachers().find(t => t.email && t.email.toLowerCase() === lowerEmail);
  },
  
  /**
   * Create a new teacher
   * @param {Object} teacher - Teacher data
   * @return {Object} Result
   */
  createTeacher: function(teacher) {
    return this.appendRow('Teachers', teacher);
  },
  
  /**
   * Add a read-only user
   * @param {Object} user - User data
   * @return {Object} Result
   */
  addReadOnlyUser: function(user) {
    return this.appendRow('ReadOnlyUsers', user);
  },
  
  /**
   * Get all admins
   * @return {Object[]} Admin objects
   */
  getAllAdmins: function() {
    return this.getAllRows('Admins');
  }
};
