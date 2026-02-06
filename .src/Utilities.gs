/**
 * Utilities Module
 * Common helper functions
 */

const Utils = {
  
  /**
   * Format a date for display
   * @param {Date|string} date - Date to format
   * @param {string} format - Format string (default: 'EEEE, MMMM d, yyyy')
   * @return {string} Formatted date string
   */
  formatDate: function(date, format = 'EEEE, MMMM d, yyyy') {
    const d = typeof date === 'string' ? new Date(date + 'T00:00:00') : date;
    return Utilities.formatDate(d, CONFIG.TIMEZONE, format);
  },
  
  /**
   * Format time for display (12-hour format)
   * @param {string} time - Time string (HH:MM or H:MM)
   * @return {string} Formatted time string
   */
  formatTime: function(time) {
    const [hours, minutes] = time.split(':').map(Number);
    const ampm = hours >= 12 ? 'PM' : 'AM';
    const hour12 = hours % 12 || 12;
    return `${hour12}:${minutes.toString().padStart(2, '0')} ${ampm}`;
  },
  
  /**
   * Get today's date as YYYY-MM-DD string
   * @return {string} Today's date
   */
  getTodayString: function() {
    return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  },
  
  /**
   * Get start of current month
   * @return {Date} First day of month
   */
  getStartOfMonth: function() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  },
  
  /**
   * Get end of current month
   * @return {Date} Last day of month
   */
  getEndOfMonth: function() {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth() + 1, 0);
  },
  
  /**
   * Check if a date is a school day (not weekend)
   * @param {Date|string} date - Date to check
   * @return {boolean} True if school day
   */
  isSchoolDay: function(date) {
    const d = typeof date === 'string' ? new Date(date + 'T00:00:00') : date;
    const day = d.getDay();
    return day !== 0 && day !== 6;
  },
  
  /**
   * Get next school day
   * @param {Date} fromDate - Starting date (default: today)
   * @return {Date} Next school day
   */
  getNextSchoolDay: function(fromDate = new Date()) {
    const date = new Date(fromDate);
    date.setDate(date.getDate() + 1);
    
    while (!this.isSchoolDay(date)) {
      date.setDate(date.getDate() + 1);
    }
    
    return date;
  },
  
  /**
   * Calculate business days between two dates
   * @param {Date} startDate - Start date
   * @param {Date} endDate - End date
   * @return {number} Number of school days
   */
  getSchoolDaysBetween: function(startDate, endDate) {
    let count = 0;
    const current = new Date(startDate);
    
    while (current <= endDate) {
      if (this.isSchoolDay(current)) {
        count++;
      }
      current.setDate(current.getDate() + 1);
    }
    
    return count;
  },
  
  /**
   * Generate a unique ID
   * @return {string} UUID
   */
  generateId: function() {
    return Utilities.getUuid();
  },
  
  /**
   * Sanitize user input
   * @param {string} input - User input
   * @return {string} Sanitized string
   */
  sanitize: function(input) {
    if (typeof input !== 'string') return input;
    return input
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },
  
  /**
   * Validate email format
   * @param {string} email - Email to validate
   * @return {boolean} True if valid email format
   */
  isValidEmail: function(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  },
  
  /**
   * Parse date string to Date object
   * @param {string} dateStr - Date string (YYYY-MM-DD)
   * @return {Date} Date object
   */
  parseDate: function(dateStr) {
    return new Date(dateStr + 'T00:00:00');
  },
  
  /**
   * Format periods array for display
   * @param {number[]} periods - Array of period numbers
   * @return {string} Formatted string (e.g., "Periods 1, 2, 3")
   */
  formatPeriods: function(periods) {
    if (!periods || periods.length === 0) return 'No periods';
    if (periods.length === 1) return `Period ${periods[0]}`;
    return `Periods ${periods.join(', ')}`;
  },
  
  /**
   * Get color class based on status
   * @param {string} status - Status string
   * @return {string} CSS class name
   */
  getStatusClass: function(status) {
    const classes = {
      'confirmed': 'status-confirmed',
      'pending_sub': 'status-pending',
      'canceled': 'status-canceled',
      'pending': 'status-pending',
      'approved': 'status-approved',
      'denied': 'status-denied'
    };
    return classes[status] || 'status-default';
  },
  
  /**
   * Get human-readable status label
   * @param {string} status - Status string
   * @return {string} Human-readable label
   */
  getStatusLabel: function(status) {
    const labels = {
      'confirmed': 'Confirmed',
      'pending_sub': 'Awaiting Sub Approval',
      'canceled': 'Canceled',
      'pending': 'Pending',
      'approved': 'Approved',
      'denied': 'Denied',
      'not_needed': 'Not Needed'
    };
    return labels[status] || status;
  },
  
  /**
   * Truncate text to specified length
   * @param {string} text - Text to truncate
   * @param {number} maxLength - Maximum length
   * @return {string} Truncated text
   */
  truncate: function(text, maxLength = 50) {
    if (!text || text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + '...';
  },
  
  /**
   * Deep clone an object
   * @param {Object} obj - Object to clone
   * @return {Object} Cloned object
   */
  deepClone: function(obj) {
    return JSON.parse(JSON.stringify(obj));
  },
  
  /**
   * Check if two date strings are the same day
   * @param {string} date1 - First date (YYYY-MM-DD)
   * @param {string} date2 - Second date (YYYY-MM-DD)
   * @return {boolean} True if same day
   */
  isSameDay: function(date1, date2) {
    return date1 === date2;
  },
  
  /**
   * Get month name from date
   * @param {Date|string} date - Date
   * @return {string} Month name (e.g., "January")
   */
  getMonthName: function(date) {
    const d = typeof date === 'string' ? new Date(date) : date;
    return Utilities.formatDate(d, CONFIG.TIMEZONE, 'MMMM');
  },
  
  /**
   * Log error with context
   * @param {string} context - Where error occurred
   * @param {Error} error - Error object
   */
  logError: function(context, error) {
    Logger.log(`[ERROR] ${context}: ${error.message}`);
    Logger.log(error.stack);
  }
};
