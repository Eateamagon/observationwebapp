/**
 * Permissions Module
 * Handles role-based access control
 */

const Permissions = {
  
  /**
   * Get user role based on email
   * @param {string} email - User's email address
   * @return {string} Role: 'admin', 'teacher', 'readonly', or 'none'
   */
  getUserRole: function(email) {
    if (!email) return 'none';
    
    const emailLower = email.toLowerCase();
    
    // Check admin first (highest priority)
    if (Data.isAdmin(emailLower)) {
      return 'admin';
    }
    
    // Check if teacher
    const teacher = Data.getTeacherByEmail(emailLower);
    if (teacher) {
      return 'teacher';
    }
    
    // Check read-only
    if (Data.isReadOnly(emailLower)) {
      return 'readonly';
    }
    
    // Default: no access
    return 'none';
  },
  
  /**
   * Check if user has permission for an action
   * @param {string} email - User's email
   * @param {string} action - Action to check
   * @return {boolean} Whether user has permission
   */
  canPerform: function(email, action) {
    const role = this.getUserRole(email);
    
    const permissions = {
      // View actions
      'view_dashboard': ['admin', 'teacher', 'readonly'],
      'view_schedule': ['admin', 'teacher', 'readonly'],
      'view_teachers': ['admin', 'teacher', 'readonly'],
      'view_observations': ['admin', 'teacher', 'readonly'],
      
      // Teacher actions
      'create_observation': ['admin', 'teacher'],
      'cancel_own_observation': ['admin', 'teacher'],
      'reschedule_own_observation': ['admin', 'teacher'],
      'request_substitute': ['admin', 'teacher'],
      
      // Admin actions
      'edit_any_observation': ['admin'],
      'delete_any_observation': ['admin'],
      'approve_substitute': ['admin'],
      'deny_substitute': ['admin'],
      'view_audit_log': ['admin'],
      'export_data': ['admin'],
      'manage_users': ['admin']
    };
    
    const allowedRoles = permissions[action];
    if (!allowedRoles) return false;
    
    return allowedRoles.includes(role);
  },
  
  /**
   * Check if user can modify a specific observation
   * @param {string} email - User's email
   * @param {Object} observation - Observation object
   * @return {boolean} Whether user can modify
   */
  canModifyObservation: function(email, observation) {
    const role = this.getUserRole(email);
    
    // Admins can modify any observation
    if (role === 'admin') return true;
    
    // Teachers can only modify their own observations
    if (role === 'teacher') {
      const teacher = Data.getTeacherByEmail(email);
      return teacher && observation.observerId === teacher.id;
    }
    
    return false;
  },
  
  /**
   * Get list of allowed actions for a role
   * @param {string} role - User role
   * @return {string[]} Array of allowed actions
   */
  getAllowedActions: function(role) {
    const allActions = {
      'admin': [
        'view_dashboard', 'view_schedule', 'view_teachers', 'view_observations',
        'create_observation', 'cancel_own_observation', 'reschedule_own_observation',
        'request_substitute', 'edit_any_observation', 'delete_any_observation',
        'approve_substitute', 'deny_substitute', 'view_audit_log', 'export_data',
        'manage_users'
      ],
      'teacher': [
        'view_dashboard', 'view_schedule', 'view_teachers', 'view_observations',
        'create_observation', 'cancel_own_observation', 'reschedule_own_observation',
        'request_substitute'
      ],
      'readonly': [
        'view_dashboard', 'view_schedule', 'view_teachers', 'view_observations'
      ],
      'none': []
    };
    
    return allActions[role] || [];
  }
};
