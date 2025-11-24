/**
 * Template Loader System for Frontend/Backend Separation
 * Handles loading and rendering of HTML templates
 */

/**
 * Include a file in a template (for CSS, JS, or HTML partials)
 * @param {string} filename - Path to the file to include
 * @return {string} The file content
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Load and render the main sidebar template
 * @param {string} initialView - The initial view to show ('control', 'quickNotes', etc.)
 * @return {HtmlOutput} The rendered sidebar
 */
function loadSidebarTemplate(initialView = 'control') {
  const template = HtmlService.createTemplateFromFile('frontend/components/sidebar-main');
  
  // Get required data
  const clientInfo = getCurrentClientInfo();
  const isEnterprise = CONFIG.version.includes('enterprise');
  
  // Set template variables
  template.clientInfo = clientInfo;
  template.isEnterprise = isEnterprise;
  template.CONFIG = CONFIG;
  template.initialView = initialView;
  
  // Get user info for enterprise mode
  if (isEnterprise) {
    try {
      const authUser = AuthManager.getCurrentUser();
      if (authUser.authenticated) {
        template.userInfo = authUser;
        template.userRole = PermissionManager.getUserRole(authUser.email) || 'tutor';
      }
    } catch (e) {
      console.log('Running in non-enterprise mode');
      template.userInfo = null;
      template.userRole = 'user';
    }
  } else {
    template.userInfo = null;
    template.userRole = 'user';
  }
  
  return template.evaluate()
    .setWidth(350)
    .setHeight(600)
    .setTitle('Client Manager');
}

/**
 * Load and render a dialog template
 * @param {string} dialogName - Name of the dialog template
 * @param {Object} data - Data to pass to the template
 * @return {HtmlOutput} The rendered dialog
 */
function loadDialogTemplate(dialogName, data = {}) {
  const template = HtmlService.createTemplateFromFile(`frontend/dialogs/${dialogName}`);
  
  // Pass data to template
  Object.keys(data).forEach(key => {
    template[key] = data[key];
  });
  
  // Add common data
  template.CONFIG = CONFIG;
  template.isEnterprise = CONFIG.version.includes('enterprise');
  
  return template.evaluate();
}

/**
 * Load Initial Setup Dialog
 */
function loadInitialSetupDialog() {
  return loadDialogTemplate('initial-setup', {
    isFirstTime: true,
    defaultCompany: CONFIG.company,
    defaultUrl: CONFIG.companyUrl
  });
}

/**
 * Load Session Recap Dialog
 */
function loadRecapDialog(clientData) {
  // Get Quick Notes data
  const quickNotes = getQuickNotes();
  let notesData = {};
  
  if (quickNotes && quickNotes.data) {
    try {
      notesData = JSON.parse(quickNotes.data);
    } catch (e) {
      console.error('Error parsing quick notes:', e);
    }
  }
  
  return loadDialogTemplate('recap-dialog', {
    clientData: clientData,
    quickNotes: notesData,
    userConfig: getUserConfig(),
    currentDate: new Date().toLocaleDateString('en-US', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    })
  });
}

/**
 * Load Client Selection Dialog
 */
function loadClientSelectionDialog(clients) {
  return loadDialogTemplate('client-selection', {
    clients: clients || getClientList(),
    hasClients: clients && clients.length > 0
  });
}

/**
 * Load Bulk Client Dialog
 */
function loadBulkClientDialog() {
  return loadDialogTemplate('bulk-client', {
    maxClients: 50,
    exampleFormat: 'John Smith, john@example.com, Math Tutoring'
  });
}

/**
 * Load Quick Notes Settings Dialog
 */
function loadQuickNotesSettingsDialog() {
  const currentSettings = JSON.parse(
    PropertiesService.getUserProperties().getProperty('quickNotesButtons') || '{}'
  );
  
  return loadDialogTemplate('quick-notes-settings', {
    settings: currentSettings,
    sections: [
      'wins',
      'skillsMastered',
      'skillsPracticed',
      'skillsIntroduced',
      'struggles',
      'parent',
      'next'
    ]
  });
}

/**
 * Load Update Client Dialog
 */
function loadUpdateClientDialog(clientData) {
  return loadDialogTemplate('update-client', {
    clientData: clientData || getCurrentClientData(),
    serviceTypes: [
      'Academic Support',
      'ACT Prep',
      'SAT Prep',
      'College Counseling',
      'Executive Function',
      'Math Tutoring',
      'Science Tutoring',
      'Writing Support'
    ]
  });
}

/**
 * Load Email Preview Dialog
 */
function loadEmailPreviewDialog(emailData) {
  return loadDialogTemplate('email-preview', {
    emailData: emailData,
    timestamp: new Date().toISOString()
  });
}

/**
 * Load Admin Dashboard
 */
function loadAdminDashboard() {
  const stats = {
    totalUsers: PermissionManager.getAllUsers().length,
    totalClients: getClientList().length,
    recapsSent: getRecapHistory().length,
    lastBackup: PropertiesService.getScriptProperties().getProperty('lastBackupDate') || 'Never'
  };
  
  return loadDialogTemplate('admin-dashboard', {
    stats: stats,
    users: PermissionManager.getAllUsers(),
    systemHealth: checkSystemHealth()
  });
}

/**
 * Load Organization Dashboard
 */
function loadOrganizationDashboard() {
  const orgData = OrganizationManager.getOrganizationData();
  
  return loadDialogTemplate('organization-dashboard', {
    organization: orgData,
    stats: OrganizationManager.getStats(),
    features: orgData.settings.features || []
  });
}

/**
 * Helper function to check system health
 */
function checkSystemHealth() {
  return {
    dataStore: UnifiedClientDataStore.validate(),
    permissions: PermissionManager.validatePermissions(),
    cache: CacheService.getScriptCache() !== null,
    properties: PropertiesService.getScriptProperties() !== null
  };
}