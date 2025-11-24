/**
 * Refactored Sidebar Functions using Template System
 * All UI functions now use separated frontend templates
 */

/**
 * Show the main sidebar using the new template system
 * Supports both CardService and HTML modes via feature flag
 */
function showSidebar() {
  // Check feature flag to determine UI mode
  const useCardService = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);

  // Performance logging
  const startTime = new Date().getTime();

  if (useCardService) {
    // Use native CardService UI
    const card = showCardServiceSidebar();
    const endTime = new Date().getTime();
    logPerformance('showSidebar', endTime - startTime);
    return card;
  } else {
    // Use classic HTML UI
    const html = loadSidebarTemplate('control');
    SpreadsheetApp.getUi().showSidebar(html);
    const endTime = new Date().getTime();
    logPerformance('showSidebar', endTime - startTime);
  }
}

/**
 * Show the enterprise sidebar
 */
function showEnterpriseSidebar() {
  const html = loadSidebarTemplate('control');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Show initial setup dialog
 */
function showInitialSetup() {
  const html = loadInitialSetupDialog()
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Welcome to Client Manager');
}

/**
 * Show session recap dialog
 */
function showRecapDialog(clientData) {
  if (!clientData) {
    clientData = getCurrentClientData();
  }
  
  const html = loadRecapDialog(clientData)
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Session Recap');
}

/**
 * Show client selection dialog (Enterprise)
 */
function showClientSelectionDialogEnterprise() {
  const clients = getClientList();
  const html = loadClientSelectionDialog(clients)
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select Client');
}

/**
 * Show bulk client creation dialog
 */
function showBulkClientDialog() {
  const html = loadBulkClientDialog()
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Multiple Clients');
}

/**
 * Show Quick Notes settings dialog
 */
function showQuickNotesSettings() {
  const html = loadQuickNotesSettingsDialog()
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Notes Settings');
}

/**
 * Show update client dialog
 */
function showFullUpdateClientDialog() {
  const html = loadUpdateClientDialog()
    .setWidth(500)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Client Information');
}

/**
 * Show email preview dialog
 */
function showUpdatePreviewDialog(emailData) {
  const html = loadEmailPreviewDialog(emailData)
    .setWidth(600)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Email Preview');
}

/**
 * Show enterprise admin dashboard
 */
function showEnterpriseAdminDashboard() {
  if (!CONFIG.version.includes('enterprise')) {
    throw new Error('Enterprise features not available');
  }
  
  const html = loadAdminDashboard()
    .setWidth(700)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Admin Dashboard');
}

/**
 * Show organization dashboard
 */
function showOrganizationDashboard() {
  if (!CONFIG.version.includes('enterprise')) {
    throw new Error('Enterprise features not available');
  }
  
  const html = loadOrganizationDashboard()
    .setWidth(700)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Organization Dashboard');
}

/**
 * Save initial setup configuration
 */
function saveInitialSetup(config) {
  const userProperties = PropertiesService.getUserProperties();
  
  // Save user configuration
  userProperties.setProperties({
    tutorName: config.tutorName,
    tutorEmail: config.tutorEmail,
    companyName: config.companyName || 'Smart College',
    companyUrl: config.companyUrl || 'https://www.smartcollege.com',
    isConfigured: 'true'
  });
  
  // Save Acuity ID if provided (Enterprise)
  if (config.acuityUserId) {
    userProperties.setProperty('acuityUserId', config.acuityUserId);
  }
  
  // Update CONFIG object
  CONFIG.tutorName = config.tutorName;
  CONFIG.tutorEmail = config.tutorEmail;
  CONFIG.company = config.companyName || CONFIG.company;
  CONFIG.companyUrl = config.companyUrl || CONFIG.companyUrl;
  CONFIG.isConfigured = true;
  
  return { success: true, message: 'Setup completed successfully' };
}

/**
 * Refresh the menu after initial setup
 */
function refreshMenu() {
  const ui = SpreadsheetApp.getUi();
  createMainMenu(ui);
  return { success: true };
}

/**
 * Get user configuration
 */
function getUserConfig() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    tutorName: userProperties.getProperty('tutorName') || '',
    tutorEmail: userProperties.getProperty('tutorEmail') || '',
    companyName: userProperties.getProperty('companyName') || CONFIG.company,
    companyUrl: userProperties.getProperty('companyUrl') || CONFIG.companyUrl,
    isConfigured: userProperties.getProperty('isConfigured') === 'true',
    acuityUserId: userProperties.getProperty('acuityUserId') || ''
  };
}

/**
 * Test connectivity for connection monitoring
 */
function testConnectivity() {
  return {
    status: 'connected',
    timestamp: new Date().getTime(),
    success: true
  };
}