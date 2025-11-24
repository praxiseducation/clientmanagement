// REPLACE THESE FUNCTIONS IN clientmanager.gs

// 1. Replace showSidebar() function (search for "function showSidebar()")
function showSidebar() {
  const html = loadSidebarTemplate('control');
  SpreadsheetApp.getUi().showSidebar(html);
}

// 2. Replace showEnterpriseSidebar() function (search for "function showEnterpriseSidebar()")
function showEnterpriseSidebar() {
  const html = loadSidebarTemplate('control');
  SpreadsheetApp.getUi().showSidebar(html);
}

// 3. Replace showUniversalSidebar() function (search for "function showUniversalSidebar(")
function showUniversalSidebar(initialView = 'control') {
  const html = loadSidebarTemplate(initialView);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 4. Replace showInitialSetup() function (search for "function showInitialSetup()")
function showInitialSetup() {
  const html = loadInitialSetupDialog()
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Welcome to Client Manager');
}

// 5. Replace showRecapDialog() function (search for "function showRecapDialog(")
function showRecapDialog(clientData) {
  if (!clientData) {
    clientData = getCurrentClientData();
  }
  
  const html = loadRecapDialog(clientData)
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Session Recap');
}

// 6. Add this new function for saving initial setup
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

// 7. Add this function for refreshing menu after setup
function refreshMenu() {
  const ui = SpreadsheetApp.getUi();
  createMainMenu(ui);
  return { success: true };
}