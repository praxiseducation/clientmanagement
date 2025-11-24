/**
 * ADD THESE FUNCTIONS TO THE END OF clientmanager.gs
 * Template Loader System for Frontend/Backend Separation
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
 * Load and render the initial setup dialog
 */
function loadInitialSetupDialog() {
  const template = HtmlService.createTemplateFromFile('frontend/dialogs/initial-setup');
  
  template.isFirstTime = true;
  template.defaultCompany = CONFIG.company;
  template.defaultUrl = CONFIG.companyUrl;
  template.CONFIG = CONFIG;
  template.isEnterprise = CONFIG.version.includes('enterprise');
  
  return template.evaluate();
}

/**
 * Load and render the session recap dialog
 */
function loadRecapDialog(clientData) {
  const template = HtmlService.createTemplateFromFile('frontend/dialogs/recap-dialog');
  
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
  
  template.clientData = clientData;
  template.quickNotes = notesData;
  template.userConfig = getUserConfig();
  template.currentDate = new Date().toLocaleDateString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  template.CONFIG = CONFIG;
  template.isEnterprise = CONFIG.version.includes('enterprise');
  
  return template.evaluate();
}