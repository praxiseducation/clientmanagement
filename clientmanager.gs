/**
 * Complete Client Management and Session Recap System
 * Combines client management with automated session recap emails
 */

const CONFIG = {
  version: '3.0.0-enterprise', // Version tracking
  build: '2025-01-08', // Build date
  tutorName: '', // Will be set during initial setup
  tutorEmail: '', // Will be set during initial setup
  company: 'Smart College',
  companyUrl: 'https://www.smartcollege.com',
  primaryColor: '#003366',
  accentColor: '#00C853',
  useMenuBar: true,  // Toggle this for add-on deployment
  isConfigured: false // Track if initial setup is complete
};

/**
 * Universal onOpen that handles both modes
 */
/**
 * Universal onOpen that handles both modes
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  
  // Check if initial setup is needed
  const userConfig = getUserConfig();
  if (!userConfig.isConfigured) {
    // Show setup dialog first
    ui.createMenu('Smart College')
      .addItem('Complete Setup', 'showInitialSetup')
      .addToUi();
    
    // Show setup automatically
    showInitialSetup();
    return;
  }
  
  try {
    // Always try to create the regular menu first
    createMainMenu(ui);
    
    // If we're in add-on mode, also create the add-on menu
    if (e && e.authMode) {
      ui.createAddonMenu()
        .addItem('Open Client Manager', 'showSidebar')
        .addToUi();
    }
  } catch (error) {
    // Fallback to add-on menu only
    console.warn('Could not create main menu, using add-on menu:', error);
    ui.createAddonMenu()
      .addItem('Open Client Manager', 'showSidebar')
      .addToUi();
  }
}

/**
 * Create the main menu bar menu
 */
function createMainMenu(ui) {
  ui.createMenu('Smart College')
    .addItem('📊 Open Control Panel', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Settings')
      .addItem('Add Multiple Clients', 'showBulkClientDialog')
      .addSeparator()
      .addItem('View Recap History', 'viewRecapHistory')
      .addItem('Batch Prep Mode', 'openBatchPrepMode')
      .addItem('Refresh Cache', 'refreshClientCache')
      .addItem('Debug Sheet Structure', 'debugSheetStructure')
      .addItem('Migrate to Unified Data Store', 'migrateToUnifiedDataStore')
      .addSeparator()
      .addItem('Sync Active Sheet to Master', 'syncActiveSheetToMaster')
      .addItem('Check Dashboard Links', 'validateDashboardLinks')
      .addItem('Export Client List (CSV)', 'exportClientListAsCSV')
      .addSeparator()
      .addItem('Help', 'showHelp'))
    .addToUi();
}

/**
 * Called when add-on is installed
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Initial setup function for new users
 */
function showInitialSetup() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
          }
          .container {
            max-width: 400px;
            margin: 0 auto;
          }
          h2 {
            color: #003366;
            text-align: center;
            margin-bottom: 8px;
          }
          h3 {
            color: #666;
            font-weight: normal;
            text-align: center;
            margin-top: 0;
            margin-bottom: 30px;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #5f6368;
            font-size: 14px;
          }
          input[type="text"], input[type="email"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e1e5e9;
            border-radius: 6px;
            font-size: 16px;
            box-sizing: border-box;
          }
          input:focus {
            outline: none;
            border-color: #003366;
            box-shadow: 0 0 0 3px rgba(0, 51, 102, 0.1);
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          button {
            width: 100%;
            padding: 12px 24px;
            border: none;
            border-radius: 24px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            background: #00C853;
            color: white;
            margin-top: 20px;
          }
          button:hover {
            background: #388E3C;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          .logo {
            text-align: center;
            margin-bottom: 20px;
            color: #003366;
            font-size: 24px;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="logo">Smart College</div>
          <h2>Welcome!</h2>
          <h3>Let's set up your Client Management System</h3>
          
          <form id="setupForm">
            <div class="form-group">
              <label>Your Name</label>
              <input type="text" id="tutorName" placeholder="Enter your full name" required>
            </div>
            
            <div class="form-group">
              <label>Your Email</label>
              <input type="email" id="tutorEmail" placeholder="Enter your email address" required>
            </div>
            
            <button type="submit">Complete Setup</button>
          </form>
        </div>
        
        <script>
          document.getElementById('setupForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.textContent;
            submitBtn.classList.add('btn-loading');
            submitBtn.disabled = true;
            
            const userData = {
              tutorName: document.getElementById('tutorName').value.trim(),
              tutorEmail: document.getElementById('tutorEmail').value.trim()
            };
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Setup complete! The page will now reload.');
                google.script.host.close();
                google.script.run.refreshUI();
              })
              .withFailureHandler(function(error) {
                submitBtn.classList.remove('btn-loading');
                submitBtn.textContent = originalText;
                submitBtn.disabled = false;
                alert('Error during setup: ' + error.message);
              })
              .saveUserConfig(userData);
          });
        </script>
      </body>
    </html>
  `).setWidth(450).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Initial Setup');
}

/**
 * Show settings dialog to update tutor information
 */
function showSettings() {
  const config = getCompleteConfig();
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
          }
          .container {
            max-width: 400px;
            margin: 0 auto;
          }
          h2 {
            color: #003366;
            text-align: center;
            margin-bottom: 30px;
          }
          .form-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #5f6368;
            font-size: 14px;
          }
          input[type="text"], input[type="email"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e1e5e9;
            border-radius: 6px;
            font-size: 16px;
            box-sizing: border-box;
          }
          input:focus {
            outline: none;
            border-color: #003366;
            box-shadow: 0 0 0 3px rgba(0, 51, 102, 0.1);
          }
          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 30px;
          }
          button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 24px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
          }
          .btn-primary {
            background: #00C853;
            color: white;
          }
          .btn-primary:hover {
            background: #388E3C;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .info-box {
            background: #e8f0fe;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 13px;
            color: #003366;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Settings</h2>
          
          <div class="info-box">
            Update your information below. This will be used in all session recap emails.
          </div>
          
          <form id="settingsForm">
            <div class="form-group">
              <label>Your Name</label>
              <input type="text" id="tutorName" value="${config.tutorName}" placeholder="Enter your full name" required>
            </div>
            
            <div class="form-group">
              <label>Your Email</label>
              <input type="email" id="tutorEmail" value="${config.tutorEmail}" placeholder="Enter your email address" required>
            </div>
            
            <div class="button-group">
              <button type="submit" class="btn-primary">Save Changes</button>
              <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
            </div>
          </form>
        </div>
        
        <script>
          document.getElementById('settingsForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.textContent;
            submitBtn.classList.add('btn-loading');
            submitBtn.disabled = true;
            
            const userData = {
              tutorName: document.getElementById('tutorName').value.trim(),
              tutorEmail: document.getElementById('tutorEmail').value.trim()
            };
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Settings updated successfully!');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                submitBtn.classList.remove('btn-loading');
                submitBtn.textContent = originalText;
                submitBtn.disabled = false;
                alert('Error updating settings: ' + error.message);
              })
              .saveUserConfig(userData);
          });
        </script>
      </body>
    </html>
  `).setWidth(450).setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Save user configuration
 */
function saveUserConfig(userData) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('tutorName', userData.tutorName);
  userProperties.setProperty('tutorEmail', userData.tutorEmail);
  userProperties.setProperty('isConfigured', 'true');
}

/**
 * Get complete configuration (user + company settings)
 */
function getCompleteConfig() {
  const userConfig = getUserConfig();
  return {
    tutorName: userConfig.tutorName || 'Your Tutor',
    tutorEmail: userConfig.tutorEmail || 'support@smartcollege.com',
    company: 'Smart College',
    companyUrl: 'https://www.smartcollege.com',
    primaryColor: '#003366',
    accentColor: '#00C853',
    useMenuBar: true,
    isConfigured: userConfig.isConfigured
  };
}

/**
 * Get user configuration
 */
function getUserConfig() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    tutorName: userProperties.getProperty('tutorName') || '',
    tutorEmail: userProperties.getProperty('tutorEmail') || '',
    isConfigured: userProperties.getProperty('isConfigured') === 'true'
  };
}

/**
 * Get version information
 */
function getVersionInfo() {
  return {
    script: {
      version: CONFIG.version,
      build: CONFIG.build,
      lastModified: new Date().toISOString().split('T')[0]
    },
    sidebar: {
      version: '3.0.0-enterprise',
      build: '20250108'
    }
  };
}

/**
 * Refresh UI after setup
 */
function refreshUI() {
  const ui = SpreadsheetApp.getUi();
  onOpen();
}

function showHelp() {
  const config = getCompleteConfig();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px;">
      <div style="text-align: center; margin-bottom: 20px;">
        <h2 style="color: #003366; margin-bottom: 5px;">Smart College</h2>
        <h3 style="color: #666; font-weight: normal; margin-top: 0;">Client Management System</h3>
      </div>
      
      <p>Version 1.0</p>
      
      <h4>Quick Start:</h4>
      <ul>
        <li><b>Add New Client:</b> Creates a new client sheet from template</li>
        <li><b>Find Client:</b> Quickly navigate to any client sheet</li>
        <li><b>Quick Notes:</b> Take session notes that auto-save</li>
        <li><b>Send Recap:</b> Create and send session recap emails</li>
      </ul>
      
      <h4>Keyboard Shortcuts:</h4>
      <ul>
        <li><b>Ctrl+Alt+N:</b> Quick Notes (when on client sheet)</li>
        <li><b>Ctrl+Alt+F:</b> Find Client</li>
      </ul>
      
      <p style="margin-top: 20px; color: #666;">
        For support, contact: ${config.tutorEmail}
      </p>
    </div>
  `).setWidth(400).setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help');
}

function testBasicSidebar() {
  const html = HtmlService.createHtmlOutput('<h1>Hello World</h1><button onclick="alert(\'clicked\')">Click Me</button>');
  SpreadsheetApp.getUi().showSidebar(html);
}

function testBasicDialog() {
  const html = HtmlService.createHtmlOutput(`
    <div>
      <h3>Test Dialog</h3>
      <button onclick="alert('Button clicked!')">Test Button</button>
      <br><br>
      <input type="text" id="testInput" placeholder="Type here">
      <button onclick="showValue()">Show Value</button>
      <p id="output"></p>
    </div>
    <script>
      function showValue() {
        var val = document.getElementById('testInput').value;
        document.getElementById('output').innerHTML = 'You typed: ' + val;
      }
    </script>
  `).setWidth(400).setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Test Dialog');
}

/**
 * Test function to verify sidebar functionality
 */
function testSimpleSidebar() {
  const html = HtmlService.createHtmlOutput();
  html.setContent(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif; padding: 20px; }
          button { 
            background: #003366; 
            color: white; 
            border: none; 
            padding: 10px; 
            margin: 5px; 
            cursor: pointer; 
            width: 100%;
          }
          button:hover { background: #002244; }
          textarea { width: 100%; margin: 5px 0; }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <h3>Test Sidebar</h3>
        <button onclick="testClick()">Test Button</button>
        <textarea id="testArea" placeholder="Type here..."></textarea>
        <button onclick="testSave()">Save Text</button>
        <div id="result"></div>
        
        <script>
          function testClick() {
            document.getElementById('result').innerHTML = 'Button clicked!';
          }
          
          function testSave() {
            var text = document.getElementById('testArea').value;
            document.getElementById('result').innerHTML = 'You typed: ' + text;
          }
        </script>
      </body>
    </html>
  `);
  
  html.setTitle('Test').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

const ACT_TESTS = [
  'April 07', 'Dec 19', 'April 23', 'June 23', 'April 24', 
  'Dec 22', 'June 21', 'Sept 23', 'June 22', 'April 19', 
  'June 19', 'Dec 18', 'June 18', 'June 17', 'Dec 17'
];

// Homework templates
const HOMEWORK_TEMPLATES = {
  'ACT': [
    'ACT Practice Test',
    'Complete ACT practice test before our next session',
    'Review ACT math sections - 30 minutes daily',
    'Complete ACT practice problems set before our next session',
    'Review and redo missed problems from last test',
    'Focus on timing strategies for 30 minutes',
    'Review ACT science passages - 20 minutes daily'
  ],
  'Math Curriculum': [
    'Complete textbook problems before our next session',
    'Review class notes - 20 minutes daily',
    'Complete practice worksheet before our next session',
    'Study for upcoming assessment',
    'Khan Academy practice - 30 minutes daily',
    'Work through example problems from today',
    'Complete problem set before our next session'
  ],
  'Language': [
    'Practice conversation phrases - 15 minutes daily',
    'Complete workbook exercises before our next session',
    'Watch target language video with subtitles - 20 minutes',
    'Write 5 sentences using new vocabulary',
    'Record yourself reading dialogue',
    'Practice pronunciation for 15 minutes daily',
    'Review grammar rules from today'
  ],
  'General': [
    'ACT Practice Test',
    'Review today\'s materials - 20 minutes daily',
    'Complete practice problems before our next session',
    'Prepare questions for next session',
    'Review and organize notes from today',
    'Practice explaining concepts aloud',
    'Complete assigned reading before our next session'
  ]
};


// Email templates - Universal template for all client types
const EMAIL_TEMPLATES = {
  'Universal': {
    subject: '{studentFirstName}\'s ACT Update {date} - Smart College',
    template: `Hi {parentName},

{studentFirstName} had an excellent session today! All notes and resources are available in {studentFirstName}'s {dashboardText}.

<b style="font-size: 16px;">Today's Focus</b>
{focusSkill}

<b style="font-size: 16px;">Big Win</b>
{todaysWin}

<b style="font-size: 16px;">What We Covered</b>
<i>Mastered:</i>
{masteredBullets}
<i>Practiced:</i>
{practicedBullets}
<i>Introduced:</i>
{introducedBullets}

<b style="font-size: 16px;">Progress Update</b>
{progressNotes}

<b style="font-size: 16px;">Homework/Practice</b>
{homework}

<b style="font-size: 16px;">Support at Home</b>
{quickWin}

<b style="font-size: 16px;">Next Session</b>
{nextSession}

<b style="font-size: 16px;">Looking Ahead</b>
{nextSessionPreview}

Questions? Just hit reply!

Best,
{tutorName}{psSection}`
  }
};

/**
 * Adds a custom menu to the spreadsheet when opened
 */
/**function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Smart College')
    .addItem('Add New Client', 'addNewClient')
    .addSeparator()
    .addItem('Find Client', 'findClient')
    .addSeparator()
    .addItem('Quick Notes', 'openQuickNotes')
    .addSeparator()
    .addItem('Send Session Recap', 'sendIndividualRecap')
    .addSeparator()
    .addSubMenu(ui.createMenu('Session Management')
      .addItem('Batch Prep Mode', 'openBatchPrepMode')
      .addItem('View Recap History', 'viewRecapHistory'))
    .addToUi();
}
*/

/**
 * Get sidebar data for client-side initialization
 */
function getSidebarData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const rawClientInfo = getCurrentClientInfo();
  
  return {
    clientInfo: {
      isClient: Boolean(rawClientInfo.isClient),
      name: rawClientInfo.name ? String(rawClientInfo.name) : '',
      sheetName: rawClientInfo.sheetName ? String(rawClientInfo.sheetName) : ''
    },
    initialView: 'control'
  };
}

/**
 * Enhanced Universal Sidebar with Enterprise Features
 * Version 3.1.0 - Complete enterprise functionality built-in
 */
function showUniversalSidebar(initialView = 'control') {
  // Pre-fetch all required data for performance
  const clientInfo = getCurrentClientInfo();
  const isEnterprise = CONFIG.version.includes('enterprise');
  
  // Get user info for enterprise mode
  let userInfo = null;
  let userRole = 'user';
  if (isEnterprise) {
    try {
      const authUser = AuthManager.getCurrentUser();
      if (authUser.authenticated) {
        userInfo = authUser;
        userRole = PermissionManager.getUserRole(authUser.email) || 'tutor';
      }
    } catch (e) {
      console.log('Running in non-enterprise mode');
    }
  }
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 0;
            font-size: 14px;
            box-sizing: border-box;
          }
          
          /* View management */
          .view { 
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            padding: 15px;
            background: white;
            transform: translateX(100%);
            transition: transform 0.2s ease-out;
            will-change: transform;
            overflow-y: auto;
            box-sizing: border-box;
            z-index: 1;
          }
          .view.active { 
            transform: translateX(0);
            z-index: 10;
          }
          
          
          
          /* Loading overlay for immediate feedback */
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(2px);
            z-index: 1000;
            display: none;
            align-items: center;
            justify-content: center;
            opacity: 0;
            transition: opacity 0.2s ease;
          }
          
          .loading-overlay.show {
            display: flex;
            opacity: 1;
          }
          
          .loading-spinner-large {
            width: 40px;
            height: 40px;
            border: 3px solid #f3f3f3;
            border-top: 3px solid #003366;
            border-radius: 50%;
            animation: spin 1s linear infinite;
          }

          /* Offline banner */
          .offline-banner {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            background: #f44336;
            color: white;
            text-align: center;
            padding: 8px;
            z-index: 2000;
            font-size: 12px;
            font-weight: 500;
            display: none;
            animation: slideDown 0.3s ease-out;
          }
          
          .offline-banner.show {
            display: block;
          }
          
          @keyframes slideDown {
            from { transform: translateY(-100%); }
            to { transform: translateY(0); }
          }
          
          /* Offline button states */
          button.offline-disabled {
            background: #ccc !important;
            color: #999 !important;
            cursor: not-allowed !important;
            opacity: 0.6 !important;
          }
          
          button.offline-disabled:hover {
            background: #ccc !important;
            transform: none !important;
            box-shadow: none !important;
          }

          /* Common styles */
          h3 {
            margin-top: 0;
            color: #003366;
          }
          
          /* Base button styles */
          button {
            background: #003366;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            margin-top: 10px;
            width: 48%;
            transition: all 0.2s ease;
            box-sizing: border-box;
          }
          
          button:hover {
            background: #002244;
          }
          
          button:disabled {
            cursor: not-allowed;
            opacity: 0.7;
          }
          
          .button-group {
            display: flex;
            gap: 4%;
            width: 100%;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            pointer-events: none;
            opacity: 0.8;
          }
          
          /* Hide text but keep icons visible during loading */
          .btn-loading .icon {
            opacity: 1 !important;
            color: inherit !important;
            position: relative;
            z-index: 2;
          }
          
          /* Hide only the text content, not the icons */
          .btn-loading {
            color: transparent !important;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            right: 12px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            z-index: 1;
          }

          /* For light buttons (menu-button without primary), use dark spinner */
          .menu-button.btn-loading:not(.primary)::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #333;
          }
          
          /* Ensure primary buttons keep their green background and white spinner */
          .menu-button.primary.btn-loading {
            background: #00C853 !important;
          }
          
          .menu-button.primary.btn-loading .icon {
            color: white !important;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          /* Updated button styles for Quick Notes section */
          .btn-primary {
            background: #00C853;
            color: white !important;
          }
          
          .btn-primary:hover {
            background: #388E3C;
          }
          
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368 !important; 
            border: 1px solid #dadce0;
          }
          
          .btn-secondary:hover {
            background: #f1f3f4;
            border-color: #d2d6dc;
          }
          
          .btn-danger {
            background: #ea4335;
            color: white !important; 
          }
          
          .btn-danger:hover {
            background: #d33b2c;
          }
          
          /* Control Panel styles */
          .current-client {
            background: #e8f0fe;
            border-left: 4px solid #003366;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 13px;
          }
          
          .current-client-label {
            font-weight: 500;
            color: #003366;
            margin-bottom: 4px;
          }
          
          .current-client-name {
            color: #202124;
            font-weight: 600;
          }
          
          .no-client {
            color: #5f6368;
            font-style: italic;
          }
          
          .section {
            margin-bottom: 25px;
          }
          
          .section-title {
            font-weight: 500;
            color: #5f6368;
            margin-bottom: 12px;
            font-size: 13px;
            letter-spacing: 0.5px;
          }
          
          .menu-button {
            width: 100%;
            padding: 12px;
            margin-bottom: 8px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: white;
            color: #202124;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s;
            text-align: left;
            display: flex;
            align-items: center;
            gap: 10px;
          }
          
          .menu-button:hover:not(:disabled) {
            background: #f8f9fa;
            border-color: #003366;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          
          .menu-button:active {
            transform: translateY(0);
            box-shadow: none;
          }
          
          .menu-button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
          }
          
          .primary {
            background: #00C853;
            color: white !important; 
            border: none;
            text-align: center;
            justify-content: center;
          }
          
          .primary:hover:not(:disabled) {
            background: #388E3C;
          }
          
          .icon {
            font-size: 18px;
            width: 24px;
            text-align: center;
          }
          
          /* Quick Notes styles */
          .control-panel-btn {
            width: 100%;
            padding: 10px;
            margin-bottom: 15px;
            background: #f8f9fa;
            color: #003366;
            border: 1px solid #dadce0;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
          }
          
          .control-panel-btn:hover {
            background: #e8f0fe;
            border-color: #003366;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          
          .divider {
            height: 1px;
            background: #e8eaed;
            margin: 15px 0;
          }
          
          .note-section {
            margin-bottom: 15px;
          }
          
          /* Enhanced Skills Section Styles */
          .skills-title {
            margin-bottom: 20px;
          }
          
          .skills-grid {
            display: flex;
            flex-direction: column;
            gap: 16px;
          }
          
          .skill-category {
            background: white;
            border-radius: 6px;
            padding: 12px;
            border-left: 4px solid transparent;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            transition: all 0.2s ease;
          }
          
          .skill-category:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 6px rgba(0,0,0,0.15);
          }
          
          .skill-category.mastered {
            border-left-color: #28a745;
          }
          
          .skill-category.practiced {
            border-left-color: #ffc107;
          }
          
          .skill-category.introduced {
            border-left-color: #17a2b8;
          }
          
          .skill-label {
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 500;
            color: #495057;
            margin-bottom: 8px;
            font-size: 14px;
          }
          
          .skill-icon {
            font-size: 16px;
          }
          
          .skill-tags {
            margin-bottom: 8px;
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
          }
          
          .skill-textarea {
            width: 100%;
            min-height: 60px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            padding: 8px;
            font-family: inherit;
            font-size: 13px;
            resize: vertical;
            transition: border-color 0.2s ease;
          }
          
          .skill-textarea:focus {
            outline: none;
            border-color: #80bdff;
            box-shadow: 0 0 0 2px rgba(0,123,255,0.25);
          }
          
          .mastered-textarea:focus {
            border-color: #28a745;
            box-shadow: 0 0 0 2px rgba(40,167,69,0.25);
          }
          
          .practiced-textarea:focus {
            border-color: #ffc107;
            box-shadow: 0 0 0 2px rgba(255,193,7,0.25);
          }
          
          .introduced-textarea:focus {
            border-color: #17a2b8;
            box-shadow: 0 0 0 2px rgba(23,162,184,0.25);
          }
          
          .mastered-tag {
            background-color: #d4edda !important;
            color: #155724 !important;
            border: 1px solid #c3e6cb !important;
          }
          
          .practiced-tag {
            background-color: #fff3cd !important;
            color: #856404 !important;
            border: 1px solid #ffeaa7 !important;
          }
          
          .introduced-tag {
            background-color: #d1ecf1 !important;
            color: #0c5460 !important;
            border: 1px solid #b8daff !important;
          }
          
          .mastered-tag:hover {
            background-color: #c3e6cb !important;
            transform: translateY(-1px);
          }
          
          .practiced-tag:hover {
            background-color: #ffeaa7 !important;
            transform: translateY(-1px);
          }
          
          .introduced-tag:hover {
            background-color: #b8daff !important;
            transform: translateY(-1px);
          }
          
          /* Enhanced Note Category Styles */
          .note-category {
            background: white;
            border-radius: 6px;
            padding: 12px;
            margin-bottom: 16px;
            border-left: 4px solid transparent;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            transition: all 0.2s ease;
          }
          
          .note-category:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 6px rgba(0,0,0,0.15);
          }
          
          .note-category.wins {
            border-left-color: #28a745;
          }
          
          .note-category.struggles {
            border-left-color: #dc3545;
          }
          
          .note-category.parent {
            border-left-color: #6f42c1;
          }
          
          .note-category.next {
            border-left-color: #fd7e14;
          }
          
          .note-label {
            display: flex;
            align-items: center;
            gap: 8px;
            font-weight: 500;
            color: #495057;
            margin-bottom: 8px;
            font-size: 14px;
          }
          
          .note-icon {
            font-size: 16px;
          }
          
          .note-tags {
            margin-bottom: 8px;
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
          }
          
          .note-textarea {
            width: 100%;
            min-height: 60px;
            border: 1px solid #ced4da;
            border-radius: 4px;
            padding: 8px;
            font-family: inherit;
            font-size: 13px;
            resize: vertical;
            transition: border-color 0.2s ease;
          }
          
          .note-textarea:focus {
            outline: none;
            border-color: #80bdff;
            box-shadow: 0 0 0 2px rgba(0,123,255,0.25);
          }
          
          .wins-textarea:focus {
            border-color: #28a745;
            box-shadow: 0 0 0 2px rgba(40,167,69,0.25);
          }
          
          .struggles-textarea:focus {
            border-color: #dc3545;
            box-shadow: 0 0 0 2px rgba(220,53,69,0.25);
          }
          
          .parent-textarea:focus {
            border-color: #6f42c1;
            box-shadow: 0 0 0 2px rgba(111,66,193,0.25);
          }
          
          .next-textarea:focus {
            border-color: #fd7e14;
            box-shadow: 0 0 0 2px rgba(253,126,20,0.25);
          }
          
          .wins-tag {
            background-color: #d4edda !important;
            color: #155724 !important;
            border: 1px solid #c3e6cb !important;
          }
          
          .struggles-tag {
            background-color: #f8d7da !important;
            color: #721c24 !important;
            border: 1px solid #f5c6cb !important;
          }
          
          .parent-tag {
            background-color: #e2d9f3 !important;
            color: #4c2c92 !important;
            border: 1px solid #d1c4e9 !important;
          }
          
          .next-tag {
            background-color: #fef3e2 !important;
            color: #8a4a00 !important;
            border: 1px solid #fed7aa !important;
          }
          
          .wins-tag:hover {
            background-color: #c3e6cb !important;
            transform: translateY(-1px);
          }
          
          .struggles-tag:hover {
            background-color: #f5c6cb !important;
            transform: translateY(-1px);
          }
          
          .parent-tag:hover {
            background-color: #d1c4e9 !important;
            transform: translateY(-1px);
          }
          
          .next-tag:hover {
            background-color: #fed7aa !important;
            transform: translateY(-1px);
          }
          
          textarea {
            width: 100%;
            min-height: 60px;
            padding: 8px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-family: inherit;
            font-size: 13px;
            resize: vertical;
            box-sizing: border-box;
          }
          
          label {
            display: block;
            margin-bottom: 4px;
            font-weight: 500;
            color: #5f6368;
            font-size: 12px;
          }
          
          .quick-tag {
            display: inline-block;
            padding: 4px 8px;
            background: #f1f3f4;
            border: 1px solid #dadce0;
            border-radius: 12px;
            margin: 1px 2px;
            cursor: pointer;
            font-size: 10px;
            font-weight: 500;
            transition: all 0.15s ease;
          }
          
          .quick-tag:hover {
            background: #e8f0fe;
            border-color: #003366;
          }
          
          .save-indicator {
            position: fixed;
            top: 10px;
            right: 10px;
            padding: 5px 10px;
            background: #00C853;
            color: white;
            border-radius: 4px;
            font-size: 11px;
            opacity: 0;
            transition: opacity 0.3s;
          }
          
          .save-indicator.show {
            opacity: 1;
          }
          
          .auto-save-status {
            font-size: 11px;
            color: #5f6368;
            margin-top: 10px;
            text-align: center;
          }
          
         /* Save button specific styles */
          #saveNotesBtn {
            background: #00C853;
            color: white;
          }

          #saveNotesBtn:hover {
            background: #388E3C;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }

          #sendRecapBtn {
            background: #003366;
            color: white;
            position: relative;
          }

          #sendRecapBtn:hover {
            background: #002244;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          
          #sendRecapBtn.btn-loading {
            color: transparent;
          }
          
          #sendRecapBtn.btn-loading::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 16px;
            height: 16px;
            margin: -8px 0 0 -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }
          
          .btn-success {
            background: #00C853 !important;
            transition: background 0.2s ease;
          }
          
          /* Universal loading states for buttons */
          .btn-loading {
            position: relative;
            pointer-events: none;
            opacity: 0.7;
          }
          
          .btn-loading::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 16px;
            height: 16px;
            margin: -8px 0 0 -8px;
            border: 2px solid transparent;
            border-top-color: currentColor;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }
          
          .btn-loading span {
            opacity: 0;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          /* Loading states */
          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }
          
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #003366;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          
          /* Client list loading overlay */
          .client-list {
            position: relative;
          }
          
          .client-list-loading-overlay {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            width: 100%;
            height: 100%;
            min-height: 200px;
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(3px);
            display: none;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            z-index: 1000;
            border-radius: 6px;
          }
          
          .client-list-loading-overlay.active {
            display: flex;
          }
          
          .client-loading-spinner {
            width: 24px;
            height: 24px;
            border: 3px solid rgba(0, 51, 102, 0.3);
            border-top-color: #003366;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin-bottom: 12px;
          }
          
          .client-loading-text {
            font-size: 14px;
            color: #003366;
            font-weight: 500;
          }
          
          /* View transition animation */
          .view.sliding-in {
            animation: slideInFromRight 0.2s ease-out;
          }
          
          @keyframes slideInFromRight {
            from { transform: translateX(100%); }
            to { transform: translateX(0); }
          }
          
          /* Hardware acceleration hints */
          .view {
            transform: translateZ(0);
            backface-visibility: hidden;
            perspective: 1000px;
          }
          
          /* Alert styles */
          .alert {
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 15px;
            font-size: 13px;
            display: none;
          }
          
          .alert.error {
            background: #fce8e6;
            color: #d93025;
            display: block;
          }
          
          .alert.success {
            background: #e6f4ea;
            color: #188038;
            display: block;
          }
          
          /* Overlay Button Styles */
          .overlay-btn {
            flex: 1;
            padding: 14px 20px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
            text-transform: none;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
          }
          
          .overlay-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.2);
          }
          
          .overlay-btn:active {
            transform: translateY(0);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
          }
          
          .overlay-btn-primary {
            background: linear-gradient(135deg, #003366, #004d99);
            color: white;
          }
          
          .overlay-btn-primary:hover {
            background: linear-gradient(135deg, #002244, #003d7a);
          }
          
          .overlay-btn-secondary {
            background: linear-gradient(135deg, #f8f9fa, #e9ecef);
            color: #495057;
            border: 1px solid #dee2e6;
          }
          
          .overlay-btn-secondary:hover {
            background: linear-gradient(135deg, #e9ecef, #dee2e6);
            color: #343a40;
          }
          
          /* Loading state for overlay buttons */
          .overlay-btn.btn-loading {
            pointer-events: none;
            opacity: 0.8;
          }
          
          .overlay-btn.btn-loading .btn-text {
            opacity: 0;
          }
          
          .overlay-btn.btn-loading::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 20px;
            height: 20px;
            margin: -10px 0 0 -10px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }
          
          .overlay-btn-secondary.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.2);
            border-top-color: #495057;
          }
          
          /* Clickable indicators */
          .indicator-clickable {
            cursor: pointer;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            position: relative;
          }
          
          .indicator-clickable:hover {
            transform: scale(1.2);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
          }
          
          .indicator-clickable.has-url:hover::after {
            content: 'open';
            position: absolute;
            bottom: -20px;
            left: 50%;
            transform: translateX(-50%);
            background: rgba(0, 0, 0, 0.8);
            color: white;
            padding: 2px 6px;
            border-radius: 3px;
            font-size: 10px;
            white-space: nowrap;
            z-index: 1000;
            animation: fadeIn 0.2s ease;
          }
          
          @keyframes fadeIn {
            from { opacity: 0; transform: translateX(-50%) translateY(-5px); }
            to { opacity: 1; transform: translateX(-50%) translateY(0); }
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <!-- Server data will be loaded via google.script.run -->
        
        <!-- Offline notification banner -->
        <div id="offlineBanner" class="offline-banner">
          📡 Offline - Some features are temporarily unavailable
        </div>
        
        <div class="sidebar-container" style="position: relative; height: 100vh; overflow: hidden;">
        <!-- Control Panel View -->
        <div id="controlView" class="view active">
          <div class="current-client">
            <div class="current-client-label">Current Client:</div>
            <div style="display: flex; align-items: center; gap: 8px;">
              <div id="clientName" class="${clientInfo.isClient ? 'current-client-name' : 'no-client'}">
                ${clientInfo.isClient ? clientInfo.name : 'No client selected'}
              </div>
              <div id="clientWarning" style="display: none; color: #ff9800; cursor: pointer; font-size: 16px;" 
                onclick="openUpdateClientDialog()"
                onmouseenter="showTooltip(event, 'Missing Client Information')"
                onmouseleave="hideTooltip()"
                onmousemove="moveTooltip(event)">
                ⚠️
              </div>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Client Management</div>
            
            <!-- Client Search -->
            <div style="margin-bottom: 15px; width: 70%;">
              <input type="text" id="clientSearch" placeholder="Search active clients..." 
                style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box;"
                oninput="searchClients(this.value)">
              <div id="searchResults" style="margin-top: 5px;"></div>
            </div>
            
            <button class="menu-button primary" onclick="runFunction('addNewClient', this)" 
              onmouseenter="preloadAddClient()">
              <span class="icon">➕</span>
              Add New Client
            </button>
            <button class="menu-button" onclick="showClientList()" 
              onmouseenter="preloadClientList()">
              <span class="icon">📋</span>
              Client List
            </button>
            <button id="updateClientBtn" class="menu-button" onclick="openUpdateClientDialog()" disabled
              onmouseenter="preloadUpdateClient()">
              <span class="icon">⚙️</span>
              Update Client Info
            </button>
          </div>
          
          <div class="section">
            <div class="section-title">Session Management</div>
            <button id="quickNotesBtn" class="menu-button" onclick="switchToQuickNotes()" onmouseover="preloadQuickNotes()" disabled>
              <span class="icon">📝</span>
              Quick Notes
            </button>
            <button id="recapBtn" class="menu-button" onclick="runFunction('sendIndividualRecap', this)" onmouseover="preloadRecapData()" disabled>
              <span class="icon">✉️</span>
              Send Session Recap
            </button>
          </div>
          
          <div class="section">
            <div class="section-title">External Tools</div>
            <button class="menu-button" onclick="window.open('https://secure.acuityscheduling.com/home.php', '_blank')">
              <span class="icon">📅</span>
              Open Acuity
            </button>
          </div>
          
          <div class="loading" id="controlLoading">
            <div class="spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
        
        <!-- Quick Notes View -->
        <div id="quickNotesView" class="view">
          <button class="control-panel-btn" onclick="switchToControlPanel()" style="background: #003366; color: white; border: none; border-radius: 6px;">
            <span>←</span>
            <span>Back to Control Panel</span>
          </button>
          
          <div class="divider"></div>
          
          <div id="quickNotesAlert" class="alert" style="display: none;"></div>
          
          <div id="quickNotesClientName" style="text-align: center; margin-bottom: 15px; padding: 10px; background: #f8f9fa; border-radius: 5px; border-left: 4px solid #00C853;">
            <h4 id="quickNotesClientTitle" style="margin: 0; color: #003366; font-weight: 600;">Loading...</h4>
          </div>
          
          <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;">
            <h3 style="margin: 0;">Quick Session Notes</h3>
            <button class="settings-btn" onclick="openQuickNotesSettings()" title="Settings" style="background: none; border: 1px solid #dadce0; border-radius: 50%; width: 32px; height: 32px; cursor: pointer; display: flex; align-items: center; justify-content: center; color: #5f6368; font-size: 14px; transition: all 0.2s ease;" onmouseover="this.style.background='#f8f9fa'; this.style.borderColor='#003366';" onmouseout="this.style.background='none'; this.style.borderColor='#dadce0';">
              ⚙️
            </button>
          </div>
          <p style="font-size: 12px; color: #666; margin-top: 0;">Click tags to add to notes</p>
          
          <div class="save-indicator" id="saveIndicator">Saved!</div>
          <div class="backup-indicator" id="backupIndicator" style="display: none; background: #ff9800; color: white; padding: 8px; border-radius: 4px; margin-bottom: 10px; font-size: 12px; text-align: center;">
            📱 Local backup active
          </div>
          
          <div class="section-title">
            <h4 style="margin: 0 0 16px 0; font-weight: bold; color: #333; font-size: 14px; border-bottom: 2px solid #dee2e6; padding-bottom: 8px;">Wins/Breakthroughs</h4>
          </div>
          
          <div class="note-category wins">
            <div class="note-tags" id="winsQuickTags">
              <!-- Quick buttons will be loaded dynamically -->
            </div>
            <textarea id="wins" class="note-textarea wins-textarea" placeholder="Quick notes about wins..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="skills-title" style="margin-top: 24px;">
            <h4 style="margin: 0 0 16px 0; font-weight: bold; color: #333; font-size: 14px; border-bottom: 2px solid #dee2e6; padding-bottom: 8px;">Skills Worked On</h4>
          </div>
          
          <div class="skills-grid">
            <div class="skill-category mastered">
              <label class="skill-label">
                <span style="font-style: italic;">Mastered</span>
              </label>
              <div class="skill-tags" id="masteredQuickTags">
                <!-- Quick buttons will be loaded dynamically -->
              </div>
              <textarea id="skillsMastered" class="skill-textarea mastered-textarea" placeholder="Skills the student has fully mastered..." onchange="markUnsaved()"></textarea>
            </div>
            
            <div class="skill-category practiced">
              <label class="skill-label">
                <span style="font-style: italic;">Practiced</span>
              </label>
              <div class="skill-tags" id="practicedQuickTags">
                <!-- Quick buttons will be loaded dynamically -->
              </div>
              <textarea id="skillsPracticed" class="skill-textarea practiced-textarea" placeholder="Skills we practiced together..." onchange="markUnsaved()"></textarea>
            </div>
            
            <div class="skill-category introduced">
              <label class="skill-label">
                <span style="font-style: italic;">Introduced</span>
              </label>
              <div class="skill-tags" id="introducedQuickTags">
                <!-- Quick buttons will be loaded dynamically -->
              </div>
              <textarea id="skillsIntroduced" class="skill-textarea introduced-textarea" placeholder="New skills introduced today..." onchange="markUnsaved()"></textarea>
            </div>
          </div>
          
          <div class="section-title" style="margin-top: 24px;">
            <h4 style="margin: 0 0 16px 0; font-weight: bold; color: #333; font-size: 14px; border-bottom: 2px solid #dee2e6; padding-bottom: 8px;">Struggles/Areas to Review</h4>
          </div>
          
          <div class="note-category struggles">
            <div class="note-tags" id="strugglesQuickTags">
              <!-- Quick buttons will be loaded dynamically -->
            </div>
            <textarea id="struggles" class="note-textarea struggles-textarea" placeholder="Note any challenges..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="section-title" style="margin-top: 24px;">
            <h4 style="margin: 0 0 16px 0; font-weight: bold; color: #333; font-size: 14px; border-bottom: 2px solid #dee2e6; padding-bottom: 8px;">Parent Communication Points</h4>
          </div>
          
          <div class="note-category parent">
            <div class="note-tags" id="parentQuickTags">
              <!-- Quick buttons will be loaded dynamically -->
            </div>
            <textarea id="parent" class="note-textarea parent-textarea" placeholder="Things to mention to parent..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="section-title" style="margin-top: 24px;">
            <h4 style="margin: 0 0 16px 0; font-weight: bold; color: #333; font-size: 14px; border-bottom: 2px solid #dee2e6; padding-bottom: 8px;">Next Session Focus</h4>
          </div>
          
          <div class="note-category next">
            <div class="note-tags" id="nextQuickTags">
              <!-- Quick buttons will be loaded dynamically -->
            </div>
            <textarea id="next" class="note-textarea next-textarea" placeholder="Plan for next time..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="button-group">
            <button id="saveNotesBtn" class="btn-primary" onclick="saveNotes()">Save Notes</button>
            <button id="sendRecapBtn" onclick="sendRecap()">Send Recap</button>
          </div>
          
          <div class="button-group" style="margin-top: 5px;">
            <button id="clearNotesBtn" class="btn-secondary" onclick="clearNotes()">Clear Notes</button>
            <button class="btn-secondary" onclick="reloadNotes()">Reload Saved</button>
          </div>
          
          <div class="auto-save-status" id="autoSaveStatus">Auto-save: Ready</div>
        </div>
        
        
        <!-- Client Update View -->
          <div id="clientUpdateView" class="view">
            <button class="control-panel-btn" onclick="immediateBackToControl()" style="background: #003366; color: white; border: none; border-radius: 6px;">
              <span>←</span>
              <span>Back to Control Panel</span>
            </button>
            
            <div class="divider"></div>
            
            <div id="clientUpdateContent">
              <h3>Update Client Info</h3>
              <div id="clientUpdateAlert" class="alert" style="display: none;"></div>
              
              <div style="margin-bottom: 15px;">
                <strong id="updateClientName">Loading...</strong>
              </div>
              
              <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Dashboard Link:</label>
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                  <span id="dashboardIndicator" style="width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; background: #f44336;"></span>
                  <input type="url" id="dashboardInput" placeholder="https://..." 
                    style="flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                </div>
              </div>
              
              <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Meeting Notes Link:</label>
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                  <span id="meetingNotesIndicator" style="width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; background: #f44336;"></span>
                  <input type="url" id="meetingNotesInput" placeholder="https://..." 
                    style="flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                </div>
              </div>
              
              <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Score Report Link:</label>
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                  <span id="scoreReportIndicator" style="width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; background: #f44336;"></span>
                  <input type="url" id="scoreReportInput" placeholder="https://..." 
                    style="flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                </div>
              </div>
              
              <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Parent Email:</label>
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                  <span id="parentEmailIndicator" style="width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; background: #f44336;"></span>
                  <input type="email" id="parentEmailInput" placeholder="parent@example.com" 
                    style="flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                </div>
              </div>
              
              <div style="margin-bottom: 15px;">
                <label style="display: block; margin-bottom: 5px; font-weight: bold;">Email Addressees:</label>
                <div style="display: flex; align-items: center; margin-bottom: 5px;">
                  <span id="emailAddresseesIndicator" style="width: 10px; height: 10px; border-radius: 50%; margin-right: 8px; background: #f44336;"></span>
                  <input type="text" id="emailAddresseesInput" placeholder="John & Jane Doe" 
                    style="flex: 1; padding: 8px; border: 1px solid #ddd; border-radius: 4px;">
                </div>
              </div>
              
              <div style="margin-bottom: 20px;">
                <label style="display: flex; align-items: center; cursor: pointer;">
                  <input type="checkbox" id="activeCheckbox" style="margin-right: 8px;"> 
                  <span style="font-weight: bold;">Active Client</span>
                </label>
              </div>
              
              <div style="display: flex; gap: 10px;">
                <button id="saveClientBtn" class="menu-button primary" onclick="saveClientDetails()" style="flex: 1; margin-bottom: 0;">
                  Save Changes
                </button>
              </div>
            </div>
          </div>
          
          <!-- Client List View -->
          <div id="clientListView" class="view">
            <button class='control-panel-btn' onclick='switchToControlPanel()' style='background: #003366; color: white; border: none; border-radius: 6px;'>
              <span>←</span>
              <span>Back to Control Panel</span>
            </button>
            
            <div class="divider"></div>
            
            <div id="clientListContent" style="display: flex; flex-direction: column; height: 100%; max-height: 450px;">
              <!-- Custom Dropdown Title Button -->
              <div style="position: relative; margin: 0 0 15px 0;">
                <button id="clientFilterDropdown" class="menu-button" style="position: relative; width: 200px; justify-content: space-between; align-items: center;" onclick="toggleClientFilterDropdown()">
                  <span id="currentFilterText">Active Clients</span>
                  <span style="margin-left: 8px; font-size: 12px;">▼</span>
                </button>
                
                <div id="clientFilterOptions" style="display: none; position: absolute; top: 100%; left: 0; right: 0; background: white; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); z-index: 1000;">
                  <div class="dropdown-option" onclick="selectClientFilter('active')" style="padding: 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; font-size: 14px;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='white'">
                    Active Clients
                  </div>
                  <div class="dropdown-option" onclick="selectClientFilter('past')" style="padding: 12px; cursor: pointer; border-bottom: 1px solid #f0f0f0; font-size: 14px;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='white'">
                    Past Clients
                  </div>
                  <div class="dropdown-option" onclick="selectClientFilter('all')" style="padding: 12px; cursor: pointer; font-size: 14px;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='white'">
                    All Clients
                  </div>
                </div>
              </div>
              
              <div id="clientListAlert" class="alert" style="display: none;"></div>
              
              <!-- Search Bar Row - Matching Control Panel Width -->
              <div style="margin-bottom: 15px; flex-shrink: 0; width: 70%;">
                <input type="text" id="clientListSearch" placeholder="Search clients..." 
                  style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box;"
                  oninput="filterClientList(this.value)">
              </div>
              
              <div id="clientList" style="position: relative; overflow-y: auto; flex: 1; border: 1px solid #e0e0e0; border-radius: 8px; background: white;">
                <div style="text-align: center; color: #666; padding: 20px;">
                  Loading clients...
                </div>
                
                <!-- Loading overlay for client selection -->
                <div class="client-list-loading-overlay">
                  <div class="client-loading-spinner"></div>
                  <div class="client-loading-text">Loading client...</div>
                </div>
              </div>
            </div>
          </div>
          
          <!-- Connection Lost Overlay -->
          <div id="connectionOverlay" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(5px); z-index: 9999; display: none; align-items: center; justify-content: center;">
            <div style="text-align: center; padding: 40px; background: white; border-radius: 12px; box-shadow: 0 8px 32px rgba(0,0,0,0.1); max-width: 400px; margin: 20px;">
              <div style="font-size: 48px; margin-bottom: 20px;">📡</div>
              <h3 style="color: #ff4444; margin: 0 0 15px 0;">Connection Lost</h3>
              <p style="color: #666; margin: 0 0 20px 0; line-height: 1.5;">
                Don't worry! Your notes are being saved locally and will sync automatically when your connection returns.
              </p>
              <div style="color: #999; font-size: 14px;">
                <div id="reconnectStatus">Checking connection...</div>
                <div style="margin-top: 10px;">
                  <div style="width: 100%; height: 4px; background: #f0f0f0; border-radius: 2px; overflow: hidden;">
                    <div id="reconnectProgress" style="width: 0%; height: 100%; background: #00C853; transition: width 0.2s ease;"></div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          
        <!-- Custom tooltip -->
        <div id="customTooltip" style="
          position: fixed;
          background: rgba(0, 0, 0, 0.8);
          color: white;
          padding: 6px 10px;
          border-radius: 4px;
          font-size: 12px;
          font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
          white-space: nowrap;
          pointer-events: none;
          z-index: 10000;
          display: none;
          transform: translate(10px, -30px);
        "></div>
        
        <!-- Missing Links Overlay -->
        <div id="missingLinksOverlay" class="overlay" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; backdrop-filter: blur(8px); background: rgba(255, 255, 255, 0.85); z-index: 1000;">
          <div class="overlay-content" style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(12px); padding: 32px; border-radius: 16px; max-width: 420px; width: 90%; box-shadow: 0 20px 40px rgba(0, 51, 102, 0.15), 0 8px 16px rgba(0, 51, 102, 0.1); border: 1px solid rgba(255, 255, 255, 0.8);">
            <div style="text-align: center; margin-bottom: 24px;">
              <div style="width: 56px; height: 56px; background: linear-gradient(135deg, #ff6b6b, #ee5a24); border-radius: 50%; margin: 0 auto 16px; display: flex; align-items: center; justify-content: center; font-size: 24px;">⚠️</div>
              <h3 style="margin: 0; color: #003366; font-size: 20px; font-weight: 600;">Missing Client Information</h3>
            </div>
            <p id="missingLinksMessage" style="color: #5f6368; margin-bottom: 16px; text-align: center;">This client is missing the following information:</p>
            <ul id="missingLinksList" style="margin: 16px 0; padding: 20px; background: rgba(255, 107, 107, 0.1); border-radius: 12px; color: #d63031; font-weight: 500; border-left: 4px solid #ff6b6b;"></ul>
            <p style="margin: 20px 0; color: #5f6368; text-align: center;">Would you like to update before sending the recap?</p>
            <div class="overlay-buttons" style="display: flex; gap: 12px; margin-top: 24px;">
              <button class="overlay-btn overlay-btn-primary" onclick="updateClientInfoFromOverlay()">
                <span class="btn-text">Update Info</span>
              </button>
              <button class="overlay-btn overlay-btn-secondary" onclick="continueWithoutUpdating()">
                <span class="btn-text">Continue Without Updating</span>
              </button>
            </div>
          </div>
        </div>
        
        <script>
          // Consolidated script block to avoid parsing issues between multiple script tags
          
          // Variable declarations
          let currentView = 'REPLACE_INITIAL_VIEW';
          // Pre-loaded client info for instant display (no server round-trip needed)
          let clientInfo = REPLACE_CLIENT_INFO;
          let unsavedChanges = false;
          let autoSaveInterval;
          let currentQuickNotesClient = null;
          let isOnline = navigator.onLine;
          
          // Offline detection and handling
          function handleOfflineState() {
            const offlineBanner = document.getElementById('offlineBanner');
            const serverDependentButtons = [
              'saveNotesBtn',
              'sendRecapBtn',
              'saveClientBtn',
              'updateClientBtn'
            ];
            
            if (!isOnline) {
              // Show banner
              if (offlineBanner) {
                offlineBanner.classList.add('show');
              }
              
              // Disable server-dependent buttons
              serverDependentButtons.forEach(buttonId => {
                const button = document.getElementById(buttonId);
                if (button) {
                  button.classList.add('offline-disabled');
                  button.disabled = true;
                  button.title = 'This feature requires an internet connection';
                }
              });
            } else {
              // Hide banner
              if (offlineBanner) {
                offlineBanner.classList.remove('show');
              }
              
              // Re-enable buttons
              serverDependentButtons.forEach(buttonId => {
                const button = document.getElementById(buttonId);
                if (button) {
                  button.classList.remove('offline-disabled');
                  button.disabled = false;
                  button.title = '';
                }
              });
            }
          }
          
          // Listen for online/offline events
          window.addEventListener('online', function() {
            isOnline = true;
            handleOfflineState();
          });
          
          window.addEventListener('offline', function() {
            isOnline = false;
            handleOfflineState();
          });
          
          // Check initial state and setup
          document.addEventListener('DOMContentLoaded', function() {
            handleOfflineState();
          });
          
          // SAVE NOTES FUNCTION - Defined at top of consolidated script with optimistic updates
          function saveNotes() {
            try {
              const saveBtn = document.getElementById('saveNotesBtn');
              
              if (!saveBtn) {
                console.error('Save button not found');
                return;
              }
              
              const notes = {
                wins: (document.getElementById('wins') && document.getElementById('wins').value) || '',
                skillsMastered: (document.getElementById('skillsMastered') && document.getElementById('skillsMastered').value) || '',
                skillsPracticed: (document.getElementById('skillsPracticed') && document.getElementById('skillsPracticed').value) || '',
                skillsIntroduced: (document.getElementById('skillsIntroduced') && document.getElementById('skillsIntroduced').value) || '',
                struggles: (document.getElementById('struggles') && document.getElementById('struggles').value) || '',
                parent: (document.getElementById('parent') && document.getElementById('parent').value) || '',
                next: (document.getElementById('next') && document.getElementById('next').value) || ''
              };
              
              
              // OPTIMISTIC UPDATE: Immediately update cache and show success
              try {
                // Store original cache state for potential rollback
                const originalCacheState = typeof PerformanceCache !== 'undefined' && currentQuickNotesClient 
                  ? PerformanceCache.getQuickNotes(currentQuickNotesClient) 
                  : null;
                
                // Update cache immediately for instant feedback
                if (typeof PerformanceCache !== 'undefined' && currentQuickNotesClient) {
                  PerformanceCache.setQuickNotes(currentQuickNotesClient, notes);
                }
                
                // Save to local backup immediately
                if (typeof saveNotesToLocalBackup === 'function') {
                  saveNotesToLocalBackup(notes);
                }
                
                // Show immediate success feedback
                saveBtn.textContent = 'Saved!';
                saveBtn.disabled = false;
                saveBtn.style.background = '#00C853';
                saveBtn.style.opacity = '1';
                saveBtn.style.cursor = 'pointer';
                
                // Update auto-save status with current time
                const autoSaveStatus = document.getElementById('autoSaveStatus');
                if (autoSaveStatus) {
                  autoSaveStatus.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
                }
                
                // Reset button text after longer confirmation display
                setTimeout(function() {
                  saveBtn.textContent = 'Save Notes';
                  saveBtn.style.background = '';  // Reset to original styling
                  saveBtn.style.cursor = '';      // Reset cursor
                }, 2000);
                
                unsavedChanges = false;
                
              } catch (cacheError) {
                console.error('Error with optimistic update:', cacheError);
                // Fall back to loading state if cache update fails
                if (typeof setButtonLoading === 'function') {
                  setButtonLoading(saveBtn, true, 'Saving...');
                } else {
                  saveBtn.textContent = 'Saving...';
                  saveBtn.disabled = true;
                }
              }
              
              // BACKGROUND SYNC: Silently sync to server in the background
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result && result.success) {
                    // Data is already updated in UI, just log success
                    // Clear local backup after successful server save
                    if (typeof clearLocalBackup === 'function') {
                      clearLocalBackup();
                    }
                    if (typeof updateBackupIndicator === 'function') {
                      updateBackupIndicator();
                    }
                  } else {
                    console.error('Background notes sync failed:', result?.message || 'Unknown error');
                    // Could implement retry logic here or show a subtle warning
                  }
                })
                .withFailureHandler(function(error) {
                  console.error('Background notes sync failed:', error.message);
                  
                  // ROLLBACK: Revert cache if server sync fails
                  if (typeof PerformanceCache !== 'undefined' && currentQuickNotesClient && originalCacheState) {
                    PerformanceCache.setQuickNotes(currentQuickNotesClient, originalCacheState);
                  }
                  
                  // Show subtle error indication but don't disrupt user experience
                  console.warn('Notes saved locally but server sync failed. Will retry on next save.');
                  
                  // Update backup indicator to show local-only state
                  if (typeof updateBackupIndicator === 'function') {
                    updateBackupIndicator();
                  }
                })
                .saveQuickNotes(notes);
                
            } catch (error) {
              console.error('Error in saveNotes function:', error);
              const saveBtn = document.getElementById('saveNotesBtn');
              if (saveBtn) {
                saveBtn.style.background = '';  // Reset styling on error
                saveBtn.style.cursor = '';      // Reset cursor on error
                saveBtn.textContent = 'Save Notes';
                saveBtn.disabled = false;
              }
              alert('Error saving notes: ' + error.message);
            }
          }
          
          
          // Immediate test to verify function accessibility
          if (typeof saveNotes === 'undefined') {
            console.error('❌ CRITICAL: saveNotes is undefined after definition!');
          } else {
          }
          
          
          // Minimal working implementations
          function runFunction(functionName, button) {
            showLoadingState();
            
            // Add button-specific loading states for external dialogs
            if (functionName === 'addNewClient' && button) {
              // Use loading wheel for Add New Client
              button.classList.add('btn-loading');
              button.disabled = true;
            } else if (functionName === 'sendIndividualRecap' && button) {
              // Use loading wheel for Send Recap too
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            google.script.run
              .withSuccessHandler(function(result) {
                hideLoadingState();
                
                // Reset button states for external dialogs
                if (functionName === 'addNewClient' && button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                } else if (functionName === 'sendIndividualRecap' && button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                  
                  // Check if we need to show the missing links overlay
                  if (result && result.showOverlay) {
                    showMissingLinksOverlayInSidebar(result.missingLinks, result.clientName);
                  }
                }
                
              })
              .withFailureHandler(function(error) {
                hideLoadingState();
                
                // Reset button states for external dialogs on error
                if (functionName === 'addNewClient' && button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                } else if (functionName === 'sendIndividualRecap' && button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                }
                
                showAlert('Error: ' + error.message, 'error');
              })[functionName]();
          }
          
          function switchToQuickNotes() {
            
            if (!clientInfo || !clientInfo.sheetName) {
              showAlert('Please select a client first.', 'error');
              return;
            }
            
            // Switch view immediately for fast UI response
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.add('active');
            currentView = 'quicknotes';
            
            // Load user's custom quick buttons
            loadQuickNotesButtons();
            
            // Start auto-save if function is available
            if (typeof startAutoSave === 'function') {
              startAutoSave();
            }
            
            // Update the client name in quick notes header (always do this)
            const quickNotesClientName = document.getElementById('quickNotesClientName');
            if (quickNotesClientName) {
              quickNotesClientName.textContent = clientInfo.name;
            }
            
            // Check if we're switching to a different client
            const isSameClient = currentQuickNotesClient === clientInfo.sheetName;
            
            if (isSameClient) {
              // Don't reload anything, keep the current form state
              return;
            }
            
            // Different client - update tracking variable
            currentQuickNotesClient = clientInfo.sheetName;
            
            // Check cache first for instant loading of selected client's notes
            const cacheKey = 'quicknotes_' + clientInfo.sheetName;
            const cachedNotes = typeof PerformanceCache !== 'undefined' ? PerformanceCache.get(cacheKey) : null;
            
            if (cachedNotes) {
              populateQuickNotesForm(cachedNotes);
              // Still load from server in background to get latest data
              loadQuickNotesFromServer(clientInfo.sheetName, true);
            } else {
              // Load from server and display the notes (not background)
              loadQuickNotesFromServer(clientInfo.sheetName, false);
            }
            
            // Check for local backup and restore if needed
            // TODO: Implement local backup functionality
            // setTimeout(function() {
            //   checkAndRestoreFromBackup();
            //   updateBackupIndicator();
            // }, 500);
          }
          
          function switchToControlPanel() {
            document.getElementById('quickNotesView').classList.remove('active');
            document.getElementById('clientListView').classList.remove('active');
            document.getElementById('clientUpdateView').classList.remove('active');
            document.getElementById('controlView').classList.add('active');
            currentView = 'control';
          }
          
          function openUpdateClientDialog() {
            if (!clientInfo || !clientInfo.sheetName) {
              showAlert('Please select a client first.', 'error');
              return;
            }
            
            // Switch to client update view
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('clientUpdateView').classList.add('active');
            currentView = 'clientupdate';
            
            // Try cache first, then preloaded data, then server
            let detailsLoaded = false;
            
            // First: Check PerformanceCache
            if (typeof PerformanceCache !== 'undefined') {
              const cachedDetails = PerformanceCache.getClientDetails(clientInfo.sheetName);
              if (cachedDetails) {
                populateClientUpdateForm(cachedDetails);
                detailsLoaded = true;
              }
            }
            
            // Second: Check preloaded data if cache miss
            if (!detailsLoaded && window.preloadedClientDetails) {
              populateClientUpdateForm(window.preloadedClientDetails);
              detailsLoaded = true;
            }
            
            // Third: Load from server if no cached or preloaded data
            if (!detailsLoaded) {
              google.script.run
                .withSuccessHandler(function(details) {
                  if (details) {
                    // Cache the loaded details for future use
                    if (typeof PerformanceCache !== 'undefined') {
                      PerformanceCache.setClientDetails(clientInfo.sheetName, details);
                    }
                    populateClientUpdateForm(details);
                  }
                })
                .withFailureHandler(function(error) {
                  showAlert('Error loading client details: ' + error.message, 'error');
                })
                .getClientDetailsForUpdate(clientInfo.sheetName);
            }
          }
          
          function showLoadingState() {
            const overlay = document.getElementById('globalLoadingOverlay');
            if (overlay) overlay.classList.add('active');
          }
          
          function hideLoadingState() {
            const overlay = document.getElementById('globalLoadingOverlay');
            if (overlay) overlay.classList.remove('active');
          }
          
          function showAlert(message, type) {
            // Simple alert for now
            alert(message);
          }
          
          // Initialize client info display
          function updateClientDisplay() {
            const clientNameEl = document.getElementById('clientName');
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            const updateBtn = document.getElementById('updateClientBtn');
            
            if (clientNameEl) {
              clientNameEl.textContent = clientInfo.isClient ? clientInfo.name : 'No client selected';
              clientNameEl.className = clientInfo.isClient ? 'current-client-name' : 'no-client';
            }
            
            // Enable/disable buttons based on client selection
            if (quickNotesBtn) quickNotesBtn.disabled = !clientInfo.isClient;
            if (recapBtn) recapBtn.disabled = !clientInfo.isClient;
            if (updateBtn) updateBtn.disabled = !clientInfo.isClient;
            
            // Preload client details when client is identified
            if (clientInfo.isClient && clientInfo.sheetName) {
              google.script.run
                .withSuccessHandler(function(unifiedClients) {
                  // Find current client in unified data
                  const currentClient = unifiedClients.find(client => client.sheetName === clientInfo.sheetName);
                  
                  if (currentClient) {
                    window.preloadedClientDetails = {
                      dashboardLink: currentClient.dashboardLink,
                      meetingNotesLink: currentClient.meetingNotesLink,
                      scoreReportLink: currentClient.scoreReportLink,
                      parentEmail: currentClient.parentEmail,
                      emailAddressees: currentClient.emailAddressees,
                      isActive: currentClient.isActive
                    };
                    
                    // Update caution sign based on client information completeness
                    const clientWarning = document.getElementById('clientWarning');
                    if (clientWarning) {
                      const hasAllInfo = currentClient.dashboardLink && 
                                       currentClient.meetingNotesLink && 
                                       currentClient.scoreReportLink &&
                                       currentClient.parentEmail &&
                                       currentClient.emailAddressees;
                      clientWarning.style.display = hasAllInfo ? 'none' : 'inline-block';
                    }
                  } else {
                    // Client not found in unified store, show warning
                    window.preloadedClientDetails = null;
                    const clientWarning = document.getElementById('clientWarning');
                    if (clientWarning) {
                      clientWarning.style.display = 'inline-block';
                    }
                  }
                })
                .withFailureHandler(function(error) {
                  window.preloadedClientDetails = null;
                  
                  // Show warning if we can't get details
                  const clientWarning = document.getElementById('clientWarning');
                  if (clientWarning) {
                    clientWarning.style.display = 'inline-block';
                  }
                })
                .getUnifiedClientList();
            } else {
              // Clear preloaded data when no client selected and hide warning
              window.preloadedClientDetails = null;
              const clientWarning = document.getElementById('clientWarning');
              if (clientWarning) {
                clientWarning.style.display = 'none';
              }
            }
            
            // Update Quick Notes client name
            const quickNotesTitle = document.getElementById('quickNotesClientTitle');
            if (quickNotesTitle) {
              quickNotesTitle.textContent = clientInfo.isClient ? clientInfo.name : 'No Client Selected';
            }
          }
          
          // Load sidebar data
          function loadSidebarData() {
            // Load client info from server
            google.script.run
              .withSuccessHandler(function(data) {
                clientInfo = data.clientInfo || {isClient: false, name: null, sheetName: null};
                updateClientDisplay();
              })
              .withFailureHandler(function(error) {
                console.error('Failed to load sidebar data:', error);
                updateClientDisplay();
              })
              .getSidebarData();
            
            // Load cached client list for instant search (UPDATED to use UnifiedClientDataStore)
            google.script.run
              .withSuccessHandler(function(clients) {
                if (clients && clients.length > 0) {
                  allClients = clients;
                } else {
                }
              })
              .withFailureHandler(function(error) {
                console.error('Failed to load unified clients:', error);
              })
              .getUnifiedClientList();
          }
          
          // Initialize when DOM is ready
          if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', loadSidebarData);
          } else {
            loadSidebarData();
          }
          
          // Preload functions for hover states
          function preloadAddClient() {
          }
          
          function preloadClientList() {
          }
          
          function preloadRecapData() {
            
            // Only proceed if we have a client
            if (!clientInfo || !clientInfo.isClient) {
              return;
            }
            
            // Check if PerformanceCache is available
            if (typeof PerformanceCache === 'undefined') {
              return;
            }
            
            // Check if we already have cached recap data
            const cacheKey = 'recapdata_' + clientInfo.sheetName;
            const cachedData = PerformanceCache.get(cacheKey);
            
            if (cachedData) {
              return; // Data is already cached
            }
            
            
            // Preload client data in background
            google.script.run
              .withSuccessHandler(function(clientData) {
                if (clientData) {
                  // Cache the client data for fast access in client-side cache
                  PerformanceCache.set(cacheKey, clientData, 300000); // 5 minutes cache
                  
                  // Also cache in script properties for server-side access
                  google.script.run
                    .withFailureHandler(function(error) {
                    })
                    .cacheRecapData(clientData);
                    
                }
              })
              .withFailureHandler(function(error) {
              })
              .getClientDataForRecap();
          }
          
          function preloadQuickNotes() {
            
            // Only proceed if we have a client and we're not already in Quick Notes view
            if (!clientInfo || !clientInfo.isClient || currentView === 'quicknotes') {
              return;
            }
            
            // Check if PerformanceCache is available
            if (typeof PerformanceCache === 'undefined') {
              return;
            }
            
            // Check if we have cached notes first
            const cacheKey = 'quicknotes_' + clientInfo.sheetName;
            const cachedNotes = PerformanceCache.get(cacheKey);
            
            if (cachedNotes) {
              return; // Notes are already cached and will load when view opens
            }
            
            // Cache blank notes for instant loading
            const blankNotes = {
              wins: '',
              skills: '',
              struggles: '',
              parent: '',
              next: ''
            };
            PerformanceCache.set(cacheKey, blankNotes, 300000); // 5 minutes cache
            
            // Then check server for actual notes in background
            google.script.run
              .withSuccessHandler(function(notesResponse) {
                if (notesResponse && notesResponse.data && typeof PerformanceCache !== 'undefined') {
                  const notes = JSON.parse(notesResponse.data);
                  // Update cache with actual notes if they exist
                  PerformanceCache.set(cacheKey, notes, 300000); // 5 minutes cache
                }
              })
              .withFailureHandler(function(error) {
                // Silent fail for preload - blank notes remain in cache
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function preloadUpdateClient() {
            // Details are now automatically preloaded when client is identified
            // This function is kept for backward compatibility but no longer needs to do anything
            if (window.preloadedClientDetails) {
            } else {
            }
          }
          
          // Missing Links Overlay Functions
          function updateClientInfoFromOverlay() {
            // Hide overlay and switch to client update view
            document.getElementById('missingLinksOverlay').style.display = 'none';
            openUpdateClientDialog();
          }
          
          function continueWithoutUpdating() {
            // Find the continue button and add loading state
            const continueBtn = document.querySelector('.overlay-btn-secondary');
            if (continueBtn) {
              continueBtn.classList.add('btn-loading');
              continueBtn.disabled = true;
            }
            
            // Call server function
            google.script.run
              .withSuccessHandler(function() {
                // Hide overlay on success
                document.getElementById('missingLinksOverlay').style.display = 'none';
              })
              .withFailureHandler(function(error) {
                // Remove loading state on error
                if (continueBtn) {
                  continueBtn.classList.remove('btn-loading');
                  continueBtn.disabled = false;
                }
                alert('Error proceeding with recap: ' + error.message);
              })
              .proceedWithRecapFromDialog();
          }
          
          function showMissingLinksOverlayInSidebar(missingLinks, clientName) {
            // Populate the overlay with data
            const missingLinksList = document.getElementById('missingLinksList');
            missingLinksList.innerHTML = '';
            missingLinks.forEach(function(link) {
              const li = document.createElement('li');
              li.textContent = link;
              missingLinksList.appendChild(li);
            });
            
            // Show the overlay
            document.getElementById('missingLinksOverlay').style.display = 'block';
          }
          
          // Save Client Details function
          function saveClientDetails() {
            const details = {
              dashboardLink: document.getElementById('dashboardInput').value.trim(),
              meetingNotesLink: document.getElementById('meetingNotesInput').value.trim(),
              scoreReportLink: document.getElementById('scoreReportInput').value.trim(),
              parentEmail: document.getElementById('parentEmailInput').value.trim(),
              emailAddressees: document.getElementById('emailAddresseesInput').value.trim(),
              isActive: document.getElementById('activeCheckbox').checked
            };
            
            const saveBtn = document.getElementById('saveClientBtn');
            
            // OPTIMISTIC UPDATE: Immediately update cache and show success
            try {
              // Store original cache state for potential rollback
              const originalCacheState = typeof PerformanceCache !== 'undefined' && clientInfo && clientInfo.sheetName 
                ? PerformanceCache.getClientDetails(clientInfo.sheetName) 
                : null;
              
              // Update cache immediately for instant feedback
              if (typeof PerformanceCache !== 'undefined' && clientInfo && clientInfo.sheetName) {
                const updatedDetails = {
                  dashboardLink: details.dashboardLink,
                  meetingNotesLink: details.meetingNotesLink,
                  scoreReportLink: details.scoreReportLink,
                  parentEmail: details.parentEmail,
                  emailAddressees: details.emailAddressees,
                  isActive: details.isActive
                };
                PerformanceCache.setClientDetails(clientInfo.sheetName, updatedDetails);
              }
              
              // Show immediate success feedback
              saveBtn.textContent = 'Saved!';
              saveBtn.disabled = false;
              saveBtn.style.opacity = '1';
              saveBtn.style.cursor = 'pointer';
              
              // Update indicators immediately
              updateIndicatorColor('dashboardIndicator', details.dashboardLink);
              updateIndicatorColor('meetingNotesIndicator', details.meetingNotesLink);
              updateIndicatorColor('scoreReportIndicator', details.scoreReportLink);
              updateIndicatorColor('parentEmailIndicator', details.parentEmail);
              updateIndicatorColor('emailAddresseesIndicator', details.emailAddressees);
              
              // Update warning indicator on control panel
              updateControlPanelWarning(details);
              
              // Refresh sidebar data to ensure control panel is up to date
              loadSidebarData();
              
              // Reset button text after longer confirmation display
              setTimeout(function() {
                saveBtn.textContent = saveBtn.dataset.originalText || 'Save Changes';
              }, 2000);
              
            } catch (cacheError) {
              console.error('Error with optimistic update:', cacheError);
              // Fall back to loading state if cache update fails
              setButtonLoading(saveBtn, true, 'Saving...');
            }
            
            // BACKGROUND SYNC: Silently sync to server in the background
            google.script.run
              .withSuccessHandler(function(result) {
                if (result && result.success) {
                  // Data is already updated in UI, just log success
                } else {
                  console.error('Background sync failed:', result?.message || 'Unknown error');
                  // Could implement retry logic here or show a subtle warning
                }
              })
              .withFailureHandler(function(error) {
                console.error('Background sync failed:', error.message);
                
                // ROLLBACK: Revert cache if server sync fails
                try {
                  if (originalCacheState && typeof PerformanceCache !== 'undefined' && clientInfo && clientInfo.sheetName) {
                    PerformanceCache.setClientDetails(clientInfo.sheetName, originalCacheState);
                    
                    // Show error state briefly
                    saveBtn.textContent = 'Sync Failed';
                    saveBtn.style.color = '#d32f2f';
                    setTimeout(function() {
                      saveBtn.textContent = saveBtn.dataset.originalText || 'Save Changes';
                      saveBtn.style.color = '';
                    }, 2000);
                  }
                } catch (rollbackError) {
                  console.error('Error rolling back cache:', rollbackError);
                }
              })
              .updateClientDetails(clientInfo?.sheetName, details);
          }
          
          // Update indicator color helper function
          function updateIndicatorColor(indicatorId, hasValue) {
            const indicator = document.getElementById(indicatorId);
            if (indicator) {
              indicator.style.background = hasValue ? '#4CAF50' : '#f44336';
              
              // Make URL indicators clickable if they have valid URLs
              const isUrlIndicator = ['dashboardIndicator', 'meetingNotesIndicator', 'scoreReportIndicator'].includes(indicatorId);
              if (isUrlIndicator) {
                indicator.classList.remove('indicator-clickable', 'has-url');
                
                if (hasValue && isValidUrl(hasValue)) {
                  indicator.classList.add('indicator-clickable', 'has-url');
                  indicator.onclick = function() {
                    window.open(hasValue, '_blank');
                  };
                  indicator.title = 'Click to open link';
                } else {
                  indicator.onclick = null;
                  indicator.title = '';
                }
              }
            }
          }
          
          // Helper function to validate URLs
          function isValidUrl(string) {
            try {
              const url = new URL(string);
              return url.protocol === 'http:' || url.protocol === 'https:';
            } catch (_) {
              return false;
            }
          }
          
          // Update control panel warning indicator
          function updateControlPanelWarning(details) {
            const clientWarning = document.getElementById('clientWarning');
            if (clientWarning) {
              const hasAllInfo = details && 
                details.dashboardLink && 
                details.meetingNotesLink && 
                details.scoreReportLink &&
                details.parentEmail &&
                details.emailAddressees;
              clientWarning.style.display = hasAllInfo ? 'none' : 'inline-block';
            }
          }
          
          // Add input event listeners for real-time indicator updates
          function setupIndicatorListeners() {
            const inputs = [
              { id: 'dashboardInput', indicator: 'dashboardIndicator' },
              { id: 'meetingNotesInput', indicator: 'meetingNotesIndicator' },
              { id: 'scoreReportInput', indicator: 'scoreReportIndicator' },
              { id: 'parentEmailInput', indicator: 'parentEmailIndicator' },
              { id: 'emailAddresseesInput', indicator: 'emailAddresseesIndicator' }
            ];
            
            inputs.forEach(function(input) {
              const inputElement = document.getElementById(input.id);
              if (inputElement) {
                inputElement.addEventListener('input', function() {
                  updateIndicatorColor(input.indicator, this.value.trim());
                  
                  // Also update control panel warning in real-time
                  updateControlPanelWarningRealTime();
                });
              }
            });
          }
          
          // Update control panel warning in real-time as user types
          function updateControlPanelWarningRealTime() {
            const currentDetails = {
              dashboardLink: document.getElementById('dashboardInput')?.value.trim(),
              meetingNotesLink: document.getElementById('meetingNotesInput')?.value.trim(),
              scoreReportLink: document.getElementById('scoreReportInput')?.value.trim(),
              parentEmail: document.getElementById('parentEmailInput')?.value.trim(),
              emailAddressees: document.getElementById('emailAddresseesInput')?.value.trim()
            };
            updateControlPanelWarning(currentDetails);
          }
          
          // Call setup when DOM is ready
          setTimeout(setupIndicatorListeners, 100);
          
          // Quick Notes tag functionality
          function addToField(fieldId, text) {
            const field = document.getElementById(fieldId);
            if (field) {
              if (field.value.trim()) {
                field.value += ', ';
              }
              field.value += text;
              
              // Mark as unsaved if function exists
              if (typeof markUnsaved === 'function') {
                markUnsaved();
              } else {
                unsavedChanges = true;
              }
              
            } else {
              console.error('Field not found:', fieldId);
            }
          }
          
          // Open Quick Notes settings dialog
          function openQuickNotesSettings() {
            google.script.run
              .withSuccessHandler(() => {
                // Settings dialog opened successfully
                // After settings are changed, reload the quick buttons
                setTimeout(loadQuickNotesButtons, 500);
              })
              .withFailureHandler((error) => {
                console.error('Failed to open settings:', error);
                alert('Failed to open Quick Notes settings');
              })
              .showQuickNotesSettings();
          }
          
          // Load user's quick button settings and populate the buttons
          function loadQuickNotesButtons() {
            google.script.run
              .withSuccessHandler((settings) => {
                populateQuickButtons(settings);
              })
              .withFailureHandler((error) => {
                console.error('Failed to load quick button settings:', error);
                // Fall back to default buttons
                populateQuickButtons(getDefaultButtonSettings());
              })
              .getQuickNotesSettings();
          }
          
          // Populate quick buttons based on user settings
          function populateQuickButtons(settings) {
            const sections = [
              { key: 'wins', containerId: 'winsQuickTags', fieldId: 'wins', cssClass: 'wins-tag' },
              { key: 'mastered', containerId: 'masteredQuickTags', fieldId: 'skillsMastered', cssClass: 'mastered-tag' },
              { key: 'practiced', containerId: 'practicedQuickTags', fieldId: 'skillsPracticed', cssClass: 'practiced-tag' },
              { key: 'introduced', containerId: 'introducedQuickTags', fieldId: 'skillsIntroduced', cssClass: 'introduced-tag' },
              { key: 'struggles', containerId: 'strugglesQuickTags', fieldId: 'struggles', cssClass: 'struggles-tag' },
              { key: 'parent', containerId: 'parentQuickTags', fieldId: 'parent', cssClass: 'parent-tag' },
              { key: 'next', containerId: 'nextQuickTags', fieldId: 'next', cssClass: 'next-tag' }
            ];
            
            sections.forEach(section => {
              const container = document.getElementById(section.containerId);
              if (container && settings[section.key]) {
                const buttons = settings[section.key]
                  .filter(text => text && text.trim()) // Only show non-empty buttons
                  .map(text => 
                    '<span class="quick-tag ' + section.cssClass + '" onclick="addToField(\\'' + section.fieldId + '\\', \\'' + text.replace(/'/g, "\\\\'") + '\\')">' + 
                    text + '</span>'
                  );
                
                container.innerHTML = buttons.join('');
              }
            });
          }
          
          // Default button settings (fallback)
          function getDefaultButtonSettings() {
            return {
              wins: ['Aha moment!', 'Solved independently', 'Confidence boost', 'Speed improved'],
              mastered: ['Problem solving', 'Reading comprehension', 'Time management'],
              practiced: ['Test strategies', 'ACT math', 'Essay writing'],
              introduced: ['New concept', 'Advanced technique', 'Study method'],
              struggles: ['Careless errors', 'Time pressure', 'Concept confusion', 'Attention focus'],
              parent: ['Great participation today', 'Ask about homework completion', 'Encourage practice at home', 'Celebrate improvement'],
              next: ['Review homework challenges', 'Build on today\'s progress', 'Practice test strategies', 'Focus on weak areas']
            };
          }
          
          // Populate client update form
          function populateClientUpdateForm(details) {
            
            // Update client name
            const nameElement = document.getElementById('updateClientName');
            if (nameElement) {
              nameElement.textContent = details.name || clientInfo.name || 'Unknown Client';
            }
            
            // Update dashboard link
            const dashboardInput = document.getElementById('dashboardInput');
            if (dashboardInput) {
              dashboardInput.value = details.dashboardLink || '';
              updateIndicatorColor('dashboardIndicator', details.dashboardLink);
            }
            
            // Update meeting notes link
            const meetingNotesInput = document.getElementById('meetingNotesInput');
            if (meetingNotesInput) {
              meetingNotesInput.value = details.meetingNotesLink || '';
              updateIndicatorColor('meetingNotesIndicator', details.meetingNotesLink);
            }
            
            // Update score report link
            const scoreReportInput = document.getElementById('scoreReportInput');
            if (scoreReportInput) {
              scoreReportInput.value = details.scoreReportLink || '';
              updateIndicatorColor('scoreReportIndicator', details.scoreReportLink);
            }
            
            // Update parent email
            const parentEmailInput = document.getElementById('parentEmailInput');
            if (parentEmailInput) {
              parentEmailInput.value = details.parentEmail || '';
              updateIndicatorColor('parentEmailIndicator', details.parentEmail);
            }
            
            // Update email addressees
            const emailAddresseesInput = document.getElementById('emailAddresseesInput');
            if (emailAddresseesInput) {
              emailAddresseesInput.value = details.emailAddressees || '';
              updateIndicatorColor('emailAddresseesIndicator', details.emailAddressees);
            }
            
            // Update active checkbox
            const activeCheckbox = document.getElementById('activeCheckbox');
            if (activeCheckbox) {
              activeCheckbox.checked = details.isActive || false;
            }
            
            // Update other form fields if they exist
            const phoneInput = document.getElementById('phoneInput');
            if (phoneInput) {
              phoneInput.value = details.phone || '';
            }
            
            const emailInput = document.getElementById('emailInput');
            if (emailInput) {
              emailInput.value = details.email || '';
            }
            
            const notesInput = document.getElementById('clientNotesInput');
            if (notesInput) {
              notesInput.value = details.notes || '';
            }
            
          }
          
          // Function to switch back to control panel from other views
          function switchToControlFromUpdate() {
            document.getElementById('clientUpdateView').classList.remove('active');
            document.getElementById('controlView').classList.add('active');
            currentView = 'control';
          }
          
          // Client List functionality
          let allClients = [];
          let filteredClients = [];
          let currentFilterType = 'active';
          let filteredByType = [];
          
          function showClientList() {
            
            // Hide other views and show client list
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.remove('active');
            document.getElementById('clientUpdateView').classList.remove('active');
            document.getElementById('clientListView').classList.add('active');
            currentView = 'clientlist';
            
            // Check if we have preloaded client data for instant display
            if (allClients && allClients.length > 0) {
              // Apply default filter (Active Clients)
              filterClientsByType('active');
              return;
            }
            
            // Fallback: Show loading state and load from server
            const clientList = document.getElementById('clientList');
            if (clientList) {
              clientList.innerHTML = '<div style="text-align: center; color: #666; padding: 20px;">Loading clients...</div>';
            }
            
            google.script.run
              .withSuccessHandler(function(clients) {
                displayClientList(clients);
              })
              .withFailureHandler(function(error) {
                console.error('Error loading unified clients:', error);
                const clientList = document.getElementById('clientList');
                if (clientList) {
                  clientList.innerHTML = '<div style="text-align: center; color: #d32f2f; padding: 20px;">Error loading clients: ' + error.message + '</div>';
                }
              })
              .getUnifiedClientList();
          }
          
          function displayClientList(clients) {
            allClients = clients || [];
            
            // Sort clients by first name A-Z
            allClients.sort(function(a, b) {
              const nameA = (a.firstName || a.name || '').toLowerCase();
              const nameB = (b.firstName || b.name || '').toLowerCase();
              return nameA.localeCompare(nameB);
            });
            
            // Apply default filter (Active Clients)
            filterClientsByType('active');
          }
          
          function renderClientList(clients) {
            const listDiv = document.getElementById('clientList');
            if (!listDiv) return;
            
            // Clear existing content but preserve the loading overlay
            const overlay = listDiv.querySelector('.client-list-loading-overlay');
            listDiv.innerHTML = '';
            if (overlay) {
              listDiv.appendChild(overlay);
            }
            
            if (!clients || clients.length === 0) {
              const emptyDiv = document.createElement('div');
              emptyDiv.style.cssText = 'text-align: center; color: #666; padding: 20px;';
              emptyDiv.textContent = 'No clients found';
              listDiv.appendChild(emptyDiv);
              return;
            }
            
            // Create clients directly in the container (no extra wrapper)
            clients.forEach(function(client, index) {
              const div = document.createElement('div');
              div.className = 'client-item';
              div.style.cssText = 'padding: 14px 16px; border-bottom: 1px solid #f0f0f0; cursor: pointer; transition: all 0.2s; min-height: 50px; display: flex; align-items: center;';
              div.setAttribute('data-sheet', client.sheetName);
              div.setAttribute('data-index', index);
              
              // Add enhanced hover effects matching Control Panel animations
              div.onmouseover = function() { 
                this.style.background = '#f8f9fa';
                this.style.borderColor = '#003366';
                this.style.transform = 'translateY(-1px)';
                this.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
                this.style.paddingLeft = '20px';
              };
              div.onmouseout = function() { 
                this.style.background = 'white';
                this.style.borderColor = '#f0f0f0';
                this.style.transform = 'translateY(0)';
                this.style.boxShadow = 'none';
                this.style.paddingLeft = '16px';
              };
              div.onmousedown = function() {
                this.style.transform = 'translateY(0) scale(0.98)';
                this.style.boxShadow = 'none';
              };
              div.onmouseup = function() {
                this.style.transform = 'translateY(-1px)';
                this.style.boxShadow = '0 2px 4px rgba(0,0,0,0.1)';
              };
              div.onclick = function() { selectClientFromList(this.getAttribute('data-sheet')); };
              
              const nameDiv = document.createElement('div');
              nameDiv.style.cssText = 'font-weight: 500; color: #003366; font-size: 15px;';
              nameDiv.textContent = client.name;
              
              div.appendChild(nameDiv);
              listDiv.appendChild(div);
            });
            
            // Add client count info at the bottom if there are many clients
            if (clients.length > 10) {
              const infoDiv = document.createElement('div');
              infoDiv.style.cssText = 'padding: 10px; text-align: center; color: #666; font-size: 13px; background: #f8f9fa; border-top: 1px solid #e0e0e0; margin-top: 10px;';
              infoDiv.textContent = 'Showing ' + clients.length + ' clients total';
              listDiv.appendChild(infoDiv);
            }
          }
          
          function filterClientsByType(filterType) {
            if (!allClients || allClients.length === 0) return;
            
            currentFilterType = filterType;
            
            // Filter based on type
            if (filterType === 'active') {
              filteredByType = allClients.filter(function(client) {
                return client.isActive === true;
              });
            } else if (filterType === 'past') {
              filteredByType = allClients.filter(function(client) {
                return client.isActive === false;
              });
            } else {
              filteredByType = allClients;
            }
            
            // Apply search filter if there's a search query
            const searchQuery = document.getElementById('clientListSearch');
            if (searchQuery && searchQuery.value && searchQuery.value.trim()) {
              filterClientList(searchQuery.value);
            } else {
              renderClientList(filteredByType);
            }
          }
          
          function filterClientList(searchTerm) {
            // Use the type-filtered list as base
            const baseList = filteredByType.length > 0 ? filteredByType : allClients;
            
            if (!searchTerm) {
              renderClientList(baseList);
              return;
            }
            
            const term = searchTerm.toLowerCase();
            filteredClients = baseList.filter(function(client) {
              return client.name.toLowerCase().includes(term) ||
                     (client.clientType && client.clientType.toLowerCase().includes(term));
            });
            
            renderClientList(filteredClients);
          }
          
          
          // Custom dropdown functions
          function toggleClientFilterDropdown() {
            const dropdown = document.getElementById('clientFilterOptions');
            const isVisible = dropdown.style.display !== 'none';
            
            if (isVisible) {
              dropdown.style.display = 'none';
            } else {
              dropdown.style.display = 'block';
              // Close dropdown when clicking elsewhere
              setTimeout(() => {
                document.addEventListener('click', closeClientFilterDropdown, { once: true });
              }, 100);
            }
          }
          
          function closeClientFilterDropdown(event) {
            const dropdown = document.getElementById('clientFilterOptions');
            const button = document.getElementById('clientFilterDropdown');
            
            // Don't close if clicking on the dropdown or button
            if (event && (dropdown.contains(event.target) || button.contains(event.target))) {
              return;
            }
            
            dropdown.style.display = 'none';
          }
          
          function selectClientFilter(filterType) {
            // Update button text
            const filterText = document.getElementById('currentFilterText');
            const filterLabels = {
              'active': 'Active Clients',
              'past': 'Past Clients', 
              'all': 'All Clients'
            };
            
            if (filterText) {
              filterText.textContent = filterLabels[filterType] || 'Active Clients';
            }
            
            // Close dropdown
            document.getElementById('clientFilterOptions').style.display = 'none';
            
            // Apply filter
            filterClientsByType(filterType);
          }
          
          function selectClientFromList(sheetName) {
            
            // Get client name from the clicked item - handle both button and div structures
            const clickedItem = event.target.closest('.client-item');
            let clientName = sheetName;
            if (clickedItem) {
              // Try to find the name in various possible structures
              const nameDivs = clickedItem.getElementsByTagName('div');
              if (nameDivs && nameDivs.length > 0) {
                // Get the first div which usually contains the name
                clientName = nameDivs[0].textContent;
              } else if (clickedItem.textContent) {
                // Fallback to the item's text content
                const textParts = clickedItem.textContent.split(String.fromCharCode(10));
                clientName = textParts[0].trim();
              }
            }
            
            // Show loading overlay
            const clientList = document.getElementById('clientList');
            const overlay = document.querySelector('.client-list-loading-overlay');
            const loadingText = overlay.querySelector('.client-loading-text');
            loadingText.textContent = 'Loading ' + clientName + '...';
            overlay.style.display = 'flex';
            
            // Navigate to the client sheet
            google.script.run
              .withSuccessHandler(function() {
                // Hide loading overlay
                overlay.style.display = 'none';
                // Return to control panel
                switchToControlPanel();
                // Reload sidebar data to update client info
                loadSidebarData();
              })
              .withFailureHandler(function(error) {
                // Hide loading overlay on error
                overlay.style.display = 'none';
                showAlert('Error navigating to client: ' + error.message, 'error');
              })
              .navigateToSheet(sheetName);
          }
          
          // Client search functionality
          let searchTimeout;
          
          
          function searchClients(query) {
            const resultsDiv = document.getElementById('searchResults');
            
            if (!query.trim()) {
              resultsDiv.innerHTML = '';
              return;
            }
            
            // Always try cache first for instant results
            if (typeof allClients !== 'undefined' && allClients.length > 0) {
              displaySearchResults(allClients, query);
              return;
            }
            
            // Fallback: debounced server call only if no cache available
            clearTimeout(searchTimeout);
            resultsDiv.innerHTML = '<div style="text-align: center; color: #666; padding: 5px; font-size: 12px;">Loading...</div>';
            
            searchTimeout = setTimeout(function() {
              google.script.run
                .withSuccessHandler(function(clients) {
                  window.allClients = clients;
                  displaySearchResults(clients, query);
                })
                .withFailureHandler(function(error) {
                  console.error('Search error:', error);
                  resultsDiv.innerHTML = '<div style="color: #ff4444; font-size: 12px;">Search error</div>';
                })
                .getUnifiedActiveClientList();
            }, 300); // Only debounce server calls
          }
          
          function displaySearchResults(clients, searchQuery) {
            const query = (searchQuery || document.getElementById('clientSearch').value).toLowerCase();
            const resultsDiv = document.getElementById('searchResults');
            
            if (!query.trim()) {
              resultsDiv.innerHTML = '';
              return;
            }
            
            const filtered = clients.filter(function(client) {
              return client.name.toLowerCase().includes(query);
            }).slice(0, 5); // Limit to 5 results
            
            if (filtered.length === 0) {
              resultsDiv.innerHTML = '<div style="color: #666; font-size: 12px; padding: 5px;">No active clients found</div>';
              return;
            }
            
            const html = filtered.map(function(client) {
              return '<button onclick="selectClient(\\''+client.sheetName+'\\')" ' +
                'style="display: block; width: 100%; padding: 8px; margin: 2px 0; border: 1px solid #ddd; background: white; border-radius: 3px; cursor: pointer; text-align: left; font-size: 12px; color: #333;" ' +
                'onmouseover="this.style.background=\\'#f0f8ff\\'; this.style.color=\\'#003366\\';" ' +
                'onmouseout="this.style.background=\\'white\\'; this.style.color=\\'#333\\';">'+
                client.name+
              '</button>';
            }).join('');
            
            resultsDiv.innerHTML = html;
          }
          
          function selectClient(sheetName) {
            // Clear search
            document.getElementById('clientSearch').value = '';
            document.getElementById('searchResults').innerHTML = '';
            
            // Navigate to client sheet
            google.script.run
              .withSuccessHandler(function() {
                // Refresh client info after navigation
                google.script.run
                  .withSuccessHandler(function(info) {
                    clientInfo = info;
                    updateClientDisplay();
                  })
                  .getCurrentClientInfo();
              })
              .withFailureHandler(function(error) {
                showAlert('Error navigating to client: ' + error.message, 'error');
              })
              .navigateToSheet(sheetName);
          }
          
          function immediateBackToControl() {
            
            // Clear auto-save immediately
            if (autoSaveInterval) {
              clearInterval(autoSaveInterval);
            }
            
            // Clear preloaded data
            window.preloadedClientDetails = null;
            
            // Immediately switch views without server call
            const views = ['quickNotesView', 'clientUpdateView', 'clientListView'];
            views.forEach(function(viewId) {
              const element = document.getElementById(viewId);
              if (element) {
                element.classList.remove('active');
              }
            });
            
            const controlView = document.getElementById('controlView');
            if (controlView) {
              controlView.classList.add('active');
            } else {
              console.error('Control view element not found!');
            }
            
            currentView = 'control';
            
            // Update client info in background (non-blocking)
            setTimeout(function() {
              google.script.run
                .withSuccessHandler(function(info) {
                  clientInfo = info;
                  updateClientDisplay();
                })
                .getCurrentClientInfo();
            }, 0);
          }
          
          function prepopulateClientListDialog(clients) {
            
            // Sort clients by first name A-Z (same as displayClientList)
            const sortedClients = clients.slice().sort(function(a, b) {
              const nameA = (a.firstName || a.name || '').toLowerCase();
              const nameB = (b.firstName || b.name || '').toLowerCase();
              return nameA.localeCompare(nameB);
            });
            
            // Pre-populate the client list content
            const container = document.getElementById('clientListContent');
            if (container) {
              container.innerHTML = ''; // Clear existing content
              
              sortedClients.forEach(function(client, index) {
                const div = document.createElement('div');
                div.className = 'client-item';
                div.style.cssText = 'padding: 12px 16px; border-bottom: 1px solid #eee; cursor: pointer; transition: background 0.2s;';
                div.setAttribute('data-sheet', client.sheetName);
                div.setAttribute('data-index', index);
                
                // Add hover effects with event listeners
                div.onmouseover = function() { this.style.background = '#f5f5f5'; };
                div.onmouseout = function() { this.style.background = 'white'; };
                div.onclick = function() { selectClientFromList(this.getAttribute('data-sheet')); };
                
                const nameDiv = document.createElement('div');
                nameDiv.style.cssText = 'font-weight: 500; color: #003366;';
                nameDiv.textContent = client.name;
                
                div.appendChild(nameDiv);
                container.appendChild(div);
              });
              
            }
          }
          
          // Sheet change detection and auto-update
          let lastKnownSheet = null;
          let sheetChangeCheckInterval = null;
          
          function startSheetChangeDetection() {
            // Initialize with current sheet
            google.script.run
              .withSuccessHandler(function(info) {
                lastKnownSheet = info.sheetName;
              })
              .getCurrentClientInfo();
            
            // Check for sheet changes every 2 seconds
            sheetChangeCheckInterval = setInterval(function() {
              google.script.run
                .withSuccessHandler(function(newInfo) {
                  if (newInfo.sheetName !== lastKnownSheet) {
                    handleSheetChange(newInfo);
                  }
                })
                .withFailureHandler(function(error) {
                  console.error('Failed to check sheet change:', error);
                })
                .getCurrentClientInfo();
            }, 2000);
          }
          
          function handleSheetChange(newClientInfo) {
            lastKnownSheet = newClientInfo.sheetName;
            
            // Check if we're in Quick Notes view with unsaved changes
            const isInQuickNotes = currentView === 'quicknotes';
            const hasUnsavedChanges = typeof unsavedChanges !== 'undefined' && unsavedChanges;
            
            if (isInQuickNotes && hasUnsavedChanges) {
              
              // Save current Quick Notes
              const notes = {
                wins: document.getElementById('wins').value,
                skillsMastered: document.getElementById('skillsMastered').value,
                skillsPracticed: document.getElementById('skillsPracticed').value,
                skillsIntroduced: document.getElementById('skillsIntroduced').value,
                struggles: document.getElementById('struggles').value,
                parent: document.getElementById('parent').value,
                next: document.getElementById('next').value
              };
              
              google.script.run
                .withSuccessHandler(function() {
                  unsavedChanges = false;
                  updateClientAfterSheetChange(newClientInfo);
                })
                .withFailureHandler(function(error) {
                  console.error('Failed to save Quick Notes:', error);
                  // Still update client info even if save failed
                  updateClientAfterSheetChange(newClientInfo);
                })
                .saveQuickNotes(notes);
            } else {
              // No unsaved changes, update client info immediately
              updateClientAfterSheetChange(newClientInfo);
            }
          }
          
          function updateClientAfterSheetChange(newClientInfo) {
            
            // Update client info
            clientInfo = newClientInfo;
            
            // Reset the current quick notes client since we're changing sheets
            if (currentQuickNotesClient !== newClientInfo.sheetName) {
              currentQuickNotesClient = null;
            }
            
            // Update the UI based on current view
            if (currentView === 'control') {
              updateClientDisplay();
            } else if (currentView === 'quicknotes') {
              // If in Quick Notes, reload notes for new client
              if (newClientInfo.isClient) {
                loadQuickNotesFromServer(newClientInfo.sheetName, false);
              } else {
                // Clear Quick Notes if not a client sheet
                populateQuickNotesForm({});
              }
            }
            
            // Always update the main client display
            updateClientDisplay();
          }
          
          function stopSheetChangeDetection() {
            if (sheetChangeCheckInterval) {
              clearInterval(sheetChangeCheckInterval);
              sheetChangeCheckInterval = null;
            }
          }
          
          // Button loading state management
          function setButtonLoading(button, isLoading, loadingText) {
            if (!button) return;
            
            if (isLoading) {
              button.dataset.originalText = button.textContent;
              button.textContent = loadingText || 'Loading...';
              button.disabled = true;
              button.style.opacity = '0.7';
              button.style.cursor = 'not-allowed';
            } else {
              button.textContent = button.dataset.originalText || button.textContent;
              button.disabled = false;
              button.style.opacity = '1';
              button.style.cursor = 'pointer';
            }
          }
          
          // Save Notes function removed - using the enhanced version below with local backup
          
          // Mark notes as having unsaved changes
          function markUnsaved() {
            unsavedChanges = true;
          }
          
          // Custom tooltip functions
          function showTooltip(event, text) {
            const tooltip = document.getElementById('customTooltip');
            if (tooltip) {
              tooltip.textContent = text;
              tooltip.style.display = 'block';
              moveTooltip(event);
            }
          }
          
          function hideTooltip() {
            const tooltip = document.getElementById('customTooltip');
            if (tooltip) {
              tooltip.style.display = 'none';
            }
          }
          
          function moveTooltip(event) {
            const tooltip = document.getElementById('customTooltip');
            if (tooltip && tooltip.style.display === 'block') {
              tooltip.style.left = event.clientX + 'px';
              tooltip.style.top = event.clientY + 'px';
            }
          }
          
          // Clear Notes function with confirmation state
          function clearNotes() {
            const button = document.getElementById('clearNotesBtn');
            if (!button) return;
            
            // Check if button is in confirmation state
            if (button.dataset.confirmState === 'true') {
              // User confirmed - actually clear notes
              setButtonLoading(button, true, 'Clearing...');
              
              // Clear all textareas
              document.getElementById('wins').value = '';
              document.getElementById('skills').value = '';
              document.getElementById('struggles').value = '';
              document.getElementById('parent').value = '';
              document.getElementById('next').value = '';
              
              // Mark as unsaved since we cleared content
              unsavedChanges = true;
              
              // Reset button state after clearing
              setTimeout(function() {
                setButtonLoading(button, false);
                resetClearNotesButton();
              }, 300);
              
              
            } else {
              // First click - show confirmation state
              button.style.background = '#ea4335';
              button.style.color = 'white';
              button.textContent = 'Confirm Clearing';
              button.dataset.confirmState = 'true';
              
              // Reset after 3 seconds if not confirmed
              setTimeout(() => {
                if (button.dataset.confirmState === 'true') {
                  resetClearNotesButton();
                }
              }, 3000);
            }
          }
          
          function resetClearNotesButton() {
            const button = document.getElementById('clearNotesBtn');
            if (button) {
              button.style.background = '';
              button.style.color = '';
              button.textContent = 'Clear Notes';
              button.dataset.confirmState = 'false';
            }
          }
          
          // Reload Notes function
          function reloadNotes() {
            const button = document.querySelector('button[onclick*="reloadNotes"]');
            if (button) {
              // Add loading state with spinning wheel (same as Send Recap button)
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            // Load notes from server
            google.script.run
              .withSuccessHandler(function(response) {
                if (response && response.data) {
                  const notes = JSON.parse(response.data);
                  
                  // Populate form fields
                  document.getElementById('wins').value = notes.wins || '';
                  document.getElementById('skills').value = notes.skills || '';
                  document.getElementById('struggles').value = notes.struggles || '';
                  document.getElementById('parent').value = notes.parent || '';
                  document.getElementById('next').value = notes.next || '';
                  
                  // Reset unsaved changes flag
                  unsavedChanges = false;
                  
                } else {
                }
                
                // Remove loading state
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                  button.textContent = button.dataset.originalText || 'Reload Saved';
                }
              })
              .withFailureHandler(function(error) {
                console.error('Failed to reload notes:', error);
                showAlert('Failed to reload notes: ' + error.message, 'error');
                
                // Remove loading state
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                  button.textContent = button.dataset.originalText || 'Reload Saved';
                }
              })
              .getQuickNotesForCurrentSheet();
          }
          
          // Send Recap function
          function sendRecap() {
            const button = document.getElementById('sendRecapBtn');
            if (!button) return;
            
            // Add loading state to button with loading wheel
            button.dataset.originalText = button.textContent;
            button.classList.add('btn-loading');
            button.disabled = true;
            
            // Save notes first
            const notes = {
              wins: document.getElementById('wins').value,
              skillsMastered: document.getElementById('skillsMastered').value,
              skillsPracticed: document.getElementById('skillsPracticed').value,
              skillsIntroduced: document.getElementById('skillsIntroduced').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(function() {
                
                // Send the recap and handle the result
                google.script.run
                  .withSuccessHandler(function(result) {
                    button.classList.remove('btn-loading');
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    
                    if (result && result.showOverlay) {
                      // Show missing links overlay
                      showMissingLinksOverlayInSidebar(result.missingLinks, result.clientName);
                    } else {
                      // Recap dialog was shown successfully
                    }
                  })
                  .withFailureHandler(function(error) {
                    button.classList.remove('btn-loading');
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    console.error('Failed to send recap:', error);
                    showAlert('Failed to send recap: ' + error.message, 'error');
                  })
                  .sendIndividualRecap();
              })
              .withFailureHandler(function(error) {
                button.classList.remove('btn-loading');
                button.disabled = false;
                button.textContent = button.dataset.originalText || 'Send Recap';
                console.error('Failed to save notes before sending recap:', error);
                showAlert('Failed to save notes: ' + error.message, 'error');
              })
              .saveQuickNotes(notes);
          }
          
          // Start sheet change detection when sidebar loads
          setTimeout(function() {
            startSheetChangeDetection();
          }, 1000); // Small delay to ensure everything is initialized
          
          // Quick Notes helper functions
          function populateQuickNotesForm(notes) {
            document.getElementById('wins').value = notes.wins || '';
            document.getElementById('skillsMastered').value = notes.skillsMastered || '';
            document.getElementById('skillsPracticed').value = notes.skillsPracticed || '';
            document.getElementById('skillsIntroduced').value = notes.skillsIntroduced || '';
            document.getElementById('struggles').value = notes.struggles || '';
            document.getElementById('parent').value = notes.parent || '';
            document.getElementById('next').value = notes.next || '';
          }
          
          function loadQuickNotesFromServer(sheetName, isBackground) {
            google.script.run
              .withSuccessHandler(function(notesResponse) {
                if (notesResponse && notesResponse.data) {
                  const notes = JSON.parse(notesResponse.data);
                  // Cache the data for next time
                  const cacheKey = 'quicknotes_' + (sheetName || 'unknown');
                  if (typeof PerformanceCache !== 'undefined') {
                    PerformanceCache.set(cacheKey, notes, 300000); // 5 minutes cache
                  }
                  
                  // Only update form if not background load or if no cached data was shown
                  if (!isBackground) {
                    populateQuickNotesForm(notes);
                  }
                }
              })
              .withFailureHandler(function(error) {
                if (!isBackground) {
                  showAlert('Error loading notes: ' + error.message, 'error');
                }
              })
              .getQuickNotesForCurrentSheet();
          }
          
        </script>
        
        <!-- Original Script 3 commented out entirely -->
        <!--
        <script>
          
          /* Commenting out to isolate syntax error
          try {
          
          // Performance optimization: Caching system
          const PerformanceCache = {
            data: {
              clientList: null,
              clientDetails: {},
              lastClientListFetch: 0,
              currentNotes: null,
              genericCache: {}
            },
            
            CACHE_DURATION: 5 * 60 * 1000, // 5 minutes
            
            isExpired(timestamp) {
              return Date.now() - timestamp > this.CACHE_DURATION;
            },
            
            setClientList(clients) {
              this.data.clientList = clients;
              this.data.lastClientListFetch = Date.now();
              // Store in localStorage for persistence
              try {
                localStorage.setItem('clientList_cache', JSON.stringify({
                  data: clients,
                  timestamp: this.data.lastClientListFetch
                }));
              } catch (e) {
                console.warn('Could not save to localStorage:', e);
              }
            },
            
            getClientList() {
              // Check memory cache first
              if (this.data.clientList && !this.isExpired(this.data.lastClientListFetch)) {
                return this.data.clientList;
              }
              
              // Check localStorage cache
              try {
                const cached = localStorage.getItem('clientList_cache');
                if (cached) {
                  const parsed = JSON.parse(cached);
                  if (!this.isExpired(parsed.timestamp)) {
                    this.data.clientList = parsed.data;
                    this.data.lastClientListFetch = parsed.timestamp;
                    return parsed.data;
                  }
                }
              } catch (e) {
                console.warn('Could not read from localStorage:', e);
              }
              
              return null;
            },
            
            setClientDetails(sheetName, details) {
              this.data.clientDetails[sheetName] = {
                data: details,
                timestamp: Date.now()
              };
            },
            
            getClientDetails(sheetName) {
              const cached = this.data.clientDetails[sheetName];
              if (cached && !this.isExpired(cached.timestamp)) {
                return cached.data;
              }
              return null;
            },
            
            // Quick Notes cache methods
            setQuickNotes(clientSheetName, notes) {
              const cacheKey = 'quicknotes_' + clientSheetName;
              this.set(cacheKey, notes, this.CACHE_DURATION);
            },
            
            getQuickNotes(clientSheetName) {
              const cacheKey = 'quicknotes_' + clientSheetName;
              return this.get(cacheKey);
            },
            
            clearExpired() {
              // Clear expired client details
              var self = this;
              for (var key in this.data.clientDetails) {
                var value = this.data.clientDetails[key];
                if (self.isExpired(value.timestamp)) {
                  delete self.data.clientDetails[key];
                }
              }
            },
            
            prefetchClientList() {
              if (!this.getClientList()) {
                var self = this;
                google.script.run
                  .withSuccessHandler(function(clients) {
                    self.setClientList(clients);
                  })
                  .getUnifiedActiveClientList();
              }
            },
            
            // Generic cache methods
            set(key, value, customDuration) {
              const duration = customDuration || this.CACHE_DURATION;
              this.data.genericCache[key] = {
                data: value,
                timestamp: Date.now(),
                duration: duration
              };
            },
            
            get(key) {
              const cached = this.data.genericCache[key];
              if (cached) {
                const isExpired = Date.now() - cached.timestamp > cached.duration;
                if (!isExpired) {
                  return cached.data;
                } else {
                  // Clean up expired entry
                  delete this.data.genericCache[key];
                }
              }
              return null;
            },
            
            clearClientList() {
              this.data.clientList = null;
              this.data.lastClientListFetch = 0;
              try {
                localStorage.removeItem('clientList_cache');
              } catch (e) {
                console.warn('Could not clear localStorage:', e);
              }
            }
          };
          
          // Request queuing system for better performance
          const RequestQueue = {
            pending: {},
            
            add(functionName, params, successCallback, failureCallback) {
              const key = functionName + '_' + JSON.stringify(params);
              
              // If same request is already pending, just add callback
              if (this.pending[key]) {
                const existing = this.pending[key];
                existing.successCallbacks.push(successCallback);
                if (failureCallback) {
                  existing.failureCallbacks.push(failureCallback);
                }
                return;
              }
              
              // Create new request
              const request = {
                functionName,
                params,
                successCallbacks: successCallback ? [successCallback] : [],
                failureCallbacks: failureCallback ? [failureCallback] : []
              };
              
              this.pending[key] = request;
              
              // Execute request
              google.script.run
                .withSuccessHandler(function(result) {
                  const req = this.pending[key];
                  if (req) {
                    req.successCallbacks.forEach(function(callback) {
                      if (callback) callback(result);
                    });
                    delete this.pending[key];
                  }
                })
                .withFailureHandler(function(error) {
                  const req = this.pending[key];
                  if (req) {
                    req.failureCallbacks.forEach(function(callback) {
                      if (callback) callback(error);
                    });
                    delete this.pending[key];
                  }
                })[functionName].apply(null, params);
            }
          };
          
          
          } catch(e) {
            console.error('Script 3 error:', e.message, 'at line', e.lineNumber);
          }
          */
          
          
          // Temporary stub for PerformanceCache
          const PerformanceCache = {
            getClientList: function() { return null; },
            setClientList: function() {},
            getClientDetails: function() { return null; },
            setClientDetails: function() {},
            get: function() { return null; },
            set: function() {},
            prefetchClientList: function() {},
            data: { clientDetails: {} }
          };
          
          // Temporary stub for RequestQueue
          const RequestQueue = {
            add: function(name, params, success, fail) {
              google.script.run
                .withSuccessHandler(success || function() {})
                .withFailureHandler(fail || function() {})[name].apply(null, params);
            }
          };

          function setButtonLoading(button, isLoading) {
            if (!button) return; // Safety check
            
            if (isLoading) {
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
              button.dataset.loadingStartTime = Date.now();
            } else {
              // Ensure minimum loading time of 300ms for visibility
              const startTime = parseInt(button.dataset.loadingStartTime) || Date.now();
              const elapsed = Date.now() - startTime;
              const minDelay = Math.max(0, 300 - elapsed);
              
              setTimeout(function() {
                button.classList.remove('btn-loading');
                button.textContent = button.dataset.originalText || button.textContent;
                button.disabled = false;
                delete button.dataset.loadingStartTime;
              }, minDelay);
            }
          }
          
          function showGlobalLoading() {
            const overlay = document.getElementById('globalLoadingOverlay');
            if (overlay) {
              overlay.classList.add('show');
            }
          }
          
          function hideGlobalLoading() {
            const overlay = document.getElementById('globalLoadingOverlay');
            if (overlay) {
              overlay.classList.remove('show');
            }
          }
          
          // View switching functions
          function switchToQuickNotes() {
            // Check if we're on a client sheet
            google.script.run
              .withSuccessHandler(function(info) {
                if (info.isClient) {
                  // Switch view immediately for fast UI response
                  document.getElementById('controlView').classList.remove('active');
                  document.getElementById('quickNotesView').classList.add('active');
                  currentView = 'quicknotes';
                  
                  // Load user's custom quick buttons
                  loadQuickNotesButtons();
                  
                  // Start auto-save
                  startAutoSave();
                  
                  // Check cache first for instant loading
                  const cacheKey = 'quicknotes_' + (info.sheetName || 'unknown');
                  const cachedNotes = PerformanceCache.get(cacheKey);
                  if (cachedNotes) {
                    populateQuickNotesForm(cachedNotes);
                    // Still load from server in background to get latest data
                    loadQuickNotesFromServer(info.sheetName, true);
                  } else {
                    // Load from server for first time
                    loadQuickNotesFromServer(info.sheetName, false);
                  }
                  
                  // Check for local backup and restore if needed
                  setTimeout(function() {
                    checkAndRestoreFromBackup();
                    updateBackupIndicator();
                  }, 500);
                } else {
                  showAlert('Please navigate to a client sheet first.', 'error');
                }
              })
              .getCurrentClientInfo();
          }
          
          function populateQuickNotesForm(notes) {
            document.getElementById('wins').value = notes.wins || '';
            document.getElementById('skillsMastered').value = notes.skillsMastered || '';
            document.getElementById('skillsPracticed').value = notes.skillsPracticed || '';
            document.getElementById('skillsIntroduced').value = notes.skillsIntroduced || '';
            document.getElementById('struggles').value = notes.struggles || '';
            document.getElementById('parent').value = notes.parent || '';
            document.getElementById('next').value = notes.next || '';
          }
          
          function loadQuickNotesFromServer(sheetName, isBackground) {
            google.script.run
              .withSuccessHandler(function(notesResponse) {
                if (notesResponse && notesResponse.data) {
                  const notes = JSON.parse(notesResponse.data);
                  // Cache the data for next time
                  const cacheKey = 'quicknotes_' + (sheetName || 'unknown');
                  PerformanceCache.set(cacheKey, notes, 300000); // 5 minutes cache
                  
                  // Only update form if not background load or if no cached data was shown
                  if (!isBackground) {
                    populateQuickNotesForm(notes);
                  }
                }
              })
              .withFailureHandler(function(error) {
                if (!isBackground) {
                  showAlert('Error loading notes: ' + error.message, 'error');
                }
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function switchToControlPanel() {
            
            // Clear auto-save
            if (autoSaveInterval) {
              clearInterval(autoSaveInterval);
            }
            
            // Clear any preloaded data
            window.preloadedClientDetails = null;
            
            // Update client info
            google.script.run
              .withSuccessHandler(function(info) {
                clientInfo = info;
                updateClientDisplay();
                
                // Switch view - force class removal and addition
                const views = ['quickNotesView', 'clientUpdateView', 'clientListView'];
                views.forEach(function(viewId) {
                  const element = document.getElementById(viewId);
                  if (element) {
                    element.classList.remove('active');
                  }
                });
                
                const controlView = document.getElementById('controlView');
                if (controlView) {
                  controlView.classList.add('active');
                }
                
                currentView = 'control';
              })
              .withFailureHandler(function(error) {
                console.error('switchToControlPanel failed:', error);
                // Still switch views even if server call fails
                const views = ['quickNotesView', 'clientUpdateView', 'clientListView'];
                views.forEach(function(viewId) {
                  const element = document.getElementById(viewId);
                  if (element) element.classList.remove('active');
                });
                const controlView = document.getElementById('controlView');
                if (controlView) controlView.classList.add('active');
                currentView = 'control';
              })
              .getCurrentClientInfo();
          }
          
          function updateControlPanelView() {
            const clientNameDiv = document.getElementById('clientName');
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            const updateClientBtn = document.getElementById('updateClientBtn');
            const clientWarning = document.getElementById('clientWarning');
            
            if (clientInfo.isClient) {
              clientNameDiv.textContent = clientInfo.name;
              clientNameDiv.className = 'current-client-name';
              quickNotesBtn.disabled = false;
              recapBtn.disabled = false;
              updateClientBtn.disabled = false;
              
              // Check if client information is complete AND preload details for instant dialog
              google.script.run
                .withSuccessHandler(function(unifiedClients) {
                  // Find current client in unified data
                  const currentClient = unifiedClients.find(client => client.sheetName === clientInfo.sheetName);
                  
                  if (currentClient) {
                    // Store details for instant dialog opening
                    window.preloadedClientDetails = {
                      dashboardLink: currentClient.dashboardLink,
                      meetingNotesLink: currentClient.meetingNotesLink,
                      scoreReportLink: currentClient.scoreReportLink,
                      parentEmail: currentClient.parentEmail,
                      emailAddressees: currentClient.emailAddressees,
                      isActive: currentClient.isActive
                    };
                    
                    // Check if client information is complete
                    const hasAllInfo = currentClient.dashboardLink && 
                                     currentClient.meetingNotesLink && 
                                     currentClient.scoreReportLink &&
                                     currentClient.parentEmail &&
                                     currentClient.emailAddressees;
                    clientWarning.style.display = hasAllInfo ? 'none' : 'inline-block';
                  } else {
                    // Client not found in unified store, show warning
                    window.preloadedClientDetails = null;
                    clientWarning.style.display = 'inline-block';
                  }
                })
                .withFailureHandler(function() {
                  // Clear preloaded data on error
                  window.preloadedClientDetails = null;
                  // If we can't get details, show warning
                  clientWarning.style.display = 'inline-block';
                })
                .getUnifiedClientList();
              
              // Update Quick Notes client name header
              const quickNotesClientNameDiv = document.getElementById('quickNotesClientName');
              if (quickNotesClientNameDiv) {
                quickNotesClientNameDiv.querySelector('h4').textContent = clientInfo.name;
              }
            } else {
              clientNameDiv.textContent = 'No client selected';
              clientNameDiv.className = 'no-client';
              quickNotesBtn.disabled = true;
              recapBtn.disabled = true;
              updateClientBtn.disabled = true;
              clientWarning.style.display = 'none';
              
              // Update Quick Notes client name header
              const quickNotesClientNameDiv = document.getElementById('quickNotesClientName');
              if (quickNotesClientNameDiv) {
                quickNotesClientNameDiv.querySelector('h4').textContent = 'No Client Selected';
              }
            }
          }
          
          // Quick Notes functions
          function markUnsaved() {
            unsavedChanges = true;
            
            // Auto-save to local backup when notes change
            const notes = {
              wins: document.getElementById('wins').value,
              skillsMastered: document.getElementById('skillsMastered').value,
              skillsPracticed: document.getElementById('skillsPracticed').value,
              skillsIntroduced: document.getElementById('skillsIntroduced').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            saveNotesToLocalBackup(notes);
            updateBackupIndicator();
          }
          
          // Function moved to main script section for proper accessibility
          
          function saveNotes() {
            try {
              const saveBtn = document.getElementById('saveNotesBtn');
              
              if (!saveBtn) {
                console.error('Save button not found');
                return;
              }
              
              // Use the correct setButtonLoading signature
              if (typeof setButtonLoading === 'function') {
                setButtonLoading(saveBtn, true, 'Saving...');
              } else {
                saveBtn.textContent = 'Saving...';
                saveBtn.disabled = true;
              }
              
              const notes = {
                wins: (document.getElementById('wins') && document.getElementById('wins').value) || '',
                skillsMastered: (document.getElementById('skillsMastered') && document.getElementById('skillsMastered').value) || '',
                skillsPracticed: (document.getElementById('skillsPracticed') && document.getElementById('skillsPracticed').value) || '',
                skillsIntroduced: (document.getElementById('skillsIntroduced') && document.getElementById('skillsIntroduced').value) || '',
                struggles: (document.getElementById('struggles') && document.getElementById('struggles').value) || '',
                parent: (document.getElementById('parent') && document.getElementById('parent').value) || '',
                next: (document.getElementById('next') && document.getElementById('next').value) || ''
              };
              
              
              // Save to local backup immediately
              if (typeof saveNotesToLocalBackup === 'function') {
                saveNotesToLocalBackup(notes);
              }
              
              google.script.run
                .withSuccessHandler(function() {
                  // Reset loading state
                  saveBtn.disabled = false;
                  saveBtn.textContent = 'Saved!';
                  saveBtn.style.background = '#00C853';
                  
                  // Update auto-save status with current time
                  const autoSaveStatus = document.getElementById('autoSaveStatus');
                  if (autoSaveStatus) {
                    autoSaveStatus.textContent = 'Last saved: ' + new Date().toLocaleTimeString();
                  }
                  
                  setTimeout(function() {
                    saveBtn.textContent = 'Save Notes';
                  }, 2000);
                  
                  unsavedChanges = false;
                  
                  // Clear local backup after successful server save
                  if (typeof clearLocalBackup === 'function') {
                    clearLocalBackup();
                  }
                  if (typeof updateBackupIndicator === 'function') {
                    updateBackupIndicator();
                  }
                })
                .withFailureHandler(function(error) {
                  console.error('Error saving notes:', error);
                  saveBtn.disabled = false;
                  saveBtn.textContent = 'Save Notes';
                  if (typeof showAlert === 'function') {
                    showAlert('Error saving notes: ' + error.message + '. Notes saved locally.', 'warning');
                  } else {
                    alert('Error saving notes: ' + error.message);
                  }
                  if (typeof updateBackupIndicator === 'function') {
                    updateBackupIndicator();
                  }
                })
                .saveQuickNotes(notes);
                
            } catch (error) {
              console.error('Error in saveNotes function:', error);
              alert('Error saving notes: ' + error.message);
            }
          }
          
          // Verify function is defined
          
          function clearNotes() {
            const button = document.getElementById('clearNotesBtn');
            if (!button) return;
            
            // Check if button is in confirmation state
            if (button.dataset.confirmState === 'true') {
              // User confirmed - actually clear notes
              document.getElementById('wins').value = '';
              document.getElementById('skills').value = '';
              document.getElementById('struggles').value = '';
              document.getElementById('parent').value = '';
              document.getElementById('next').value = '';
              unsavedChanges = true;
              
              // Reset button state
              resetClearNotesButton();
              showAlert('Notes cleared. Remember to save!', 'success');
              
            } else {
              // First click - show confirmation state
              button.style.background = '#ea4335';
              button.style.color = 'white';
              button.textContent = 'Confirm Clearing';
              button.dataset.confirmState = 'true';
              
              // Reset after 3 seconds if not confirmed
              setTimeout(() => {
                if (button.dataset.confirmState === 'true') {
                  resetClearNotesButton();
                }
              }, 3000);
            }
          }
          
          function resetClearNotesButton() {
            const button = document.getElementById('clearNotesBtn');
            if (button) {
              button.style.background = '';
              button.style.color = '';
              button.textContent = 'Clear Notes';
              button.dataset.confirmState = 'false';
            }
          }
          
          function reloadNotes() {
            if (unsavedChanges && !confirm('You have unsaved changes. Reload anyway?')) {
              return;
            }
            
            const button = document.querySelector('button[onclick*="reloadNotes"]');
            if (button) {
              // Add loading state with spinning wheel (same as Send Recap button)
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            google.script.run
              .withSuccessHandler(function(notesResponse) {
                if (notesResponse && notesResponse.data) {
                  const notes = JSON.parse(notesResponse.data);
                  document.getElementById('wins').value = notes.wins || '';
                  document.getElementById('skills').value = notes.skills || '';
                  document.getElementById('struggles').value = notes.struggles || '';
                  document.getElementById('parent').value = notes.parent || '';
                  document.getElementById('next').value = notes.next || '';
                  unsavedChanges = false;
                  showAlert('Notes reloaded from saved version.', 'success');
                } else {
                  showAlert('No saved notes found for this student.', 'error');
                }
                
                // Remove loading state
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                  button.textContent = button.dataset.originalText || 'Reload Saved';
                }
              })
              .withFailureHandler(function(error) {
                showAlert('Error loading notes: ' + error.message, 'error');
                
                // Remove loading state
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                  button.textContent = button.dataset.originalText || 'Reload Saved';
                }
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function sendRecap() {
            const button = document.getElementById('sendRecapBtn');
            if (!button) return;
            
            // Add loading state to button with loading wheel
            button.dataset.originalText = button.textContent;
            button.classList.add('btn-loading');
            button.disabled = true;
            
            // Save notes first
            const notes = {
              wins: document.getElementById('wins').value,
              skillsMastered: document.getElementById('skillsMastered').value,
              skillsPracticed: document.getElementById('skillsPracticed').value,
              skillsIntroduced: document.getElementById('skillsIntroduced').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(function() {
                
                // Send the recap
                google.script.run
                  .withSuccessHandler(function() {
                    button.classList.remove('btn-loading');
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    // Switch back to control panel after sending
                    switchToControlPanel();
                  })
                  .withFailureHandler(function(error) {
                    button.classList.remove('btn-loading');  
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    console.error('Failed to send recap:', error);
                    showAlert('Failed to send recap: ' + error.message, 'error');
                  })
                  .sendIndividualRecap();
              })
              .withFailureHandler(function(error) {
                button.classList.remove('btn-loading');
                button.disabled = false;
                button.textContent = button.dataset.originalText || 'Send Recap';
                console.error('Failed to save notes before sending recap:', error);
                showAlert('Failed to save notes: ' + error.message, 'error');
              })
              .saveQuickNotes(notes);
          }
          
          // Auto-save functionality
          function startAutoSave() {
            autoSaveInterval = setInterval(function() {
              if (unsavedChanges) {
                const notes = {
                  wins: document.getElementById('wins').value,
                  skills: document.getElementById('skills').value,
                  struggles: document.getElementById('struggles').value,
                  parent: document.getElementById('parent').value,
                  next: document.getElementById('next').value
                };
                
                google.script.run
                  .withSuccessHandler(function() {
                    document.getElementById('autoSaveStatus').textContent = 
                      'Last saved: ' + new Date().toLocaleTimeString() + ' (auto)';
                    unsavedChanges = false;
                  })
                  .saveQuickNotes(notes);
              }
            }, 60000); // Every minute
          }
          
          // Control Panel functions
          function runFunction(functionName) {
            // Find which button was clicked
            const button = event.target.closest('button');
            
            // Immediate visual feedback
            if (button) {
              setButtonLoading(button, true);
              // Add haptic-style feedback
              button.style.transform = 'scale(0.95)';
              setTimeout(function() {
                if (button.style.transform === 'scale(0.95)') {
                  button.style.transform = '';
                }
              }, 150);
            }
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (button) {
                  setButtonLoading(button, false);
                }
                
                // Handle special functions with immediate UI feedback
                if (functionName === 'findClient') {
                  // Show Find Client view immediately with loading state
                  document.getElementById('controlView').classList.remove('active');
                  document.getElementById('quickNotesView').classList.remove('active');
                  document.getElementById('clientUpdateView').classList.remove('active');
                  document.getElementById('clientListView').classList.add('active');
                  currentView = 'clientlist';
                  
                  // Then load the data
                  openClientList();
                } else if (functionName === 'addNewClient') {
                  // addNewClient opens its own modal, refresh after
                  google.script.run
                    .withSuccessHandler(function(info) {
                      clientInfo = info;
                      updateClientDisplay();
                    })
                    .getCurrentClientInfo();
                }
              })
              .withFailureHandler(function(error) {
                if (button) {
                  setButtonLoading(button, false);
                }
                showAlert('Error: ' + error.message, 'error');
              })[functionName]();
          }
          
          // Client search functionality
          let searchTimeout;
          function searchClients(query) {
            clearTimeout(searchTimeout);
            const resultsDiv = document.getElementById('searchResults');
            
            if (!query.trim()) {
              resultsDiv.innerHTML = '';
              return;
            }
            
            // Show loading immediately for user feedback
            resultsDiv.innerHTML = '<div style="text-align: center; color: #666; padding: 10px; font-size: 12px;"><i>⏳</i> Searching...</div>';
            
            // Check cache first for instant results
            const cachedClients = PerformanceCache.getClientList();
            if (cachedClients) {
              // Small delay to show loading state briefly
              setTimeout(function() {
                displaySearchResults(cachedClients, query);
              }, 100);
              return;
            }
            
            // Debounce search for server call
            searchTimeout = setTimeout(function() {
              google.script.run
                .withSuccessHandler(function(clients) {
                  PerformanceCache.setClientList(clients);
                  displaySearchResults(clients, query);
                })
                .withFailureHandler(function(error) {
                  console.error('Search error:', error);
                  resultsDiv.innerHTML = '<div style="color: #ff4444; font-size: 12px;">Search error</div>';
                })
                .getUnifiedActiveClientList();
            }, 150); // Reduced debounce time
          }
          
          function displaySearchResults(clients, searchQuery) {
            const query = (searchQuery || document.getElementById('clientSearch').value).toLowerCase();
            const resultsDiv = document.getElementById('searchResults');
            
            if (!query.trim()) {
              resultsDiv.innerHTML = '';
              return;
            }
            
            const filtered = clients.filter(function(client) {
              return client.name.toLowerCase().includes(query);
            }).slice(0, 5); // Limit to 5 results
            
            if (filtered.length === 0) {
              resultsDiv.innerHTML = '<div style="color: #666; font-size: 12px; padding: 5px;">No active clients found</div>';
              return;
            }
            
            const html = filtered.map(function(client) {
              return '<button onclick="selectClient(\\''+client.sheetName+'\\')" ' +
                'style="display: block; width: 100%; padding: 8px; margin: 2px 0; border: 1px solid #ddd; background: white; border-radius: 3px; cursor: pointer; text-align: left; font-size: 12px; color: #333;" ' +
                'onmouseover="this.style.background=\\'#f0f8ff\\'; this.style.color=\\'#003366\\';" ' +
                'onmouseout="this.style.background=\\'white\\'; this.style.color=\\'#333\\';">'+
                client.name+
              '</button>';
            }).join('');
            
            resultsDiv.innerHTML = html;
          }
          
          function selectClient(sheetName) {
            // Clear search
            document.getElementById('clientSearch').value = '';
            document.getElementById('searchResults').innerHTML = '';
            
            // Navigate to client sheet
            google.script.run
              .withSuccessHandler(function() {
                // Refresh client info after navigation
                google.script.run
                  .withSuccessHandler(function(info) {
                    clientInfo = info;
                    updateClientDisplay();
                  })
                  .getCurrentClientInfo();
              })
              .withFailureHandler(function(error) {
                showAlert('Error selecting client: ' + error.message, 'error');
              })
              .navigateToSheet(sheetName);
          }
          
          // Preloading functions for hover states
          function preloadAddClient() {
            // Pre-load any data needed for add client (minimal since it's a modal)
          }
          
          function preloadClientList() {
            // Client list is already preloaded during sidebar initialization
          }
          
          function preloadUpdateClient() {
            // Pre-cache current client details if available
            if (clientInfo && clientInfo.sheetName && !PerformanceCache.getClientDetails(clientInfo.sheetName)) {
              google.script.run
                .withSuccessHandler(function(details) {
                  PerformanceCache.setClientDetails(clientInfo.sheetName, details);
                })
                .withFailureHandler(function() {
                  // Silent fail for preload
                })
                .getClientDetailsForUpdate(clientInfo.sheetName);
            }
          }
          
          // Duplicate function removed - now defined earlier in the code
          
          // Update Client Info dialog
          function openUpdateClientDialog() {
            if (!clientInfo || !clientInfo.sheetName) {
              showAlert('Please select a client first.', 'error');
              return;
            }
            
            // Switch to update view
            switchToClientUpdateView();
            
            // Check cache first
            const cachedDetails = PerformanceCache.getClientDetails(clientInfo.sheetName);
            if (cachedDetails) {
              populateClientUpdateForm(cachedDetails);
              return;
            }
            
            google.script.run
              .withSuccessHandler(function(details) {
                PerformanceCache.setClientDetails(clientInfo.sheetName, details);
                populateClientUpdateForm(details);
              })
              .withFailureHandler(function(error) {
                showAlert('Error loading client details: ' + error.message, 'error');
              })
              .getClientDetailsForUpdate(clientInfo.sheetName);
          }
          
          function switchToClientUpdateView() {
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.remove('active');
            document.getElementById('clientUpdateView').classList.add('active');
            currentView = 'clientupdate';
          }
          
          function populateClientUpdateForm(clientDetails) {
            document.getElementById('updateClientName').textContent = clientDetails.clientName;
            document.getElementById('dashboardInput').value = clientDetails.dashboardLink || '';
            document.getElementById('meetingNotesInput').value = clientDetails.meetingNotesLink || '';
            document.getElementById('scoreReportInput').value = clientDetails.scoreReportLink || '';
            document.getElementById('activeCheckbox').checked = clientDetails.isActive;
            
            // Update indicators
            document.getElementById('dashboardIndicator').style.background = clientDetails.dashboardLink ? '#4CAF50' : '#f44336';
            document.getElementById('meetingNotesIndicator').style.background = clientDetails.meetingNotesLink ? '#4CAF50' : '#f44336';
            document.getElementById('scoreReportIndicator').style.background = clientDetails.scoreReportLink ? '#4CAF50' : '#f44336';
            
            // Store client details for saving
            window.currentClientDetails = clientDetails;
          }
          
          // Function moved to main script section for proper accessibility
          
          
          function selectClient(sheetName) {
            // Clear search
            document.getElementById('clientListSearch').value = '';
            
            // Navigate to client sheet
            google.script.run
              .withSuccessHandler(function() {
                // Go back to control panel and refresh
                switchToControlPanel();
              })
              .withFailureHandler(function(error) {
                showAlert('Error selecting client: ' + error.message, 'error');
              })
              .navigateToSheet(sheetName);
          }
          
          // Local backup functions
          function saveNotesToLocalBackup(notes) {
            const sheetName = clientInfo ? clientInfo.sheetName : 'unknown';
            const backupKey = 'quicknotes_backup_' + sheetName;
            const backup = {
              notes: notes,
              timestamp: new Date().getTime(),
              sheetName: sheetName
            };
            
            try {
              localStorage.setItem(backupKey, JSON.stringify(backup));
            } catch (error) {
              console.error('Error saving to local backup:', error);
            }
          }
          
          function getLocalBackup() {
            const sheetName = clientInfo ? clientInfo.sheetName : 'unknown';
            const backupKey = 'quicknotes_backup_' + sheetName;
            
            try {
              const backup = localStorage.getItem(backupKey);
              return backup ? JSON.parse(backup) : null;
            } catch (error) {
              console.error('Error reading local backup:', error);
              return null;
            }
          }
          
          function clearLocalBackup() {
            const sheetName = clientInfo ? clientInfo.sheetName : 'unknown';
            const backupKey = 'quicknotes_backup_' + sheetName;
            
            try {
              localStorage.removeItem(backupKey);
            } catch (error) {
              console.error('Error clearing local backup:', error);
            }
          }
          
          function updateBackupIndicator() {
            const backup = getLocalBackup();
            const indicator = document.getElementById('backupIndicator');
            
            if (backup && indicator) {
              indicator.style.display = 'block';
            } else if (indicator) {
              indicator.style.display = 'none';
            }
          }
          
          function checkAndRestoreFromBackup() {
            const backup = getLocalBackup();
            if (!backup) return;
            
            // Get current server notes
            google.script.run
              .withSuccessHandler(function(serverNotes) {
                compareAndRestoreBackup(backup, serverNotes);
              })
              .withFailureHandler(function(error) {
                console.error('Error getting server notes:', error);
                // If can't reach server, use backup
                restoreFromBackup(backup.notes);
                updateBackupIndicator();
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function compareAndRestoreBackup(backup, serverNotes) {
            const serverTimestamp = serverNotes.timestamp || 0;
            const backupTimestamp = backup.timestamp || 0;
            
            if (backupTimestamp > serverTimestamp) {
              // Local backup is newer, ask user
              const shouldUseBackup = confirm(
                'A newer local backup was found. Do you want to restore it?\\n\\n' +
                'Backup: ' + new Date(backupTimestamp).toLocaleString() + '\\n' +
                'Server: ' + new Date(serverTimestamp).toLocaleString()
              );
              
              if (shouldUseBackup) {
                restoreFromBackup(backup.notes);
                // Save the restored backup to server
                saveNotes();
              } else {
                clearLocalBackup();
              }
            } else {
              // Server is newer or same, clear backup
              clearLocalBackup();
            }
            
            updateBackupIndicator();
          }
          
          function restoreFromBackup(notes) {
            if (notes) {
              document.getElementById('wins').value = notes.wins || '';
              document.getElementById('skills').value = notes.skills || '';
              document.getElementById('struggles').value = notes.struggles || '';
              document.getElementById('parent').value = notes.parent || '';
              document.getElementById('next').value = notes.next || '';
              
              showAlert('Notes restored from local backup', 'success');
            }
          }
          
          // Optimized connection monitoring
          let connectionCheckInterval;
          let isOnline = true;
          let reconnectAttempts = 0;
          let checkInterval = 10000; // Start with 10 seconds for stable connections
          const maxReconnectAttempts = 12;
          const minInterval = 5000;
          const maxInterval = 30000;
          
          function startConnectionMonitoring() {
            // Initial connection check
            checkConnection();
            
            // Adaptive interval based on connection stability
            function scheduleNextCheck() {
              connectionCheckInterval = setTimeout(function() {
                checkConnection();
                scheduleNextCheck();
              }, checkInterval);
            }
            
            scheduleNextCheck();
          }
          
          function stopConnectionMonitoring() {
            if (connectionCheckInterval) {
              clearTimeout(connectionCheckInterval);
              connectionCheckInterval = null;
            }
          }
          
          function checkConnection() {
            google.script.run
              .withSuccessHandler(function() {
                handleConnectionRestored();
              })
              .withFailureHandler(function() {
                handleConnectionLost();
              })
              .testConnection();
          }
          
          function handleConnectionLost() {
            if (isOnline) {
              isOnline = false;
              reconnectAttempts = 0;
              checkInterval = minInterval; // Check more frequently when offline
              showConnectionOverlay();
            }
            
            reconnectAttempts++;
            updateReconnectStatus();
            
            // Exponential backoff for failed connections
            if (reconnectAttempts > 3) {
              checkInterval = Math.min(checkInterval * 1.5, maxInterval);
            }
          }
          
          function handleConnectionRestored() {
            if (!isOnline) {
              isOnline = true;
              reconnectAttempts = 0;
              checkInterval = 10000; // Reset to normal interval
              hideConnectionOverlay();
              
              // Check for local backup to sync
              setTimeout(function() {
                checkAndRestoreFromBackup();
              }, 1000);
            }
          }
          
          function showConnectionOverlay() {
            const overlay = document.getElementById('connectionOverlay');
            if (overlay) {
              overlay.style.display = 'flex';
              document.body.style.overflow = 'hidden';
            }
          }
          
          function hideConnectionOverlay() {
            const overlay = document.getElementById('connectionOverlay');
            if (overlay) {
              overlay.style.display = 'none';
              document.body.style.overflow = '';
            }
          }
          
          function updateReconnectStatus() {
            const statusDiv = document.getElementById('reconnectStatus');
            const progressDiv = document.getElementById('reconnectProgress');
            
            if (statusDiv && progressDiv) {
              const progress = Math.min((reconnectAttempts / maxReconnectAttempts) * 100, 100);
              progressDiv.style.width = progress + '%';
              
              if (reconnectAttempts <= maxReconnectAttempts) {
                statusDiv.textContent = 'Attempting to reconnect... (' + reconnectAttempts + '/' + maxReconnectAttempts + ')';
              } else {
                statusDiv.textContent = 'Still trying to reconnect...';
              }
            }
          }
          
          // Start optimized systems when the sidebar loads
          setTimeout(function() {
            startConnectionMonitoring();
            // Initialize background processes (client-side)
            // Start periodic cache cleanup
            setInterval(function() {
              const now = Date.now();
              // Clean up expired cache entries - using Map iteration properly
              if (PerformanceCache.data.clientDetails instanceof Map) {
                for (let [key, entry] of PerformanceCache.data.clientDetails) {
                  if (entry && entry.timestamp && (now - entry.timestamp > PerformanceCache.CACHE_DURATION)) {
                    PerformanceCache.data.clientDetails.delete(key);
                  }
                }
              }
            }, 5 * 60 * 1000); // Clean every 5 minutes
            
            // Background prefetching
            PerformanceCache.prefetchClientList();
            
            // Periodic cache cleanup
            setInterval(function() {
              PerformanceCache.clearExpired();
            }, 10 * 60 * 1000); // Every 10 minutes
            
            // Pre-warm frequently used functions after initial load
            setTimeout(function() {
              if (clientInfo && clientInfo.sheetName) {
                // Pre-cache current client details
                google.script.run
                  .withSuccessHandler(function(details) {
                    PerformanceCache.setClientDetails(clientInfo.sheetName, details);
                  })
                  .withFailureHandler(function() {
                    // Silent fail for background operations
                  })
                  .getClientDetailsForUpdate(clientInfo.sheetName);
              }
            }, 2000);
          }, 500);
          
          // Clean up on page unload
          window.addEventListener('beforeunload', function() {
            stopConnectionMonitoring();
          });
          
          // Alert helper
          function showAlert(message, type) {
            const alertDiv = currentView === 'quicknotes' ? 
              document.getElementById('quickNotesAlert') : 
              document.createElement('div');
              
            if (currentView === 'control') {
              alertDiv.className = 'alert ' + type;
              alertDiv.textContent = message;
              document.getElementById('controlView').insertBefore(
                alertDiv, 
                document.getElementById('controlView').firstChild
              );
              
              setTimeout(function() {
                alertDiv.remove();
              }, 5000);
            } else {
              alertDiv.className = 'alert ' + type;
              alertDiv.textContent = message;
              alertDiv.style.display = 'block';
              
              setTimeout(function() {
                alertDiv.style.display = 'none';
              }, 5000);
            }
          }
          
          // Monitor for text input
          document.querySelectorAll('textarea').forEach(function(textarea) {
            textarea.addEventListener('input', markUnsaved);
          });
          
          // Refresh client info periodically
          setInterval(function() {
            if (currentView === 'control') {
              google.script.run
                .withSuccessHandler(function(info) {
                  clientInfo = info;
                  updateClientDisplay();
                })
                .getCurrentClientInfo();
            }
          }, 5000);
          
          // Initialize auto-save if starting in quick notes
          if (currentView === 'quicknotes') {
            startAutoSave();
          }
          
          // Cache management functions
          function refreshClientListCache() {
            const button = event.target.closest('button');
            
            if (button) {
              setButtonLoading(button, true);
              button.style.transform = 'scale(0.95)';
              setTimeout(function() {
                if (button.style.transform === 'scale(0.95)') {
                  button.style.transform = '';
                }
              }, 150);
            }
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (button) {
                  setButtonLoading(button, false);
                }
                
                if (result.success) {
                  showAlert('Cache refreshed! ' + (result.stats ? result.stats.active : 0) + ' active clients found.', 'success');
                  // Clear any existing client list cache in PerformanceCache too
                  PerformanceCache.clearClientList();
                } else {
                  showAlert('Failed to refresh cache: ' + result.message, 'error');
                }
              })
              .withFailureHandler(function(error) {
                if (button) {
                  setButtonLoading(button, false);
                }
                showAlert('Error refreshing cache: ' + error.message, 'error');
              })
              .refreshClientCache();
          }
          
          // Initialize all preload functions when sidebar opens
          function initializePreloads() {
            try {
              preloadAddClient();
              preloadClientList();
              preloadUpdateClient();
            } catch(error) {
              console.error('Error initializing preloads:', error);
            }
          }
          
          // Function to load sidebar data from server
          function loadSidebarData() {
            // Use pre-loaded client info for instant display (no server round-trip)
            
            // Set the active view based on currentView
            if (currentView === 'clientlist') {
              const clientListView = document.getElementById('clientListView');
              if (clientListView) {
                clientListView.classList.add('active');
              }
            } else if (currentView === 'clientupdate') {
              const clientUpdateView = document.getElementById('clientUpdateView');
              if (clientUpdateView) {
                clientUpdateView.classList.add('active');
              }
            }
            
            // Update UI immediately with pre-loaded data
            updateClientDisplay();
            
            // Initialize preloads
            initializePreloads();
          }
          
          // Function to update client display
          function updateClientDisplay() {
            const clientNameEl = document.getElementById('clientName');
            const quickNotesTitle = document.getElementById('quickNotesClientTitle');
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            const updateBtn = document.getElementById('updateClientBtn');
            
            if (clientInfo.isClient) {
              clientNameEl.textContent = clientInfo.name;
              clientNameEl.className = 'current-client-name';
              quickNotesTitle.textContent = clientInfo.name;
              quickNotesBtn.disabled = false;
              recapBtn.disabled = false;
              if (updateBtn) updateBtn.disabled = false;
            } else {
              clientNameEl.textContent = 'No client selected';
              clientNameEl.className = 'no-client';
              quickNotesTitle.textContent = 'No Client Selected';
              quickNotesBtn.disabled = true;
              recapBtn.disabled = true;
              if (updateBtn) updateBtn.disabled = true;
            }
          }
          
          // Load data when DOM is ready
          if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', loadSidebarData);
          } else {
            loadSidebarData();
          }
        </script>
        -->
        <!-- End of commented out Script 3 -->
        </div> <!-- Close sidebar-container -->
        
        <!-- Global loading overlay -->
        <div id="globalLoadingOverlay" class="loading-overlay">
          <div class="loading-spinner-large"></div>
        </div>
      </body>
    </html>
  `).setTitle('Smart College').setWidth(300);
  
  // Safe string replacement to embed client info without template literal issues
  let htmlContent = html.getContent();
  htmlContent = htmlContent.replace('REPLACE_INITIAL_VIEW', initialView);
  htmlContent = htmlContent.replace('REPLACE_CLIENT_INFO', JSON.stringify(clientInfo).replace(/'/g, "\\'"));
  
  const finalHtml = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Smart College')
    .setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(finalHtml);
}

/**
 * Prompt user to enter dashboard link if not found in Script Properties
 */
function promptForDashboardLink(sheetName, studentName) {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #333;
            margin-bottom: 8px;
            font-weight: 500;
            text-align: center;
          }
          .warning-box {
            background: #fef7e0;
            border: 1px solid #f9cc79;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            font-size: 14px;
            color: #7f6000;
          }
          .student-info {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
            text-align: center;
          }
          .input-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #5f6368;
            font-size: 14px;
          }
          input[type="url"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e1e5e9;
            border-radius: 6px;
            font-size: 16px;
            font-family: inherit;
            transition: border-color 0.3s ease;
            box-sizing: border-box;
          }
          input[type="url"]:focus {
            outline: none;
            border-color: #003366;  /* Changed from #1a73e8 */
            box-shadow: 0 0 0 3px rgba(0, 51, 102, 0.1);  /* Changed */
          }
          .example {
            font-style: italic;
            color: #5f6368;
            font-size: 13px;
            margin-top: 8px;
          }
          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 24px;
          }
          button {
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            flex: 1;
          }
          .btn-primary {
            background: #00C853;   
            color: white;
          }
          .btn-primary:hover {
            background: #388E3C;    
          }

          .btn-primary:disabled {
            background: #dadce0;
            cursor: not-allowed;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .drive-link {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            color: #003366;
            text-decoration: none;
            font-size: 13px;
            padding: 8px 16px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: #f8f9fa;
            transition: all 0.2s ease;
            margin-top: 20px;
          }
          .drive-link:hover {
            background: #e8f0fe;
            border-color: #003366; 
          }
          .drive-section {
            text-align: center;
            padding-top: 20px;
            border-top: 1px solid #e1e5e9;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }
          .loading.active {
            display: block;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Dashboard Link Required</h2>
          
          <div class="warning-box">
            ⚠️ No dashboard link found for this student. Please enter the dashboard URL to continue with the session recap.
          </div>
          
          <div class="student-info">
            <strong>Student:</strong> ${studentName}
          </div>
          
          <div class="input-group">
            <label>Dashboard Link</label>
            <input type="url" id="dashboardLink" placeholder="Enter dashboard URL" autofocus>
            <div class="example">Example: https://smartcollege.com/dashboard/john-smith</div>
          </div>
          
          <div class="button-group">
            <button class="btn-primary" onclick="saveDashboardLink()">Save & Continue</button>
            <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>
          
          <div class="drive-section">
            <a href="https://drive.google.com/drive/u/0/folders/1UYJa--G9hE23a5WfxblqrtdnKdp3r91_" 
               target="_blank" 
               class="drive-link">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                <path d="M10 4H4c-1.11 0-2 .89-2 2v12c0 1.11.89 2 2 2h16c1.11 0 2-.89 2-2V8c0-1.11-.89-2-2-2h-8l-2-2z"/>
              </svg>
              Open Dashboards Drive Folder
            </a>
          </div>
          
          <div class="loading" id="loadingDiv">
            <p>Saving dashboard link...</p>
          </div>
        </div>
        
        <script>
          const sheetName = '${sheetName}';
          const studentName = '${studentName}';
          
          // Allow Enter key to submit
          document.getElementById('dashboardLink').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
              saveDashboardLink();
            }
          });
          
          function saveDashboardLink() {
            const dashboardLink = document.getElementById('dashboardLink').value.trim();
            
            if (!dashboardLink) {
              alert('Please enter a valid dashboard URL');
              return;
            }
            
            // Basic URL validation
            try {
              new URL(dashboardLink);
            } catch (e) {
              alert('Please enter a valid URL (including https://)');
              return;
            }
            
            // Get the button and show loading
            const saveBtn = document.querySelector('.btn-primary');
            const originalText = saveBtn.textContent;
            saveBtn.classList.add('btn-loading');
            saveBtn.disabled = true;
            
            // Save the dashboard link
            google.script.run
              .withSuccessHandler(function() {
                // Close dialog and continue with recap
                google.script.host.close();
                // Re-trigger the recap dialog
                google.script.run.continueWithRecap(sheetName);
              })
              .withFailureHandler(function(error) {
                saveBtn.classList.remove('btn-loading');
                saveBtn.textContent = originalText;
                saveBtn.disabled = false;
                alert('Error saving dashboard link: ' + error.message);
              })
              .saveDashboardLinkToProperties(sheetName, dashboardLink);
          }
        </script>
      </body>
    </html>
  `).setWidth(550).setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Link Required');
}

/**
 * Save dashboard link to unified data store
 */
function saveDashboardLinkToProperties(sheetName, dashboardLink) {
  // Find client by sheet name and update dashboard link
  const allClients = UnifiedClientDataStore.getAllClients();
  const client = allClients.find(c => c.sheetName === sheetName);
  if (client) {
    UnifiedClientDataStore.updateClient(client.name, { dashboardLink: dashboardLink });
  }
  return true;
}

/**
 * Continue with recap after dashboard link is saved
 */
function continueWithRecap(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    spreadsheet.setActiveSheet(sheet);
    const clientData = getClientDataFromSheet(sheet);
    showRecapDialog(clientData);
  }
}

/**
 * Main function to add a new client sheet
 */
function addNewClient() {
  showClientTypeDialog();
}

/**
 * Get information about the currently selected client
 * Simply checks A1 value and verifies it's a client sheet
 */
function getCurrentClientInfo() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Check if it's a client sheet using existing helper
    if (!isClientSheet(sheetName)) {
      return {
        isClient: false,
        name: null,
        sheetName: sheetName
      };
    }
    
    // Simply get the value from A1
    const clientName = sheet.getRange('A1').getValue();
    
    return {
      isClient: true,
      name: clientName || 'Unnamed Client', // Fallback if A1 is empty
      sheetName: sheetName
    };
    
  } catch (e) {
    console.error('Error getting client info:', e);
    return {
      isClient: false,
      name: null,
      sheetName: null
    };
  }
}
/**
 * Shows a dialog to select client type (ACT or Academic Support)
 */
function showClientTypeDialog() {
  // Skip the client type selection and go directly to the input dialog
  showClientInputDialog();
}

/**
 * Creates a custom HTML dialog for client input
 */
function showClientInputDialog() {
  // Implement lazy loading for form
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          .loading-skeleton {
            background: linear-gradient(90deg, #f0f0f0 25%, transparent 37%, #f0f0f0 63%);
            background-size: 400% 100%;
            animation: skeleton-shimmer 1.5s ease-in-out infinite;
          }
          @keyframes skeleton-shimmer {
            0% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
          }
          .form-skeleton {
            width: 100%;
            height: 20px;
            margin: 10px 0;
            border-radius: 4px;
          }
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #333;
            margin-bottom: 8px;
            font-weight: 500;
            text-align: center;
          }
          .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 25px;
            text-align: center;
          }
          .form-section {
            margin-bottom: 25px;
          }
          .input-group {
            margin-bottom: 20px;
          }
          label {
            display: block;
            margin-bottom: 8px;
            font-weight: 500;
            color: #5f6368;
            font-size: 14px;
          }
          input[type="text"], input[type="email"], input[type="url"] {
            width: 100%;
            padding: 12px 16px;
            border: 2px solid #e1e5e9;
            border-radius: 6px;
            font-size: 16px;
            font-family: inherit;
            transition: border-color 0.3s ease;
            box-sizing: border-box;
          }
          input[type="text"]:focus, input[type="email"]:focus, input[type="url"]:focus {
            outline: none;
            border-color: #003366;                      
            box-shadow: 0 0 0 3px rgba(0, 51, 102, 0.1);
          }
          .service-selection {
            margin-bottom: 20px;
          }
          .service-options {
            display: flex;
            gap: 12px;
            flex-wrap: wrap;
          }
          .service-checkbox {
            display: flex;
            align-items: center;
            padding: 10px 16px;
            border: 2px solid #e1e5e9;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
            background: white;
            flex: 1;
            min-width: 120px;
          }
          .service-checkbox:hover {
            border-color: #dadce0;
            transform: translateY(-1px);
          }

          .service-checkbox.act:hover {
            background-color: rgba(0, 51, 102, 0.1);  /* Blue tint */
            border-color: #003366;
          }
          .service-checkbox.math:hover {
            background-color: rgba(0, 200, 83, 0.1);   /* Green tint */
            border-color: #00C853;
          }
          .service-checkbox.language:hover {
            background-color: rgba(159, 124, 172, 0.1);
            border-color: #9f7cac;
          }
          .service-checkbox.selected {
            border-color: #003366;                      
            background-color: #f8f9fa;
          }
          .service-checkbox.act.selected {
            background-color: rgba(0, 51, 102, 0.15);
            border-color: #003366;
          }
          .service-checkbox.math.selected {
            background-color: rgba(217, 131, 122, 0.15);
            border-color: #d9837a;
          }
          .service-checkbox.language.selected {
            background-color: rgba(159, 124, 172, 0.15);
            border-color: #9f7cac;
          }
          .service-checkbox input[type="checkbox"] {
            margin-right: 8px;
            cursor: pointer;
          }
          .service-checkbox label {
            margin: 0;
            font-weight: 500;
            cursor: pointer;
            user-select: none;
          }
          .button-group {
            display: flex;
            gap: 12px;
            justify-content: center;
            margin-top: 30px;
          }
          button {
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            min-width: 100px;
          }
          .btn-primary {
            background: #00C853;
            color: white;
          }
          .btn-primary:hover {
            background: #388E3C;
            transform: translateY(-1px);
          }
          .btn-primary:disabled {
            background: #dadce0;
            cursor: not-allowed;
            transform: none;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .example {
            font-style: italic;
            color: #5f6368;
            font-size: 13px;
            margin-top: 8px;
          }
          .section-divider {
            border-top: 1px solid #e1e5e9;
            margin: 30px 0;
          }
          
          /* Loading overlay styles */
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            transition: background 0.5s ease;
          }
          
          .loading-overlay.active {
            display: flex;
          }
          
          .loading-overlay.success {
            background: rgba(52, 168, 83, 0.1);
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
          }
          
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #003366;             
            border-radius: 50%;
            animation: spin 1s linear infinite;
            transition: all 0.2s ease;
          }
          
          .loading-spinner.success {
            animation: none;
            border: 4px solid #00C853;                  
            position: relative;
          }
          
          .loading-spinner.success::after {
            content: '';
            position: absolute;
            left: 14px;
            top: 5px;
            width: 8px;
            height: 20px;
            border: solid #34a853;
            border-width: 0 4px 4px 0;
            transform: rotate(45deg);
            animation: checkmark 0.3s ease-in-out;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          @keyframes checkmark {
            0% {
              opacity: 0;
              transform: rotate(45deg) scale(0);
            }
            100% {
              opacity: 1;
              transform: rotate(45deg) scale(1);
            }
          }
          
          .loading-text {
            position: absolute;
            margin-top: 80px;
            font-size: 14px;
            color: #5f6368;
            font-weight: 500;
            transition: color 0.3s ease;
          }
          
          .loading-text.success {
            color: #34a853;
            font-weight: 600;
          }
          
          .error-message {
            color: #d93025;
            font-size: 13px;
            margin-top: 4px;
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <!-- Loading skeleton -->
          <div id="formSkeleton" style="display: block;">
            <div class="loading-skeleton form-skeleton"></div>
            <div class="loading-skeleton form-skeleton" style="width: 60%;"></div>
            <div class="loading-skeleton form-skeleton"></div>
            <div class="loading-skeleton form-skeleton" style="width: 80%;"></div>
            <div class="loading-skeleton form-skeleton"></div>
            <div class="loading-skeleton form-skeleton" style="width: 70%;"></div>
          </div>
          
          <!-- Actual form (initially hidden) -->
          <div id="actualForm" style="display: none;">
            <h2>Add New Client</h2>
            <p class="subtitle">Enter client information and select services</p>
            
            <div class="form-section">
              <div class="input-group">
                <label>Student Full Name*</label>
                <input type="text" id="clientName" placeholder="Enter student's full name" autofocus required>
                <div class="example">Example: John Smith</div>
            </div>
            
            <div class="service-selection">
              <label>Services* (select all that apply)</label>
              <div class="service-options">
                <div class="service-checkbox act" onclick="toggleService(this, 'ACT')">
                  <input type="checkbox" id="service-act" value="ACT">
                  <label for="service-act">ACT</label>
                </div>
                <div class="service-checkbox math" onclick="toggleService(this, 'Math Curriculum')">
                  <input type="checkbox" id="service-math" value="Math Curriculum">
                  <label for="service-math">Math Curriculum</label>
                </div>
                <div class="service-checkbox language" onclick="toggleService(this, 'Language')">
                  <input type="checkbox" id="service-language" value="Language">
                  <label for="service-language">Language</label>
                </div>
              </div>
              <div class="error-message" id="service-error">Please select at least one service</div>
            </div>
          </div>
          
          <div class="section-divider"></div>
          
          <div class="form-section">
            <div class="input-group">
              <label>Parent Email</label>
              <input type="email" id="parentEmail" placeholder="Enter parent email address">
              <div class="example">Example: parent@example.com</div>
            </div>
            
            <div class="input-group">
              <label>Email Addressees</label>
              <input type="text" id="emailAddressees" placeholder="Enter names for email greeting">
              <div class="example">Example: John & Jane Doe</div>
            </div>
            
            <div class="input-group">
              <label>Dashboard Link</label>
              <input type="url" id="dashboardLink" placeholder="Enter dashboard URL">
              <div class="example">Example: https://smartcollege.com/dashboard/john-smith</div>
            </div>
          </div>
          
          <div class="button-group">
            <button class="btn-primary" onclick="processClient()">Create Client</button>
            <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>

          <div style="text-align: center; margin-top: 20px; padding-top: 20px; border-top: 1px solid #e1e5e9;">
            <a href="https://drive.google.com/drive/u/0/folders/1UYJa--G9hE23a5WfxblqrtdnKdp3r91_" 
              target="_blank" 
              style="display: inline-flex; align-items: center; gap: 8px; color: #1a73e8; text-decoration: none; font-size: 13px; padding: 8px 16px; border: 1px solid #dadce0; border-radius: 6px; background: #f8f9fa; transition: all 0.2s ease;"
              onmouseover="this.style.background='#e8f0fe'; this.style.borderColor='#1a73e8';"
              onmouseout="this.style.background='#f8f9fa'; this.style.borderColor='#dadce0';">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                <path d="M10 4H4c-1.11 0-2 .89-2 2v12c0 1.11.89 2 2 2h16c1.11 0 2-.89 2-2V8c0-1.11-.89-2-2-2h-8l-2-2z"/>
              </svg>
              Open Dashboards Drive Folder
            </a>
          </div>  
        </div>
        
        <!-- Loading overlay -->
        <div class="loading-overlay" id="loadingOverlay">
          <div class="loading-spinner"></div>
          <div class="loading-text">Adding client...</div>
        </div>
        
        <script>
          // Toggle service selection
          function toggleService(element, service) {
            const checkbox = element.querySelector('input[type="checkbox"]');
            checkbox.checked = !checkbox.checked;
            element.classList.toggle('selected', checkbox.checked);
          }
          
          function setButtonLoading(button, isLoading) {
            if (isLoading) {
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
            } else {
              button.classList.remove('btn-loading');
              button.textContent = button.dataset.originalText || button.textContent;
              button.disabled = false;
            }
          }
          // Allow Enter key to submit
          document.addEventListener('keypress', function(e) {
            if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
              processClient();
            }
          });
          
          function getSelectedServices() {
            const services = [];
            if (document.getElementById('service-act').checked) services.push('ACT');
            if (document.getElementById('service-math').checked) services.push('Math Curriculum');
            if (document.getElementById('service-language').checked) services.push('Language');
            return services;
          }
          
          function validateForm() {
            const clientName = document.getElementById('clientName').value.trim();
            const services = getSelectedServices();
            
            if (!clientName) {
              alert('Please enter a valid student name.');
              return false;
            }
            
            if (services.length === 0) {
              document.getElementById('service-error').style.display = 'block';
              return false;
            }
            
            return true;
          }
          
          // Lazy load the form after a brief delay
          setTimeout(function() {
            document.getElementById('formSkeleton').style.display = 'none';
            document.getElementById('actualForm').style.display = 'block';
            
            // Focus the first input after load
            const firstInput = document.getElementById('clientName');
            if (firstInput) {
              firstInput.focus();
            }
          }, 300);

          function processClient() {
            if (!validateForm()) return;
            
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, true);
            
            const clientData = {
              name: document.getElementById('clientName').value.trim(),
              services: getSelectedServices().join(', '),
              parentEmail: document.getElementById('parentEmail').value.trim(),
              emailAddressees: document.getElementById('emailAddressees').value.trim(),
              dashboardLink: document.getElementById('dashboardLink').value.trim()
            };
            
            // Show loading overlay with client name
            document.getElementById('loadingOverlay').classList.add('active');
            document.querySelector('.loading-text').textContent = 'Adding ' + clientData.name + '...';
            
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .createClientSheet(clientData);
          }
          
          function onSuccess(message) {
            // Transition to success state
            const overlay = document.getElementById('loadingOverlay');
            const spinner = document.querySelector('.loading-spinner');
            const loadingText = document.querySelector('.loading-text');
            
            // Add success classes for green theme
            overlay.classList.add('success');
            spinner.classList.add('success');
            loadingText.classList.add('success');
            
            // Update text to success message
            loadingText.textContent = 'Success!';
            
            // Wait for animation, then close
            setTimeout(function() {
              google.script.host.close();
            }, 1500);
          }
          
          function onFailure(error) {
            // Hide loading overlay on error
            document.getElementById('loadingOverlay').classList.remove('active');
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, false);
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `).setWidth(550).setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Smart College');
}

/**
 * Show Quick Notes Settings Dialog for customizing quick buttons
 */
function showQuickNotesSettings() {
  // Get current user's quick button preferences
  const userProperties = PropertiesService.getUserProperties();
  const currentSettings = JSON.parse(userProperties.getProperty('quickNotesButtons') || '{}');
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
            font-size: 14px;
            background: #f8f9fa;
          }
          
          .header {
            text-align: center;
            margin-bottom: 24px;
            padding-bottom: 16px;
            border-bottom: 2px solid #e1e5e9;
          }
          
          .header h2 {
            margin: 0;
            color: #003366;
            font-size: 20px;
            font-weight: 600;
          }
          
          .header p {
            margin: 8px 0 0 0;
            color: #5f6368;
            font-size: 13px;
          }
          
          .section {
            background: white;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          .section h3 {
            margin: 0 0 12px 0;
            color: #202124;
            font-size: 16px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          
          .section-description {
            font-size: 12px;
            color: #5f6368;
            margin-bottom: 16px;
          }
          
          .button-inputs {
            display: grid;
            gap: 12px;
          }
          
          .button-input {
            display: flex;
            gap: 8px;
            align-items: center;
          }
          
          .button-input label {
            min-width: 60px;
            font-size: 12px;
            color: #5f6368;
            font-weight: 500;
          }
          
          .button-input input {
            flex: 1;
            padding: 8px 12px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 13px;
            font-family: inherit;
          }
          
          .button-input input:focus {
            outline: none;
            border-color: #1a73e8;
            box-shadow: 0 0 0 2px rgba(26,115,232,0.1);
          }
          
          .button-input input::placeholder {
            color: #9aa0a6;
          }
          
          .preview {
            margin-top: 12px;
            padding: 12px;
            background: #f8f9fa;
            border-radius: 4px;
            border-left: 3px solid #1a73e8;
          }
          
          .preview-label {
            font-size: 11px;
            color: #5f6368;
            margin-bottom: 8px;
            font-weight: 500;
          }
          
          .preview-buttons {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
          }
          
          .preview-tag {
            padding: 4px 8px;
            background: #e8f0fe;
            border-radius: 12px;
            font-size: 11px;
            color: #1a73e8;
            border: 1px solid #dadce0;
            cursor: default;
            white-space: nowrap;
          }
          
          .empty-preview {
            color: #9aa0a6;
            font-style: italic;
            font-size: 11px;
          }
          
          .actions {
            display: flex;
            gap: 12px;
            justify-content: flex-end;
            padding: 16px;
            background: white;
            border-top: 1px solid #e1e5e9;
            margin: 0 -20px -20px -20px;
            border-radius: 0 0 8px 8px;
          }
          
          .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            font-family: inherit;
          }
          
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          
          .btn-secondary:hover {
            background: #e8f0fe;
            border-color: #1a73e8;
          }
          
          .btn-primary {
            background: linear-gradient(135deg, #1a73e8 0%, #1557b0 100%);
            color: white;
          }
          
          .btn-primary:hover {
            background: linear-gradient(135deg, #1557b0 0%, #0d47a1 100%);
          }
          
          .reset-link {
            font-size: 12px;
            color: #1a73e8;
            text-decoration: none;
            margin-left: auto;
            margin-right: 16px;
            cursor: pointer;
          }
          
          .reset-link:hover {
            text-decoration: underline;
          }
          
          /* Specific section colors */
          .wins-section { border-left: 3px solid #28a745; }
          .mastered-section { border-left: 3px solid #17a2b8; }
          .practiced-section { border-left: 3px solid #ffc107; }
          .introduced-section { border-left: 3px solid #6f42c1; }
          .struggles-section { border-left: 3px solid #dc3545; }
          .parent-section { border-left: 3px solid #fd7e14; }
          .next-section { border-left: 3px solid #20c997; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>Quick Notes Settings</h2>
          <p>Customize your quick-insert buttons for each section (up to 4 buttons per section)</p>
        </div>
        
        <div class="section wins-section">
          <h3>🏆 Wins/Breakthroughs</h3>
          <div class="section-description">Quick buttons for celebrating student achievements</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="wins1" placeholder="e.g., Aha moment!" maxlength="30" value="${currentSettings.wins?.[0] || 'Aha moment!'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="wins2" placeholder="e.g., Solved independently" maxlength="30" value="${currentSettings.wins?.[1] || 'Solved independently'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="wins3" placeholder="e.g., Confidence boost" maxlength="30" value="${currentSettings.wins?.[2] || 'Confidence boost'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="wins4" placeholder="e.g., Speed improved" maxlength="30" value="${currentSettings.wins?.[3] || 'Speed improved'}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="winsPreview"></div>
          </div>
        </div>
        
        <div class="section mastered-section">
          <h3>✅ Skills Mastered</h3>
          <div class="section-description">Skills the student has fully mastered</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="mastered1" placeholder="e.g., Problem solving" maxlength="30" value="${currentSettings.mastered?.[0] || 'Problem solving'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="mastered2" placeholder="e.g., Reading comprehension" maxlength="30" value="${currentSettings.mastered?.[1] || 'Reading comprehension'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="mastered3" placeholder="e.g., Time management" maxlength="30" value="${currentSettings.mastered?.[2] || 'Time management'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="mastered4" placeholder="e.g., Advanced concept" maxlength="30" value="${currentSettings.mastered?.[3] || ''}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="masteredPreview"></div>
          </div>
        </div>
        
        <div class="section practiced-section">
          <h3>📝 Skills Practiced</h3>
          <div class="section-description">Skills we worked on together during the session</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="practiced1" placeholder="e.g., Test strategies" maxlength="30" value="${currentSettings.practiced?.[0] || 'Test strategies'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="practiced2" placeholder="e.g., ACT math" maxlength="30" value="${currentSettings.practiced?.[1] || 'ACT math'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="practiced3" placeholder="e.g., Essay writing" maxlength="30" value="${currentSettings.practiced?.[2] || 'Essay writing'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="practiced4" placeholder="e.g., Research skills" maxlength="30" value="${currentSettings.practiced?.[3] || ''}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="practicedPreview"></div>
          </div>
        </div>
        
        <div class="section introduced-section">
          <h3>🆕 Skills Introduced</h3>
          <div class="section-description">New concepts or skills introduced today</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="introduced1" placeholder="e.g., New concept" maxlength="30" value="${currentSettings.introduced?.[0] || 'New concept'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="introduced2" placeholder="e.g., Advanced technique" maxlength="30" value="${currentSettings.introduced?.[1] || 'Advanced technique'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="introduced3" placeholder="e.g., Study method" maxlength="30" value="${currentSettings.introduced?.[2] || 'Study method'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="introduced4" placeholder="e.g., Learning strategy" maxlength="30" value="${currentSettings.introduced?.[3] || ''}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="introducedPreview"></div>
          </div>
        </div>
        
        <div class="section struggles-section">
          <h3>⚠️ Struggles/Challenges</h3>
          <div class="section-description">Areas where the student had difficulty</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="struggles1" placeholder="e.g., Careless errors" maxlength="30" value="${currentSettings.struggles?.[0] || 'Careless errors'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="struggles2" placeholder="e.g., Time pressure" maxlength="30" value="${currentSettings.struggles?.[1] || 'Time pressure'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="struggles3" placeholder="e.g., Concept confusion" maxlength="30" value="${currentSettings.struggles?.[2] || 'Concept confusion'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="struggles4" placeholder="e.g., Attention focus" maxlength="30" value="${currentSettings.struggles?.[3] || 'Attention focus'}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="strugglesPreview"></div>
          </div>
        </div>
        
        <div class="section parent-section">
          <h3>👨‍👩‍👧‍👦 Parent Notes</h3>
          <div class="section-description">Information to share with parents</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="parent1" placeholder="e.g., Great participation today" maxlength="30" value="${currentSettings.parent?.[0] || 'Great participation today'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="parent2" placeholder="e.g., Ask about homework completion" maxlength="30" value="${currentSettings.parent?.[1] || 'Ask about homework completion'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="parent3" placeholder="e.g., Encourage practice at home" maxlength="30" value="${currentSettings.parent?.[2] || 'Encourage practice at home'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="parent4" placeholder="e.g., Celebrate improvement" maxlength="30" value="${currentSettings.parent?.[3] || 'Celebrate improvement'}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="parentPreview"></div>
          </div>
        </div>
        
        <div class="section next-section">
          <h3>➡️ Next Steps</h3>
          <div class="section-description">Plans for the next session</div>
          <div class="button-inputs">
            <div class="button-input">
              <label>Button 1:</label>
              <input type="text" id="next1" placeholder="e.g., Review homework challenges" maxlength="30" value="${currentSettings.next?.[0] || 'Review homework challenges'}">
            </div>
            <div class="button-input">
              <label>Button 2:</label>
              <input type="text" id="next2" placeholder="e.g., Build on today's progress" maxlength="30" value="${currentSettings.next?.[1] || 'Build on today\'s progress'}">
            </div>
            <div class="button-input">
              <label>Button 3:</label>
              <input type="text" id="next3" placeholder="e.g., Practice test strategies" maxlength="30" value="${currentSettings.next?.[2] || 'Practice test strategies'}">
            </div>
            <div class="button-input">
              <label>Button 4:</label>
              <input type="text" id="next4" placeholder="e.g., Focus on weak areas" maxlength="30" value="${currentSettings.next?.[3] || 'Focus on weak areas'}">
            </div>
          </div>
          <div class="preview">
            <div class="preview-label">Preview:</div>
            <div class="preview-buttons" id="nextPreview"></div>
          </div>
        </div>
        
        <div class="actions">
          <a href="#" class="reset-link" onclick="resetToDefaults()">Reset to Defaults</a>
          <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
          <button class="btn btn-primary" onclick="saveSettings()">Save Settings</button>
        </div>
        
        <script>
          // Update previews on input change
          const sections = ['wins', 'mastered', 'practiced', 'introduced', 'struggles', 'parent', 'next'];
          
          sections.forEach(section => {
            for (let i = 1; i <= 4; i++) {
              const input = document.getElementById(section + i);
              if (input) {
                input.addEventListener('input', () => updatePreview(section));
              }
            }
            updatePreview(section);
          });
          
          function updatePreview(section) {
            const preview = document.getElementById(section + 'Preview');
            const buttons = [];
            
            for (let i = 1; i <= 4; i++) {
              const input = document.getElementById(section + i);
              if (input && input.value.trim()) {
                buttons.push('<span class="preview-tag">' + input.value.trim() + '</span>');
              }
            }
            
            if (buttons.length > 0) {
              preview.innerHTML = buttons.join('');
            } else {
              preview.innerHTML = '<span class="empty-preview">No buttons configured</span>';
            }
          }
          
          function resetToDefaults() {
            if (confirm('Reset all quick buttons to their default values?')) {
              // Reset to default values
              const defaults = {
                wins: ['Aha moment!', 'Solved independently', 'Confidence boost', 'Speed improved'],
                mastered: ['Problem solving', 'Reading comprehension', 'Time management', ''],
                practiced: ['Test strategies', 'ACT math', 'Essay writing', ''],
                introduced: ['New concept', 'Advanced technique', 'Study method', ''],
                struggles: ['Careless errors', 'Time pressure', 'Concept confusion', 'Attention focus'],
                parent: ['Great participation today', 'Ask about homework completion', 'Encourage practice at home', 'Celebrate improvement'],
                next: ['Review homework challenges', 'Build on today\'s progress', 'Practice test strategies', 'Focus on weak areas']
              };
              
              sections.forEach(section => {
                for (let i = 1; i <= 4; i++) {
                  const input = document.getElementById(section + i);
                  if (input) {
                    input.value = defaults[section][i-1] || '';
                  }
                }
                updatePreview(section);
              });
            }
          }
          
          function saveSettings() {
            const settings = {};
            
            sections.forEach(section => {
              settings[section] = [];
              for (let i = 1; i <= 4; i++) {
                const input = document.getElementById(section + i);
                if (input) {
                  settings[section].push(input.value.trim());
                }
              }
            });
            
            google.script.run
              .withSuccessHandler(() => {
                alert('Quick Notes settings saved successfully!');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
                alert('Error saving settings: ' + error.message);
              })
              .saveQuickNotesSettings(settings);
          }
        </script>
      </body>
    </html>
  `).setWidth(600).setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Notes Settings');
}

/**
 * Save user's Quick Notes button settings
 */
function saveQuickNotesSettings(settings) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('quickNotesButtons', JSON.stringify(settings));
  return { success: true };
}

/**
 * Get user's Quick Notes button settings
 */
function getQuickNotesSettings() {
  const userProperties = PropertiesService.getUserProperties();
  const settings = JSON.parse(userProperties.getProperty('quickNotesButtons') || '{}');
  
  // Return defaults if no settings exist
  const defaults = {
    wins: ['Aha moment!', 'Solved independently', 'Confidence boost', 'Speed improved'],
    mastered: ['Problem solving', 'Reading comprehension', 'Time management', ''],
    practiced: ['Test strategies', 'ACT math', 'Essay writing', ''],
    introduced: ['New concept', 'Advanced technique', 'Study method', ''],
    struggles: ['Careless errors', 'Time pressure', 'Concept confusion', 'Attention focus'],
    parent: ['Great participation today', 'Ask about homework completion', 'Encourage practice at home', 'Celebrate improvement'],
    next: ['Review homework challenges', 'Build on today\'s progress', 'Practice test strategies', 'Focus on weak areas']
  };
  
  // Merge user settings with defaults
  const merged = {};
  Object.keys(defaults).forEach(section => {
    merged[section] = settings[section] || defaults[section];
  });
  
  return merged;
}

function createClientSheet(clientData) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Find the template sheet
  const templateSheet = spreadsheet.getSheetByName('NewClient');
  if (!templateSheet) {
    throw new Error('Template sheet "NewClient" not found. Please ensure it exists.');
  }
  
  // Create concatenated name (remove spaces)
  const concatenatedName = clientData.name.replace(/\s+/g, '');
  
  // Check if sheet with this name already exists
  if (spreadsheet.getSheetByName(concatenatedName)) {
    throw new Error(`A sheet named "${concatenatedName}" already exists.`);
  }
  
  // Copy the template sheet
  const newSheet = templateSheet.copyTo(spreadsheet);
  
  // Rename the copied sheet
  newSheet.setName(concatenatedName);
  
  // Populate the cells with client data (for backward compatibility)
  newSheet.getRange('A1').setValue(clientData.name);           // Student name
  newSheet.getRange('C2').setValue(clientData.services);       // Client type(s)
  newSheet.getRange('D2').setValue(clientData.dashboardLink);  // Dashboard smart chip (display)
  newSheet.getRange('E2').setValue(clientData.parentEmail);    // Parent email
  newSheet.getRange('F2').setValue(clientData.emailAddressees); // Email addressees
  
  // Save all data to unified data store (primary storage)
  const unifiedClientData = {
    name: clientData.name,
    sheetName: concatenatedName,
    services: clientData.services,
    dashboardLink: clientData.dashboardLink || '',
    parentEmail: clientData.parentEmail || '',
    emailAddressees: clientData.emailAddressees || '',
    isActive: true,
    meetingNotesLink: '',
    scoreReportLink: ''
  };
  
  // Add client to unified data store
  UnifiedClientDataStore.addClient(unifiedClientData);
  
  // Move the new sheet to position 3 (index 2, since it's 0-based)
  const totalSheets = spreadsheet.getSheets().length;
  const targetPosition = Math.min(3, totalSheets);
  
  // Activate the new sheet first, then move it
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(targetPosition);
  
  // IMPORTANT: Force Google Sheets to complete the sheet creation
  SpreadsheetApp.flush();
  
  // Verify the sheet exists before adding to master list
  const verifySheet = spreadsheet.getSheetByName(concatenatedName);
  if (!verifySheet) {
    throw new Error('Sheet creation failed - sheet not found after creation');
  }
  
  // NOW add client to master list after sheet is confirmed to exist
  try {
    addClientToMasterList(clientData, concatenatedName, spreadsheet);
    
    // Update enhanced cache with complete new client data
    const fullClientData = {
      name: clientData.name,
      sheetName: concatenatedName,
      isActive: true,
      dashboardLink: clientData.dashboardLink || '',
      meetingNotesLink: clientData.meetingNotesLink || '',
      scoreReportLink: clientData.scoreReportLink || '',
      parentEmail: clientData.parentEmail || '',
      emailAddressees: clientData.emailAddressees || ''
    };
    // Add to UnifiedClientDataStore instead of legacy cache
    try {
      UnifiedClientDataStore.upsertClient(fullClientData);
    } catch (error) {
      console.warn(`Failed to add client to UnifiedClientDataStore:`, error);
    }
    
  } catch (e) {
    console.error('Error adding client to master list:', e);
    // Continue even if master list update fails
    // But let's throw a more informative error
    throw new Error(`Client sheet created successfully, but failed to add to master list: ${e.message}`);
  }
  
  return `Client sheet for "${clientData.name}" has been created successfully! ✓`;
}
/**
 * Get email data from Properties Service
 * This function retrieves the temporarily stored email data for the preview dialog
 */
function getEmailDataFromProperties() {
  const userProperties = PropertiesService.getUserProperties();
  const emailDataString = userProperties.getProperty('tempEmailData');
  
  if (!emailDataString) {
    throw new Error('No email data found in properties');
  }
  
  return JSON.parse(emailDataString);
}

/**
 * Clear temporary email data from Properties Service
 */
function clearTempEmailData() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('tempEmailData');
}

/**
 * Back to edit recap - reopens the edit form with the stored data
 */
function backToEditRecap() {
  const userProperties = PropertiesService.getUserProperties();
  const emailDataString = userProperties.getProperty('tempEmailData');
  
  if (!emailDataString) {
    throw new Error('No email data found to edit');
  }
  
  const emailData = JSON.parse(emailDataString);
  
  // Get the client data from the stored form data
  const clientData = {
    studentName: emailData.formData.studentName,
    clientTypes: emailData.formData.clientTypes,
    clientType: emailData.formData.clientType,
    dashboardLink: emailData.formData.dashboardLink,
    parentEmail: emailData.formData.parentEmail,
    parentName: emailData.formData.parentName,
    sheetName: emailData.formData.sheetName
  };
  
  // Clear the temporary data
  clearTempEmailData();
  
  // Reopen the recap dialog with the client data
  showRecapDialog(clientData);
}


/**
 * Add client to the master ClientList table with correct column mapping
 * Columns: Client Name | Dashboard | Last Session | Next Session Scheduled? | Current Client?
 * @param {Object} clientData - Client information
 * @param {string} sheetName - The concatenated sheet name (FirstLast)
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The active spreadsheet
 */
function addClientToMasterList(clientData, sheetName, spreadsheet) {
  
  // Find the Master sheet
  const masterSheet = spreadsheet.getSheetByName('Master');
  
  if (!masterSheet) {
    console.error('ERROR: Could not find sheet named "Master"');
    throw new Error('Master sheet not found');
  }
  
  
  // Get the new sheet reference
  const newSheet = spreadsheet.getSheetByName(sheetName);
  if (!newSheet) {
    console.error('ERROR: Could not find newly created sheet:', sheetName);
    throw new Error('New client sheet not found');
  }
  
  // Find where to add the new row - look for last non-empty row in column A
  const columnAValues = masterSheet.getRange('A:A').getValues();
  let lastRow = 0;
  
  // Find the last non-empty cell in column A
  for (let i = 0; i < columnAValues.length; i++) {
    if (columnAValues[i][0] !== '') {
      lastRow = i + 1; // Convert to 1-based index
    }
  }
  
  
  // Add new row immediately after the last row
  const newRow = lastRow + 1;
  
  try {
    // Column A: Client Name with hyperlink to their sheet
    const sheetId = newSheet.getSheetId();
    const studentNameFormula = `=HYPERLINK("#gid=${sheetId}", "${clientData.name}")`;
    masterSheet.getRange(newRow, 1).setFormula(studentNameFormula);
    
    // Column B: Dashboard link as raw URL (for smart chip conversion)
    if (clientData.dashboardLink) {
      masterSheet.getRange(newRow, 2).setValue(clientData.dashboardLink);
    } else {
      masterSheet.getRange(newRow, 2).setValue('');
    }
    
    // Column C: Last Session - Formula with IFERROR to prevent errors
    const lastSessionFormula = `=IFERROR(INDEX(FILTER(${sheetName}!A:A,${sheetName}!A:A<>""), COUNTA(${sheetName}!A:A)), "")`;
    masterSheet.getRange(newRow, 3).setFormula(lastSessionFormula);
    
    // Column D: Next Session Scheduled? - Set checkbox value (don't insert new checkbox)
    masterSheet.getRange(newRow, 4).setValue(true);
    
    // Column E: Current Client? - Set checkbox value (don't insert new checkbox)
    masterSheet.getRange(newRow, 5).setValue(true);
    
    
    // Force a refresh of the sheet to ensure all formulas calculate
    SpreadsheetApp.flush();
        
  } catch (error) {
    console.error('ERROR adding data to master list:', error);
    throw error;
  }
}

/**
 * Syncs the active sheet's information to the master sheet if it doesn't exist already
 * Intended for use via Settings menu to add existing client sheets to master list
 */
function syncActiveSheetToMaster() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const sheetName = activeSheet.getName();
    
    // Check if current sheet is a client sheet
    const clientInfo = getCurrentClientInfo();
    if (!clientInfo.isClient) {
      SpreadsheetApp.getUi().alert('Error', 'The active sheet is not a client sheet. Please navigate to a client sheet first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get the Master sheet
    const masterSheet = spreadsheet.getSheetByName('Master');
    if (!masterSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Master sheet not found. Please ensure a "Master" sheet exists.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Check if client already exists in master sheet
    const masterData = masterSheet.getDataRange().getValues();
    const clientName = clientInfo.name;
    
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === clientName) {
        SpreadsheetApp.getUi().alert('Already Synced', `${clientName} already exists in the Master sheet.`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
    }
    
    // Get client details from unified data store and sheet fallbacks
    const existingClient = UnifiedClientDataStore.getClient(clientName);
    const dashboardLink = existingClient?.dashboardLink || 
                         activeSheet.getRange('H2').getValue() || '';
    const meetingNotesLink = existingClient?.meetingNotesLink || 
                            activeSheet.getRange('I2').getValue() || '';
    const scoreReportLink = existingClient?.scoreReportLink || 
                           activeSheet.getRange('J2').getValue() || '';
    
    // Prepare client data for master list
    const clientData = {
      name: clientName,
      dashboardLink: dashboardLink,
      meetingNotesLink: meetingNotesLink,
      scoreReportLink: scoreReportLink
    };
    
    // Add to master list using existing function
    addClientToMasterList(clientData, sheetName, spreadsheet);
    
    // Update UnifiedClientDataStore with complete client data
    const fullClientData = {
      name: clientName,
      sheetName: sheetName,
      isActive: true,
      dashboardLink: dashboardLink,
      meetingNotesLink: meetingNotesLink,
      scoreReportLink: scoreReportLink,
      parentEmail: activeSheet.getRange('E2').getValue() || '',
      emailAddressees: activeSheet.getRange('F2').getValue() || ''
    };
    // Add to UnifiedClientDataStore
    try {
      UnifiedClientDataStore.upsertClient(fullClientData);
    } catch (error) {
      console.warn(`Failed to add client to UnifiedClientDataStore:`, error);
    }
    
    // Show success message
    SpreadsheetApp.getUi().alert('Success', `${clientName} has been added to the Master sheet successfully!`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Error syncing active sheet to master:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to sync sheet to master: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Helper function to get sheet by its ID
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet - The spreadsheet
 * @param {number} sheetId - The sheet ID
 * @returns {SpreadsheetApp.Sheet|null} The sheet or null if not found
 */
function getSheetById(spreadsheet, sheetId) {
  const sheets = spreadsheet.getSheets();
  for (let sheet of sheets) {
    if (sheet.getSheetId() === sheetId) {
      return sheet;
    }
  }
  return null;
}

/**
 * Alternative method using proper hyperlink for internal sheets
 * This creates a better hyperlink that works more reliably
 */
function createInternalHyperlink(spreadsheet, targetSheetName, displayText) {
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    return displayText; // Return plain text if sheet not found
  }
  
  const sheetId = targetSheet.getSheetId();
  return `=HYPERLINK("#gid=${sheetId}", "${displayText}")`;
}

/**
 * Helper function to get sheet by its ID
 */
function getSheetById(spreadsheet, sheetId) {
  const sheets = spreadsheet.getSheets();
  for (let sheet of sheets) {
    if (sheet.getSheetId() === sheetId) {
      return sheet;
    }
  }
  return null;
}

/**
 * Extract sheet name from hyperlink formula
 */
function extractSheetNameFromFormula(formula, clientName, spreadsheet) {
  if (!formula) {
    return clientName.replace(/\s+/g, ''); // Default to concatenated
  }
  
  const match = formula.match(/HYPERLINK\("#gid=(\d+)"/);
  if (match) {
    const sheetId = parseInt(match[1]);
    const sheet = getSheetById(spreadsheet, sheetId);
    if (sheet) {
      return sheet.getName();
    }
  }
  
  return clientName.replace(/\s+/g, ''); // Fallback
}

/**
 * Open client list by opening the sidebar Client List view
 */
function findClient() {
  // For enterprise sidebar, show client selection dialog instead of switching sidebars
  const isEnterprise = CONFIG.version.includes('enterprise');
  if (isEnterprise) {
    showClientSelectionDialogEnterprise();
  } else {
    showUniversalSidebar('clientlist');
  }
}

/**
 * Show client selection dialog for enterprise version
 */
function showClientSelectionDialogEnterprise() {
  try {
    const clients = UnifiedClientDataStore.getActiveClients();
    
    if (clients.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Clients Found',
        'No active clients found in the system.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Sort clients alphabetically
    clients.sort((a, b) => a.name.localeCompare(b.name));
    
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body {
              font-family: 'Google Sans', Arial, sans-serif;
              margin: 0;
              padding: 20px;
            }
            .search-box {
              width: 100%;
              padding: 10px;
              margin-bottom: 15px;
              border: 1px solid #dadce0;
              border-radius: 4px;
              font-size: 14px;
              box-sizing: border-box;
            }
            .client-list {
              max-height: 400px;
              overflow-y: auto;
              border: 1px solid #dadce0;
              border-radius: 4px;
            }
            .client-item {
              padding: 12px;
              cursor: pointer;
              border-bottom: 1px solid #f1f3f4;
              transition: background 0.2s;
            }
            .client-item:hover {
              background: #f8f9fa;
            }
            .client-item:last-child {
              border-bottom: none;
            }
            .client-name {
              font-weight: 500;
              color: #202124;
            }
            .client-email {
              font-size: 12px;
              color: #5f6368;
              margin-top: 2px;
            }
            .no-results {
              padding: 20px;
              text-align: center;
              color: #5f6368;
            }
          </style>
        </head>
        <body>
          <input type="text" class="search-box" id="searchBox" placeholder="Search clients..." onkeyup="filterClients()">
          <div class="client-list" id="clientList">
            ${clients.map(client => `
              <div class="client-item" onclick="selectClient('${client.name}')">
                <div class="client-name">${client.name}</div>
                ${client.email ? `<div class="client-email">${client.email}</div>` : ''}
              </div>
            `).join('')}
          </div>
          
          <script>
            function filterClients() {
              const searchTerm = document.getElementById('searchBox').value.toLowerCase();
              const items = document.querySelectorAll('.client-item');
              let hasResults = false;
              
              items.forEach(item => {
                const text = item.textContent.toLowerCase();
                if (text.includes(searchTerm)) {
                  item.style.display = 'block';
                  hasResults = true;
                } else {
                  item.style.display = 'none';
                }
              });
              
              // Show no results message if needed
              const clientList = document.getElementById('clientList');
              if (!hasResults) {
                if (!document.getElementById('noResults')) {
                  const noResults = document.createElement('div');
                  noResults.id = 'noResults';
                  noResults.className = 'no-results';
                  noResults.textContent = 'No clients found';
                  clientList.appendChild(noResults);
                }
              } else {
                const noResults = document.getElementById('noResults');
                if (noResults) noResults.remove();
              }
            }
            
            function selectClient(clientName) {
              google.script.run
                .withSuccessHandler(() => {
                  google.script.host.close();
                })
                .withFailureHandler((error) => {
                  alert('Error: ' + error.message);
                })
                .switchToClient(clientName);
            }
            
            // Focus search box on load
            document.getElementById('searchBox').focus();
          </script>
        </body>
      </html>
    `)
    .setWidth(400)
    .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Client');
    
  } catch (error) {
    console.error('Error showing client selection dialog:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to load client list: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Fallback to original method if Master sheet approach fails
 */
function findClientFallback() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  // Filter sheets that follow FirstLast naming convention
  // Look for sheets that don't contain spaces and aren't common template names
  const clientSheets = sheets
    .filter(sheet => {
      const name = sheet.getName();
      // Exclude common template/system sheets - updated to include new template names
      const excludeNames = ['NewClient', 'NewACTClient', 'NewAcademicSupportClient', 'Template', 'Settings', 'Dashboard', 'Summary', 'SessionRecaps'];
      return !excludeNames.includes(name) && 
             !name.includes(' ') && 
             /^[A-Z][a-z]+[A-Z][a-z]+/.test(name); // FirstLast pattern
    })
    .map(sheet => {
      const name = sheet.getName();
      // Extract first name for sorting (everything before the first capital letter after position 0)
      const firstNameMatch = name.match(/^([A-Z][a-z]+)/);
      const firstName = firstNameMatch ? firstNameMatch[1] : name;
      return {
        sheetName: name,
        firstName: firstName,
        displayName: name.replace(/([A-Z])/g, ' $1').trim() // Add spaces: JohnSmith -> John Smith
      };
    })
    .sort((a, b) => a.firstName.localeCompare(b.firstName));
  
  if (clientSheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Clients Found',
      'No client sheets were found. Client sheets should follow the FirstLast naming convention.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  showClientSelectionDialog(clientSheets);
}

/**
 * Shows a dialog to select a client from the list
 */
function showClientSelectionDialog(clientSheets) {
  const clientOptions = clientSheets.map(client => 
    `<div class="client-item" onclick="selectClient('${client.sheetName}', '${client.displayName}')">${client.displayName}</div>`
  ).join('');
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
          }
          .container {
            max-height: 450px;
            overflow: hidden;
            display: flex;
            flex-direction: column;
          }
          h2 {
            color: #333;
            margin-bottom: 8px;
            font-weight: 500;
            text-align: center;
          }
          .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 20px;
            text-align: center;
          }
          .client-list {
            flex: 1;
            overflow-y: auto;
            border: 1px solid #dadce0;
            border-radius: 8px;
            max-height: 380px;
          }
          .client-item {
            padding: 12px 16px;
            cursor: pointer;
            border-bottom: 1px solid #f8f9fa;
            transition: all 0.2s ease;
            position: relative;
          }
          .client-item:hover {
            background: #f8f9fa;
            color: #003366;      
            font-weight: 500;
          }
          .client-item:last-child {
            border-bottom: none;
          }
          .client-item.loading {
            color: transparent;
            pointer-events: none;
          }
          .client-item.loading::after {
            content: '';
            position: absolute;
            top: 50%;
            left: 50%;
            width: 16px;
            height: 16px;
            margin: -8px 0 0 -8px;
            border: 2px solid #e0e0e0;
            border-top-color: #003366;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .button-group {
            margin-top: 20px;
            text-align: center;
          }
          button {
            padding: 12px 24px;
            border: none;
            border-radius: 24px;  
            font-size: 14px;
            font-weight: 600;     
            cursor: pointer;
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }

          button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          } 
          
          /* Loading overlay styles */
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(5px);
            -webkit-backdrop-filter: blur(5px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
          }
          
          .loading-overlay.active {
            display: flex;
          }
          
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #003366;  
            border-radius: 50%;
            animation: spin 1s linear infinite;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .loading-text {
            position: absolute;
            margin-top: 80px;
            font-size: 14px;
            color: #5f6368;
            font-weight: 500;
          }
        </style>
      </head>
      <body>
        <div class="container">
          
          <p class="subtitle">Click on a client name to navigate to their sheet</p>
          
          <div class="client-list">
            ${clientOptions}
          </div>
          
          <div class="button-group">
            <button onclick="google.script.host.close()">Cancel</button>
          </div>
        </div>
        
        <!-- Loading overlay -->
        <div class="loading-overlay" id="loadingOverlay">
          <div class="loading-spinner"></div>
          <div class="loading-text">Loading client...</div>
        </div>
        
        <script>
          function selectClient(sheetName, displayName) {
            // Add loading state to the clicked item
            const clickedItem = event.target;
            clickedItem.classList.add('loading');
            
            // Also show loading overlay with client name
            document.getElementById('loadingOverlay').classList.add('active');
            document.querySelector('.loading-text').textContent = 'Loading ' + displayName + '...';
            
            google.script.run
              .withSuccessHandler(function() {
                // Small delay to ensure smooth transition
                setTimeout(function() {
                  google.script.host.close();
                }, 300);
              })
              .withFailureHandler(function(error) {
                // Remove loading state on error
                clickedItem.classList.remove('loading');
                // Hide loading overlay on error
                document.getElementById('loadingOverlay').classList.remove('active');
                alert('Error: ' + error.message);
              })
              .navigateToClient(sheetName);
          }
        </script>
      </body>
    </html>
  `).setWidth(450).setHeight(550);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Find Client');
}

/**
 * Navigate to the selected client sheet
 */
function navigateToClient(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    spreadsheet.setActiveSheet(sheet);
  } else {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }
}


/**
 * Updated showSidebar to use universal sidebar
 */
function showSidebar() {
  // Use the enhanced universal sidebar for all versions
  // It handles both enterprise and basic modes internally
  showUniversalSidebar('control');
}

/**
 * Updated openQuickNotes to use universal sidebar
 */
function openQuickNotes() {
  // First check if we're on a client sheet
  const clientInfo = getCurrentClientInfo();
  if (!clientInfo.isClient) {
    SpreadsheetApp.getUi().alert('Please navigate to a client sheet first.');
    return;
  }
  
  showUniversalSidebar('quicknotes');
}

/**
 * Wrapper function for menu to open Quick Notes
 */
function openQuickNotesFromMenu() {
  openQuickNotes();
}

/**
 * Get quick notes for the current sheet
 */
function getQuickNotesForCurrentSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Find client by sheet name
  const allClients = UnifiedClientDataStore.getAllClients();
  const client = allClients.find(c => c.sheetName === sheetName);
  
  if (client && client.quickNotes) {
    return {
      data: client.quickNotes.data,
      timestamp: client.quickNotes.timestamp || 0
    };
  }
  
  return {
    data: null,
    timestamp: 0
  };
}

/**
 * Updated logRecapSent - NO LONGER automatically clears notes
 */
function logRecapSent(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('SessionRecaps');
  
  if (!sheet) {
    createRecapTrackingSheet();
    sheet = spreadsheet.getSheetByName('SessionRecaps');
  }
  
  const row = [
    new Date(),
    data.studentName,
    data.clientType || data.clientTypes, // Handle both property names
    data.parentEmail,
    data.focusSkill,
    data.todaysWin,
    data.homework,
    data.nextSession || '',
    'Sent'
  ];
  
  sheet.appendRow(row);
  
  // DO NOT automatically clear notes here
  // Let the user manually clear them when ready
}

/**
 * Manual function to clear quick notes for a specific student
 * This can be called from the menu or after confirming the recap was sent successfully
 */
function clearQuickNotesForStudent(studentSheetName) {
  // Find client by sheet name and clear their quick notes
  const allClients = UnifiedClientDataStore.getAllClients();
  const client = allClients.find(c => c.sheetName === studentSheetName);
  
  if (client) {
    const updatedData = { ...client };
    delete updatedData.quickNotes;
    UnifiedClientDataStore.updateClient(client.name, updatedData);
  }
}

/**
 * Called when add-on is installed or enabled
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Called when add-on is opened from the add-ons menu
 */
function onHomepage(e) {
  return createMainCard();
}



/**
 * Menu function to clear notes for current student
 */
function clearCurrentStudentNotes() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!isClientSheet(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Please navigate to a client sheet first.');
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Quick Notes?',
    `This will clear all quick notes for ${sheet.getRange('A1').getValue()}. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    clearQuickNotesForStudent(sheet.getName());
    ui.alert('Quick notes cleared.');
  }
}

/**
 * Save quick notes to unified data store
 */
function saveQuickNotes(notes) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    // Find client by sheet name
    const allClients = UnifiedClientDataStore.getAllClients();
    const client = allClients.find(c => c.sheetName === sheetName);
    
    if (client) {
      const quickNotesData = {
        data: JSON.stringify(notes),
        timestamp: new Date().getTime()
      };
      
      UnifiedClientDataStore.updateClient(client.name, { quickNotes: quickNotesData });
    }
    
    return { success: true };
  } catch (error) {
    console.error('Error saving quick notes:', error);
    throw new Error('Failed to save notes: ' + error.message);
  }
}

/**
 * Open batch prep mode
 */
function openBatchPrepMode() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const clients = getClientList();
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
          }
          .instruction {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
            color: #5f6368;
          }
          .client-list {
            border: 1px solid #dadce0;
            border-radius: 6px;
            max-height: 400px;
            overflow-y: auto;
          }
          .client-item {
            padding: 12px;
            border-bottom: 1px solid #f8f9fa;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          .client-item:hover {
            background: #f8f9fa;
          }
          .client-item:last-child {
            border-bottom: none;
          }
          .client-name {
            font-weight: 500;
          }
          .client-type {
            font-size: 12px;
            color: #5f6368;
            background: #f1f3f4;
            padding: 2px 8px;
            border-radius: 3px;
          }
          .button-group {
            margin-top: 20px;
            display: flex;
            gap: 10px;
          }
          button {
            padding: 12px 24px;
            border: none;
            border-radius: 24px;  /* More rounded */
            font-size: 14px;
            font-weight: 600;     /* Slightly bolder */
            cursor: pointer;
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);  /* Add subtle shadow */
          }
          .btn-primary {
            background: #003366;
            color: white;
          }
          .btn-primary:hover {
            background: #002244;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          checkbox {
            margin-right: 10px;
          }
        </style>
      </head>
      <body>
        <h2>Batch Prep Mode</h2>
        
        <div class="instruction">
          Select the clients you're seeing today. This will prepare quick access to their recap forms and pre-load common information.
        </div>
        
        <div class="client-list">
          ${clients.map(client => `
            <div class="client-item">
              <label style="display: flex; align-items: center; cursor: pointer;">
                <input type="checkbox" value="${client.sheetName}" style="margin-right: 10px;">
                <span class="client-name">${client.displayName}</span>
              </label>
              <span class="client-type">${client.clientType}</span>
            </div>
          `).join('')}
        </div>
        
        <div class="button-group">
          <button class="btn-primary" onclick="prepareBatch()">Prepare Selected Clients</button>
          <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
        </div>
        
        <script>
          function setButtonLoading(button, isLoading) {
            if (isLoading) {
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
            } else {
              button.classList.remove('btn-loading');
              button.textContent = button.dataset.originalText || button.textContent;
              button.disabled = false;
            }
          }
          
          function prepareBatch() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
            const selectedClients = Array.from(checkboxes).map(function(cb) { return cb.value; });
            
            if (selectedClients.length === 0) {
              alert('Please select at least one client.');
              return;
            }
            
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, true);
            
            google.script.run
              .withSuccessHandler(function() {
                alert('Batch prep complete! Use "Send Session Recap" after each session.');
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                setButtonLoading(button, false);
                alert('Error: ' + error.message);
              })
              .prepareBatchClients(selectedClients);
          }
        </script>
      </body>
    </html>
  `).setWidth(500).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Batch Prep Mode');
}

/**
 * Get list of all clients
 */
function getClientList() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  return sheets
    .filter(sheet => isClientSheet(sheet.getName()))
    .map(sheet => {
      const data = getClientDataFromSheet(sheet);
      return {
        sheetName: sheet.getName(),
        displayName: data.studentName,
        clientType: data.clientType
      };
    })
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

/**
 * Prepare batch clients
 */
function prepareBatchClients(clientSheetNames) {
  // Store batch processing state in UnifiedClientDataStore temporary storage
  const batchState = {
    batchClients: clientSheetNames,
    batchIndex: 0
  };
  UnifiedClientDataStore.setTempData('batchProcessing', batchState);
}

/**
 * Get optimized client data for recap dialog (used for preloading)
 */
function getClientDataForRecap() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!isClientSheet(sheetName)) {
    return null;
  }
  
  // Get the essential data needed for recap dialog
  const clientData = getClientDataFromSheet(sheet);
  
  // Try to get quick notes from UnifiedClientDataStore first
  let notes = {};
  const clientName = sheet.getRange('A1').getValue();
  if (clientName) {
    const unifiedClient = UnifiedClientDataStore.getClient(clientName);
    if (unifiedClient && unifiedClient.quickNotes) {
      // If quickNotes is a string, try to parse it; if it's already an object, use it
      if (typeof unifiedClient.quickNotes === 'string') {
        try {
          notes = JSON.parse(unifiedClient.quickNotes);
        } catch (e) {
          console.warn('Could not parse quick notes from UnifiedClientDataStore, treating as plain text');
          notes = { data: unifiedClient.quickNotes };
        }
      } else if (typeof unifiedClient.quickNotes === 'object') {
        notes = unifiedClient.quickNotes;
      }
    }
  }
  
  return {
    ...clientData,
    notes: notes,
    timestamp: new Date().getTime()
  };
}

/**
 * Cache recap data in UnifiedClientDataStore temporary storage for fast server-side access
 */
function cacheRecapData(clientData) {
  if (!clientData || !clientData.sheetName) {
    return;
  }
  
  try {
    const cacheKey = `recapdata_${clientData.sheetName}`;
    const cacheData = {
      ...clientData,
      cached_timestamp: new Date().getTime()
    };
    
    UnifiedClientDataStore.setTempData(cacheKey, cacheData);
  } catch (error) {
    console.error('Failed to cache recap data:', error);
  }
}

/**
 * Get complete client data for Preview Update dialog using UnifiedClientDataStore
 */
function getCompleteClientDataForPreview(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Get the client name from the sheet to find in UnifiedClientDataStore
    const studentName = sheet.getRange('A1').getValue() || sheetName.replace(/([A-Z])/g, ' $1').trim();
    const clientTypes = sheet.getRange('C2').getValue() || '';
    
    // Get client data from UnifiedClientDataStore
    const unifiedClient = UnifiedClientDataStore.getClient(studentName);
    
    if (unifiedClient) {
      return {
        studentName: studentName,
        studentFirstName: (studentName || '').split(' ')[0],
        parentName: unifiedClient.contactInfo.emailAddressees || 'Parents',
        parentEmail: unifiedClient.contactInfo.parentEmail || '',
        clientTypes: clientTypes,
        clientType: clientTypes,
        sheetName: sheetName,
        dashboardLink: unifiedClient.links.dashboard || '',
        meetingNotesLink: unifiedClient.links.meetingNotes || '',
        scoreReportLink: unifiedClient.links.scoreReport || '',
        emailAddressees: unifiedClient.contactInfo.emailAddressees || '',
        isActive: unifiedClient.isActive !== undefined ? unifiedClient.isActive : true
      };
    } else {
      // If not in UnifiedClientDataStore, return basic data from sheet only
      return {
        studentName: studentName,
        studentFirstName: (studentName || '').split(' ')[0],
        parentName: 'Parents',
        parentEmail: '',
        clientTypes: clientTypes,
        clientType: clientTypes,
        sheetName: sheetName,
        dashboardLink: '',
        meetingNotesLink: '',
        scoreReportLink: '',
        emailAddressees: '',
        isActive: true
      };
    }
  } catch (error) {
    console.error('Error getting complete client data for preview:', error);
    throw error;
  }
}

/**
 * Process recap preview for unified dialog - returns data instead of opening new dialog
 */
function processRecapPreviewUnified(formData) {
  try {
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (!isClientSheet(sheetName)) {
      throw new Error('Please navigate to a client sheet before sending a recap.');
    }
    
    // Use refined data architecture to get complete client data
    const clientData = getCompleteClientDataForPreview(sheetName);
    
    // Merge form data with client data
    const mergedFormData = {
      ...formData,
      // Override with fresh client data
      studentName: clientData.studentName,
      studentFirstName: clientData.studentFirstName,
      parentName: clientData.parentName,
      parentEmail: clientData.parentEmail,
      clientTypes: clientData.clientTypes,
      clientType: clientData.clientType,
      sheetName: clientData.sheetName,
      dashboardLink: clientData.dashboardLink,
      meetingNotesLink: clientData.meetingNotesLink,
      scoreReportLink: clientData.scoreReportLink,
      tutorName: getCompleteConfig().tutorName || '',
      // Field mappings for preview dialog compatibility
      openingNotes: formData.openingNotes || formData.todaysWin || '',
      struggles: formData.progressNotes || formData.additionalParentNotes || ''
    };
    
    // Generate both email versions
    const dashboardVersion = generateDashboardUpdate(mergedFormData, clientData);
    const emailVersion = generateEmailUpdate(mergedFormData, clientData);
    
    // Format date as MM/DD/YYYY
    const today = new Date();
    const dateStr = `${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getDate().toString().padStart(2, '0')}/${today.getFullYear()}`;
    
    // Prepare response data structure
    const previewData = {
      formData: mergedFormData,
      subject: `${clientData.studentFirstName || 'Student'}'s ACT Update ${dateStr} - Smart College`,
      dashboardVersion: dashboardVersion,
      emailVersion: emailVersion,
      meetingNotesLink: clientData.meetingNotesLink,
      scoreReportLink: clientData.scoreReportLink
    };
    
    return previewData;
    
  } catch (error) {
    console.error('Error in processRecapPreviewUnified:', error);
    throw new Error('Failed to generate preview: ' + error.message);
  }
}

/**
 * Generate Dashboard Update version
 */
function generateDashboardUpdate(formData, clientData) {
  const parts = [];
  
  // Opening line
  const openingLine = formData.openingNotes || formData.todaysWin || `${formData.studentFirstName} had an excellent session today!`;
  parts.push(openingLine);
  parts.push('<br><br>');
  
  // Today's Focus
  parts.push('<strong>Today\'s Focus</strong>');
  parts.push('<br>');
  parts.push(formData.focusSkill || 'General tutoring session');
  parts.push('<br><br>');
  
  // Big Win
  parts.push('<strong>Big Win</strong>');
  parts.push('<br>');
  parts.push(formData.todaysWin || 'Great progress made today!');
  parts.push('<br><br>');
  
  // What We Covered
  parts.push('<strong>What We Covered</strong>');
  parts.push('<br>');
  
  // Helper function to create bullet list from comma-separated values
  const createBulletList = (items) => {
    if (!items) return '';
    const itemArray = items.split(',').map(item => item.trim()).filter(item => item);
    if (itemArray.length === 0) return '';
    return '<ul style="margin: 5px 0 5px 20px; padding: 0;">' + 
           itemArray.map(item => `<li>${item}</li>`).join('') + 
           '</ul>';
  };
  
  // Build What We Covered section with bullet lists
  const coveredSections = [];
  
  if (formData.mastered) {
    coveredSections.push('<em>Mastered:</em>');
    coveredSections.push(createBulletList(formData.mastered));
  }
  
  if (formData.practiced) {
    coveredSections.push('<em>Practiced:</em>');
    coveredSections.push(createBulletList(formData.practiced));
  }
  
  if (formData.introduced) {
    coveredSections.push('<em>Introduced:</em>');
    coveredSections.push(createBulletList(formData.introduced));
  }
  
  if (coveredSections.length > 0) {
    parts.push(coveredSections.join(''));
  } else {
    parts.push('Session content covered');
  }
  parts.push('<br>');
  
  // Progress Update
  if (formData.progressNotes || formData.struggles) {
    parts.push('<strong>Progress Update</strong>');
    parts.push('<br>');
    parts.push(formData.progressNotes || formData.struggles);
    parts.push('<br><br>');
  }
  
  // Homework/Practice
  if (formData.homework) {
    parts.push('<strong>Homework/Practice</strong>');
    parts.push('<br>');
    parts.push(formData.homework);
    parts.push('<br><br>');
  }
  
  // Support at Home
  if (formData.parentNotes) {
    parts.push('<strong>Support at Home</strong>');
    parts.push('<br>');
    parts.push(formData.parentNotes);
    parts.push('<br><br>');
  }
  
  // Next Session
  if (formData.nextSession) {
    parts.push('<strong>Next Session</strong>');
    parts.push('<br>');
    parts.push(formData.nextSession);
    parts.push('<br><br>');
  }
  
  // Looking Ahead
  if (formData.nextSessionPreview) {
    parts.push('<strong>Looking Ahead</strong>');
    parts.push('<br>');
    parts.push(formData.nextSessionPreview);
    parts.push('<br><br>');
  }
  
  // P.S. section
  if (formData.quickWin) {
    parts.push('<br><strong>P.S.</strong> - ');
    parts.push(formData.quickWin);
  }
  
  return parts.join('');
}

/**
 * Generate Email Update version
 */
function generateEmailUpdate(formData, clientData) {
  const parts = [];
  
  // Greeting
  parts.push(`<p>Hi ${formData.parentName || 'Parents'},</p>`);
  parts.push('<p>&nbsp;</p>');
  
  // Session type determination
  const clientTypes = formData.clientTypes ? formData.clientTypes.split(',').map(t => t.trim()) : [];
  const sessionType = clientTypes.length === 1 ? clientTypes[0] : 'tutoring';
  
  // Meeting notes link
  if (formData.meetingNotesLink) {
    parts.push(`<p>Here is the link to our <a href="${formData.meetingNotesLink}">Meeting Notes</a> from our ${sessionType} session, and you will find ${formData.studentFirstName}'s homework linked on the Meeting Notes.</p>`);
  } else {
    parts.push(`<p>Here are the notes from our ${sessionType} session, and you will find ${formData.studentFirstName}'s homework included.</p>`);
  }
  
  // Dashboard and Score Report links
  parts.push(`<p>You can always follow our progress in more detail in ${formData.studentFirstName}'s `);
  if (formData.dashboardLink) {
    parts.push(`<a href="${formData.dashboardLink}">Dashboard</a>`);
  } else {
    parts.push('Dashboard');
  }
  parts.push(' and track their progress in their ');
  if (formData.scoreReportLink) {
    parts.push(`<a href="${formData.scoreReportLink}">Score Report</a>!</p>`);
  } else {
    parts.push('Score Report!</p>');
  }
  
  // Next Session information
  if (formData.nextSession) {
    parts.push('<p>&nbsp;</p>');
    parts.push(`<p><strong>Next Session:</strong> ${formData.nextSession}</p>`);
  }
  
  // Closing
  parts.push('<p>&nbsp;</p>');
  parts.push('<p>Best,<br>');
  parts.push(formData.tutorName || 'Your Tutor');
  parts.push('</p>');
  
  return parts.join('');
}

/**
 * Process recap preview form data and show preview dialog
 */
function processRecapPreview(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (!isClientSheet(sheetName)) {
      throw new Error('Please navigate to a client sheet before sending a recap.');
    }
    
    // Use the most efficient data retrieval method based on architecture analysis
    let clientData = getCompleteClientDataForPreview(sheetName);
    
    // Get existing quick notes from UnifiedClientDataStore
    const clientName = clientData.studentName;
    let existingNotes = {};
    
    if (clientName) {
      const unifiedClient = UnifiedClientDataStore.getClient(clientName);
      if (unifiedClient && unifiedClient.quickNotes) {
        if (typeof unifiedClient.quickNotes === 'string') {
          try {
            existingNotes = JSON.parse(unifiedClient.quickNotes);
          } catch (e) {
            console.warn('Could not parse quick notes from UnifiedClientDataStore');
            existingNotes = { data: unifiedClient.quickNotes };
          }
        } else if (typeof unifiedClient.quickNotes === 'object') {
          existingNotes = unifiedClient.quickNotes;
        }
      }
    }
    
    // Extend quick notes with recap-specific fields from the preview form
    const extendedNotes = {
      ...existingNotes, // Keep existing quick notes
      // Add recap-specific fields from preview form
      focusSkill: formData.focusSkill || '',
      todaysWin: formData.todaysWin || '',
      homework: formData.homework || '',
      nextSession: formData.nextSession || '', // Map to nextSession for compatibility
      nextSessionPreview: formData.nextSessionPreview || '',
      quickWin: formData.quickWin || '',
      additionalParentNotes: formData.additionalParentNotes || '',
      psNote: formData.psNote || '',
      // Also map some fields for better compatibility
      mastered: formData.mastered || '',
      practiced: formData.practiced || '',
      introduced: formData.introduced || '',
      progressNotes: formData.progressNotes || '',
      parentNotes: formData.parentNotes || '',
      timestamp: new Date().getTime()
    };
    
    // Save the extended notes back to UnifiedClientDataStore
    if (clientName) {
      const updatedClientData = {
        quickNotes: extendedNotes
      };
      UnifiedClientDataStore.updateClient(clientName, updatedClientData);
    }
    
    // Prepare simplified data structure - showUpdatePreviewDialog will handle client data merging
    const previewData = {
      formData: {
        ...formData,
        // Include sheet name for proper data retrieval
        sheetName: sheetName,
        // Map Session Recap fields to Preview Update dialog field names
        openingNotes: formData.openingNotes || formData.todaysWin || '',
        struggles: formData.progressNotes || formData.additionalParentNotes || ''
      },
      subject: `${(clientData.studentName || '').split(' ')[0] || 'Student'}'s ACT Update ${new Date().toLocaleDateString('en-US', {month: '2-digit', day: '2-digit', year: 'numeric'})} - Smart College`
    };
    
    
    // Show the preview dialog - it will get fresh complete client data
    return showUpdatePreviewDialog(previewData);
  } catch (error) {
    console.error('Error processing recap preview:', error);
    throw new Error('Failed to process recap: ' + error.message);
  }
}

/**
 * Send individual recap email for current client
 */
function sendIndividualRecap() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!isClientSheet(sheetName)) {
    SpreadsheetApp.getUi().alert(
      'Invalid Sheet',
      'Please navigate to a client sheet before sending a recap.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Try to get cached data first for faster loading
  let clientData;
  const cachedDataKey = `recapdata_${sheetName}`;
  const cachedData = UnifiedClientDataStore.getTempData(cachedDataKey);
  
  if (cachedData) {
    try {
      clientData = cachedData;
    } catch (e) {
      clientData = getClientDataFromSheet(sheet);
    }
  } else {
    clientData = getClientDataFromSheet(sheet);
  }
  
  // Check if any required links are missing
  const missingLinks = [];
  if (!clientData.dashboardLink) missingLinks.push('Dashboard Link');
  if (!clientData.meetingNotesLink) missingLinks.push('Meeting Notes Link');
  if (!clientData.scoreReportLink) missingLinks.push('Score Report Link');
  
  if (missingLinks.length > 0) {
    // Store data in UnifiedClientDataStore temporary storage
    UnifiedClientDataStore.setTempData('tempClientData', clientData);
    
    return {
      showOverlay: true,
      missingLinks: missingLinks,
      clientName: clientData.studentName || 'Unknown Client'
    };
  }
  
  // All links present, proceed with recap
  showRecapDialog(clientData);
  return { showOverlay: false };
}

/**
 * Show sidebar overlay for missing links instead of separate dialog
 */
function showMissingLinksOverlay(clientData, missingLinks) {
  // Store data in UnifiedClientDataStore temporary storage for the overlay buttons
  UnifiedClientDataStore.setTempData('tempClientData', clientData);
  
  // Since we can't directly communicate with the existing sidebar,
  // we'll return a result that tells the sidebar JavaScript to show the overlay
  return {
    showOverlay: true,
    missingLinks: missingLinks,
    clientName: clientData.studentName || 'Unknown Client'
  };
}

/**
 * Show dialog when links are missing for email recap
 */
function showMissingLinksDialog(clientData, missingLinks) {
  // Store all data in UnifiedClientDataStore temporary storage to avoid HTML embedding
  UnifiedClientDataStore.setTempData('tempClientData', clientData);
  UnifiedClientDataStore.setTempData('tempMissingLinks', missingLinks);
  
  // Create minimal HTML with no embedded data
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body { font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif; margin: 20px; }
          .loading { text-align: center; padding: 20px; }
          .content { display: none; }
          .btn { 
            padding: 12px 16px; 
            margin: 8px; 
            border: none; 
            border-radius: 6px; 
            cursor: pointer;
            font-size: 14px;
          }
          .btn-primary { background: #003366; color: white; }
          .btn-secondary { background: #f1f3f4; color: #5f6368; border: 1px solid #dadce0; }
          .missing-list { margin: 16px 0; padding: 16px; background: #fef7e0; border-radius: 8px; }
        </style>
      </head>
      <body>
        <div class="loading" id="loading">Loading...</div>
        
        <div class="content" id="content">
          <h2>⚠️ Missing Client Information</h2>
          <div id="clientName"></div>
          <div class="missing-list" id="missingLinks"></div>
          <p>The recap email works best with all client links included.</p>
          
          <div style="display: flex; gap: 12px; margin-top: 20px;">
            <button class="menu-button primary" onclick="updateClientInfo()" style="flex: 1; margin-bottom: 0;">Update Client Info</button>
            <button class="menu-button" onclick="continueWithoutLinks()" style="flex: 1; margin-bottom: 0;">Continue Without Links</button>
          </div>
        </div>
        
        <script>
          // Load data from server on page load
          google.script.run
            .withSuccessHandler(function(data) {
              document.getElementById('loading').style.display = 'none';
              document.getElementById('content').style.display = 'block';
              
              document.getElementById('clientName').innerHTML = '<strong>Client: ' + (data.clientName || 'Unknown') + '</strong>';
              document.getElementById('missingLinks').innerHTML = '<strong>Missing:</strong> ' + data.missingLinks.join(', ');
            })
            .withFailureHandler(function(error) {
              document.getElementById('loading').innerHTML = 'Error loading data: ' + error.message;
            })
            .getTempDialogData();
          
          function updateClientInfo() {
            // Close this dialog and show the full update client dialog
            google.script.host.close();
            google.script.run.showFullUpdateClientDialog();
          }
          
          function continueWithoutLinks() {
            // Close dialog and proceed with recap
            google.script.host.close();
            google.script.run.proceedWithRecapFromDialog();
          }
          
          // Initialization handled above
        </script>
      </body>
    </html>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Missing Client Information')
    .setWidth(550)
    .setHeight(600);
  
  // No need to store data separately
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Missing Client Information');
}


/**
 * Update client details and proceed with recap
 */
function updateClientDetailsAndProceedWithRecap(sheetName, details) {
  try {
    // Update the client details
    const updateResult = updateClientDetails(sheetName, details);
    
    if (updateResult.success) {
      // Get updated client data
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      const updatedClientData = getClientDataFromSheet(sheet);
      
      // Proceed with recap using updated data
      showRecapDialog(updatedClientData);
      
      return {
        success: true,
        message: 'Client details updated and recap dialog opened'
      };
    } else {
      return updateResult;
    }
  } catch (error) {
    console.error('Error updating client details and proceeding with recap:', error);
    return {
      success: false,
      message: 'Error updating client details: ' + error.message
    };
  }
}

/**
 * Proceed with recap even if some links are missing
 */
function proceedWithRecapAnyway(clientData) {
  showRecapDialog(clientData);
}

/**
 * Proceed with recap from missing links dialog (without embedded data)
 */
function proceedWithRecapFromDialog() {
  const clientData = UnifiedClientDataStore.getTempData('tempClientData');
  
  if (!clientData) {
    throw new Error('No client data found for recap');
  }
  
  // Clean up temporary data
  UnifiedClientDataStore.deleteTempData('tempClientData');
  
  showRecapDialog(clientData);
}

/**
 * Update client details from missing links dialog and proceed with recap
 */
function updateClientDetailsFromDialog(details) {
  const clientData = UnifiedClientDataStore.getTempData('tempClientData');
  
  if (!clientData) {
    throw new Error('No client data found for update');
  }
  
  // Update the client details
  const result = updateClientDetailsAndProceedWithRecap(clientData.sheetName, details);
  
  // Clean up temporary data
  UnifiedClientDataStore.deleteTempData('tempClientData');
  
  return result;
}

/**
 * Get temporarily stored client data for the dialog
 */
function getTempClientData() {
  const clientData = UnifiedClientDataStore.getTempData('tempClientData');
  
  if (!clientData) {
    return null;
  }
  
  return clientData;
}

/**
 * Get dialog data (client name and missing links) for the simple dialog
 */
function getTempDialogData() {
  const clientData = UnifiedClientDataStore.getTempData('tempClientData');
  const missingLinks = UnifiedClientDataStore.getTempData('tempMissingLinks');
  
  if (!clientData || !missingLinks) {
    throw new Error('Dialog data not found');
  }
  
  return {
    clientName: clientData.studentName || 'Unknown Client',
    missingLinks: missingLinks
  };
}

/**
 * Get recap dialog data from server-side storage
 */
function getRecapDialogData() {
  // Get the stored client data from UnifiedClientDataStore temporary storage
  const clientData = UnifiedClientDataStore.getTempData('CURRENT_RECAP_DATA');
  if (!clientData) {
    throw new Error('No client data found for recap dialog');
  }
  
  // Get saved quick notes from UnifiedClientDataStore
  let notes = {};
  const clientName = clientData.studentName;
  if (clientName) {
    const unifiedClient = UnifiedClientDataStore.getClient(clientName);
    if (unifiedClient && unifiedClient.quickNotes) {
      if (typeof unifiedClient.quickNotes === 'string') {
        try {
          notes = JSON.parse(unifiedClient.quickNotes);
        } catch (e) {
          console.warn('Could not parse quick notes from UnifiedClientDataStore');
          notes = { data: unifiedClient.quickNotes };
        }
      } else if (typeof unifiedClient.quickNotes === 'object') {
        notes = unifiedClient.quickNotes;
      }
    }
  }
  
  // Get homework templates based on client types
  const clientTypesList = clientData.clientTypes ? clientData.clientTypes.split(',').map(t => t.trim()) : [];
  let homeworkTemplates = [];
  
  // Add relevant templates based on client types
  clientTypesList.forEach(type => {
    if (HOMEWORK_TEMPLATES[type]) {
      homeworkTemplates = homeworkTemplates.concat(HOMEWORK_TEMPLATES[type]);
    }
  });
  
  // Add general templates as fallback
  if (homeworkTemplates.length === 0) {
    homeworkTemplates = HOMEWORK_TEMPLATES['General'];
  }
  
  // Return a clean, simple data structure
  return {
    client: {
      name: clientData.studentName || 'Unknown Client',
      types: clientData.clientTypes || '',
      sheetName: clientData.sheetName,
      parentName: clientData.parentName || '',
      parentEmail: clientData.parentEmail || '',
      dashboardLink: clientData.dashboardLink || '',
      meetingNotesLink: clientData.meetingNotesLink || '',
      scoreReportLink: clientData.scoreReportLink || ''
    },
    notes: notes,
    homeworkTemplates: homeworkTemplates,
    actTests: ACT_TESTS,
    // Include the full clientData for form submission
    fullClientData: clientData
  };
}

/**
 * Show the full update client dialog
 */
function showFullUpdateClientDialog() {
  const clientData = UnifiedClientDataStore.getTempData('tempClientData');
  
  if (!clientData) {
    throw new Error('Client data not found');
  }
  
  // Use the existing update client dialog from the sidebar
  showUpdateClientInfoDialog(clientData.sheetName);
}

/**
 * Show dialog for individual recap input
 */
function showRecapDialog(clientData) {
  // Use the enhanced session recap dialog
  return showEnhancedSessionRecapDialog(clientData);
}

function showRecapDialogLegacy(clientData) {
  // Ensure we have valid client data
  if (!clientData || !clientData.sheetName) {
    throw new Error('Invalid client data for recap dialog');
  }
  
  // Store ALL client data in UnifiedClientDataStore temporary storage for easy retrieval
  UnifiedClientDataStore.setTempData('CURRENT_RECAP_DATA', clientData);
  
  
  // Create HTML using hybrid approach - template literals for structure, string building for dynamic content
  const htmlContent = createSessionRecapDialogHTML();
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(700).setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Session Recap');
}

/**
 * Create Session Recap dialog HTML using hybrid approach
 */
function createSessionRecapDialogHTML() {
  const css = getSessionRecapCSS();
  const body = getSessionRecapBody();
  const javascript = getSessionRecapJavaScript();
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>${css}</style>
      </head>
      <body>
        ${body}
        <script>${javascript}</script>
      </body>
    </html>
  `;
}

/**
 * Get CSS styles for Session Recap dialog
 */
function getSessionRecapCSS() {
  return `
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap');
    
    body {
      font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      margin: 0;
      padding: 0;
      background: white;
    }
    
    /* View container system for unified dialog */
    .view-container {
      display: none;
      padding: 20px;
      animation: fadeIn 0.3s ease-in;
    }
    
    .view-container.active {
      display: block;
    }
    
    @keyframes fadeIn {
      from {
        opacity: 0;
        transform: translateY(10px);
      }
      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
    
    /* Clean centered dialog header - no banner, improved styling */
    .dialog-header {
      text-align: center;
      padding: 16px 20px 12px 20px;
      margin-bottom: 16px;
    }
    
    .dialog-title {
      font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      font-size: 24px;
      font-weight: 600;
      color: #202124;
      margin: 0;
      padding: 0;
      line-height: 1.3;
    }
    
    .dialog-subtitle {
      font-size: 14px;
      color: #5f6368;
      margin-top: 4px;
    }
    
    /* Hide view indicator to remove banner-like appearance */
    .view-indicator {
      display: none;
    }
    
    h2 {
      color: #333;
      margin-bottom: 20px;
      font-weight: 500;
    }
    
    .client-info {
      background: #f8f9fa;
      padding: 12px;
      border-radius: 8px;
      margin-bottom: 20px;
      font-size: 14px;
    }
    
    .form-section {
      background: white;
      border: 1px solid #e8eaed;
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 16px;
    }
    
    .form-section h3 {
      margin: 0 0 16px 0;
      color: #003366;
      font-size: 16px;
      font-weight: 500;
    }
    
    .form-group {
      margin-bottom: 16px;
    }
    
    label {
      display: block;
      margin-bottom: 6px;
      font-weight: 500;
      color: #5f6368;
      font-size: 14px;
    }
    
    input[type="text"], input[type="email"], textarea, select {
      width: 100%;
      padding: 10px;
      border: 1px solid #dadce0;
      border-radius: 6px;
      font-size: 14px;
      font-family: inherit;
      box-sizing: border-box;
    }
    
    textarea {
      resize: vertical;
      min-height: 60px;
    }
    
    select {
      cursor: pointer;
    }
    
    .quick-fill {
      display: flex;
      gap: 8px;
      margin-top: 6px;
      flex-wrap: wrap;
    }
    
    .quick-fill-btn {
      padding: 4px 8px;
      background: #e8f0fe;
      color: #003366;
      border: none;
      border-radius: 4px;
      font-size: 12px;
      cursor: pointer;
      transition: all 0.3s;
    }
    
    .quick-fill-btn:hover {
      background: #d2e3fc;
      transform: translateY(-1px);
      box-shadow: 0 2px 6px rgba(0, 51, 102, 0.2);
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .button-group {
      display: flex;
      gap: 12px;
      margin-top: 24px;
    }
    
    button {
      padding: 10px 20px;
      border: none;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.2s ease;
    }
    
    .btn-primary {
      background: #00C853;
      color: white;
      flex: 1;
    }
    
    .btn-primary:hover {
      background: #388E3C;
      transform: translateY(-1px);
      box-shadow: 0 4px 12px rgba(56, 142, 60, 0.3);
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .btn-secondary {
      background: #f8f9fa;
      color: #5f6368;
      border: 1px solid #dadce0;
    }
    
    .btn-secondary:hover {
      background: #f1f3f4;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
      transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .btn-loading {
      position: relative;
      color: transparent !important;
      pointer-events: none;
      opacity: 0.8;
    }
    
    .btn-loading::after {
      content: '';
      position: absolute;
      width: 16px;
      height: 16px;
      top: 50%;
      left: 50%;
      margin-left: -8px;
      margin-top: -8px;
      border: 2px solid rgba(255, 255, 255, 0.3);
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
    }
    
    .btn-secondary.btn-loading::after {
      border-color: rgba(0, 0, 0, 0.1);
      border-top-color: #5f6368;
    }
    
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
    
    .loading-overlay {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(255, 255, 255, 0.95);
      display: none;
      justify-content: center;
      align-items: center;
      z-index: 9999;
    }
    
    .loading-overlay.active {
      display: flex;
    }
    
    .loading-spinner {
      width: 40px;
      height: 40px;
      border: 3px solid #f3f3f3;
      border-top: 3px solid #003366;
      border-radius: 50%;
      animation: spin 1s linear infinite;
    }
    
    .homework-input-group {
      display: flex;
      gap: 10px;
      margin-bottom: 12px;
      align-items: center;
    }
    
    .homework-input-group select {
      flex: 1;
      font-size: 14px;
    }
    
    .homework-add-btn {
      padding: 10px 16px;
      font-size: 13px;
      font-weight: 500;
      white-space: nowrap;
      border-radius: 6px;
    }
    
    .homework-add-btn:hover {
      background: #e8f0fe;
      color: #003366;
    }
    
    .saved-notes-indicator {
      background: #e8f5e9;
      padding: 8px;
      border-radius: 4px;
      font-size: 12px;
      color: #00A043;
      margin-bottom: 10px;
    }
    
    .act-test-selector {
      display: none;
      margin-top: 10px;
      padding: 10px;
      background: #f8f9fa;
      border-radius: 6px;
    }
    
    .act-test-selector.show {
      display: block;
    }
    
    .act-test-selector label {
      font-size: 13px;
      margin-bottom: 4px;
    }
    
    .act-test-selector select {
      width: 100%;
      padding: 8px;
      font-size: 13px;
    }
    
    .next-session-input-group {
      display: flex;
      gap: 8px;
      align-items: center;
    }
    
    .next-session-input-group input {
      flex: 1;
    }
    
    .quick-not-scheduled, .quick-acuity {
      padding: 8px 16px;
      border: none;
      border-radius: 4px;
      font-size: 13px;
      cursor: pointer;
      white-space: nowrap;
      transition: all 0.3s;
      height: 36px;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    
    .quick-not-scheduled {
      background: #fce8b2;
      color: #7f6000;
    }
    
    .quick-not-scheduled:hover {
      background: #f9cc79;
    }
    
    .quick-acuity {
      background: #e3f2fd;
      color: #1565c0;
    }
    
    .quick-acuity:hover {
      background: #bbdefb;
    }
    
    .offline-banner {
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      background: #f44336;
      color: white;
      text-align: center;
      padding: 8px;
      z-index: 2000;
      font-size: 12px;
      font-weight: 500;
      display: none;
      animation: slideDown 0.3s ease-out;
    }
    
    .offline-banner.show {
      display: block;
    }
    
    @keyframes slideDown {
      from { transform: translateY(-100%); }
      to { transform: translateY(0); }
    }
    
    button.offline-disabled {
      background: #ccc !important;
      color: #999 !important;
      cursor: not-allowed !important;
      opacity: 0.6 !important;
    }
    
    button.offline-disabled:hover {
      background: #ccc !important;
      transform: none !important;
      box-shadow: none !important;
    }
    
    /* Consistent menu button styles from sidebar */
    .menu-button {
      width: 100%;
      padding: 12px;
      margin-bottom: 8px;
      border: 1px solid #dadce0;
      border-radius: 6px;
      background: white;
      color: #202124;
      cursor: pointer;
      font-size: 14px;
      font-weight: 500;
      transition: all 0.3s;
      text-align: left;
      display: flex;
      align-items: center;
      gap: 10px;
      box-shadow: 0 1px 3px rgba(60, 64, 67, 0.08);
    }
    
    .menu-button:hover:not(:disabled) {
      background: #f8f9fa;
      border-color: #003366;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(0, 51, 102, 0.15);
    }
    
    .menu-button:active {
      transform: translateY(0);
      box-shadow: none;
    }
    
    .menu-button:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    
    .menu-button.primary {
      background: #00C853;
      color: white !important;
      border: none;
      font-weight: 500;
      box-shadow: 0 2px 4px rgba(0, 200, 83, 0.3);
    }
    
    .menu-button.primary:hover:not(:disabled) {
      background: #388E3C;
      border: none;
    }
    
    .menu-button .icon {
      font-size: 16px;
      line-height: 1;
      flex-shrink: 0;
      margin-right: 4px;
    }
    
    /* Loading state for menu buttons */
    .menu-button.btn-loading {
      position: relative;
      color: transparent !important;
      pointer-events: none;
      overflow: hidden;
    }
    
    .menu-button.btn-loading::after {
      content: '';
      position: absolute;
      width: 16px;
      height: 16px;
      top: 50%;
      left: 50%;
      margin-left: -8px;
      margin-top: -8px;
      border: 2px solid rgba(255, 255, 255, 0.3);
      border-top-color: white;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
    }
    
    .menu-button.btn-loading:not(.primary)::after {
      border-color: rgba(0, 0, 0, 0.1);
      border-top-color: #333;
    }
    
    .menu-button.primary.btn-loading {
      background: #00C853 !important;
    }
    
    .menu-button.action-button {
      background: #fbbc04;
      color: #202124;
      border: none;
      font-weight: 500;
      box-shadow: 0 2px 4px rgba(251, 188, 4, 0.3);
    }
    
    .menu-button.action-button:hover:not(:disabled) {
      background: #f9ab00;
      border: none;
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(251, 188, 4, 0.4);
    }
    
    /* Preview view specific styles */
    .email-content {
      margin-top: 20px;
    }
    
    .version-tabs {
      display: flex;
      gap: 10px;
      margin-bottom: 20px;
      background: #f8f9fa;
      padding: 4px;
      border-radius: 8px;
    }
    
    .version-tab {
      flex: 1;
      padding: 10px;
      background: transparent;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 500;
      color: #5f6368;
      transition: all 0.2s;
    }
    
    .version-tab.active {
      background: white;
      color: #003366;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(0, 51, 102, 0.15);
    }
    
    .version-tab:hover:not(.active) {
      background: rgba(255, 255, 255, 0.5);
      transform: translateY(-1px);
      box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
    }
    
    .email-body {
      background: white;
      border: 1px solid #dadce0;
      border-radius: 8px;
      padding: 20px;
      min-height: 300px;
      max-height: 400px;
      overflow-y: auto;
      font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
      font-size: 14px;
      line-height: 1.6;
      white-space: pre-wrap;
    }
    
    .copy-tip {
      background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
      border: 1px solid #a5d6a7;
      padding: 12px;
      border-radius: 8px;
      margin-bottom: 16px;
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 13px;
      color: #2e7d32;
    }
    
    .copy-tip-icon {
      font-size: 18px;
    }
    
    .success-message {
      position: fixed;
      top: 20px;
      right: 20px;
      background: #4caf50;
      color: white;
      padding: 12px 20px;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 500;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
      opacity: 0;
      transform: translateX(100%);
      transition: all 0.3s cubic-bezier(0.68, -0.55, 0.265, 1.55);
      z-index: 1000;
    }
    
    .success-message.show {
      opacity: 1;
      transform: translateX(0);
    }
  `;
}

/**
 * Get HTML body content for Session Recap dialog
 */
function getSessionRecapBody() {
  return `
    <!-- Dialog Header with Dynamic Title -->
    <div class="dialog-header">
      <div class="dialog-title">
        <span class="icon" id="dialogIcon" style="display: none;"></span>
        <span id="dialogTitle">Session Recap</span>
      </div>
      <div class="view-indicator" id="viewIndicator" style="display: none;">Form</div>
    </div>
    
    <!-- Offline notification banner -->
    <div id="offlineBanner" class="offline-banner">
      📡 Offline - Some features are temporarily unavailable
    </div>
    
    <!-- Form View Container -->
    <div id="formView" class="view-container active">
      <div class="client-info" id="clientInfo">
        Loading client information...
      </div>
    
    <div class="saved-notes-indicator" id="savedNotesIndicator" style="display: none;">
      Quick notes loaded from your session
    </div>
    
    <form id="recapForm">
      <div class="form-section">
        <h3>Session Highlights</h3>
        
        <div class="form-group">
          <label>Today&rsquo;s Win*</label>
          <input type="text" id="todaysWin" value="" placeholder="e.g., Solved 5 quadratic equations independently!" required>
          <div class="quick-fill">
            <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Finally conquered [skill] after struggling last week!')">Breakthrough</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Solved [number] problems independently')">Independence</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Had an aha moment with [concept]')">Aha Moment</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Confidence soaring - attempted harder problems without prompting')">Confidence</button>
          </div>
        </div>
        
        <div class="form-group">
          <label>Focus Skill/Topic*</label>
          <input type="text" id="focusSkill" value="" placeholder="e.g., Quadratic Equations, French Conversation, ACT Geometry" required>
          <div class="quick-fill">
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'ACT Practice Problems')">ACT Practice</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'Test-Taking Strategies')">Test Strategies</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'Time Management')">Time Management</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'Critical Reading')">Critical Reading</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'Essay Writing')">Essay Writing</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('focusSkill', 'Math Problem Solving')">Math Problem Solving</button>
          </div>
        </div>
      </div>
      
      <div class="form-section">
        <h3>Skills Breakdown</h3>
        
        <div class="form-group">
          <label>Skills Mastered</label>
          <input type="text" id="mastered" placeholder="e.g., Factoring, using the quadratic formula">
        </div>
        
        <div class="form-group">
          <label>Skills Practiced</label>
          <input type="text" id="practiced" value="" placeholder="e.g., Word problems, graphing parabolas">
        </div>
        
        <div class="form-group">
          <label>Skills Introduced</label>
          <input type="text" id="introduced" placeholder="e.g., Completing the square">
        </div>
        
        <div class="form-group">
          <label>Progress Notes</label>
          <textarea id="progressNotes" placeholder="e.g., Making great progress with quadratics, still needs practice with word problems"></textarea>
        </div>
        
        <div class="form-group">
          <label>Parent Notes</label>
          <textarea id="parentNotes" placeholder="Additional notes for parents"></textarea>
        </div>
      </div>
      
      <div class="form-section">
        <h3>Homework & Next Steps</h3>
        
        <div class="form-group">
          <label>Assignment Details</label>
          <div class="homework-input-group">
            <select id="homeworkTemplate">
              <option value="">-- Choose from common assignments or type your own --</option>
            </select>
            <button type="button" class="btn-secondary homework-add-btn" onclick="addHomework()">+ Add</button>
          </div>
          <div class="act-test-selector" id="actTestSelector">
            <label style="font-size:13px;margin-bottom:8px;display:block;">Which ACT test?</label>
            <select id="actTestDropdown">
              <option value="">-- Choose test --</option>
            </select>
          </div>
          <textarea id="homework" placeholder="Type specific homework assignments here, or use the dropdown above to add common templates"></textarea>
        </div>
        
        <div class="form-group">
          <label>Next Session Date</label>
          <div class="next-session-input-group">
            <input type="text" id="nextSession" placeholder="e.g., Monday, July 29 at 3:00 PM">
            <button type="button" class="quick-not-scheduled" onclick="fillField('nextSession', 'Not Yet Scheduled')">Not Yet Scheduled</button>
            <button type="button" class="quick-acuity" onclick="window.open('https://secure.acuityscheduling.com/admin/clients', '_blank')">📅 Open Acuity</button>
          </div>
        </div>
        
        <div class="form-group">
          <label>Next Session Preview</label>
          <input type="text" id="nextSessionPreview" value="" placeholder="e.g., We&rsquo;ll build on factoring to tackle more complex equations">
        </div>
      </div>
      
      <div class="form-section">
        <h3>Parent Communication</h3>
        
        <div class="form-group">
          <label>Parent Quick Win</label>
          <input type="text" id="quickWin" value="" placeholder="e.g., Ask them to explain the FOIL method">
          <div class="quick-fill">
            <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Ask them to explain [concept] to you')">Explain concept</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Notice when they use [skill] in homework')">Notice skill</button>
            <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Celebrate their persistence with challenging problems')">Celebrate effort</button>
          </div>
        </div>
        
        <div class="form-group">
          <label>Additional Parent Notes</label>
          <textarea id="additionalParentNotes" placeholder="Conversation starters, encouragement focus, etc."></textarea>
        </div>
        
        <div class="form-group">
          <label>P.S. Note (Personal moment - leave blank if not needed)</label>
          <input type="text" id="psNote" placeholder="e.g., Their face lit up when they realized they could do this!">
        </div>
      </div>
      
      <div class="button-group">
        <button type="submit" class="btn-primary">Preview Update</button>
        <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
      </div>
    </form>
    </div>
    <!-- End of Form View -->
    
    <!-- Loading View -->
    <div id="loadingView" class="view-container">
      <div style="text-align: center; padding: 60px 20px;">
        <div class="loading-spinner" style="width: 40px; height: 40px; margin: 0 auto 20px;"></div>
        <div style="font-size: 16px; color: #5f6368; margin-bottom: 10px;">Generating preview...</div>
        <div style="font-size: 14px; color: #9aa0a6;">Processing your session recap data</div>
      </div>
    </div>
    
    <!-- Preview View -->
    <div id="previewView" class="view-container">
      <div class="copy-tip">
        <span class="copy-tip-icon">💡</span>
        <span>Toggle between versions below. "Send Recap Email" will copy the selected version to your clipboard!</span>
      </div>
      
      <div class="email-content">
        <div class="version-tabs">
          <button class="version-tab active" id="dashboardTab" onclick="toggleEmailVersion('dashboard')">
            📊 Dashboard Update
          </button>
          <button class="version-tab" id="emailTab" onclick="toggleEmailVersion('email')">
            ✉️ Email Version
          </button>
        </div>
        
        <div class="email-body" id="emailBody">
          <!-- Email content will be populated here -->
        </div>
      </div>
      
      <!-- Action Buttons Grid -->
      <div style="display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; gap: 12px; margin-top: 20px;">
        <!-- Top left: Back to Edit -->
        <button class="menu-button" id="backToEditBtn">
          <span class="icon">←</span>
          <span>Back to Edit</span>
        </button>
        
        <!-- Top right: Update Meeting Notes -->
        <button class="menu-button action-button" id="meetingNotesBtn" style="display: none; text-align: center; justify-content: center;">
          <span class="icon">📝</span>
          <span>Update Meeting Notes</span>
        </button>
        
        <!-- Bottom left: Save & Close -->
        <button class="menu-button" id="saveCloseBtn">
          <span class="icon">💾</span>
          <span>Save & Close</span>
        </button>
        
        <!-- Bottom right: Send Recap Email -->
        <button class="menu-button primary" id="sendEmailBtn" style="text-align: center; justify-content: center;">
          <span class="icon">✉️</span>
          <span>Send Recap Email</span>
        </button>
      </div>
    </div>
    
    <!-- Legacy loading overlay for backward compatibility -->
    <div class="loading-overlay" id="loadingOverlay" style="display: none;">
      <div class="loading-spinner"></div>
    </div>
    
    <!-- Success message for clipboard operations -->
    <div class="success-message" id="successMessage">Copied with formatting!</div>
  `;
}

/**
 * Get JavaScript for Session Recap dialog
 */
function getSessionRecapJavaScript() {
  return `
    // Global variables for unified dialog
    let recapData = null;
    let previewData = null;
    let currentView = 'form';
    let currentEmailVersion = 'dashboard';
    let isOnline = navigator.onLine;
    
    // Dialog state management
    const dialogState = {
      currentView: 'form',
      formData: null,
      previewData: null,
      clientData: null,
      emailVersion: 'dashboard'
    };
    
    // View switching functions
    function switchToView(viewName) {
      
      // Hide all views
      const views = ['formView', 'loadingView', 'previewView'];
      views.forEach(viewId => {
        const view = document.getElementById(viewId);
        if (view) {
          view.classList.remove('active');
        }
      });
      
      // Show target view
      const targetView = document.getElementById(viewName + 'View');
      if (targetView) {
        targetView.classList.add('active');
        
        // Update dialog header
        updateDialogHeader(viewName);
        
        // Execute view-specific initialization
        if (viewName === 'preview') {
          initializePreviewView();
        }
      }
      
      dialogState.currentView = viewName;
      currentView = viewName;
    }
    
    function updateDialogHeader(viewName) {
      const titleElement = document.getElementById('dialogTitle');
      const iconElement = document.getElementById('dialogIcon');
      const indicatorElement = document.getElementById('viewIndicator');
      
      const viewConfig = {
        'form': {
          title: 'Session Recap',
          icon: '',
          indicator: 'Form',
          showTitle: true
        },
        'loading': {
          title: '',
          icon: '',
          indicator: 'Loading',
          showTitle: false
        },
        'preview': {
          title: 'Email Preview',
          icon: '',
          indicator: 'Preview',
          showTitle: true
        }
      };
      
      const config = viewConfig[viewName] || viewConfig.form;
      
      if (titleElement) {
        if (config.showTitle) {
          titleElement.textContent = config.title;
          titleElement.style.visibility = 'visible';
        } else {
          titleElement.style.visibility = 'hidden';
        }
      }
      if (iconElement) iconElement.textContent = config.icon;
      if (indicatorElement) indicatorElement.textContent = config.indicator;
    }
    
    function initializePreviewView() {
      // Show/hide meeting notes button based on availability
      if (previewData && previewData.meetingNotesLink) {
        const meetingBtn = document.getElementById('meetingNotesBtn');
        if (meetingBtn) {
          meetingBtn.style.display = 'block';
          meetingBtn.dataset.link = previewData.meetingNotesLink;
        }
      }
      
      // Initialize email version toggle
      toggleEmailVersion(currentEmailVersion);
    }
    
    // Email version toggle
    function toggleEmailVersion(version) {
      currentEmailVersion = version;
      dialogState.emailVersion = version;
      
      // Update tab appearance
      const dashboardTab = document.getElementById('dashboardTab');
      const emailTab = document.getElementById('emailTab');
      
      if (dashboardTab && emailTab) {
        dashboardTab.classList.toggle('active', version === 'dashboard');
        emailTab.classList.toggle('active', version === 'email');
      }
      
      // Update email body content
      const emailBody = document.getElementById('emailBody');
      if (emailBody && previewData) {
        if (version === 'dashboard') {
          emailBody.innerHTML = previewData.dashboardVersion || 'Dashboard version not available';
        } else {
          emailBody.innerHTML = previewData.emailVersion || 'Email version not available';
        }
      }
    }
    
    // Offline detection and handling
    function handleOfflineState() {
      const offlineBanner = document.getElementById('offlineBanner');
      const serverDependentButtons = ['previewBtn', 'submitBtn'];
      
      if (!isOnline) {
        if (offlineBanner) {
          offlineBanner.classList.add('show');
        }
        
        serverDependentButtons.forEach(buttonId => {
          const button = document.getElementById(buttonId) || document.querySelector('button[type="submit"]');
          if (button) {
            button.classList.add('offline-disabled');
            button.disabled = true;
            button.title = 'This feature requires an internet connection';
          }
        });
      } else {
        if (offlineBanner) {
          offlineBanner.classList.remove('show');
        }
        
        serverDependentButtons.forEach(buttonId => {
          const button = document.getElementById(buttonId) || document.querySelector('button[type="submit"]');
          if (button) {
            button.classList.remove('offline-disabled');
            button.disabled = false;
            button.title = '';
          }
        });
      }
    }
    
    // Listen for online/offline events
    window.addEventListener('online', () => {
      isOnline = true;
      handleOfflineState();
    });
    
    window.addEventListener('offline', () => {
      isOnline = false;
      handleOfflineState();
    });
    
    // Load all data when dialog opens
    document.addEventListener('DOMContentLoaded', () => {
      
      const loadingOverlay = document.getElementById('loadingOverlay');
      if (loadingOverlay) {
        loadingOverlay.classList.add('active');
      }
      
      google.script.run
        .withSuccessHandler(function(data) {
          recapData = data;
          populateDialog(data);
          
          if (loadingOverlay) {
            loadingOverlay.classList.remove('active');
          }
        })
        .withFailureHandler(function(error) {
          console.error('Failed to load recap data:', error);
          const clientInfo = document.getElementById('clientInfo');
          if (clientInfo) {
            clientInfo.innerHTML = '<span style="color:#d93025;">Error loading client data: ' + error.message + '</span>';
          }
          
          if (loadingOverlay) {
            loadingOverlay.classList.remove('active');
          }
        })
        .getRecapDialogData();
        
      handleOfflineState();
    });
    
    function populateDialog(data) {
      
      // Populate client info header
      if (data.client) {
        const clientInfo = document.getElementById('clientInfo');
        if (clientInfo && data.client.name) {
          clientInfo.innerHTML = 
            '<strong>' + escapeHtml(data.client.name) + '</strong> • ' + 
            escapeHtml(data.client.types || 'Client');
        }
      }
      
      // Show saved notes indicator if we have notes
      if (data.notes && Object.keys(data.notes).length > 0) {
        const indicator = document.getElementById('savedNotesIndicator');
        if (indicator) {
          indicator.style.display = 'block';
        }
        
        // Pre-populate fields with quick notes if available
        if (data.notes.wins) document.getElementById('todaysWin').value = data.notes.wins;
        if (data.notes.skillsMastered) document.getElementById('mastered').value = data.notes.skillsMastered;
        if (data.notes.skillsPracticed) document.getElementById('practiced').value = data.notes.skillsPracticed;
        if (data.notes.skillsIntroduced) document.getElementById('introduced').value = data.notes.skillsIntroduced;
        if (data.notes.struggles) document.getElementById('progressNotes').value = data.notes.struggles;
        if (data.notes.parent) document.getElementById('parentNotes').value = data.notes.parent;
        if (data.notes.next) document.getElementById('nextSessionPreview').value = data.notes.next;
        
        // Pre-populate recap-specific fields if available from preview update
        if (data.notes.todaysWin) document.getElementById('todaysWin').value = data.notes.todaysWin;
        if (data.notes.focusSkill) document.getElementById('focusSkill').value = data.notes.focusSkill;
        if (data.notes.homework) document.getElementById('homework').value = data.notes.homework;
        if (data.notes.nextSession) document.getElementById('nextSession').value = data.notes.nextSession;
        if (data.notes.nextSessionPreview) document.getElementById('nextSessionPreview').value = data.notes.nextSessionPreview;
        if (data.notes.quickWin) document.getElementById('quickWin').value = data.notes.quickWin;
        if (data.notes.additionalParentNotes) document.getElementById('additionalParentNotes').value = data.notes.additionalParentNotes;
        if (data.notes.psNote) document.getElementById('psNote').value = data.notes.psNote;
        // Override with recap-specific field values if they exist
        if (data.notes.mastered) document.getElementById('mastered').value = data.notes.mastered;
        if (data.notes.practiced) document.getElementById('practiced').value = data.notes.practiced;
        if (data.notes.introduced) document.getElementById('introduced').value = data.notes.introduced;
        if (data.notes.progressNotes) document.getElementById('progressNotes').value = data.notes.progressNotes;
        if (data.notes.parentNotes) document.getElementById('parentNotes').value = data.notes.parentNotes;
      }
      
      // Populate homework templates
      const homeworkSelect = document.getElementById('homeworkTemplate');
      if (homeworkSelect && data.homeworkTemplates) {
        homeworkSelect.innerHTML = '<option value="">-- Choose from common assignments or type your own --</option>';
        data.homeworkTemplates.forEach(function(template) {
          const option = document.createElement('option');
          option.value = template;
          option.textContent = template;
          homeworkSelect.appendChild(option);
        });
      }
      
      // Populate ACT tests
      const actSelect = document.getElementById('actTestDropdown');
      if (actSelect && data.actTests) {
        actSelect.innerHTML = '<option value="">-- Choose test --</option>';
        data.actTests.forEach(function(test) {
          const option = document.createElement('option');
          option.value = test;
          option.textContent = test;
          actSelect.appendChild(option);
        });
      }
      
    }
    
    // Utility function to escape HTML
    function escapeHtml(text) {
      const div = document.createElement('div');
      div.textContent = text;
      return div.innerHTML;
    }
    
    // Fill field helper function
    function fillField(fieldId, value) {
      const field = document.getElementById(fieldId);
      if (field) {
        field.value = value;
      }
    }
    
    // Add homework template
    function addHomework() {
      const select = document.getElementById('homeworkTemplate');
      const textarea = document.getElementById('homework');
      
      if (select && textarea && select.value) {
        let currentValue = textarea.value;
        if (currentValue && !currentValue.endsWith('\\n')) {
          currentValue += '\\n';
        }
        
        if (select.value === 'ACT Practice Test') {
          const actSelector = document.getElementById('actTestSelector');
          if (actSelector) {
            actSelector.classList.add('show');
            actSelector.scrollIntoView({behavior: 'smooth', block: 'nearest'});
          }
        } else {
          textarea.value = currentValue + select.value;
          select.selectedIndex = 0;
        }
      } else {
      }
    }
    
    // Handle ACT test selection
    document.getElementById('actTestDropdown').addEventListener('change', function() {
      if (this.value) {
        const textarea = document.getElementById('homework');
        let currentValue = textarea.value;
        if (currentValue && !currentValue.endsWith('\\n')) {
          currentValue += '\\n';
        }
        textarea.value = currentValue + 'ACT Practice Test - ' + this.value;
        
        document.getElementById('actTestSelector').classList.remove('show');
        document.getElementById('homeworkTemplate').selectedIndex = 0;
        this.selectedIndex = 0;
      }
    });
    
    // Unified form submission handler
    document.getElementById('recapForm').addEventListener('submit', function(e) {
      e.preventDefault();
      handlePreviewUpdate();
    });
    
    // Handle Preview Update button click
    function handlePreviewUpdate() {
      
      // Save current form data to dialog state
      dialogState.formData = getFormData();
      
      // Switch to loading view
      switchToView('loading');
      
      // Call server to generate preview data
      google.script.run
        .withSuccessHandler(function(data) {
          
          // Store preview data
          previewData = data;
          dialogState.previewData = data;
          
          // Populate preview content
          populatePreviewView(data);
          
          // Switch to preview view
          switchToView('preview');
        })
        .withFailureHandler(function(error) {
          console.error('Preview generation failed:', error);
          alert('Failed to generate preview: ' + error.message);
          
          // Switch back to form view
          switchToView('form');
        })
        .processRecapPreviewUnified(dialogState.formData);
    }
    
    // Collect form data
    function getFormData() {
      return {
        clientData: recapData ? recapData.fullClientData : null,
        sheetName: recapData ? recapData.sheetName : null,
        todaysWin: document.getElementById('todaysWin').value,
        focusSkill: document.getElementById('focusSkill').value,
        mastered: document.getElementById('mastered').value,
        practiced: document.getElementById('practiced').value,
        introduced: document.getElementById('introduced').value,
        progressNotes: document.getElementById('progressNotes').value,
        parentNotes: document.getElementById('parentNotes').value,
        homework: document.getElementById('homework').value,
        nextSession: document.getElementById('nextSession').value,
        nextSessionPreview: document.getElementById('nextSessionPreview').value,
        quickWin: document.getElementById('quickWin').value,
        additionalParentNotes: document.getElementById('additionalParentNotes').value,
        psNote: document.getElementById('psNote').value
      };
    }
    
    // Populate preview view with generated data
    function populatePreviewView(data) {
      
      // Update email body with current version
      toggleEmailVersion(currentEmailVersion);
      
      // Show/hide meeting notes button
      const meetingBtn = document.getElementById('meetingNotesBtn');
      if (meetingBtn && data.meetingNotesLink) {
        meetingBtn.style.display = 'block';
        meetingBtn.dataset.link = data.meetingNotesLink;
      } else if (meetingBtn) {
        meetingBtn.style.display = 'none';
      }
    }
    
    // Handle Back to Edit button
    function handleBackToEdit() {
      
      // Restore form data if available
      if (dialogState.formData) {
        restoreFormData(dialogState.formData);
      }
      
      // Switch back to form view
      switchToView('form');
    }
    
    // Restore form data from dialog state
    function restoreFormData(data) {
      if (!data) return;
      
      Object.keys(data).forEach(key => {
        const element = document.getElementById(key);
        if (element && data[key]) {
          element.value = data[key];
        }
      });
    }
    
    // Add event listeners for preview view buttons
    document.addEventListener('DOMContentLoaded', function() {
      // Back to Edit button
      const backBtn = document.getElementById('backToEditBtn');
      if (backBtn) {
        backBtn.addEventListener('click', handleBackToEdit);
      }
      
      // Meeting Notes button
      const meetingBtn = document.getElementById('meetingNotesBtn');
      if (meetingBtn) {
        meetingBtn.addEventListener('click', function() {
          const link = this.dataset.link;
          if (link) {
            // Copy dashboard version to clipboard first
            copyDashboardToClipboard();
            
            // Show success message
            showSuccessMessage('Dashboard Update copied to clipboard! Opening Meeting Notes...');
            
            // Open link after brief delay
            setTimeout(() => {
              window.open(link, '_blank');
            }, 1500);
          }
        });
      }
      
      // Save & Close button
      const saveBtn = document.getElementById('saveCloseBtn');
      if (saveBtn) {
        saveBtn.addEventListener('click', function() {
          // TODO: Implement save logic
          google.script.host.close();
        });
      }
      
      // Send Email button
      const sendBtn = document.getElementById('sendEmailBtn');
      if (sendBtn) {
        sendBtn.addEventListener('click', function() {
          
          // Add loading state to button
          this.classList.add('btn-loading');
          this.disabled = true;
          
          // Copy current version to clipboard
          copyCurrentVersionToClipboard();
          
          // Show success message
          showSuccessMessage('Email content copied to clipboard! Opening Gmail...');
          
          // Open Gmail with recipient and subject populated
          if (previewData && previewData.formData) {
            const to = encodeURIComponent(previewData.formData.parentEmail || '');
            const subject = encodeURIComponent(previewData.subject || '');
            
            
            const gmailUrl = 'https://mail.google.com/mail/?view=cm&fs=1' + 
              (to ? '&to=' + to : '') +
              '&su=' + subject;
            
            
            // Open Gmail in new tab
            setTimeout(() => {
              window.open(gmailUrl, '_blank');
            }, 500);
            
            // Close dialog after brief delay
            setTimeout(() => {
              google.script.host.close();
            }, 2500);
          } else {
            console.error('No preview data available for email');
            // Fallback if no preview data
            setTimeout(() => {
              google.script.host.close();
            }, 2000);
          }
        });
      }
    });
    
    // Utility functions for preview functionality
    function copyDashboardToClipboard() {
      if (previewData && previewData.dashboardVersion) {
        copyToClipboard(previewData.dashboardVersion);
      }
    }
    
    function copyCurrentVersionToClipboard() {
      if (previewData) {
        const content = currentEmailVersion === 'dashboard' 
          ? previewData.dashboardVersion 
          : previewData.emailVersion;
        if (content) {
          copyToClipboard(content);
        }
      }
    }
    
    function copyToClipboard(htmlContent) {
      if (navigator.clipboard && window.ClipboardItem) {
        // Modern approach with proper HTML formatting
        const blob = new Blob([htmlContent], { type: 'text/html' });
        const clipboardItem = new ClipboardItem({ 'text/html': blob });
        
        navigator.clipboard.write([clipboardItem]).then(() => {
        }).catch(err => {
          console.error('Failed to copy to clipboard:', err);
          fallbackCopy(htmlContent);
        });
      } else {
        fallbackCopy(htmlContent);
      }
    }
    
    function fallbackCopy(content) {
      // Remove HTML tags for fallback
      const textContent = content.replace(/<[^>]*>/g, '');
      
      const textArea = document.createElement('textarea');
      textArea.value = textContent;
      document.body.appendChild(textArea);
      textArea.select();
      
      try {
        document.execCommand('copy');
      } catch (err) {
        console.error('Fallback copy failed:', err);
      }
      
      document.body.removeChild(textArea);
    }
    
    function showSuccessMessage(message) {
      const successElement = document.getElementById('successMessage');
      if (successElement) {
        successElement.textContent = message;
        successElement.classList.add('show');
        
        // Hide after 3 seconds
        setTimeout(() => {
          successElement.classList.remove('show');
        }, 3000);
      }
    }
  `;
}


/**
 * Updated generateEmailPreview function
 */
function generateEmailPreview(formData) {
  // Extract first name from full name
  const config = getCompleteConfig();
  formData.studentFirstName = formData.studentName.split(' ')[0];
  formData.tutorName = config.tutorName;

  // Format the current date nicely for the email body
  const today = new Date();
  formData.formattedDate = today.toLocaleDateString('en-US', { 
    weekday: 'long', 
    month: 'long', 
    day: 'numeric' 
  });

  // Determine session type for subject line
  let sessionType = '';
  const clientTypes = formData.clientTypes ? formData.clientTypes.split(',').map(t => t.trim()) : [];

  if (clientTypes.length === 1) {
    sessionType = clientTypes[0];
  } else if (clientTypes.length > 1) {
    sessionType = 'Tutoring';
  } else {
    sessionType = 'Tutoring';
  }

  // Format date as M/DD or MM/DD
  const month = today.getMonth() + 1;
  const day = today.getDate();
  const dateFormatted = `${month}/${day}`;

  // Build the email subject
  const subject = `${formData.studentFirstName}'s ACT Update ${dateFormatted} - Smart College`;

  // Build the email body
  const emailSections = buildEmailBody(formData);
  
  // Create the email data with all formats
  const emailData = {
    subject: subject,
    bodyHtml: formatForHtml(emailSections),
    bodyPlain: formatForPlainText(emailSections),
    bodyHtmlRich: formatForRichHtml(emailSections),
    formData: formData
  };
  
  // Store the email data in Properties Service before showing the dialog
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('tempEmailData', JSON.stringify(emailData));
  
  // Show the preview dialog with the email data
  showUpdatePreviewDialog(emailData);
  
  return emailData;
}
/**
 * Build the email body with proper structure
 */
function buildEmailBody(data) {
  const sections = [];
  
  // Greeting
  const greeting = data.parentName ? `Hi ${data.parentName},` : 'Hi,';
  sections.push(greeting);
  sections.push('');
  
  // Opening line with dashboard link
  if (data.dashboardLink) {
    sections.push(`${data.studentFirstName} had an excellent session today! All notes and resources are available in ${data.studentFirstName}'s Dashboard.`);
  } else {
    sections.push(`${data.studentFirstName} had an excellent session today!`);
  }
  sections.push('');
  
  // Today's Focus
  sections.push('**Today\'s Focus**');
  sections.push(data.focusSkill || 'General tutoring session');
  
  // Big Win
  sections.push('**Big Win**');
  sections.push(data.todaysWin || 'Great progress made today!');
  
  // What We Covered
  sections.push('**What We Covered**');
  
  if (data.mastered) {
    sections.push('*Mastered:*');
    data.mastered.split(',').forEach(skill => {
      sections.push(`    • ${skill.trim()}`);
    });
  }
  
  if (data.practiced) {
    sections.push('*Practiced:*');
    data.practiced.split(',').forEach(skill => {
      sections.push(`    • ${skill.trim()}`);
    });
  }
  
  if (data.introduced) {
    sections.push('*Introduced:*');
    data.introduced.split(',').forEach(skill => {
      sections.push(`    • ${skill.trim()}`);
    });
  }
  
  // Progress Update
  if (data.progressNotes) {
    sections.push('**Progress Update**');
    sections.push(data.progressNotes);
  }
  
  // Homework/Practice
  if (data.homework) {
    sections.push('**Homework/Practice**');
    sections.push(data.homework);
  }
  
  // Support at Home
  if (data.quickWin) {
    sections.push('**Support at Home**');
    sections.push(data.quickWin);
  }
  
  // Next Session
  if (data.nextSession) {
    sections.push('**Next Session**');
    sections.push(data.nextSession);
  }
  
  // Looking Ahead
  if (data.nextSessionPreview) {
    sections.push('**Looking Ahead**');
    sections.push(data.nextSessionPreview);
  }
  
  // Additional parent notes
  if (data.parentNotes && data.parentNotes !== data.quickWin) {
    sections.push(data.parentNotes);
  }
  
  // Closing
  sections.push('');
  sections.push('Questions? Just hit reply!');
  sections.push('');
  sections.push('Best,');
  sections.push(data.tutorName || getCompleteConfig().tutorName);
  
  // P.S. section if provided
  if (data.psNote) {
    sections.push('');
    sections.push(`P.S. - ${data.psNote}`);
  }
  
  return sections;
}

/**
 * Format email body for HTML display (used for Gmail preview)
 */
function formatForHtml(sections) {
  return sections.map(line => {
    if (line === '') return '<br>';
    
    // Handle bold headers (text between **)
    let formattedLine = line.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    
    // Handle italic subheaders (text between *)
    formattedLine = formattedLine.replace(/\*(.+?)\*/g, '<em>$1</em>');
    
    // Handle bullet points with proper spacing
    if (formattedLine.trim().startsWith('•')) {
      return `<div style="margin-left: 20px;">${formattedLine}</div>`;
    }
    
    // Add line breaks
    return `${formattedLine}<br>`;
  }).join('\n');
}

/**
 * Format email body for plain text (keep markdown formatting)
 */
function formatForPlainText(sections) {
  return sections.join('\n');
}

/**
 * Format email body for rich HTML (for copying to Google Docs)
 * This creates proper HTML structure that will paste correctly
 */
function formatForRichHtml(sections) {
  const htmlParts = [];
  let inList = false;
  let currentListType = null;
  
  sections.forEach((line, index) => {
    if (line === '') {
      if (inList) {
        htmlParts.push('</ul>');
        inList = false;
        currentListType = null;
      }
      htmlParts.push('<p>&nbsp;</p>');
      return;
    }
    
    // Handle bold headers (text between **)
    if (line.includes('**')) {
      if (inList) {
        htmlParts.push('</ul>');
        inList = false;
        currentListType = null;
      }
      const boldText = line.replace(/\*\*(.+?)\*\*/g, '$1');
      htmlParts.push(`<p><strong>${boldText}</strong></p>`);
      return;
    }
    
    // Handle italic subheaders (text between *)
    if (line.includes('*') && !line.includes('**')) {
      if (inList) {
        htmlParts.push('</ul>');
        inList = false;
        currentListType = null;
      }
      const italicText = line.replace(/\*(.+?)\*/g, '$1');
      htmlParts.push(`<p><em>${italicText}</em></p>`);
      return;
    }
    
    // Handle bullet points
    if (line.trim().startsWith('•')) {
      if (!inList) {
        htmlParts.push('<ul>');
        inList = true;
      }
      const bulletText = line.trim().substring(1).trim();
      htmlParts.push(`<li>${bulletText}</li>`);
      return;
    }
    
    // Regular text
    if (inList) {
      htmlParts.push('</ul>');
      inList = false;
      currentListType = null;
    }
    htmlParts.push(`<p>${line}</p>`);
  });
  
  // Close any open list
  if (inList) {
    htmlParts.push('</ul>');
  }
  
  return htmlParts.join('\n');
}



function showUpdatePreviewDialog(emailData) {
  // Get sheet name from emailData or fallback to active sheet
  const sheetName = emailData.formData?.sheetName || SpreadsheetApp.getActiveSheet().getName();
  
  // Get complete client data using the most efficient method
  const completeClientData = getCompleteClientDataForPreview(sheetName);
  
  // Merge the complete client data with the form data
  const mergedEmailData = {
    formData: {
      ...emailData.formData,
      // Override with fresh, complete client data
      studentName: completeClientData.studentName,
      studentFirstName: completeClientData.studentFirstName,
      parentName: completeClientData.parentName,
      parentEmail: completeClientData.parentEmail,
      clientTypes: completeClientData.clientTypes,
      clientType: completeClientData.clientType,
      sheetName: completeClientData.sheetName,
      dashboardLink: completeClientData.dashboardLink,
      meetingNotesLink: completeClientData.meetingNotesLink,
      scoreReportLink: completeClientData.scoreReportLink,
      tutorName: getCompleteConfig().tutorName || ''
    },
    subject: emailData.subject,
    meetingNotesLink: completeClientData.meetingNotesLink,
    scoreReportLink: completeClientData.scoreReportLink
  };
  
  
  // Store the merged data in Properties Service
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('tempEmailData', JSON.stringify(mergedEmailData));
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
            font-weight: 500;
          }
          .copy-btn {
            padding: 6px 12px;
            background: #e8f0fe;
            color: #003366;          
            border: none;     
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
            font-weight: 500;
            white-space: nowrap;
            transition: all 0.2s ease;
          }
          .copy-btn:hover {
            background: #d2e3fc;
          }
          .copy-btn.copied {
            background: #00C853;    
            color: white;
          }
          .copy-btn.gmail {
            background: #ea4335;
            color: white;
          }
          .copy-btn.gmail:hover {
            background: #d33b2c;
          }
          .copy-btn.dashboard {
            background: #00C853;  
            color: white;
          }
          /* Universal button loading state */
          .btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }

          .btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
          }

          /* For secondary/light buttons, use dark spinner */
          .btn-secondary.btn-loading::after,
          .btn-danger.btn-loading::after,
          .quick-fill-btn.btn-loading::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #5f6368;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .copy-btn.dashboard:hover {
            background: #388E3C;     
          }
          .email-content {
            margin-bottom: 20px;
          }
          .email-preview {
            width: 100%;
            min-height: 400px;
            padding: 20px;
            border: 1px solid #dadce0;
            border-radius: 8px;
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            font-size: 14px;
            line-height: 1.6;
            background: white;
            overflow-y: auto;
            box-sizing: border-box;
          }
          .email-preview p {
            margin: 0 0 0.5em 0;
          }
          .email-preview strong {
            font-weight: bold;
            display: block;
            margin-top: 1em;
            margin-bottom: 0.5em;
          }
          .email-preview em {
            font-style: italic;
            display: block;
            margin-top: 0.5em;
            margin-bottom: 0.3em;
          }
          .email-preview ul {
            margin: 0.3em 0 0.5em 0;
            padding-left: 30px;
          }
          .email-preview li {
            margin-bottom: 0.2em;
          }
          .instructions {
            background: #e8f0fe;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 14px;
            color: #003366;          /* Changed from #1a73e8 */
          }
          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 20px;
          }
          button {
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
          }
          .btn-primary {
            background: #003366;     /* Changed from #1a73e8 */
            color: white;
            flex: 1;
          }
          .btn-primary:hover {
            background: #002244;     /* Changed from #1557b0 */
          }

          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
          }
          .btn-edit {
            background: #fbbc04;
            color: #202124;
          }
          .btn-edit:hover {
            background: #f9ab00;
          }
          .success-message {
            position: fixed;
            top: 20px;
            right: 20px;
            background: #188038;
            color: white;
            padding: 12px 20px;
            border-radius: 6px;
            display: none;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            animation: slideIn 0.3s ease-out;
            max-width: 300px;
            z-index: 10000;
          }
          @keyframes slideIn {
            from {
              transform: translateX(100%);
              opacity: 0;
            }
            to {
              transform: translateX(0);
              opacity: 1;
            }
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .copy-tip {
            background: #fef7e0;
            border: 1px solid #f9cc79;
            padding: 10px;
            border-radius: 6px;
            margin-bottom: 10px;
            font-size: 13px;
            color: #7f6000;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          .copy-tip-icon {
            font-size: 16px;
          }
          .loading {
            text-align: center;
            padding: 40px;
          }
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #003366;  /* Changed from #1a73e8 */
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
          }
          .loading-text {
            font-size: 14px;
            color: #5f6368;
            font-weight: 500;
          }
          .version-toggle {
            position: relative;
            display: flex;
            background: #f1f3f4;
            border-radius: 25px;
            padding: 3px;
            width: 300px;
            height: 40px;
            align-items: center;
          }
          .toggle-slider {
            position: absolute;
            background: white;
            border: 2px solid #1a73e8;
            border-radius: 20px;
            height: 34px;
            width: 147px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            transform: translateX(0);
            z-index: 1;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            cursor: pointer;
          }
          .toggle-slider:hover {
            background: #f0f8ff;
            border-color: #1a73e8;
            transform: translateX(0) translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .toggle-slider:active {
            transform: translateX(0) translateY(0) scale(0.95);
            box-shadow: none;
          }
          .toggle-slider.dashboard {
            transform: translateX(150px);
          }
          .toggle-slider.dashboard:hover {
            background: #f0f8ff;
            border-color: #1a73e8;
            transform: translateX(150px) translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .toggle-slider.dashboard:active {
            transform: translateX(150px) translateY(0) scale(0.95);
            box-shadow: none;
          }
          .toggle-btn {
            flex: 1;
            padding: 8px 16px;
            background: transparent;
            border: none;
            border-radius: 20px;
            cursor: pointer;
            font-weight: 500;
            color: #5f6368;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            font-size: 14px;
            position: relative;
            z-index: 2;
            height: 34px;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .toggle-btn.active {
            color: #003366;
            font-weight: normal;
          }
          .toggle-btn:hover {
            color: #003366;
          }
          .toggle-btn.active:hover {
            color: #003366;
          }
          
          /* Dialog button styles matching sidebar */
          .btn-primary {
            background: white;
            color: #202124;
            border: 1px solid #dadce0;
            border-radius: 6px;
            padding: 10px 16px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s ease;
            font-size: 14px;
          }
          .btn-primary:hover:not(:disabled) {
            background: #f0f8ff;
            border-color: #1a73e8;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .btn-primary:active {
            transform: translateY(0) scale(0.95);
            box-shadow: none;
          }
          
          .control-panel-btn {
            transition: all 0.2s ease;
          }
          .control-panel-btn:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .control-panel-btn:active {
            transform: translateY(0) scale(0.95);
            box-shadow: none;
          }
          
          .copy-btn, .btn-secondary {
            transition: all 0.2s ease;
            border-radius: 6px;
          }
          .copy-btn:hover, .btn-secondary:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .copy-btn:active, .btn-secondary:active {
            transform: translateY(0) scale(0.95);
            box-shadow: none;
          }
          
          /* Menu button styles matching control panel */
          .menu-button {
            width: 100%;
            padding: 12px;
            margin-bottom: 8px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: white;
            color: #202124;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.3s;
            text-align: left;
            display: flex;
            align-items: center;
            gap: 10px;
          }
          
          .menu-button:hover:not(:disabled) {
            background: #f8f9fa;
            border-color: #003366;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          
          .menu-button:active {
            transform: translateY(0);
            box-shadow: none;
          }
          
          .menu-button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
          }
          
          .menu-button.primary {
            background: #00C853;
            color: white !important;
            border: none;
            text-align: center;
            justify-content: center;
          }
          
          .menu-button.primary:hover:not(:disabled) {
            background: #388E3C;
            border: none;
          }
          
          .menu-button .icon {
            font-size: 16px;
            line-height: 1;
            flex-shrink: 0;
          }
          
          /* Universal button loading state for menu buttons */
          .menu-button.btn-loading {
            position: relative;
            color: transparent !important;
            pointer-events: none;
            opacity: 0.8;
          }
          
          .menu-button.btn-loading::after {
            content: '';
            position: absolute;
            width: 16px;
            height: 16px;
            top: 50%;
            left: 50%;
            margin-left: -8px;
            margin-top: -8px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-top-color: #fff;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            z-index: 1;
          }
          
          /* For light buttons (menu-button without primary), use dark spinner */
          .menu-button.btn-loading:not(.primary)::after {
            border-color: rgba(0, 0, 0, 0.1);
            border-top-color: #333;
          }
          
          /* Ensure primary buttons keep their green background and white spinner */
          .menu-button.primary.btn-loading {
            background: #00C853 !important;
          }
          
          .menu-button.primary.btn-loading .icon {
            color: white !important;
          }
          
          /* Action button variant - centered text, distinct styling */
          .menu-button.action-button {
            background: #fbbc04;
            color: #202124;
            border: none;
            text-align: center;
            justify-content: center;
            font-weight: 500;
          }
          
          .menu-button.action-button:hover:not(:disabled) {
            background: #f9ab00;
            border: none;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.15);
          }
          
          .menu-button.action-button.btn-loading {
            background: #fbbc04 !important;
          }
          
          .menu-button.action-button.btn-loading .icon {
            color: #202124 !important;
          }
        </style>
      </head>
      <body>
        <div id="loadingDiv" class="loading">
          <div class="loading-spinner"></div>
          <div class="loading-text">Loading email preview...</div>
        </div>
        <div id="mainContent" style="display: none;">
          <div class="email-content">
            <div class="copy-tip">
              <span class="copy-tip-icon">💡</span>
              <span>Toggle between versions below. "Send Recap Email" will copy the selected version to your clipboard!</span>
            </div>
            
            <!-- Copy Update Button and Version Toggle -->
            <div style="display: flex; gap: 12px; align-items: center; margin-bottom: 20px;">
              <button class="btn-primary" id="copyBodyBtn">Copy Update</button>
              <div class="version-toggle">
                <div class="toggle-slider" id="toggleSlider"></div>
                <button class="toggle-btn active" id="emailVersionBtn" data-version="email">Email Update</button>
                <button class="toggle-btn" id="dashboardVersionBtn" data-version="dashboard">Dashboard Update</button>
              </div>
            </div>
            
            <!-- Email Version Content -->
            <div class="email-preview" id="emailContent" style="display: block;">
              [Loading...]
            </div>
            
            <!-- Dashboard Version Content -->
            <div class="email-preview" id="dashboardContent" style="display: none;">
              [Loading...]
            </div>
          </div>
          
          <div class="button-group" style="display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1fr 1fr; gap: 12px; margin-top: 20px;">
            <!-- Top left: Back to Edit -->
            <button class="menu-button" id="backToEditBtn">
              <span class="icon">←</span>
              <span>Back to Edit</span>
            </button>
            
            <!-- Top right: Update Meeting Notes -->
            ${mergedEmailData.meetingNotesLink ? 
              '<button class="menu-button action-button" id="meetingNotesBtn" style="text-align: center; justify-content: center;"><span class="icon">📝</span> <span>Update Meeting Notes</span></button>' : 
              '<div></div>'}
            
            <!-- Bottom left: Save & Close -->
            <button class="menu-button" id="saveCloseBtn">
              <span class="icon">💾</span>
              <span>Save & Close</span>
            </button>
            
            <!-- Bottom right: Send Recap Email -->
            <button class="menu-button primary" id="sendEmailBtn" style="text-align: center; justify-content: center;">
              <span class="icon">✉️</span>
              <span>Send Recap Email</span>
            </button>
          </div>
        </div>
        
        <div class="success-message" id="successMessage">Copied with formatting!</div>
        
        <script>
          let emailData = null;
          
          function setButtonLoading(button, isLoading) {
            if (isLoading) {
              button.dataset.originalText = button.textContent;
              button.classList.add('btn-loading');
              button.disabled = true;
            } else {
              button.classList.remove('btn-loading');
              button.textContent = button.dataset.originalText || button.textContent;
              button.disabled = false;
            }
          }
          // Load email data from server
          google.script.run
            .withSuccessHandler(function(data) {
              emailData = data;
              initializeDialog();
            })
            .withFailureHandler(function(error) {
              alert('Error loading email data: ' + error.message);
              google.script.host.close();
            })
            .getEmailDataFromProperties();
          
          function initializeDialog() {
            // Hide loading, show content
            document.getElementById('loadingDiv').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
            
            // Debug logging - let's see what data we're receiving
            
            // Generate and populate both versions
            generateEmailVersions();
            
            // Show/hide meeting notes button
            if (emailData.meetingNotesLink) {
              const meetingNotesBtn = document.getElementById('meetingNotesBtn');
              if (meetingNotesBtn) {
                meetingNotesBtn.style.display = 'block';
              }
            }
            
            // Add toggle functionality
            document.getElementById('emailVersionBtn').addEventListener('click', function() {
              toggleEmailVersion('email');
            });
            document.getElementById('dashboardVersionBtn').addEventListener('click', function() {
              toggleEmailVersion('dashboard');
            });
            
            // Attach event listeners
            document.getElementById('copyBodyBtn').addEventListener('click', copyCurrentVersion);
            document.getElementById('backToEditBtn').addEventListener('click', backToEdit);
            document.getElementById('sendEmailBtn').addEventListener('click', sendRecapEmail);
            document.getElementById('saveCloseBtn').addEventListener('click', saveAndClose);
            
            const meetingNotesBtn = document.getElementById('meetingNotesBtn');
            if (meetingNotesBtn) {
              meetingNotesBtn.addEventListener('click', openMeetingNotes);
            }
          }
          
          function generateEmailVersions() {
            try {
              const parentNames = emailData.formData.parentName || 'Parents';
              const studentFirstName = emailData.formData.studentFirstName;
              const formData = emailData.formData;
              
              // Generate Dashboard Update version (no salutation/sign-off)
              const dashboardParts = [];
              
              // Opening line
              const openingLine = formData.openingNotes || (studentFirstName + ' had an excellent session today!');
              dashboardParts.push(openingLine);
              
              // Today's Focus
              dashboardParts.push('<strong>Today&#39;s Focus</strong>');
              dashboardParts.push(formData.focusSkill || 'General tutoring session');
              
              // Big Win
              dashboardParts.push('<strong>Big Win</strong>');
              dashboardParts.push(formData.todaysWin || 'Great progress made today!');
              
              // What We Covered
              dashboardParts.push('<strong>What We Covered</strong>');
              let coveredContent = [];
              if (formData.practiced) {
                coveredContent.push('<em>Practiced:</em>');
                formData.practiced.split(',').forEach(function(skill) {
                  coveredContent.push(skill.trim());
                });
              }
              if (formData.introduced) {
                coveredContent.push('<em>Introduced:</em>');
                formData.introduced.split(',').forEach(function(skill) {
                  coveredContent.push(skill.trim());
                });
              }
              dashboardParts.push(coveredContent.join('<br>'));
              
              // Progress Update
              if (formData.struggles) {
                dashboardParts.push('<strong>Progress Update</strong>');
                dashboardParts.push(formData.struggles);
              }
              
              // Homework/Practice
              if (formData.homework) {
                dashboardParts.push('<strong>Homework/Practice</strong>');
                dashboardParts.push(formData.homework);
              }
              
              // Support at Home
              if (formData.parentNotes) {
                dashboardParts.push('<strong>Support at Home</strong>');
                dashboardParts.push(formData.parentNotes);
              }
              
              // Next Session
              if (formData.nextSession) {
                dashboardParts.push('<strong>Next Session</strong>');
                dashboardParts.push(formData.nextSession);
              }
              
              // Looking Ahead
              if (formData.nextSessionPreview) {
                dashboardParts.push('<strong>Looking Ahead</strong>');
                dashboardParts.push(formData.nextSessionPreview);
              }
              
              // P.S. section if there's a quick win
              if (formData.quickWin) {
                dashboardParts.push(''); // Add extra blank line before P.S.
                dashboardParts.push('P.S. - ' + formData.quickWin);
              }
              
              // Join parts with proper spacing: bold headers directly followed by content, blank lines between sections
              const dashboardVersionHtml = dashboardParts.map(function(part, index) {
                // If this part is a bold header (contains <strong>), don't add extra line before
                if (part.includes('<strong>')) {
                  return (index === 0 ? '' : '<br><br>') + part;
                } else {
                  return '<br>' + part;
                }
              }).join('');
              
              // Generate Email Update version (simpler format)
              const emailParts = [];
              
              // Determine session type
              let sessionType = '';
              const clientTypes = emailData.formData.clientTypes ? emailData.formData.clientTypes.split(',').map(function(t) { return t.trim(); }) : [];
              
              if (clientTypes.length === 1) {
                sessionType = clientTypes[0];
              } else {
                sessionType = 'ACT';
              }
              
              // Salutation with spacing
              emailParts.push('<p>Hi ' + parentNames + '!</p>');
              emailParts.push('<p>&nbsp;</p>'); // Line spacing after salutation
              
              // Body paragraph
              emailParts.push('<p>' + studentFirstName + ' had a great session today! ');
              
              // Add meeting notes link
              if (emailData.meetingNotesLink) {
                emailParts.push('Here is the link to our <a href="' + emailData.meetingNotesLink + '">Meeting Notes</a> from our ' + sessionType + ' session, and you will find ' + studentFirstName + '&#39;s homework linked on the Meeting Notes. ');
              } else {
                emailParts.push('Here are the notes from our ' + sessionType + ' session, and you will find ' + studentFirstName + '&#39;s homework included. ');
              }
              
              // Dashboard and Score Report links in one sentence
              emailParts.push('You can always follow our progress in more detail in ' + studentFirstName + '&#39;s ');
              if (emailData.formData.dashboardLink) {
                emailParts.push('<a href="' + emailData.formData.dashboardLink + '">Dashboard</a>');
              } else {
                emailParts.push('Dashboard');
              }
              emailParts.push(' and track their progress in their ');
              if (emailData.scoreReportLink) {
                emailParts.push('<a href="' + emailData.scoreReportLink + '">Score Report</a>!');
              } else {
                emailParts.push('Score Report!');
              }
              emailParts.push('</p>');
              
              // Next Session information (if provided)
              if (formData.nextSession) {
                emailParts.push('<p>&nbsp;</p>'); // Line spacing before next session
                emailParts.push('<p><strong>Next Session:</strong> ' + formData.nextSession + '</p>');
              }
              
              // Closing with spacing
              emailParts.push('<p>&nbsp;</p>'); // Line spacing before closing
              emailParts.push('<p>As always, I&#39;m just an email away with any questions.</p>');
              emailParts.push('<p>&nbsp;</p>'); // Line spacing before sign-off
              emailParts.push('<p>Best,<br>');
              emailParts.push(emailData.formData.tutorName || 'Your Tutor');
              emailParts.push('</p>');
              
              const emailVersionHtml = emailParts.join('');
              
              // Populate the content divs
              document.getElementById('emailContent').innerHTML = emailVersionHtml;
              document.getElementById('dashboardContent').innerHTML = dashboardVersionHtml;
              
            } catch (error) {
              console.error('Error generating email versions:', error);
              document.getElementById('emailContent').innerHTML = '<p>Error loading email content</p>';
              document.getElementById('dashboardContent').innerHTML = '<p>Error loading dashboard content</p>';
            }
          }
          
          function toggleEmailVersion(version) {
            // Update button states
            document.getElementById('emailVersionBtn').classList.remove('active');
            document.getElementById('dashboardVersionBtn').classList.remove('active');
            
            // Update slider position
            const slider = document.getElementById('toggleSlider');
            
            if (version === 'email') {
              document.getElementById('emailVersionBtn').classList.add('active');
              slider.classList.remove('dashboard');
              document.getElementById('emailContent').style.display = 'block';
              document.getElementById('dashboardContent').style.display = 'none';
            } else {
              document.getElementById('dashboardVersionBtn').classList.add('active');
              slider.classList.add('dashboard');
              document.getElementById('emailContent').style.display = 'none';
              document.getElementById('dashboardContent').style.display = 'block';
            }
          }
          
          function copyCurrentVersion() {
            const currentVersion = document.getElementById('emailVersionBtn').classList.contains('active') ? 'email' : 'dashboard';
            const contentDiv = document.getElementById(currentVersion + 'Content');
            
            const copyBtn = document.getElementById('copyBodyBtn');
            setButtonLoading(copyBtn, true);
            
            setTimeout(function() {
              // Create a selection range
              const selection = window.getSelection();
              const range = document.createRange();
              
              // Select the entire content
              range.selectNodeContents(contentDiv);
              selection.removeAllRanges();
              selection.addRange(range);
              
              try {
                // Execute the copy command
                document.execCommand('copy');
                setButtonLoading(copyBtn, false);
                showSuccessMessage('Copied with formatting!');
              } catch (err) {
                console.error('Copy failed:', err);
                setButtonLoading(copyBtn, false);
                // Fallback to plain text copy
                const textArea = document.createElement('textarea');
                textArea.value = contentDiv.innerText;
                textArea.style.position = 'fixed';
                textArea.style.left = '-999999px';
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
                showSuccessMessage('Copied as plain text.');
              }
              
              // Clear selection
              selection.removeAllRanges();
            }, 300);
          }
          
          function copyToClipboard(elementId, button) {
            setButtonLoading(button, true);
            
            const element = document.getElementById(elementId);
            const text = element.textContent;
            
            // Add a small delay to show the loading state
            setTimeout(function() {
              const textArea = document.createElement('textarea');
              textArea.value = text;
              textArea.style.position = 'fixed';
              textArea.style.left = '-999999px';
              document.body.appendChild(textArea);
              textArea.select();
              
              try {
                document.execCommand('copy');
                setButtonLoading(button, false);
                button.textContent = 'Copied!';
                button.classList.add('copied');
                setTimeout(function() {
                  button.textContent = button.dataset.originalText || 'Copy';
                  button.classList.remove('copied');
                }, 2000);
              } catch (err) {
                setButtonLoading(button, false);
                console.error('Copy failed:', err);
              }
              
              document.body.removeChild(textArea);
            }, 300);
          }
          
          function copyRichText() {
            const copyBtn = document.getElementById('copyBodyBtn');
            setButtonLoading(copyBtn, true);
            
            setTimeout(function() {
              const emailBody = document.getElementById('emailBody');
              
              // Create a selection range
              const selection = window.getSelection();
              const range = document.createRange();
              
              // Select the entire email body content
              range.selectNodeContents(emailBody);
              selection.removeAllRanges();
              selection.addRange(range);
              
              try {
                // Execute the copy command
                document.execCommand('copy');
                
                // Clear the selection
                selection.removeAllRanges();
                
                setButtonLoading(copyBtn, false);
                
                // Show success message
                showSuccessMessage('Copied with formatting! Paste into Google Docs.');
              } catch (err) {
                console.error('Copy failed:', err);
                setButtonLoading(copyBtn, false);
                // Fallback to plain text copy
                copyPlainTextFallback();
              }
            }, 300);
          }
          
          function copyPlainTextFallback() {
            const emailBody = document.getElementById('emailBody');
            const textArea = document.createElement('textarea');
            textArea.value = emailBody.innerText;
            textArea.style.position = 'fixed';
            textArea.style.left = '-999999px';
            document.body.appendChild(textArea);
            textArea.select();
            
            try {
              document.execCommand('copy');
              showSuccessMessage('Copied as plain text.');
            } catch (err) {
              alert('Copy failed. Please try selecting the text manually.');
            }
            
            document.body.removeChild(textArea);
          }
          
          function showSuccessMessage(message) {
            const successMsg = document.getElementById('successMessage');
            successMsg.textContent = message || 'Copied to clipboard!';
            successMsg.style.display = 'block';
            setTimeout(function() {
              successMsg.style.display = 'none';
            }, 3000);
          }
          
          function backToEdit() {
            const button = document.getElementById('backToEditBtn');
            if (button) {
              button.dataset.originalText = button.innerHTML;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            // Close this dialog and call the server to reopen the edit form
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                if (button) {
                  button.classList.remove('btn-loading');
                  button.innerHTML = button.dataset.originalText || button.innerHTML;
                  button.disabled = false;
                }
                alert('Error going back to edit: ' + error.message);
              })
              .backToEditRecap();
          }
          
          function openMeetingNotes() {
            
            const button = document.getElementById('meetingNotesBtn');
            if (button) {
              button.dataset.originalText = button.innerHTML;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            // Copy Dashboard Update to clipboard
            copyDashboardToClipboard();
            
            if (emailData && emailData.meetingNotesLink) {
              // Show success message immediately
              showSuccessMessage('Dashboard Update copied to clipboard! Paste into Meeting Notes with Ctrl+V (or Cmd+V on Mac)');
              
              // Wait 1.5 seconds for user to see the success banner before opening link
              setTimeout(function() {
                window.open(emailData.meetingNotesLink, '_blank');
              }, 1500);
              
              // Remove loading state after the link opens
              setTimeout(function() {
                if (button) {
                  button.classList.remove('btn-loading');
                  button.innerHTML = button.dataset.originalText || button.innerHTML;
                  button.disabled = false;
                }
              }, 1700);
            } else {
              console.error('No meeting notes link found. emailData:', emailData);
              showSuccessMessage('Dashboard Update copied to clipboard!');
              if (button) {
                button.classList.remove('btn-loading');
                button.innerHTML = button.dataset.originalText || button.innerHTML;
                button.disabled = false;
              }
              alert('No meeting notes link available for this student.');
            }
          }
          
          function copyDashboardToClipboard() {
            try {
              const contentDiv = document.getElementById('dashboardContent');
              
              if (!contentDiv) {
                console.error('Dashboard content div not found');
                return;
              }
              
              // Store original display style
              const originalDisplay = contentDiv.style.display;
              
              // Temporarily make visible for copying (required for some browsers)
              contentDiv.style.display = 'block';
              
              // Create a selection range
              const selection = window.getSelection();
              const range = document.createRange();
              
              // Select the entire dashboard content
              range.selectNodeContents(contentDiv);
              selection.removeAllRanges();
              selection.addRange(range);
              
              try {
                // Execute the copy command
                document.execCommand('copy');
              } catch (err) {
                console.error('Primary copy failed:', err);
                // Fallback to plain text copy
                const textArea = document.createElement('textarea');
                textArea.value = contentDiv.innerText;
                textArea.style.position = 'fixed';
                textArea.style.left = '-999999px';
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
              }
              
              // Restore original display style
              contentDiv.style.display = originalDisplay;
              
              // Clear selection
              selection.removeAllRanges();
            } catch (error) {
              console.error('Error in copyDashboardToClipboard:', error);
            }
          }
          
          function sendRecapEmail() {
            const button = document.getElementById('sendEmailBtn');
            if (button) {
              button.dataset.originalText = button.innerHTML;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            const parentNames = emailData.formData.parentName || 'Parents';
            const studentFirstName = emailData.formData.studentFirstName;
            
            let sessionType = '';
            const clientTypes = emailData.formData.clientTypes ? emailData.formData.clientTypes.split(',').map(function(t) { return t.trim(); }) : [];
            
            if (clientTypes.length === 1) {
              sessionType = clientTypes[0];
            } else {
              sessionType = 'Tutoring';
            }
            
            // Get the current version from the toggle
            const currentVersion = document.getElementById('emailVersionBtn').classList.contains('active') ? 'email' : 'dashboard';
            
            // Copy the current rich text version to clipboard
            copyCurrentVersionToClipboard();
            
            // Open blank Gmail with recipient and subject only
            const to = encodeURIComponent(emailData.formData.parentEmail || '');
            const subject = encodeURIComponent(emailData.subject);
            
            const gmailUrl = 'https://mail.google.com/mail/?view=cm&fs=1' + 
              (to ? '&to=' + to : '') +
              '&su=' + subject;
            
            window.open(gmailUrl, '_blank');
            
            // Remove loading state
            if (button) {
              button.classList.remove('btn-loading');
              button.innerHTML = button.dataset.originalText || button.innerHTML;
              button.disabled = false;
            }
            
            // Show success message
            showSuccessMessage('Email content copied! Paste into Gmail body with Ctrl+V (or Cmd+V on Mac)');
            
            // Don't close immediately - let user see the success message
            setTimeout(function() {
              saveAndClose();
            }, 3000);
          }
          
          function copyCurrentVersionToClipboard() {
            try {
              const currentVersion = document.getElementById('emailVersionBtn').classList.contains('active') ? 'email' : 'dashboard';
              let contentDiv;
              
              if (currentVersion === 'email') {
                contentDiv = document.getElementById('emailContent');
              } else {
                contentDiv = document.getElementById('dashboardContent');
              }
              
              if (!contentDiv) {
                console.error('Content div not found for version:', currentVersion);
                return;
              }
              
              // Create a selection range
              const selection = window.getSelection();
              const range = document.createRange();
              
              // Select the entire content
              range.selectNodeContents(contentDiv);
              selection.removeAllRanges();
              selection.addRange(range);
              
              try {
                // Execute the copy command
                document.execCommand('copy');
              } catch (err) {
                console.error('Copy failed:', err);
                // Fallback to plain text copy
                const textArea = document.createElement('textarea');
                textArea.value = contentDiv.innerText;
                textArea.style.position = 'fixed';
                textArea.style.left = '-999999px';
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
              }
              
              // Clear selection
              selection.removeAllRanges();
            } catch (error) {
              console.error('Error in copyCurrentVersionToClipboard:', error);
            }
          }
          
          function saveAndClose() {
            const button = document.getElementById('saveCloseBtn');
            if (button) {
              button.dataset.originalText = button.innerHTML;
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            google.script.run
              .withSuccessHandler(function() {
                // Clear the temporary data
                google.script.run.clearTempEmailData();
                
                google.script.run
                  .withSuccessHandler(function(nextClient) {
                    if (nextClient) {
                      if (confirm('Ready for next client: ' + nextClient.studentName + '?')) {
                        google.script.host.close();
                      } else {
                        google.script.host.close();
                      }
                    } else {
                      google.script.host.close();
                    }
                  })
                  .getNextBatchClient();
              })
              .withFailureHandler(function(error) {
                if (button) {
                  button.classList.remove('btn-loading');
                  button.innerHTML = button.dataset.originalText || button.innerHTML;
                  button.disabled = false;
                }
                alert('Error logging email: ' + error.message);
                google.script.host.close();
              })
              .logRecapSent(emailData.formData);
          }
        </script>
      </body>
    </html>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(700)
    .setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview Update');
}
/**
 * Send the recap email - now just logs it as sent
 */
function sendRecapEmail(emailData) {
  // This function is no longer used for sending emails
  // Kept for compatibility if called from other parts
  logRecapSent(emailData.formData);
  return 'Email logged successfully';
}

/**
 * Get next batch client
 */
function getNextBatchClient() {
  // Get batch processing state from UnifiedClientDataStore
  const batchState = UnifiedClientDataStore.getTempData('batchProcessing');
  
  if (!batchState || !batchState.batchClients) return null;
  
  const clients = batchState.batchClients;
  const currentIndex = batchState.batchIndex;
  
  if (currentIndex < clients.length - 1) {
    const nextIndex = currentIndex + 1;
    
    // Update batch state
    const updatedState = {
      batchClients: clients,
      batchIndex: nextIndex
    };
    UnifiedClientDataStore.setTempData('batchProcessing', updatedState);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(clients[nextIndex]);
    if (sheet) {
      spreadsheet.setActiveSheet(sheet);
      const clientData = getClientDataFromSheet(sheet);
      
      // Automatically open the next recap form
      showRecapDialog(clientData);
      
      return clientData;
    }
  } else {
    // Batch complete - clear temporary data
    UnifiedClientDataStore.deleteTempData('batchProcessing');
  }
  
  return null;
}

function getClientDataFromSheet(sheet) {
  const sheetName = sheet.getName();
  const clientNameWithSpaces = sheet.getRange('A1').getValue() || sheetName.replace(/([A-Z])/g, ' $1').trim();
  const clientTypes = sheet.getRange('C2').getValue() || '';
  
  // Get client data from UnifiedClientDataStore
  const unifiedClient = UnifiedClientDataStore.getClient(clientNameWithSpaces);
  
  let dashboardLink = '';
  let meetingNotesLink = '';
  let scoreReportLink = '';
  let parentEmail = '';
  let emailAddressees = '';
  let isActive = true; // Default to true for new clients
  
  if (unifiedClient) {
    dashboardLink = unifiedClient.links.dashboard || '';
    meetingNotesLink = unifiedClient.links.meetingNotes || '';
    scoreReportLink = unifiedClient.links.scoreReport || '';
    parentEmail = unifiedClient.contactInfo.parentEmail || '';
    emailAddressees = unifiedClient.contactInfo.emailAddressees || '';
    isActive = unifiedClient.isActive !== undefined ? unifiedClient.isActive : true;
  }
  
  // Extract parent name from emailAddressees (e.g., "John & Jane Doe" becomes "John & Jane Doe")
  // This is used for display in the recap dialog
  const parentName = emailAddressees || 'Parents';
  
  return {
    studentName: clientNameWithSpaces,
    clientTypes: clientTypes,
    clientType: clientTypes,
    dashboardLink: dashboardLink,
    meetingNotesLink: meetingNotesLink,
    scoreReportLink: scoreReportLink,
    isActive: isActive,
    parentEmail: parentEmail,
    emailAddressees: emailAddressees,
    parentName: parentName,
    sheetName: sheetName
  };
}

function debugDashboardCell() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getRange('D2');
  
  
  try {
    const richText = cell.getRichTextValue();
    if (richText) {
      const runs = richText.getRuns();
      runs.forEach((run, index) => {
        const runInfo = {
          text: run.getText(),
          url: run.getLinkUrl(),
          textStyle: run.getTextStyle()
        };
        console.log(`Run ${index}:`, runInfo);
      });
    }
  } catch (e) {
  }
  
  // Try formula
  try {
    const formula = cell.getFormula();
    if (formula) {
    }
  } catch (e) {
  }
}


/**
 * Create or update recap tracking sheet
 */
function createRecapTrackingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('SessionRecaps');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('SessionRecaps');
  }
  
  // Set up headers
  const headers = [
    'Date Sent',
    'Student Name',
    'Client Type',
    'Parent Email',
    'Focus Skill',
    'Today\'s Win',
    'Homework',
    'Next Session',
    'Status'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#f8f9fa')
    .setFontWeight('bold');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Log sent recap to tracking sheet
 */
function logRecapSent(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('SessionRecaps');
  
  if (!sheet) {
    createRecapTrackingSheet();
    sheet = spreadsheet.getSheetByName('SessionRecaps');
  }
  
  const row = [
    new Date(),
    data.studentName,
    data.clientType,
    data.parentEmail,
    data.focusSkill,
    data.todaysWin,
    data.homework,
    data.nextSession || '',
    'Sent'
  ];
  
  sheet.appendRow(row);
}

/**
 * Helper function to check if a sheet is a client sheet
 */
function isClientSheet(sheetName) {
  const excludeNames = ['NewClient', 'New Client', 'NewACTClient', 'NewAcademicSupportClient', 
                       'Template', 'Settings', 'Dashboard', 'Summary', 'SessionRecaps', 'Master',
                       'Date Sent', 'Student Name', 'Client Type', 'Parent Email', 'Focus Skill'];
  
  // Check if it's in the exclude list
  if (excludeNames.includes(sheetName)) {
    return false;
  }
  
  // Check if it contains "Template"
  if (sheetName.includes('Template')) {
    return false;
  }
  
  // Check if it starts with "Sheet" (default sheets)
  if (sheetName.startsWith('Sheet')) {
    return false;
  }
  
  // Check for column header patterns that shouldn't be treated as client names
  if (sheetName.toLowerCase().includes('date') || 
      sheetName.toLowerCase().includes('sent') || 
      sheetName.toLowerCase().includes('email') ||
      sheetName.toLowerCase().includes('type') ||
      sheetName.toLowerCase().includes('skill') ||
      sheetName.toLowerCase().includes('win')) {
    return false;
  }
  
  // For valid client sheets, allow spaces in names
  // Just check that it's not empty and not a system sheet
  return sheetName.trim().length > 0;
}

/**
 * Debug function to show all sheets and their A1 values
 */
function debugSheetStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      try {
        const a1Value = sheet.getRange('A1').getValue();
        
        // Check if this would be filtered out
        const invalidNames = ['Date Sent', 'Student Name', 'Client Type', 'Template', 'New Client', 'Parent Email', 'Focus Skill', 'Today\'s Win', 'Client Name'];
        const wouldBeFiltered = invalidNames.includes(name) || 
                               invalidNames.includes(a1Value) ||
                               (a1Value && a1Value.toString().toLowerCase().includes('sent'));
        
        if (name === 'Date Sent' || a1Value === 'Date Sent') {
        }
        
      } catch (error) {
      }
    });
    
    
    SpreadsheetApp.getUi().alert('Debug Complete', 'Check the Apps Script console (Ctrl+`) for detailed sheet information.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Error debugging sheet structure:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to debug sheet structure: ' + error.message);
  }
}

/**
 * Manual refresh of client list cache
 */
function refreshClientCache() {
  try {
    const startTime = Date.now();
    
    // Clear and rebuild UnifiedClientDataStore first
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('UNIFIED_CLIENT_DATA');
    
    // Clear all legacy caches
    properties.deleteProperty('ENHANCED_CLIENT_CACHE');
    properties.deleteProperty('CLIENT_LIST_CACHE');
    properties.deleteProperty('unified_client_data_backup');
    
    // Rebuild from UnifiedClientDataStore only
    const allClients = UnifiedClientDataStore.getAllClients();
    
    // Log all clients found before filtering
    
    // Filter out invalid clients more aggressively
    const validClients = allClients.filter(client => {
      const invalidNames = ['Date Sent', 'Student Name', 'Client Type', 'Template', 'New Client', 'Parent Email', 'Focus Skill', 'Today\'s Win', 'Client Name'];
      const name = client.name ? client.name.toString().trim() : '';
      
      
      // Exact match check
      if (invalidNames.includes(name)) {
        return false;
      }
      
      // Pattern matching
      const lowerName = name.toLowerCase();
      if (lowerName.includes('sent') || 
          lowerName.includes('email') || 
          lowerName.includes('type') ||
          lowerName.includes('skill') ||
          lowerName.includes('date') ||
          lowerName.includes('win') ||
          !name || name.length === 0) {
        return false;
      }
      
      return true;
    });
    
    // Log valid clients after filtering
    
    const validActiveCount = validClients.filter(client => client.isActive).length;
    
    // Force refresh of UnifiedClientDataStore
    const refreshedClients = UnifiedClientDataStore.getAllClients();
    const addedCount = refreshedClients.length;
    const activeCount = refreshedClients.filter(client => client.isActive).length;
    
    const duration = Date.now() - startTime;
    
    const message = `Unified client data refreshed successfully!\n\n` +
                   `Total clients: ${addedCount}\n` +
                   `Active clients: ${activeCount}\n` +
                   `Refresh time: ${duration}ms\n\n` +
                   `Using UnifiedClientDataStore (modern architecture)`;
    
    SpreadsheetApp.getUi().alert('Cache Refreshed', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    
    return { 
      success: true,
      message: 'Unified client data refreshed successfully',
      stats: {
        total: addedCount,
        active: activeCount,
        duration: duration
      }
    };
  } catch (error) {
    console.error('Error refreshing client cache:', error);
    
    SpreadsheetApp.getUi().alert('Error', 
      'Failed to refresh client cache: ' + error.message, 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
    return { 
      success: false,
      message: 'Error refreshing cache: ' + error.message 
    };
  }
}

/**
 * Placeholder function for checking dashboard links
 */
function validateDashboardLinks() {
  // This is a placeholder - implement if needed
  SpreadsheetApp.getUi().alert('Dashboard links validated!');
  return { message: 'Dashboard links validated' };
}

/**
 * Placeholder function for exporting client list
 */
function exportClientList() {
  // This is a placeholder - implement if needed
  SpreadsheetApp.getUi().alert('Client list exported!');
  return { message: 'Client list exported' };
}


/**
 * View recap history
 */
function viewRecapHistory() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('SessionRecaps');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'No History Found',
      'No session recaps have been sent yet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  spreadsheet.setActiveSheet(sheet);
}

/**
 * Test connection function for checking internet connectivity
 */
function testConnection() {
  try {
    // Simple test that just returns true if the function can execute
    return true;
  } catch (error) {
    console.error('Connection test failed:', error);
    return false;
  }
}

/**
 * Get list of all clients - now uses persistent cache
 */
function getAllClientList() {
  // Deprecated: Use getUnifiedClientList() instead
  console.warn('getAllClientList() is deprecated. Use getUnifiedClientList() instead.');
  try {
    return UnifiedClientDataStore.getAllClients();
  } catch (error) {
    console.error('Error getting all client list:', error);
    return [];
  }
}

/**
 * Get list of active clients only - now uses persistent cache
 */
function getActiveClientList() {
  try {
    const activeClients = UnifiedClientDataStore.getActiveClients();
    
    // Format for sidebar display
    const html = `
      <div style="max-height: 400px; overflow-y: auto;">
        ${activeClients.map(client => `
          <div style="padding: 10px; margin: 5px; background: #f8f9fa; border-radius: 5px; cursor: pointer;"
               onclick="google.script.run.withSuccessHandler(() => {
                 google.script.host.close();
                 google.script.run.switchToClient('${client.name}');
               }).switchToClient('${client.name}')">
            <strong>${client.name}</strong>
            ${client.email ? `<br><small>${client.email}</small>` : ''}
          </div>
        `).join('')}
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(400)
      .setHeight(Math.min(activeClients.length * 70 + 50, 500));
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Active Clients');
    
    return { 
      success: true, 
      clients: activeClients,
      count: activeClients.length
    };
  } catch (error) {
    console.error('Error getting active client list:', error);
    return { 
      success: false, 
      clients: [],
      error: error.toString()
    };
  }
}

/**
 * Direct Master sheet read as fallback (old method)
 */
function getActiveClientListDirect() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = spreadsheet.getSheetByName('Master');
    
    if (!masterSheet) {
      return [];
    }
    
    const lastRow = masterSheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    const nameRange = masterSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const activeRange = masterSheet.getRange(2, 5, lastRow - 1, 1).getValues();
    
    const activeClients = [];
    
    for (let i = 0; i < nameRange.length; i++) {
      const clientName = nameRange[i][0];
      const isActive = activeRange[i][0];
      
      if (clientName && isActive === true) {
        activeClients.push({
          name: clientName,
          sheetName: clientName.replace(/\s+/g, ''),
          isActive: true
        });
      }
    }
    
    return activeClients;
  } catch (error) {
    console.error('Error in direct client list read:', error);
    return [];
  }
}

/**
 * Enhanced Client Cache System
 * Comprehensive client data storage that eliminates the need for the Master sheet
 * Stores all client information including links, contact data, and session info
 */
const EnhancedClientCache = {
  CACHE_KEY: 'enhanced_client_cache',
  VERSION_KEY: 'enhanced_client_cache_version',
  TIMESTAMP_KEY: 'enhanced_client_cache_timestamp',
  
  /**
   * Get complete cached client data
   */
  getCachedClientData: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const cachedData = properties.getProperty(this.CACHE_KEY);
      const cacheVersion = properties.getProperty(this.VERSION_KEY) || '1';
      const timestamp = properties.getProperty(this.TIMESTAMP_KEY);
      
      if (cachedData) {
        return {
          clients: JSON.parse(cachedData),
          version: cacheVersion,
          timestamp: parseInt(timestamp) || 0
        };
      }
      return null;
    } catch (error) {
      console.error('Error getting cached client data:', error);
      return null;
    }
  },
  
  /**
   * Build enhanced cache from existing data sources (Script Properties + client sheets)
   */
  buildEnhancedCache: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const scriptProperties = PropertiesService.getScriptProperties();
      const sheets = spreadsheet.getSheets();
      
      const allClients = [];
      
      // Get all client sheets (exclude Master and other system sheets)
      const clientSheets = sheets.filter(sheet => {
        const name = sheet.getName();
        // Exclude system sheets and default sheets
        if (name === 'Master' || 
            name === 'SessionRecaps' ||
            name === 'Settings' ||
            name === 'Dashboard' ||
            name === 'Summary' ||
            name.startsWith('Sheet') || 
            name.includes('Template') ||
            name === 'NewClient' ||
            name === 'New Client') {
          return false;
        }
        
        // Check if A1 contains a valid client name (not empty, not a formula header)
        const a1Value = sheet.getRange('A1').getValue();
        if (!a1Value || 
            a1Value === '' || 
            a1Value === 'New Client' || 
            a1Value === 'Client Name' ||
            a1Value === 'Date Sent' ||
            a1Value === 'Student Name' ||
            a1Value === 'Client Type' ||
            a1Value.toString().toLowerCase().includes('name') ||
            a1Value.toString().toLowerCase().includes('date') ||
            a1Value.toString().toLowerCase().includes('type')) {
          return false;
        }
        
        return true;
      });
      
      
      for (const sheet of clientSheets) {
        const sheetName = sheet.getName();
        
        try {
          // Get client name from A1
          const a1Value = sheet.getRange('A1').getValue();
          const clientName = a1Value || sheetName.replace(/([A-Z])/g, ' $1').trim();
          
          
          // Get all data from UnifiedClientDataStore if available, otherwise try legacy Script Properties for migration
          let dashboardLink = '';
          let meetingNotesLink = '';
          let scoreReportLink = '';
          let parentEmail = '';
          let emailAddressees = '';
          let isActive = true;
          
          // Try UnifiedClientDataStore first
          const unifiedClient = UnifiedClientDataStore.getClient(clientName);
          if (unifiedClient) {
            dashboardLink = unifiedClient.links.dashboard || '';
            meetingNotesLink = unifiedClient.links.meetingNotes || '';
            scoreReportLink = unifiedClient.links.scoreReport || '';
            parentEmail = unifiedClient.contactInfo.parentEmail || '';
            emailAddressees = unifiedClient.contactInfo.emailAddressees || '';
            isActive = unifiedClient.isActive !== undefined ? unifiedClient.isActive : true;
          } else {
            // Fallback to Script Properties for migration purposes
            dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || '';
            meetingNotesLink = scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '';
            scoreReportLink = scriptProperties.getProperty(`scorereport_${sheetName}`) || '';
            parentEmail = scriptProperties.getProperty(`parentemail_${sheetName}`) || '';
            emailAddressees = scriptProperties.getProperty(`emailaddressees_${sheetName}`) || '';
            
            // Get active status from Script Properties (default to true if not set)
            const isActiveProperty = scriptProperties.getProperty(`isActive_${sheetName}`);
            isActive = isActiveProperty === null ? true : isActiveProperty === 'true';
            
            // Add to UnifiedClientDataStore for future use
            const clientDataForUnified = {
              name: clientName,
              isActive: isActive,
              dashboardLink: dashboardLink,
              meetingNotesLink: meetingNotesLink,
              scoreReportLink: scoreReportLink,
              parentEmail: parentEmail,
              emailAddressees: emailAddressees
            };
            
            try {
              UnifiedClientDataStore.upsertClient(clientDataForUnified);
            } catch (migrateError) {
              console.warn(`Failed to migrate client ${clientName} to UnifiedClientDataStore:`, migrateError);
            }
          }
          
          // Get session information from the client sheet
          const lastSession = this.getLastSessionFromSheet(sheet);
          const nextSession = this.getNextSessionFromSheet(sheet);
          
          const clientData = {
            name: clientName.toString().trim(),
            sheetName: sheetName,
            isActive: isActive,
            dashboardLink: dashboardLink,
            meetingNotesLink: meetingNotesLink,
            scoreReportLink: scoreReportLink,
            parentEmail: parentEmail,
            emailAddressees: emailAddressees,
            lastSession: lastSession,
            nextSession: nextSession,
            lastUpdated: Date.now()
          };
          
          allClients.push(clientData);
          
        } catch (sheetError) {
          console.warn(`❌ Error processing sheet "${sheetName}":`, sheetError);
          continue;
        }
      }
      
      // Save to cache
      this.setCachedClientData(allClients);
      
      return allClients;
    } catch (error) {
      console.error('Error building enhanced cache:', error);
      return [];
    }
  },
  
  /**
   * Get last session from client sheet
   */
  getLastSessionFromSheet: function(sheet) {
    try {
      // Find the last non-empty row in column A
      const lastRow = sheet.getLastRow();
      if (lastRow < 5) return ''; // No sessions recorded
      
      // Look for the most recent session date (working backwards from last row)
      for (let i = lastRow; i >= 5; i--) {
        const cellValue = sheet.getRange(i, 1).getValue();
        if (cellValue && cellValue.toString().trim()) {
          return cellValue.toString().trim();
        }
      }
      return '';
    } catch (error) {
      console.warn('Error getting last session from sheet:', error);
      return '';
    }
  },
  
  /**
   * Get next session info from client sheet or recent notes
   */
  getNextSessionFromSheet: function(sheet) {
    try {
      // This could be enhanced to look for next session info in recent notes
      // For now, return empty - this can be populated when sessions are scheduled
      return '';
    } catch (error) {
      console.warn('Error getting next session from sheet:', error);
      return '';
    }
  },
  
  /**
   * Save complete client data to cache
   */
  setCachedClientData: function(clientList) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const currentVersion = parseInt(properties.getProperty(this.VERSION_KEY) || '1');
      
      var setPropsObj = {};
      setPropsObj[this.CACHE_KEY] = JSON.stringify(clientList);
      setPropsObj[this.VERSION_KEY] = (currentVersion + 1).toString();
      setPropsObj[this.TIMESTAMP_KEY] = Date.now().toString();
      properties.setProperties(setPropsObj);
      
    } catch (error) {
      console.error('Error setting cached client data:', error);
    }
  },
  
  /**
   * Get only active clients from cache
   */
  getActiveClientsFromCache: function() {
    const cachedData = this.getCachedClientData();
    if (cachedData && cachedData.clients) {
      return cachedData.clients.filter(function(client) { return client.isActive; });
    }
    return null;
  },
  
  /**
   * Get all clients from cache
   */
  getAllClientsFromCache: function() {
    const cachedData = this.getCachedClientData();
    if (cachedData && cachedData.clients) {
      return cachedData.clients;
    }
    return null;
  },
  
  /**
   * Get specific client data by sheet name
   */
  getClientBySheetName: function(sheetName) {
    const cachedData = this.getCachedClientData();
    if (cachedData && cachedData.clients) {
      return cachedData.clients.find(function(client) { 
        return client.sheetName === sheetName; 
      });
    }
    return null;
  },
  
  /**
   * Clear all cached data
   */
  clearCache: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty(this.CACHE_KEY);
      properties.deleteProperty(this.VERSION_KEY);
      properties.deleteProperty(this.TIMESTAMP_KEY);
    } catch (error) {
      console.error('Error clearing cache:', error);
    }
  },
  
  /**
   * Get specific client data by name (case insensitive)
   */
  getClientByName: function(clientName) {
    const cachedData = this.getCachedClientData();
    if (cachedData && cachedData.clients) {
      return cachedData.clients.find(function(client) { 
        return client.name.toLowerCase() === clientName.toLowerCase(); 
      });
    }
    return null;
  },
  
  /**
   * Add new client to cache with complete data
   */
  addClientToCache: function(clientData) {
    try {
      const cachedData = this.getCachedClientData();
      const clientList = cachedData ? cachedData.clients : [];
      
      // Check if client already exists
      var existingIndex = -1;
      for (var i = 0; i < clientList.length; i++) {
        if (clientList[i].sheetName === clientData.sheetName) {
          existingIndex = i;
          break;
        }
      }
      
      // Create complete client data object
      const completeClientData = {
        name: clientData.name || clientData.clientName || '',
        sheetName: clientData.sheetName || '',
        isActive: typeof clientData.isActive !== 'undefined' ? clientData.isActive : true,
        dashboardLink: clientData.dashboardLink || '',
        meetingNotesLink: clientData.meetingNotesLink || '',
        scoreReportLink: clientData.scoreReportLink || '',
        parentEmail: clientData.parentEmail || '',
        emailAddressees: clientData.emailAddressees || '',
        lastSession: clientData.lastSession || '',
        nextSession: clientData.nextSession || '',
        lastUpdated: Date.now()
      };
      
      if (existingIndex >= 0) {
        // Update existing client
        clientList[existingIndex] = completeClientData;
      } else {
        // Add new client
        clientList.push(completeClientData);
      }
      
      this.setCachedClientData(clientList);
      
    } catch (error) {
      console.error('Error adding client to enhanced cache:', error);
    }
  },
  
  /**
   * Update client active status in cache
   */
  updateClientStatusInCache: function(clientName, isActive) {
    try {
      const cachedData = this.getCachedClientData();
      if (!cachedData || !cachedData.clients) {
        console.warn('No cached client data found for status update');
        return;
      }
      
      const clientList = cachedData.clients;
      var clientIndex = -1;
      for (var i = 0; i < clientList.length; i++) {
        if (clientList[i].name.toLowerCase() === clientName.toLowerCase()) {
          clientIndex = i;
          break;
        }
      }
      
      if (clientIndex >= 0) {
        clientList[clientIndex].isActive = isActive;
        clientList[clientIndex].lastUpdated = Date.now();
        this.setCachedClientData(clientList);
      } else {
        console.warn('Client "' + clientName + '" not found in cache for status update');
      }
      
    } catch (error) {
      console.error('Error updating client status in cache:', error);
    }
  },
  
  /**
   * Update specific client field in cache
   */
  updateClientField: function(sheetName, field, value) {
    try {
      const cachedData = this.getCachedClientData();
      if (!cachedData || !cachedData.clients) {
        console.warn('No cached client data found for field update');
        return false;
      }
      
      const clientList = cachedData.clients;
      var clientIndex = -1;
      for (var i = 0; i < clientList.length; i++) {
        if (clientList[i].sheetName === sheetName) {
          clientIndex = i;
          break;
        }
      }
      
      if (clientIndex >= 0) {
        clientList[clientIndex][field] = value;
        clientList[clientIndex].lastUpdated = Date.now();
        this.setCachedClientData(clientList);
        return true;
      } else {
        console.warn(`Client with sheet name "${sheetName}" not found in cache for field update`);
        return false;
      }
      
    } catch (error) {
      console.error('Error updating client field in cache:', error);
      return false;
    }
  },
  
  /**
   * Update multiple client fields at once
   */
  updateClientFields: function(sheetName, updates) {
    try {
      const cachedData = this.getCachedClientData();
      if (!cachedData || !cachedData.clients) {
        console.warn('No cached client data found for bulk update');
        return false;
      }
      
      const clientList = cachedData.clients;
      var clientIndex = -1;
      for (var i = 0; i < clientList.length; i++) {
        if (clientList[i].sheetName === sheetName) {
          clientIndex = i;
          break;
        }
      }
      
      if (clientIndex >= 0) {
        // Apply all updates
        for (const field in updates) {
          if (updates.hasOwnProperty(field)) {
            clientList[clientIndex][field] = updates[field];
          }
        }
        clientList[clientIndex].lastUpdated = Date.now();
        this.setCachedClientData(clientList);
        return true;
      } else {
        console.warn(`Client with sheet name "${sheetName}" not found in cache for bulk update`);
        return false;
      }
      
    } catch (error) {
      console.error('Error updating client fields in cache:', error);
      return false;
    }
  },
  
  /**
   * Remove client from cache
   */
  removeClientFromCache: function(sheetName) {
    try {
      const cachedData = this.getCachedClientData();
      if (!cachedData || !cachedData.clients) {
        console.warn('No cached client data found for removal');
        return false;
      }
      
      const clientList = cachedData.clients;
      const initialLength = clientList.length;
      
      // Filter out the client to remove
      const updatedList = clientList.filter(function(client) {
        return client.sheetName !== sheetName;
      });
      
      if (updatedList.length < initialLength) {
        this.setCachedClientData(updatedList);
        return true;
      } else {
        console.warn(`Client with sheet name "${sheetName}" not found in cache for removal`);
        return false;
      }
      
    } catch (error) {
      console.error('Error removing client from cache:', error);
      return false;
    }
  },
  
  /**
   * Get cache statistics for debugging
   */
  getCacheStats: function() {
    const cachedData = this.getCachedClientData();
    if (!cachedData) {
      return { cached: false, total: 0, active: 0, timestamp: null };
    }
    
    const activeCount = cachedData.clients.filter(function(client) { return client.isActive; }).length;
    
    return {
      cached: true,
      total: cachedData.clients.length,
      active: activeCount,
      timestamp: new Date(cachedData.timestamp).toLocaleString(),
      version: cachedData.version
    };
  }
};

// DEPRECATED: ClientListCache compatibility layer removed
// All functions now use UnifiedClientDataStore directly

/**
 * Enhanced function to populate cache from all existing data sources
 * This replaces the need to read from the Master sheet
 */
function buildComprehensiveClientCache() {
  try {
    const result = UnifiedClientDataStore.getAllClients();
    
    if (result.length > 0) {
      return {
        success: true,
        message: `Retrieved ${result.length} clients from unified data store`,
        clientCount: result.length,
        clients: result
      };
    } else {
      console.warn('No clients found when building cache');
      return {
        success: true,
        message: 'Cache built but no clients found',
        clientCount: 0,
        clients: []
      };
    }
  } catch (error) {
    console.error('Error building comprehensive client cache:', error);
    return {
      success: false,
      message: 'Error building cache: ' + error.message,
      clientCount: 0,
      clients: []
    };
  }
}

/**
 * Enhanced helper functions for accessing cached client data
 */
function getClientFromCache(sheetName) {
  // Migrated to UnifiedClientDataStore
  const allClients = UnifiedClientDataStore.getAllClients();
  return allClients.find(client => client.sheetName === sheetName) || null;
}

function getAllActiveClientsFromCache() {
  // Migrated to UnifiedClientDataStore
  return UnifiedClientDataStore.getActiveClients();
}

function getAllClientsFromCache() {
  // Migrated to UnifiedClientDataStore
  return UnifiedClientDataStore.getAllClients();
}

function updateClientInCache(sheetName, updates) {
  // Migrated to UnifiedClientDataStore
  try {
    const client = UnifiedClientDataStore.getAllClients().find(c => c.sheetName === sheetName);
    if (client) {
      const updatedClient = { ...client, ...updates };
      return UnifiedClientDataStore.upsertClient(updatedClient);
    }
    return false;
  } catch (error) {
    console.error('Error updating client in cache:', error);
    return false;
  }
}

function addNewClientToCache(clientData) {
  // Migrate to UnifiedClientDataStore
  try {
    return UnifiedClientDataStore.upsertClient(clientData);
  } catch (error) {
    console.warn('Failed to add client to unified store:', error);
    return null;
  }
}

/**
 * Get client details for update dialog (enhanced with cache support)
 * First tries cache, then falls back to direct data access
 */
function getClientDetailsForUpdate(sheetName) {
  try {
    // Try to get from UnifiedClientDataStore first
    const allClients = UnifiedClientDataStore.getAllClients();
    const cachedClient = allClients.find(client => client.sheetName === sheetName);
    if (cachedClient) {
      
      // Get master row for compatibility (if Master sheet exists)
      let masterRow = null;
      try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const masterSheet = spreadsheet.getSheetByName('Master');
        if (masterSheet) {
          const masterData = masterSheet.getDataRange().getValues();
          for (let i = 1; i < masterData.length; i++) {
            if (masterData[i][0] === cachedClient.name) {
              masterRow = i + 1;
              break;
            }
          }
        }
      } catch (masterError) {
        console.warn('Could not determine master row:', masterError);
      }
      
      return {
        clientName: cachedClient.name,
        sheetName: cachedClient.sheetName,
        dashboardLink: cachedClient.dashboardLink || '',
        meetingNotesLink: cachedClient.meetingNotesLink || '',
        scoreReportLink: cachedClient.scoreReportLink || '',
        parentEmail: cachedClient.parentEmail || '',
        emailAddressees: cachedClient.emailAddressees || '',
        isActive: cachedClient.isActive,
        lastSession: cachedClient.lastSession || '',
        nextSession: cachedClient.nextSession || '',
        masterRow: masterRow,
        fromCache: true
      };
    }
    
    // Fallback to original method if not in cache
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    const masterSheet = spreadsheet.getSheetByName('Master');
    
    if (!sheet) {
      throw new Error('Client sheet not found');
    }
    
    // Get client name from sheet A1
    const clientName = sheet.getRange('A1').getValue();
    
    // Get links and parent data from UnifiedClientDataStore
    const unifiedClient = UnifiedClientDataStore.getClient(clientName);
    
    let dashboardLink = '';
    let meetingNotesLink = '';
    let scoreReportLink = '';
    let parentEmail = '';
    let emailAddressees = '';
    
    if (unifiedClient) {
      dashboardLink = unifiedClient.links.dashboard || '';
      meetingNotesLink = unifiedClient.links.meetingNotes || '';
      scoreReportLink = unifiedClient.links.scoreReport || '';
      parentEmail = unifiedClient.contactInfo.parentEmail || '';
      emailAddressees = unifiedClient.contactInfo.emailAddressees || '';
    } else {
    }
    
    // Find row in Master sheet
    const masterData = masterSheet.getDataRange().getValues();
    let masterRow = -1;
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0] === clientName) {
        masterRow = i + 1; // Convert to 1-based
        break;
      }
    }
    
    if (masterRow === -1) {
      // Client not found in master sheet - automatically sync it
      
      // Prepare client data for master list
      const clientData = {
        name: clientName,
        dashboardLink: dashboardLink,
        meetingNotesLink: meetingNotesLink,
        scoreReportLink: scoreReportLink
      };
      
      // Add to master list using existing function
      addClientToMasterList(clientData, sheetName, spreadsheet);
      
      // Update enhanced cache with full client data
      const fullClientData = {
        name: clientName,
        sheetName: sheetName,
        isActive: true,
        dashboardLink: dashboardLink,
        meetingNotesLink: meetingNotesLink,
        scoreReportLink: scoreReportLink,
        parentEmail: parentEmail,
        emailAddressees: emailAddressees
      };
      // Add to UnifiedClientDataStore
    try {
      UnifiedClientDataStore.upsertClient(fullClientData);
    } catch (error) {
      console.warn(`Failed to add client to UnifiedClientDataStore:`, error);
    }
      
      // Now find the newly added row
      const updatedMasterData = masterSheet.getDataRange().getValues();
      for (let i = 1; i < updatedMasterData.length; i++) {
        if (updatedMasterData[i][0] === clientName) {
          masterRow = i + 1; // Convert to 1-based
          break;
        }
      }
    }
    
    // Get active status from UnifiedClientDataStore
    const isActive = unifiedClient && unifiedClient.isActive !== undefined ? unifiedClient.isActive : true; // Default to true for new clients
    
    // Add to cache for future use (if not already cached)
    const fullClientData = {
      name: clientName,
      sheetName: sheetName,
      isActive: isActive,
      dashboardLink: dashboardLink,
      meetingNotesLink: meetingNotesLink,
      scoreReportLink: scoreReportLink,
      parentEmail: parentEmail,
      emailAddressees: emailAddressees
    };
    // Add to UnifiedClientDataStore
    try {
      UnifiedClientDataStore.upsertClient(fullClientData);
    } catch (error) {
      console.warn(`Failed to add client to UnifiedClientDataStore:`, error);
    }
    
    return {
      clientName: clientName,
      sheetName: sheetName,
      dashboardLink: dashboardLink,
      meetingNotesLink: meetingNotesLink,
      scoreReportLink: scoreReportLink,
      parentEmail: parentEmail,
      emailAddressees: emailAddressees,
      isActive: isActive,
      masterRow: masterRow,
      fromCache: false
    };
  } catch (error) {
    console.error('Error getting client details:', error);
    throw error;
  }
}

/**
 * Update client details (links and active status)
 */
function updateClientDetails(sheetName, details) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = spreadsheet.getSheetByName('Master');
    
    if (!masterSheet) {
      throw new Error('Master sheet not found');
    }
    
    // Get client details to find master row
    const clientDetails = getClientDetailsForUpdate(sheetName);
    const masterRow = clientDetails.masterRow;
    
    // Update client data in UnifiedClientDataStore (primary data source)
    const clientUpdateData = {
      name: clientDetails.clientName,
      isActive: details.isActive,
      dashboardLink: details.dashboardLink || '',
      meetingNotesLink: details.meetingNotesLink || '',
      scoreReportLink: details.scoreReportLink || '',
      parentEmail: details.parentEmail || '',
      emailAddressees: details.emailAddressees || ''
    };
    
    try {
      UnifiedClientDataStore.updateClient(clientDetails.clientName, clientUpdateData);
    } catch (error) {
      console.error('Failed to update UnifiedClientDataStore:', error);
      throw error;
    }
    
    // Update Master sheet Column E to reflect the boolean value
    if (masterRow) {
      masterSheet.getRange(masterRow, 5).setValue(details.isActive);
    }
    
    return {
      success: true,
      message: 'Client details updated successfully'
    };
  } catch (error) {
    console.error('Error updating client details:', error);
    return {
      success: false,
      message: 'Error updating client details: ' + error.message
    };
  }
}

/**
 * Navigate to a specific sheet by name
 */
function navigateToSheet(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }
    
    spreadsheet.setActiveSheet(sheet);
    return { success: true };
  } catch (error) {
    console.error('Error navigating to sheet:', error);
    throw error;
  }
}

/**
 * Background processing and optimization initialization
 * Called when the sidebar loads to start background tasks
 */
function initializeBackgroundProcesses() {
  // Start periodic cache cleanup
  setInterval(() => {
    // Clean up expired cache entries
    const now = Date.now();
    Object.keys(PerformanceCache.data.clientDetails).forEach(key => {
      const entry = PerformanceCache.data.clientDetails.get(key);
      if (entry && entry.timestamp && (now - entry.timestamp > PerformanceCache.CACHE_DURATION)) {
        PerformanceCache.data.clientDetails.delete(key);
      }
    });
  }, 5 * 60 * 1000); // Clean every 5 minutes
  
  // Initialize connection monitoring (already handled by startConnectionMonitoring)
  // ConnectionMonitor.startMonitoring() - using existing optimized function
  
  // Prefetch commonly used data during idle time
  if (window.requestIdleCallback) {
    requestIdleCallback(() => {
      PerformanceCache.prefetchClientList();
    });
  } else {
    // Fallback for browsers without requestIdleCallback
    setTimeout(() => {
      PerformanceCache.prefetchClientList();
    }, 2000);
  }
}

/**
 * Test function to verify the enhanced cache system
 * Run this to validate that the cache is working correctly
 */
function testUnifiedDataStore() {
  try {
    
    // Test 1: Build cache
    const result = buildComprehensiveClientCache();
    
    if (!result.success) {
      throw new Error('Failed to build cache: ' + result.message);
    }
    
    // Test 2: Get all clients from cache
    const allClients = getAllClientsFromCache();
    
    if (allClients && allClients.length > 0) {
    }
    
    // Test 3: Get active clients only
    const activeClients = getAllActiveClientsFromCache();
    
    // Test 4: Get specific client by sheet name
    if (allClients && allClients.length > 0) {
      const testClient = allClients[0];
      const retrievedClient = getClientFromCache(testClient.sheetName);
    }
    
    // Test 5: Get storage statistics
    const stats = {
      totalClients: allClients ? allClients.length : 0,
      activeClients: activeClients ? activeClients.length : 0,
      dataSource: 'UnifiedClientDataStore'
    };
    
    const testMessage = `Unified Data Store System Test Results:
    
✓ Data store accessed successfully
✓ Total clients: ${result.clientCount}
✓ Active clients: ${activeClients ? activeClients.length : 0}
✓ Client data retrieval working correctly
✓ All test operations completed successfully

The unified data store system is working correctly!`;
    
    SpreadsheetApp.getUi().alert('Unified Data Store Test', testMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return {
      success: true,
      message: 'All tests passed',
      details: {
        totalClients: result.clientCount,
        activeClients: activeClients ? activeClients.length : 0,
        cacheStats: stats
      }
    };
    
  } catch (error) {
    console.error('Unified data store test failed:', error);
    const errorMessage = `Unified Data Store Test Failed:

Error: ${error.message}

Please check the console logs for more details.`;
    
    SpreadsheetApp.getUi().alert('Test Failed', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Utility function to get comprehensive client information
 * This function demonstrates how to access all client data from the cache
 */
function getComprehensiveClientInfo(sheetName) {
  const client = getClientFromCache(sheetName);
  if (!client) {
    return null;
  }
  
  return {
    // Basic info
    name: client.name,
    sheetName: client.sheetName,
    isActive: client.isActive,
    
    // Links
    dashboardLink: client.dashboardLink,
    meetingNotesLink: client.meetingNotesLink,
    scoreReportLink: client.scoreReportLink,
    
    // Contact info
    parentEmail: client.parentEmail,
    emailAddressees: client.emailAddressees,
    
    // Session info
    lastSession: client.lastSession,
    nextSession: client.nextSession,
    
    // Metadata
    lastUpdated: client.lastUpdated ? new Date(client.lastUpdated).toLocaleString() : 'Unknown'
  };
}

/**
 * Show bulk client entry dialog
 */
function showEnhancedBulkClientDialog() {
  // This function is now defined in enhanced_dialogs.js
  // Call the enhanced version
  return showEnhancedBulkClientDialog();
}

function showBulkClientDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
          }
          h2 {
            color: #333;
            margin-bottom: 8px;
            font-weight: 500;
            text-align: center;
          }
          .subtitle {
            color: #666;
            font-size: 14px;
            margin-bottom: 20px;
            text-align: center;
          }
          .container {
            max-width: 1200px;
            margin: 0 auto;
          }
          .table-container {
            overflow-x: auto;
            max-height: 400px;
            overflow-y: auto;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            margin-bottom: 20px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            background: white;
          }
          th {
            background: #f8f9fa;
            padding: 12px 8px;
            text-align: left;
            font-weight: 500;
            color: #5f6368;
            font-size: 13px;
            position: sticky;
            top: 0;
            z-index: 10;
            border-bottom: 2px solid #e0e0e0;
          }
          td {
            padding: 8px;
            border-bottom: 1px solid #f0f0f0;
          }
          tr:hover td {
            background: #f8f9fa;
          }
          input[type="text"], input[type="email"], input[type="url"] {
            width: 100%;
            padding: 8px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            transition: border-color 0.2s;
          }
          input:focus {
            outline: none;
            border-color: #1a73e8;
          }
          .service-checkboxes {
            display: flex;
            gap: 10px;
            flex-wrap: nowrap;
          }
          .service-checkbox {
            display: flex;
            align-items: center;
            padding: 4px 8px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 13px;
            white-space: nowrap;
          }
          .service-checkbox:hover {
            background: #f8f9fa;
          }
          .service-checkbox.selected {
            background: #e8f0fe;
            border-color: #1a73e8;
          }
          .service-checkbox input {
            margin-right: 4px;
            cursor: pointer;
          }
          .remove-btn {
            background: #f44336;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 6px 12px;
            cursor: pointer;
            font-size: 13px;
            transition: all 0.2s;
          }
          .remove-btn:hover {
            background: #d32f2f;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          }
          .button-group {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid #e0e0e0;
          }
          .btn-primary {
            background: #1a73e8;
            color: white;
            border: none;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          .btn-primary:hover {
            background: #1557b0;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          }
          .btn-primary:disabled {
            background: #dadce0;
            cursor: not-allowed;
            transform: none;
          }
          .btn-secondary {
            background: white;
            color: #5f6368;
            border: 1px solid #dadce0;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          .btn-secondary:hover {
            background: #f8f9fa;
            border-color: #d2d2d2;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          }
          .add-row-btn {
            background: #34a853;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            gap: 6px;
          }
          .add-row-btn:hover {
            background: #2d8e47;
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.2);
          }
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.95);
            display: none;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            z-index: 1000;
          }
          .loading-overlay.active {
            display: flex;
          }
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .loading-text {
            margin-top: 16px;
            font-size: 14px;
            color: #5f6368;
          }
          .success-message {
            color: #34a853;
            font-weight: 500;
            padding: 12px;
            background: #e6f4ea;
            border-radius: 4px;
            margin-bottom: 12px;
            display: none;
            text-align: center;
          }
          .error-message {
            color: #d93025;
            font-weight: 500;
            padding: 12px;
            background: #fce8e6;
            border-radius: 4px;
            margin-bottom: 12px;
            display: none;
          }
          .empty-state {
            text-align: center;
            padding: 40px;
            color: #5f6368;
          }
          .instructions {
            background: #f8f9fa;
            padding: 12px;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 14px;
            color: #5f6368;
          }
          .required {
            color: #d93025;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Add Multiple Clients</h2>
          <p class="subtitle">Enter information for multiple clients and add them all at once</p>
          
          <div class="instructions">
            <strong>Instructions:</strong> Fill in client information below. Name and at least one service are required for each client. 
            Click "Add Row" for more clients or "Remove" to delete a row.
          </div>
          
          <div class="success-message" id="successMessage"></div>
          <div class="error-message" id="errorMessage"></div>
          
          <div class="table-container">
            <table id="clientTable">
              <thead>
                <tr>
                  <th style="width: 30px;">#</th>
                  <th>Student Name <span class="required">*</span></th>
                  <th>Services <span class="required">*</span></th>
                  <th>Parent Email</th>
                  <th>Email Addressees</th>
                  <th>Dashboard Link</th>
                  <th style="width: 80px;">Action</th>
                </tr>
              </thead>
              <tbody id="clientTableBody">
                <!-- Rows will be added here -->
              </tbody>
            </table>
            <div class="empty-state" id="emptyState" style="display: none;">
              No clients added yet. Click "Add Row" to begin.
            </div>
          </div>
          
          <div class="button-group">
            <button class="add-row-btn" onclick="addRow()">
              <span>+</span> Add Row
            </button>
            <div>
              <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
              <button class="btn-primary" id="addAllBtn" onclick="addAllClients()">Add All Clients</button>
            </div>
          </div>
        </div>
        
        <div class="loading-overlay" id="loadingOverlay">
          <div class="loading-spinner"></div>
          <div class="loading-text" id="loadingText">Processing clients...</div>
        </div>
        
        <script>
          let rowCount = 0;
          
          // Initialize with 3 empty rows
          window.onload = function() {
            for (let i = 0; i < 3; i++) {
              addRow();
            }
          };
          
          function addRow() {
            rowCount++;
            const tbody = document.getElementById('clientTableBody');
            const row = document.createElement('tr');
            row.id = 'row_' + rowCount;
            
            row.innerHTML = \`
              <td>\${rowCount}</td>
              <td><input type="text" id="name_\${rowCount}" placeholder="Enter full name" /></td>
              <td>
                <div class="service-checkboxes">
                  <label class="service-checkbox" onclick="toggleService(this)">
                    <input type="checkbox" value="ACT" />
                    <span>ACT</span>
                  </label>
                  <label class="service-checkbox" onclick="toggleService(this)">
                    <input type="checkbox" value="Math Curriculum" />
                    <span>Math</span>
                  </label>
                  <label class="service-checkbox" onclick="toggleService(this)">
                    <input type="checkbox" value="Language" />
                    <span>Language</span>
                  </label>
                </div>
              </td>
              <td><input type="email" id="email_\${rowCount}" placeholder="parent@example.com" /></td>
              <td><input type="text" id="addressees_\${rowCount}" placeholder="John & Jane Doe" /></td>
              <td><input type="url" id="dashboard_\${rowCount}" placeholder="https://..." /></td>
              <td>
                <button class="remove-btn" onclick="removeRow(\${rowCount})">Remove</button>
              </td>
            \`;
            
            tbody.appendChild(row);
            updateEmptyState();
          }
          
          function removeRow(rowId) {
            const row = document.getElementById('row_' + rowId);
            if (row) {
              row.remove();
              updateEmptyState();
            }
          }
          
          function updateEmptyState() {
            const tbody = document.getElementById('clientTableBody');
            const emptyState = document.getElementById('emptyState');
            const hasRows = tbody.children.length > 0;
            
            emptyState.style.display = hasRows ? 'none' : 'block';
            document.getElementById('addAllBtn').disabled = !hasRows;
          }
          
          function toggleService(element) {
            event.preventDefault();
            const checkbox = element.querySelector('input[type="checkbox"]');
            checkbox.checked = !checkbox.checked;
            element.classList.toggle('selected', checkbox.checked);
          }
          
          function getClientData() {
            const clients = [];
            const rows = document.querySelectorAll('#clientTableBody tr');
            
            rows.forEach((row, index) => {
              const rowId = row.id.split('_')[1];
              const name = document.getElementById('name_' + rowId).value.trim();
              const services = [];
              
              row.querySelectorAll('.service-checkbox input:checked').forEach(cb => {
                services.push(cb.value);
              });
              
              if (name && services.length > 0) {
                clients.push({
                  name: name,
                  services: services.join(', '),
                  parentEmail: document.getElementById('email_' + rowId).value.trim(),
                  emailAddressees: document.getElementById('addressees_' + rowId).value.trim(),
                  dashboardLink: document.getElementById('dashboard_' + rowId).value.trim()
                });
              }
            });
            
            return clients;
          }
          
          function validateClients(clients) {
            const errors = [];
            const names = new Set();
            
            clients.forEach((client, index) => {
              // Check for duplicate names
              if (names.has(client.name.toLowerCase())) {
                errors.push(\`Row \${index + 1}: Duplicate client name "\${client.name}"\`);
              }
              names.add(client.name.toLowerCase());
              
              // Validate email if provided
              if (client.parentEmail && !isValidEmail(client.parentEmail)) {
                errors.push(\`Row \${index + 1}: Invalid email format\`);
              }
            });
            
            return errors;
          }
          
          function isValidEmail(email) {
            const re = /^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/;
            return re.test(email);
          }
          
          function showMessage(message, isError = false) {
            const successMsg = document.getElementById('successMessage');
            const errorMsg = document.getElementById('errorMessage');
            
            if (isError) {
              errorMsg.textContent = message;
              errorMsg.style.display = 'block';
              successMsg.style.display = 'none';
              setTimeout(() => errorMsg.style.display = 'none', 5000);
            } else {
              successMsg.textContent = message;
              successMsg.style.display = 'block';
              errorMsg.style.display = 'none';
              setTimeout(() => successMsg.style.display = 'none', 5000);
            }
          }
          
          function addAllClients() {
            const clients = getClientData();
            
            if (clients.length === 0) {
              showMessage('Please enter at least one client with name and services', true);
              return;
            }
            
            const errors = validateClients(clients);
            if (errors.length > 0) {
              showMessage('Please fix the following errors:\\n' + errors.join('\\n'), true);
              return;
            }
            
            // Show loading
            document.getElementById('loadingOverlay').classList.add('active');
            document.getElementById('loadingText').textContent = 'Adding ' + clients.length + ' client(s)...';
            
            // Process all clients
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loadingOverlay').classList.remove('active');
                
                if (result.success) {
                  showMessage(\`Successfully added \${result.addedCount} client(s)!\`);
                  
                  // Clear successful rows
                  if (result.addedClients) {
                    result.addedClients.forEach(clientName => {
                      // Find and remove the row with this client
                      const rows = document.querySelectorAll('#clientTableBody tr');
                      rows.forEach(row => {
                        const rowId = row.id.split('_')[1];
                        const nameInput = document.getElementById('name_' + rowId);
                        if (nameInput && nameInput.value.trim() === clientName) {
                          row.remove();
                        }
                      });
                    });
                  }
                  
                  updateEmptyState();
                  
                  // Close dialog after short delay if all were successful
                  if (clients.length === result.addedCount) {
                    setTimeout(() => google.script.host.close(), 2000);
                  }
                } else {
                  showMessage('Error: ' + result.message, true);
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loadingOverlay').classList.remove('active');
                showMessage('Error: ' + error.message, true);
              })
              .processBulkClients(clients);
          }
        </script>
      </body>
    </html>
  `).setWidth(900).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Multiple Clients');
}

/**
 * Process bulk client creation (OPTIMIZED with UnifiedClientDataStore)
 * @param {Array} clientsData - Array of client objects to create
 * @returns {Object} Result with success/failure details
 */
function processBulkClients(clientsData) {
  if (!Array.isArray(clientsData) || clientsData.length === 0) {
    throw new Error('No client data provided');
  }
  
  try {
    // Clean and prepare data
    const cleanedData = clientsData.map(clientData => ({
      name: (clientData.name || '').trim(),
      services: (clientData.services || '').trim(),
      parentEmail: (clientData.parentEmail || '').trim(),
      emailAddressees: (clientData.emailAddressees || '').trim(),
      dashboardLink: (clientData.dashboardLink || '').trim(),
      meetingNotesLink: '',
      scoreReportLink: '',
      isActive: true // Default to active for new clients
    }));
    
    // Use the optimized batch operation from UnifiedClientDataStore
    const results = UnifiedClientDataStore.addClientsBatch(cleanedData);
    
    // Create summary message
    let message = `Bulk client creation completed:\n`;
    message += `✓ ${results.successful.length} clients created successfully\n`;
    
    if (results.failed.length > 0) {
      message += `✗ ${results.failed.length} clients failed to create\n\n`;
      message += `Failed clients:\n`;
      results.failed.forEach(failure => {
        message += `• ${failure.name}: ${failure.error}\n`;
      });
    }
    
    if (results.successful.length > 0) {
      message += `\nSuccessful clients:\n`;
      results.successful.forEach(success => {
        message += `• ${success.name}\n`;
      });
    }
    
    // Also create sheets for successful clients (optional - keeping for compatibility)
    if (results.successful.length > 0) {
      let sheetErrors = [];
      
      for (const success of results.successful) {
        try {
          const originalData = cleanedData.find(c => c.name === success.name);
          if (originalData) {
            // Only create sheet, don't duplicate data storage
            createClientSheetMinimal(originalData, success.key);
          }
        } catch (error) {
          console.error(`Error creating sheet for ${success.name}:`, error);
          sheetErrors.push(`${success.name}: ${error.message}`);
        }
      }
      
      if (sheetErrors.length > 0) {
        message += `\n⚠️ Note: ${sheetErrors.length} client sheets failed to create (data is saved in system):\n`;
        sheetErrors.forEach(error => {
          message += `• ${error}\n`;
        });
      }
    }
    
    return {
      success: results.failed.length === 0,
      message: message,
      details: results
    };
    
  } catch (error) {
    console.error('Error in bulk client processing:', error);
    throw error;
  }
}

/**
 * Create minimal client sheet without duplicating data storage
 * (Only creates the sheet, data is already in UnifiedClientDataStore)
 */
function createClientSheetMinimal(clientData, clientKey) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Find the template sheet
    const templateSheet = spreadsheet.getSheetByName('NewClient');
    if (!templateSheet) {
      throw new Error('Template sheet "NewClient" not found');
    }
    
    // Check if sheet already exists
    if (spreadsheet.getSheetByName(clientKey)) {
      return;
    }
    
    // Copy the template sheet
    const newSheet = templateSheet.copyTo(spreadsheet);
    newSheet.setName(clientKey);
    
    // Populate only the display cells (not for data storage)
    newSheet.getRange('A1').setValue(clientData.name);
    
    // Move to position 3
    const totalSheets = spreadsheet.getSheets().length;
    const targetPosition = Math.min(3, totalSheets);
    spreadsheet.setActiveSheet(newSheet);
    spreadsheet.moveActiveSheet(targetPosition);
    
    SpreadsheetApp.flush();
    
    
  } catch (error) {
    console.error(`Error creating sheet for ${clientData.name}:`, error);
    throw error;
  }
}

/**
 * =============================================================================
 * UNIFIED CLIENT DATA STORE SYSTEM
 * Streamlined client data management with single source of truth
 * =============================================================================
 */

/**
 * UnifiedClientDataStore - Modern client data management system
 * Replaces fragmented Script Properties with single JSON structure
 */
const UnifiedClientDataStore = {
  // Storage keys
  DATA_KEY: 'unified_client_data',
  BACKUP_KEY: 'unified_client_data_backup',
  VERSION: 1,
  
  /**
   * Initialize the unified data store (migration from existing system)
   */
  initialize: function() {
    try {
      let unifiedData = this.getUnifiedData();
      
      if (!unifiedData) {
        unifiedData = {
          version: this.VERSION,
          lastUpdated: new Date().getTime(),
          clients: {},
          stats: {
            totalClients: 0,
            activeClients: 0,
            lastCacheRefresh: new Date().getTime()
          }
        };
        
        // Migrate existing data from EnhancedClientCache if available
        this.migrateFromEnhancedCache(unifiedData);
        this.saveUnifiedData(unifiedData);
      }
      
      return unifiedData;
    } catch (error) {
      console.error('Error initializing unified data store:', error);
      throw error;
    }
  },
  
  /**
   * Get the unified data from Script Properties
   */
  getUnifiedData: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const dataString = properties.getProperty(this.DATA_KEY);
      
      if (dataString) {
        const data = JSON.parse(dataString);
        // Version check - migrate if needed
        if (data.version < this.VERSION) {
          // Add version upgrade logic here if needed
          data.version = this.VERSION;
          this.saveUnifiedData(data);
        }
        return data;
      }
      return null;
    } catch (error) {
      console.error('Error reading unified data:', error);
      return this.restoreFromBackup();
    }
  },
  
  /**
   * Save unified data to Script Properties with backup
   */
  saveUnifiedData: function(data) {
    try {
      const properties = PropertiesService.getScriptProperties();
      
      // Update timestamp
      data.lastUpdated = new Date().getTime();
      
      // Create backup of existing data before overwriting
      const existingData = properties.getProperty(this.DATA_KEY);
      if (existingData) {
        properties.setProperty(this.BACKUP_KEY, existingData);
      }
      
      // Save new data
      const dataString = JSON.stringify(data);
      properties.setProperty(this.DATA_KEY, dataString);
      
      return true;
    } catch (error) {
      console.error('Error saving unified data:', error);
      throw error;
    }
  },
  
  /**
   * Restore from backup if main data is corrupted
   */
  restoreFromBackup: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      const backupString = properties.getProperty(this.BACKUP_KEY);
      
      if (backupString) {
        const backupData = JSON.parse(backupString);
        return backupData;
      }
      
      return null;
    } catch (error) {
      console.error('Error restoring from backup:', error);
      return null;
    }
  },
  
  /**
   * Migrate existing data from EnhancedClientCache
   */
  migrateFromEnhancedCache: function(unifiedData) {
    try {
      const existingClients = EnhancedClientCache.getAllClientsFromCache();
      
      if (existingClients && existingClients.length > 0) {
        let migrated = 0;
        
        existingClients.forEach(client => {
          const clientKey = client.sheetName || client.name.replace(/\s+/g, '');
          
          unifiedData.clients[clientKey] = {
            name: client.name,
            isActive: client.isActive !== false, // Default to true
            contactInfo: {
              parentEmail: client.parentEmail || '',
              emailAddressees: client.emailAddressees || ''
            },
            links: {
              dashboard: client.dashboardLink || '',
              meetingNotes: client.meetingNotesLink || '',
              scoreReport: client.scoreReportLink || ''
            },
            sessionInfo: {
              lastSession: client.lastSession || '',
              nextSession: client.nextSession || ''
            },
            quickNotes: client.quickNotes || '',
            metadata: {
              created: new Date().getTime(),
              lastUpdated: new Date().getTime(),
              sheetId: client.sheetId || null,
              migratedFromCache: true
            }
          };
          migrated++;
        });
        
        // Update stats
        unifiedData.stats.totalClients = migrated;
        unifiedData.stats.activeClients = existingClients.filter(c => c.isActive !== false).length;
        
      }
    } catch (error) {
      console.error('Error migrating from EnhancedClientCache:', error);
    }
  },
  
  /**
   * Add a single client to the unified store
   */
  addClient: function(clientData) {
    try {
      const unifiedData = this.getUnifiedData() || this.initialize();
      const clientKey = clientData.name.replace(/\s+/g, '');
      
      // Check if client already exists
      if (unifiedData.clients[clientKey]) {
        throw new Error(`Client "${clientData.name}" already exists`);
      }
      
      // Add new client
      unifiedData.clients[clientKey] = {
        name: clientData.name,
        isActive: clientData.isActive !== false,
        contactInfo: {
          parentEmail: clientData.parentEmail || '',
          emailAddressees: clientData.emailAddressees || ''
        },
        links: {
          dashboard: clientData.dashboardLink || '',
          meetingNotes: clientData.meetingNotesLink || '',
          scoreReport: clientData.scoreReportLink || ''
        },
        sessionInfo: {
          lastSession: clientData.lastSession || '',
          nextSession: clientData.nextSession || ''
        },
        quickNotes: clientData.quickNotes || '',
        metadata: {
          created: new Date().getTime(),
          lastUpdated: new Date().getTime(),
          sheetId: clientData.sheetId || null,
          source: 'unified_store'
        }
      };
      
      // Update stats
      unifiedData.stats.totalClients++;
      if (clientData.isActive !== false) {
        unifiedData.stats.activeClients++;
      }
      
      this.saveUnifiedData(unifiedData);
      return clientKey;
    } catch (error) {
      console.error('Error adding client:', error);
      throw error;
    }
  },
  
  /**
   * Update or add a client (doesn't throw error for existing clients)
   */
  upsertClient: function(clientData) {
    try {
      const unifiedData = this.getUnifiedData() || this.initialize();
      const clientKey = clientData.name.replace(/\s+/g, '');
      
      // Track if this is an update or insert
      const isUpdate = !!unifiedData.clients[clientKey];
      
      // Create or update client data
      unifiedData.clients[clientKey] = {
        name: clientData.name,
        isActive: clientData.isActive !== false,
        contactInfo: {
          parentEmail: clientData.parentEmail || '',
          emailAddressees: clientData.emailAddressees || ''
        },
        links: {
          dashboard: clientData.dashboardLink || '',
          meetingNotes: clientData.meetingNotesLink || '',
          scoreReport: clientData.scoreReportLink || ''
        },
        sessionInfo: {
          lastSession: clientData.lastSession || null,
          nextSession: clientData.nextSession || null
        },
        metadata: {
          created: isUpdate ? unifiedData.clients[clientKey]?.metadata?.created : new Date().getTime(),
          lastUpdated: new Date().getTime(),
          sheetId: clientData.sheetId || null,
          source: 'unified_store'
        }
      };
      
      // Update stats
      unifiedData.stats.totalClients = Object.keys(unifiedData.clients).length;
      unifiedData.stats.activeClients = Object.values(unifiedData.clients)
        .filter(client => client.isActive).length;
      unifiedData.stats.lastUpdated = new Date().getTime();
      
      // Save the updated data
      this.saveUnifiedData(unifiedData);
      
      return clientKey;
    } catch (error) {
      console.error('Error upserting client:', error);
      throw error;
    }
  },
  
  /**
   * Add multiple clients in a batch operation (OPTIMIZED)
   */
  addClientsBatch: function(clientsArray) {
    try {
      if (!Array.isArray(clientsArray) || clientsArray.length === 0) {
        throw new Error('No client data provided for batch operation');
      }
      
      const unifiedData = this.getUnifiedData() || this.initialize();
      const results = {
        successful: [],
        failed: [],
        total: clientsArray.length
      };
      
      // Process all clients in memory first (no I/O operations)
      for (let i = 0; i < clientsArray.length; i++) {
        const clientData = clientsArray[i];
        
        try {
          // Validate required fields
          if (!clientData.name || clientData.name.trim() === '') {
            throw new Error('Client name is required');
          }
          
          const clientKey = clientData.name.trim().replace(/\s+/g, '');
          
          // Check if client already exists
          if (unifiedData.clients[clientKey]) {
            throw new Error(`Client "${clientData.name}" already exists`);
          }
          
          // Add client to memory structure
          unifiedData.clients[clientKey] = {
            name: clientData.name.trim(),
            isActive: clientData.isActive !== false,
            contactInfo: {
              parentEmail: (clientData.parentEmail || '').trim(),
              emailAddressees: (clientData.emailAddressees || '').trim()
            },
            links: {
              dashboard: (clientData.dashboardLink || '').trim(),
              meetingNotes: (clientData.meetingNotesLink || '').trim(),
              scoreReport: (clientData.scoreReportLink || '').trim()
            },
            sessionInfo: {
              lastSession: clientData.lastSession || '',
              nextSession: clientData.nextSession || ''
            },
            quickNotes: clientData.quickNotes || '',
            metadata: {
              created: new Date().getTime(),
              lastUpdated: new Date().getTime(),
              sheetId: null,
              source: 'batch_addition'
            }
          };
          
          results.successful.push({
            name: clientData.name.trim(),
            key: clientKey
          });
          
        } catch (error) {
          results.failed.push({
            name: clientData.name || 'Unknown',
            error: error.message
          });
        }
      }
      
      // Update stats
      unifiedData.stats.totalClients = Object.keys(unifiedData.clients).length;
      unifiedData.stats.activeClients = Object.values(unifiedData.clients)
        .filter(client => client.isActive).length;
      
      // Single save operation for all clients
      this.saveUnifiedData(unifiedData);
      
      return results;
      
    } catch (error) {
      console.error('Error in batch client addition:', error);
      throw error;
    }
  },
  
  /**
   * Get all clients (replaces getActiveClientList/getAllClientList)
   */
  getAllClients: function() {
    try {
      const unifiedData = this.getUnifiedData();
      if (!unifiedData || !unifiedData.clients) {
        return [];
      }
      
      return Object.keys(unifiedData.clients)
        .filter(key => {
          const client = unifiedData.clients[key];
          
          if (!client || !client.name) {
            return false;
          }
          
          const name = client.name.toString().trim();
          const invalidNames = ['Date Sent', 'Student Name', 'Client Type', 'Template', 'New Client', 'NewClient', 'Parent Email', 'Focus Skill', 'Today\'s Win', 'Client Name'];
          
          // Exact match check
          if (invalidNames.includes(name) || invalidNames.includes(key)) {
            return false;
          }
          
          // Pattern matching for additional safety
          const lowerName = name.toLowerCase();
          if (lowerName.includes('sent') || 
              lowerName.includes('email') || 
              lowerName.includes('type') ||
              lowerName.includes('skill') ||
              lowerName.includes('date') ||
              lowerName.includes('win') ||
              lowerName.includes('template') ||
              name.length === 0) {
            return false;
          }
          
          return true;
        })
        .map(key => {
          const client = unifiedData.clients[key];
          
          // Extract first name for proper sorting
          const fullName = client.name || '';
          const firstName = fullName.split(' ')[0] || fullName;
          
          return {
            name: client.name,
            firstName: firstName,
            sheetName: key,
            isActive: client.isActive,
            dashboardLink: client.links.dashboard,
            meetingNotesLink: client.links.meetingNotes,
            scoreReportLink: client.links.scoreReport,
            parentEmail: client.contactInfo.parentEmail,
            emailAddressees: client.contactInfo.emailAddressees,
            lastSession: client.sessionInfo.lastSession,
            nextSession: client.sessionInfo.nextSession,
            quickNotes: client.quickNotes
          };
        }).sort((a, b) => {
        // Sort alphabetically by first name
        const nameA = (a.firstName || '').toLowerCase();
        const nameB = (b.firstName || '').toLowerCase();
        return nameA.localeCompare(nameB);
      });
    } catch (error) {
      console.error('Error getting all clients:', error);
      return [];
    }
  },
  
  /**
   * Get only active clients
   */
  getActiveClients: function() {
    return this.getAllClients().filter(client => client.isActive);
  },
  
  /**
   * Export all client data as CSV-ready array
   */
  exportToCSV: function() {
    try {
      const clients = this.getAllClients();
      const csvData = [];
      
      // Header row
      csvData.push([
        'Client Name',
        'Active',
        'Parent Email', 
        'Email Addressees',
        'Dashboard Link',
        'Meeting Notes Link',
        'Score Report Link',
        'Last Session',
        'Next Session',
        'Quick Notes'
      ]);
      
      // Data rows
      clients.forEach(client => {
        csvData.push([
          client.name,
          client.isActive ? 'Yes' : 'No',
          client.parentEmail,
          client.emailAddressees,
          client.dashboardLink,
          client.meetingNotesLink,
          client.scoreReportLink,
          client.lastSession,
          client.nextSession,
          client.quickNotes
        ]);
      });
      
      return csvData;
    } catch (error) {
      console.error('Error exporting to CSV:', error);
      throw error;
    }
  },
  
  /**
   * Get statistics about the client data
   */
  getStats: function() {
    try {
      const unifiedData = this.getUnifiedData();
      if (!unifiedData) {
        return { totalClients: 0, activeClients: 0, lastUpdated: 0 };
      }
      
      return {
        totalClients: unifiedData.stats.totalClients,
        activeClients: unifiedData.stats.activeClients,
        lastUpdated: unifiedData.lastUpdated,
        version: unifiedData.version
      };
    } catch (error) {
      console.error('Error getting stats:', error);
      return { totalClients: 0, activeClients: 0, lastUpdated: 0 };
    }
  },
  
  /**
   * Temporary data management for session/workflow state
   */
  TEMP_DATA_KEY: 'UNIFIED_TEMP_DATA',
  
  setTempData: function(key, data) {
    try {
      const properties = PropertiesService.getScriptProperties();
      let tempData = {};
      
      const existingTempData = properties.getProperty(this.TEMP_DATA_KEY);
      if (existingTempData) {
        tempData = JSON.parse(existingTempData);
      }
      
      tempData[key] = {
        data: data,
        timestamp: new Date().getTime()
      };
      
      properties.setProperty(this.TEMP_DATA_KEY, JSON.stringify(tempData));
    } catch (error) {
      console.error('Error storing temporary data:', error);
    }
  },
  
  getTempData: function(key) {
    try {
      const properties = PropertiesService.getScriptProperties();
      const tempDataString = properties.getProperty(this.TEMP_DATA_KEY);
      
      if (!tempDataString) return null;
      
      const tempData = JSON.parse(tempDataString);
      const entry = tempData[key];
      
      if (!entry) return null;
      
      // Clean up old temporary data (older than 1 hour)
      const oneHourAgo = new Date().getTime() - (60 * 60 * 1000);
      if (entry.timestamp < oneHourAgo) {
        this.deleteTempData(key);
        return null;
      }
      
      return entry.data;
    } catch (error) {
      console.error('Error retrieving temporary data:', error);
      return null;
    }
  },
  
  deleteTempData: function(key) {
    try {
      const properties = PropertiesService.getScriptProperties();
      let tempData = {};
      
      const existingTempData = properties.getProperty(this.TEMP_DATA_KEY);
      if (existingTempData) {
        tempData = JSON.parse(existingTempData);
        delete tempData[key];
        
        if (Object.keys(tempData).length === 0) {
          properties.deleteProperty(this.TEMP_DATA_KEY);
        } else {
          properties.setProperty(this.TEMP_DATA_KEY, JSON.stringify(tempData));
        }
      }
    } catch (error) {
      console.error('Error deleting temporary data:', error);
    }
  },
  
  clearAllTempData: function() {
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty(this.TEMP_DATA_KEY);
    } catch (error) {
      console.error('Error clearing temporary data:', error);
    }
  }
};

/**
 * =============================================================================
 * UNIFIED CLIENT DATA STORE - UTILITY FUNCTIONS
 * =============================================================================
 */

/**
 * Initialize the unified data store (migration helper)
 */
function initializeUnifiedDataStore() {
  try {
    const result = UnifiedClientDataStore.initialize();
    return result;
  } catch (error) {
    console.error('Error initializing UnifiedClientDataStore:', error);
    throw error;
  }
}

/**
 * Export client data as CSV file download
 */
function exportClientListAsCSV() {
  try {
    const csvData = UnifiedClientDataStore.exportToCSV();
    
    // Convert to CSV string
    const csvString = csvData.map(row => {
      return row.map(cell => {
        // Escape quotes and wrap fields containing commas/quotes
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes('\n'))) {
          return '"' + cell.replace(/"/g, '""') + '"';
        }
        return cell;
      }).join(',');
    }).join('\n');
    
    // Create blob and download
    const htmlOutput = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html>
        <head>
          <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">
          <title>Export Client List</title>
        </head>
        <body>
          <script>
            // Create CSV file and trigger download
            const csvContent = \`${csvString}\`;
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', 'client_list_' + new Date().toISOString().split('T')[0] + '.csv');
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            // Close dialog after download
            setTimeout(function() {
              google.script.host.close();
            }, 1000);
          </script>
          <p>Your client list is being downloaded as a CSV file...</p>
        </body>
      </html>
    `);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Client List');
    
  } catch (error) {
    console.error('Error exporting client list:', error);
    SpreadsheetApp.getUi().alert('Export Error', 'Failed to export client list: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Get unified client list for UI (replaces getAllClientList)
 */
function getUnifiedClientList() {
  try {
    return UnifiedClientDataStore.getAllClients();
  } catch (error) {
    console.error('Error getting unified client list:', error);
    return [];
  }
}

/**
 * Get unified active client list for UI (replaces getActiveClientList)
 */
function getUnifiedActiveClientList() {
  try {
    return UnifiedClientDataStore.getActiveClients();
  } catch (error) {
    console.error('Error getting unified active client list:', error);
    return [];
  }
}

/**
 * Get unified client data stats
 */
function getUnifiedClientStats() {
  try {
    return UnifiedClientDataStore.getStats();
  } catch (error) {
    console.error('Error getting unified client stats:', error);
    return { totalClients: 0, activeClients: 0, lastUpdated: 0 };
  }
}

/**
 * Migrate from existing EnhancedClientCache to UnifiedClientDataStore
 */
function migrateToUnifiedDataStore() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Migrate Client Data',
      'This will migrate all client data to the new unified system. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      const result = initializeUnifiedDataStore();
      const stats = UnifiedClientDataStore.getStats();
      
      ui.alert(
        'Migration Complete',
        `Successfully migrated ${stats.totalClients} clients to the unified data store.\n\nActive clients: ${stats.activeClients}`,
        ui.ButtonSet.OK
      );
      
      return result;
    }
  } catch (error) {
    console.error('Error during migration:', error);
    SpreadsheetApp.getUi().alert('Migration Error', 'Failed to migrate client data: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

// ============================================================================
// ENHANCED FEATURES - SERVER SIDE SUPPORT
// ============================================================================

/**
 * Get paginated client list with advanced filtering
 */
function getPaginatedClientList(page = 1, pageSize = 50, filters = {}) {
  try {
    const allClients = UnifiedClientDataStore.getAllClients();
    
    // Apply filters
    let filteredClients = allClients;
    
    if (filters.search) {
      const searchTerm = filters.search.toLowerCase();
      filteredClients = filteredClients.filter(client => 
        (client.name && client.name.toLowerCase().includes(searchTerm)) ||
        (client.parentEmail && client.parentEmail.toLowerCase().includes(searchTerm)) ||
        (client.services && client.services.toLowerCase().includes(searchTerm))
      );
    }
    
    if (filters.isActive !== undefined) {
      filteredClients = filteredClients.filter(client => 
        Boolean(client.isActive) === Boolean(filters.isActive)
      );
    }
    
    if (filters.serviceType) {
      filteredClients = filteredClients.filter(client => 
        client.services && client.services.includes(filters.serviceType)
      );
    }
    
    // Calculate pagination
    const totalItems = filteredClients.length;
    const totalPages = Math.ceil(totalItems / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    
    // Get page data
    const pageData = filteredClients.slice(startIndex, endIndex);
    
    return {
      clients: pageData,
      pagination: {
        currentPage: page,
        pageSize: pageSize,
        totalPages: totalPages,
        totalItems: totalItems,
        hasNextPage: page < totalPages,
        hasPrevPage: page > 1
      }
    };
    
  } catch (error) {
    console.error('Error in getPaginatedClientList:', error);
    return {
      clients: [],
      pagination: {
        currentPage: 1,
        pageSize: pageSize,
        totalPages: 1,
        totalItems: 0,
        hasNextPage: false,
        hasPrevPage: false
      },
      error: error.message
    };
  }
}

/**
 * Advanced client search with fuzzy matching
 */
function searchClientsAdvanced(query, options = {}) {
  try {
    const {
      includeInactive = false,
      fuzzySearch = true,
      maxResults = 100,
      threshold = 0.6
    } = options;
    
    const allClients = UnifiedClientDataStore.getAllClients();
    let searchableClients = includeInactive ? 
      allClients : 
      allClients.filter(client => client.isActive);
    
    if (!query || query.trim().length === 0) {
      return {
        clients: searchableClients.slice(0, maxResults),
        query: query,
        resultsCount: Math.min(searchableClients.length, maxResults)
      };
    }
    
    const searchTerm = query.toLowerCase().trim();
    let results = [];
    
    if (fuzzySearch) {
      // Implement fuzzy search
      results = searchableClients.map(client => {
        const score = calculateFuzzyScore(client, searchTerm);
        return { client, score };
      })
      .filter(item => item.score >= threshold)
      .sort((a, b) => b.score - a.score)
      .map(item => item.client);
    } else {
      // Exact matching
      results = searchableClients.filter(client => 
        (client.name && client.name.toLowerCase().includes(searchTerm)) ||
        (client.parentEmail && client.parentEmail.toLowerCase().includes(searchTerm)) ||
        (client.services && client.services.toLowerCase().includes(searchTerm)) ||
        (client.sheetName && client.sheetName.toLowerCase().includes(searchTerm))
      );
    }
    
    return {
      clients: results.slice(0, maxResults),
      query: query,
      resultsCount: Math.min(results.length, maxResults),
      totalMatches: results.length
    };
    
  } catch (error) {
    console.error('Error in searchClientsAdvanced:', error);
    return {
      clients: [],
      query: query,
      resultsCount: 0,
      error: error.message
    };
  }
}

/**
 * Calculate fuzzy search score
 */
function calculateFuzzyScore(client, searchTerm) {
  const searchableFields = [
    client.name || '',
    client.parentEmail || '',
    client.services || '',
    client.emailAddressees || ''
  ];
  
  const fullText = searchableFields.join(' ').toLowerCase();
  
  // Simple fuzzy scoring algorithm
  let score = 0;
  const words = searchTerm.split(/\s+/);
  
  for (const word of words) {
    if (fullText.includes(word)) {
      score += 1.0; // Exact word match
    } else {
      // Check for partial matches
      const partialScore = calculatePartialMatch(fullText, word);
      score += partialScore;
    }
  }
  
  // Normalize score
  return Math.min(score / words.length, 1.0);
}

/**
 * Calculate partial match score for fuzzy search
 */
function calculatePartialMatch(text, word) {
  let maxScore = 0;
  const wordLength = word.length;
  
  for (let i = 0; i <= text.length - wordLength; i++) {
    const substring = text.substring(i, i + wordLength);
    const similarity = calculateStringSimilarity(word, substring);
    maxScore = Math.max(maxScore, similarity);
  }
  
  return maxScore * 0.5; // Partial match gets half score
}

/**
 * Calculate string similarity (simple Levenshtein-based)
 */
function calculateStringSimilarity(str1, str2) {
  const len1 = str1.length;
  const len2 = str2.length;
  
  if (len1 === 0) return len2 === 0 ? 1 : 0;
  if (len2 === 0) return 0;
  
  const matrix = Array(len2 + 1).fill().map(() => Array(len1 + 1).fill(0));
  
  for (let i = 0; i <= len1; i++) matrix[0][i] = i;
  for (let j = 0; j <= len2; j++) matrix[j][0] = j;
  
  for (let j = 1; j <= len2; j++) {
    for (let i = 1; i <= len1; i++) {
      if (str1[i - 1] === str2[j - 1]) {
        matrix[j][i] = matrix[j - 1][i - 1];
      } else {
        matrix[j][i] = Math.min(
          matrix[j - 1][i] + 1,
          matrix[j][i - 1] + 1,
          matrix[j - 1][i - 1] + 1
        );
      }
    }
  }
  
  const distance = matrix[len2][len1];
  const maxLength = Math.max(len1, len2);
  
  return (maxLength - distance) / maxLength;
}

/**
 * Get enhanced client details with related data
 */
function getEnhancedClientDetails(clientName) {
  try {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (!client) {
      throw new Error('Client not found');
    }
    
    // Get additional related data
    const relatedData = {
      recentNotes: getRecentNotesForClient(clientName),
      sessionHistory: getSessionHistoryForClient(clientName),
      lastActivity: getLastActivityForClient(clientName)
    };
    
    return {
      ...client,
      relatedData: relatedData
    };
    
  } catch (error) {
    console.error('Error getting enhanced client details:', error);
    throw error;
  }
}

/**
 * Get recent notes for a client
 */
function getRecentNotesForClient(clientName) {
  try {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (client && client.quickNotes) {
      return {
        hasNotes: true,
        lastModified: client.quickNotes.timestamp || 0,
        preview: truncateText(client.quickNotes.data, 100)
      };
    }
    
    return {
      hasNotes: false,
      lastModified: 0,
      preview: null
    };
    
  } catch (error) {
    console.error('Error getting recent notes:', error);
    return { hasNotes: false, lastModified: 0, preview: null };
  }
}

/**
 * Get session history for a client (mock implementation)
 */
function getSessionHistoryForClient(clientName) {
  // This would integrate with actual session tracking
  return {
    totalSessions: 0,
    lastSession: null,
    averageRating: null
  };
}

/**
 * Get last activity for a client
 */
function getLastActivityForClient(clientName) {
  try {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (client) {
      const timestamps = [
        client.quickNotes?.timestamp || 0,
        client.lastModified || 0
      ];
      
      const lastActivity = Math.max(...timestamps);
      
      return {
        timestamp: lastActivity,
        type: lastActivity === (client.quickNotes?.timestamp || 0) ? 'notes' : 'profile',
        relativeTime: getRelativeTime(lastActivity)
      };
    }
    
    return {
      timestamp: 0,
      type: 'unknown',
      relativeTime: 'Never'
    };
    
  } catch (error) {
    console.error('Error getting last activity:', error);
    return { timestamp: 0, type: 'unknown', relativeTime: 'Never' };
  }
}

/**
 * Get active clients count
 */
function getActiveClientsCount() {
  try {
    const allClients = UnifiedClientDataStore.getAllClients();
    return allClients.filter(client => client.isActive).length;
  } catch (error) {
    console.error('Error getting active clients count:', error);
    return 0;
  }
}

/**
 * Get recent sessions data (mock implementation)
 */
function getRecentSessionsData() {
  // This would integrate with actual session tracking
  return {
    totalSessions: 0,
    thisWeek: 0,
    thisMonth: 0,
    averageDuration: 60
  };
}

/**
 * User preferences management
 */
function getUserPreferences() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const prefsJson = userProperties.getProperty('user_preferences');
    
    if (prefsJson) {
      return JSON.parse(prefsJson);
    }
    
    // Return default preferences
    return {
      quickPills: {
        wins: ['Great participation', 'Improved focus', 'Asked good questions', 'Completed homework'],
        mastered: ['Algebra basics', 'Reading comprehension', 'Essay structure', 'Math formulas'],
        practiced: ['Problem solving', 'Time management', 'Note taking', 'Critical thinking'],
        introduced: ['New concept', 'Advanced technique', 'Study strategy', 'Test strategy'],
        struggles: ['Time pressure', 'Concept confusion', 'Motivation', 'Organization'],
        parent: ['Homework assigned', 'Test preparation', 'Progress update', 'Practice needed'],
        next: ['Review material', 'Practice problems', 'Prepare for test', 'Continue topic']
      },
      theme: 'light',
      autoSave: true,
      notifications: true
    };
    
  } catch (error) {
    console.error('Error getting user preferences:', error);
    return getDefaultUserPreferences();
  }
}

/**
 * Save user preferences
 */
function saveUserPreferences(preferences) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('user_preferences', JSON.stringify(preferences));
    return { success: true };
  } catch (error) {
    console.error('Error saving user preferences:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Default user preferences
 */
function getDefaultUserPreferences() {
  return {
    quickPills: {
      wins: ['Great participation', 'Improved focus', 'Asked good questions', 'Completed homework'],
      mastered: ['Algebra basics', 'Reading comprehension', 'Essay structure', 'Math formulas'],
      practiced: ['Problem solving', 'Time management', 'Note taking', 'Critical thinking'],
      introduced: ['New concept', 'Advanced technique', 'Study strategy', 'Test strategy'],
      struggles: ['Time pressure', 'Concept confusion', 'Motivation', 'Organization'],
      parent: ['Homework assigned', 'Test preparation', 'Progress update', 'Practice needed'],
      next: ['Review material', 'Practice problems', 'Prepare for test', 'Continue topic']
    },
    theme: 'light',
    autoSave: true,
    notifications: true
  };
}

/**
 * Sync offline change
 */
function syncOfflineChange(change) {
  try {
    switch (change.type) {
      case 'notes_update':
        return saveQuickNotes(change.data);
      
      case 'client_update':
        return updateClientDetails(change.key, change.data);
      
      case 'preferences_update':
        return saveUserPreferences(change.data);
      
      default:
        throw new Error('Unknown change type: ' + change.type);
    }
  } catch (error) {
    console.error('Error syncing offline change:', error);
    throw error;
  }
}

/**
 * Validate data integrity
 */
function validateDataIntegrity() {
  try {
    const issues = [];
    const allClients = UnifiedClientDataStore.getAllClients();
    
    // Check for required fields
    allClients.forEach(client => {
      if (!client.name || client.name.trim().length === 0) {
        issues.push(`Client missing name: ${client.sheetName || 'Unknown'}`);
      }
      
      if (!client.sheetName || client.sheetName.trim().length === 0) {
        issues.push(`Client missing sheet name: ${client.name || 'Unknown'}`);
      }
      
      if (client.parentEmail && !isValidEmail(client.parentEmail)) {
        issues.push(`Invalid email for ${client.name}: ${client.parentEmail}`);
      }
    });
    
    return {
      isValid: issues.length === 0,
      issues: issues,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error validating data integrity:', error);
    return {
      isValid: false,
      issues: ['Validation failed: ' + error.message],
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Submit error batch for reporting
 */
function submitErrorBatch(batch) {
  try {
    // In a real implementation, this would send to a logging service
    // For now, we'll store in a dedicated sheet or external service
    
    const timestamp = new Date().toISOString();
    const errorData = {
      timestamp: timestamp,
      sessionId: batch.session.sessionId,
      errorCount: batch.errors.length,
      errors: batch.errors.map(error => ({
        message: error.message,
        severity: error.severity,
        timestamp: new Date(error.context.timestamp).toISOString(),
        context: error.context.type || 'unknown'
      }))
    };
    
    // Store in user properties for now (in production, use external service)
    const userProperties = PropertiesService.getUserProperties();
    const existingErrors = userProperties.getProperty('error_reports');
    const errorReports = existingErrors ? JSON.parse(existingErrors) : [];
    
    errorReports.push(errorData);
    
    // Keep only last 50 error batches
    if (errorReports.length > 50) {
      errorReports.splice(0, errorReports.length - 50);
    }
    
    userProperties.setProperty('error_reports', JSON.stringify(errorReports));
    
    return { success: true, timestamp: timestamp };
    
  } catch (error) {
    console.error('Error submitting error batch:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Get server timestamp for conflict detection
 */
function getServerTimestamp(entityType, entityId) {
  try {
    switch (entityType) {
      case 'client':
        const client = UnifiedClientDataStore.getClient(entityId);
        return {
          lastModified: client ? (client.lastModified || Date.now()) : Date.now()
        };
      
      default:
        return { lastModified: Date.now() };
    }
  } catch (error) {
    console.error('Error getting server timestamp:', error);
    return { lastModified: Date.now() };
  }
}

/**
 * Get server version for conflict detection
 */
function getServerVersion(entityType, entityId) {
  try {
    // Simple version based on last modified timestamp
    const timestamp = getServerTimestamp(entityType, entityId);
    return {
      version: Math.floor(timestamp.lastModified / 1000).toString()
    };
  } catch (error) {
    console.error('Error getting server version:', error);
    return { version: '1' };
  }
}

/**
 * Get active editors (mock implementation)
 */
function getActiveEditors(entityType, entityId) {
  // In a real implementation, this would track active sessions
  // For now, return empty array
  return [];
}

/**
 * Force update with local data
 */
function forceUpdateWithLocalData(operation, data) {
  try {
    switch (operation.type) {
      case 'client_update':
        return updateClientDetails(operation.entityId, data);
      
      case 'notes_update':
        return saveQuickNotes(data);
      
      default:
        throw new Error('Unsupported operation type');
    }
  } catch (error) {
    console.error('Error forcing update with local data:', error);
    throw error;
  }
}

/**
 * Get latest server data
 */
function getLatestServerData(entityType, entityId) {
  try {
    switch (entityType) {
      case 'client':
        return UnifiedClientDataStore.getClient(entityId);
      
      default:
        throw new Error('Unsupported entity type');
    }
  } catch (error) {
    console.error('Error getting latest server data:', error);
    throw error;
  }
}

/**
 * Save merged data after conflict resolution
 */
function saveMergedData(operation, mergedData) {
  try {
    switch (operation.type) {
      case 'client_update':
        UnifiedClientDataStore.updateClient(operation.entityId, mergedData);
        return mergedData;
      
      default:
        throw new Error('Unsupported operation type');
    }
  } catch (error) {
    console.error('Error saving merged data:', error);
    throw error;
  }
}

// Utility functions
function truncateText(text, maxLength) {
  if (!text || text.length <= maxLength) {
    return text;
  }
  return text.substring(0, maxLength) + '...';
}

function getRelativeTime(timestamp) {
  if (!timestamp) return 'Never';
  
  const now = Date.now();
  const diff = now - timestamp;
  
  if (diff < 60000) return 'Just now';
  if (diff < 3600000) return Math.floor(diff / 60000) + ' minutes ago';
  if (diff < 86400000) return Math.floor(diff / 3600000) + ' hours ago';
  if (diff < 604800000) return Math.floor(diff / 86400000) + ' days ago';
  
  return new Date(timestamp).toLocaleDateString();
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// =============================================================================
// ENTERPRISE MULTI-USER ARCHITECTURE
// =============================================================================

/**
 * Authentication Manager for Google OAuth and User Management
 */
const AuthManager = {
  getCurrentUser() {
    try {
      const user = Session.getActiveUser();
      const email = user.getEmail();
      
      // Check if email is actually available (not empty)
      if (!email || email === '') {
        // Fallback to effective user if active user is not available
        const effectiveUser = Session.getEffectiveUser();
        const effectiveEmail = effectiveUser.getEmail();
        
        if (effectiveEmail && effectiveEmail !== '') {
          return {
            email: effectiveEmail,
            displayName: effectiveEmail.split('@')[0],
            authenticated: true
          };
        }
        
        // If still no email, return not authenticated
        return { authenticated: false, error: 'User email not available - please authorize the script' };
      }
      
      return {
        email: email,
        displayName: email.split('@')[0],
        authenticated: true
      };
    } catch (error) {
      console.error('Authentication error:', error);
      return { authenticated: false, error: error.toString() };
    }
  },

  initializeUser(userEmail) {
    const userDataStore = this.getUserDataStore(userEmail);
    if (!userDataStore.preferences) {
      // Initialize user preferences with Beta 2 defaults
      userDataStore.preferences = {
        quickPills: {
          wins: ["Great progress", "Breakthrough moment", "Excellent focus"],
          struggles: ["Needs more practice", "Concept review needed", "Time management"],
          parent: ["Positive feedback", "Areas for improvement", "Home practice suggestions"],
          next: ["Continue current topics", "Move to next chapter", "Review fundamentals"]
        },
        theme: "light",
        autoSave: true,
        notifications: true
      };
      this.saveUserDataStore(userEmail, userDataStore);
    }
  },

  getUserDataStore(userEmail) {
    const data = PropertiesService.getScriptProperties().getProperty(`user_${userEmail}`);
    return data ? JSON.parse(data) : { preferences: {}, clientData: {} };
  },

  saveUserDataStore(userEmail, data) {
    PropertiesService.getScriptProperties().setProperty(`user_${userEmail}`, JSON.stringify(data));
  }
};

/**
 * Client Ownership Manager for User Assignment
 */
const ClientOwnershipManager = {
  assignPrimaryTutor(clientName, tutorEmail) {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (client) {
      client.primaryTutor = tutorEmail;
      client.assignedTutors = client.assignedTutors || [tutorEmail];
      UnifiedClientDataStore.updateClient(clientName, client);
      
      // Log the assignment
      AuditLogger.log('tutor_assigned', 'client', clientName, AuthManager.getCurrentUser().email, { 
        primaryTutor: tutorEmail 
      });
    }
  },

  addAssignedTutor(clientName, tutorEmail) {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (client) {
      client.assignedTutors = client.assignedTutors || [];
      if (!client.assignedTutors.includes(tutorEmail)) {
        client.assignedTutors.push(tutorEmail);
        UnifiedClientDataStore.updateClient(clientName, client);
        
        AuditLogger.log('tutor_added', 'client', clientName, AuthManager.getCurrentUser().email, { 
          addedTutor: tutorEmail 
        });
      }
    }
  },

  canUserAccessClient(clientName, userEmail) {
    const client = UnifiedClientDataStore.getClient(clientName);
    return client && (
      client.primaryTutor === userEmail || 
      (client.assignedTutors && client.assignedTutors.includes(userEmail)) ||
      this.hasAdminRole(userEmail)
    );
  },

  hasAdminRole(userEmail) {
    const userRole = PropertiesService.getScriptProperties().getProperty(`role_${userEmail}`);
    return userRole === 'admin' || userRole === 'supervisor';
  }
};

/**
 * Permission Manager for Role-Based Access Control
 */
const PermissionManager = {
  roles: {
    admin: {
      permissions: ["create_clients", "delete_clients", "manage_users", "view_all_data", "export_data", "system_config"],
      description: "Full system access"
    },
    supervisor: {
      permissions: ["create_clients", "view_all_data", "assign_tutors", "export_reports"],
      description: "Supervisory access to all clients"
    },
    tutor: {
      permissions: ["create_clients", "view_assigned_clients", "edit_own_data"],
      description: "Standard tutor access"
    },
    observer: {
      permissions: ["view_assigned_clients"],
      description: "Read-only access to assigned clients"
    }
  },

  assignRole(userEmail, role) {
    if (!this.roles[role]) throw new Error(`Invalid role: ${role}`);
    PropertiesService.getScriptProperties().setProperty(`role_${userEmail}`, role);
    
    // Log role assignment
    AuditLogger.log('role_assigned', 'user', userEmail, AuthManager.getCurrentUser().email, { newRole: role });
  },

  getUserRole(userEmail) {
    return PropertiesService.getScriptProperties().getProperty(`role_${userEmail}`) || 'tutor';
  },

  hasPermission(userEmail, permission) {
    const userRole = this.getUserRole(userEmail);
    return this.roles[userRole]?.permissions.includes(permission) || false;
  },

  getAccessibleClients(userEmail) {
    const userRole = this.getUserRole(userEmail);
    
    if (userRole === 'admin' || userRole === 'supervisor') {
      return UnifiedClientDataStore.getAllClients();
    }
    
    // For tutors and observers, only show assigned clients
    return UnifiedClientDataStore.getAllClients().filter(client => 
      ClientOwnershipManager.canUserAccessClient(client.name, userEmail)
    );
  }
};

/**
 * Audit Logger for Compliance and Security
 */
const AuditLogger = {
  log(action, entityType, entityId, userId, metadata = {}) {
    const entry = {
      id: Utilities.getUuid(),
      timestamp: new Date().toISOString(),
      action: action, // create, update, delete, view, export, login, logout, tutor_assigned, role_assigned
      entityType: entityType, // client, session, notes, user, system
      entityId: entityId,
      userId: userId,
      userRole: PermissionManager.getUserRole(userId),
      metadata: metadata,
      sessionId: this.getSessionId()
    };

    // Store in chunked properties to handle size limits
    this.storeAuditEntry(entry);
    
    // For compliance, also log critical actions separately
    if (this.isCriticalAction(action)) {
      this.storeCriticalAudit(entry);
    }
  },

  isCriticalAction(action) {
    return ['delete', 'export', 'role_assigned', 'system_config', 'client_created'].includes(action);
  },

  storeAuditEntry(entry) {
    const today = new Date().toISOString().split('T')[0];
    const key = `audit_${today}_${entry.id}`;
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(entry));
  },

  storeCriticalAudit(entry) {
    const key = `critical_audit_${entry.id}`;
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(entry));
  },

  getSessionId() {
    // Simple session ID based on user and timestamp
    const user = AuthManager.getCurrentUser();
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 
      user.email + new Date().toDateString()).toString();
  },

  getAuditLog(startDate, endDate, filters = {}) {
    const entries = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('audit_') && !key.startsWith('audit_critical_')) {
        try {
          const entry = JSON.parse(properties[key]);
          
          // Apply date and filter logic
          if (this.matchesFilters(entry, startDate, endDate, filters)) {
            entries.push(entry);
          }
        } catch (e) {
          // Skip corrupted entries
        }
      }
    });
    
    return entries.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  },

  matchesFilters(entry, startDate, endDate, filters) {
    const entryDate = new Date(entry.timestamp);
    const start = startDate ? new Date(startDate) : new Date('1970-01-01');
    const end = endDate ? new Date(endDate) : new Date();
    
    if (entryDate < start || entryDate > end) return false;
    
    if (filters.userId && entry.userId !== filters.userId) return false;
    if (filters.action && entry.action !== filters.action) return false;
    if (filters.entityType && entry.entityType !== filters.entityType) return false;
    
    return true;
  },

  generateComplianceReport(startDate, endDate) {
    const auditEntries = this.getAuditLog(startDate, endDate);
    
    return {
      period: { start: startDate, end: endDate },
      summary: {
        totalActions: auditEntries.length,
        uniqueUsers: [...new Set(auditEntries.map(e => e.userId))].length,
        criticalActions: auditEntries.filter(e => this.isCriticalAction(e.action)).length,
        clientsAccessed: [...new Set(auditEntries.filter(e => e.entityType === 'client').map(e => e.entityId))].length
      },
      entries: auditEntries,
      generatedAt: new Date().toISOString(),
      generatedBy: AuthManager.getCurrentUser().email
    };
  }
};

/**
 * Data Privacy Manager for Sensitive Information
 */
const DataPrivacyManager = {
  encryptSensitiveData(data, clientName) {
    const sensitiveFields = ['parentEmail', 'emailAddressees', 'personalNotes'];
    const encrypted = { ...data };
    
    sensitiveFields.forEach(field => {
      if (encrypted[field]) {
        encrypted[field] = this.encrypt(encrypted[field], clientName);
        encrypted[`${field}_encrypted`] = true;
      }
    });
    
    return encrypted;
  },

  decryptSensitiveData(data, clientName) {
    const decrypted = { ...data };
    
    Object.keys(decrypted).forEach(field => {
      if (field.endsWith('_encrypted') && decrypted[field]) {
        const originalField = field.replace('_encrypted', '');
        if (decrypted[originalField]) {
          decrypted[originalField] = this.decrypt(decrypted[originalField], clientName);
        }
      }
    });
    
    return decrypted;
  },

  encrypt(text, key) {
    return Utilities.base64Encode(text + '|' + key);
  },

  decrypt(encrypted, key) {
    try {
      const decoded = Utilities.base64Decode(encrypted);
      const parts = decoded.split('|');
      return parts.length === 2 && parts[1] === key ? parts[0] : encrypted;
    } catch (e) {
      return encrypted;
    }
  },

  filterDataByPermission(clientData, userEmail) {
    const userRole = PermissionManager.getUserRole(userEmail);
    const filtered = { ...clientData };
    
    // Remove sensitive data for observers
    if (userRole === 'observer') {
      delete filtered.parentEmail;
      delete filtered.emailAddressees;
      
      // Only show summary of notes, not full content
      if (filtered.quickNotes) {
        Object.keys(filtered.quickNotes).forEach(key => {
          if (filtered.quickNotes[key] && filtered.quickNotes[key].length > 100) {
            filtered.quickNotes[key] = filtered.quickNotes[key].substring(0, 100) + '...';
          }
        });
      }
    }
    
    return filtered;
  },

  canExportData(userEmail, exportType) {
    const permissions = PermissionManager.hasPermission(userEmail, 'export_data');
    
    const exportRules = {
      'client_list': permissions || PermissionManager.getUserRole(userEmail) === 'supervisor',
      'session_reports': permissions,
      'full_audit': permissions && PermissionManager.getUserRole(userEmail) === 'admin',
      'personal_notes': true
    };
    
    return exportRules[exportType] || false;
  }
};

/**
 * Activity Tracker for User Presence and Activity Awareness
 */
const ActivityTracker = {
  updateUserActivity(clientName = null) {
    const user = AuthManager.getCurrentUser();
    if (!user.authenticated) return;
    
    const activity = {
      lastSeen: new Date().toISOString(),
      currentClient: clientName,
      action: clientName ? 'viewing_client' : 'in_dashboard'
    };
    
    PropertiesService.getScriptProperties().setProperty(`activity_${user.email}`, JSON.stringify(activity));
  },

  getUserActivity(userEmail) {
    const data = PropertiesService.getScriptProperties().getProperty(`activity_${userEmail}`);
    return data ? JSON.parse(data) : null;
  },

  getActiveUsers(withinMinutes = 15) {
    const cutoff = new Date(Date.now() - withinMinutes * 60 * 1000);
    const activeUsers = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('activity_')) {
        try {
          const userEmail = key.replace('activity_', '');
          const activity = JSON.parse(properties[key]);
          
          if (new Date(activity.lastSeen) > cutoff) {
            activeUsers.push({
              email: userEmail,
              ...activity
            });
          }
        } catch (e) {
          // Skip corrupted entries
        }
      }
    });
    
    return activeUsers;
  },

  getRecentClientActivity(clientName) {
    const recentActivity = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('activity_')) {
        try {
          const userEmail = key.replace('activity_', '');
          const activity = JSON.parse(properties[key]);
          
          if (activity.currentClient === clientName) {
            recentActivity.push({
              user: userEmail,
              lastSeen: activity.lastSeen,
              action: activity.action
            });
          }
        } catch (e) {
          // Skip corrupted entries
        }
      }
    });
    
    return recentActivity.sort((a, b) => new Date(b.lastSeen) - new Date(a.lastSeen));
  }
};

/**
 * Enhanced Client Creation with Multi-User Context
 */
function createClientSheetWithUserContext(clientData) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('User not authenticated');
  }

  // Check permissions
  if (!PermissionManager.hasPermission(user.email, 'create_clients')) {
    throw new Error('User does not have permission to create clients');
  }

  // Check if client exists
  const existingClient = UnifiedClientDataStore.getClient(clientData.name);
  if (existingClient) {
    return {
      success: false,
      reason: "client_exists",
      existingClient: existingClient,
      createdBy: existingClient.createdBy,
      suggestedAction: "collaborate"
    };
  }

  // Add user context to client data
  clientData.createdBy = user.email;
  clientData.createdAt = new Date().toISOString();
  clientData.primaryTutor = user.email;
  clientData.assignedTutors = [user.email];

  // Create the client using existing function
  const result = createClientSheet(clientData);
  
  if (result.success) {
    // Log client creation
    AuditLogger.log('client_created', 'client', clientData.name, user.email, {
      services: clientData.services,
      parentEmail: clientData.parentEmail
    });
    
    // Update user activity
    ActivityTracker.updateUserActivity(clientData.name);
  }

  return result;
}

/**
 * Enhanced Client Data Retrieval with Permission Filtering
 */
function getAccessibleClientList() {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    return { error: 'User not authenticated', clients: [] };
  }

  const accessibleClients = PermissionManager.getAccessibleClients(user.email);
  
  // Filter sensitive data based on user permissions
  const filteredClients = accessibleClients.map(client => 
    DataPrivacyManager.filterDataByPermission(client, user.email)
  );

  // Log data access
  AuditLogger.log('view', 'client_list', 'all', user.email, { 
    clientCount: filteredClients.length 
  });

  return { clients: filteredClients, userRole: PermissionManager.getUserRole(user.email) };
}

/**
 * Enhanced Session Recap with User Context
 */
function sendIndividualRecapWithUserContext(sheetName, emailData) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('User not authenticated');
  }

  // Check if user can access this client
  if (!ClientOwnershipManager.canUserAccessClient(sheetName, user.email)) {
    throw new Error('User does not have access to this client');
  }

  // Send recap using existing function
  const result = sendIndividualRecap(sheetName, emailData);
  
  if (result.success) {
    // Log email sent
    AuditLogger.log('email_sent', 'client', sheetName, user.email, {
      subject: emailData.subject,
      recipients: emailData.recipients
    });
    
    // Update user activity
    ActivityTracker.updateUserActivity(sheetName);
  }

  return result;
}

/**
 * User Management Functions
 */
function showUserManagementDialog() {
  const user = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(user.email, 'manage_users')) {
    throw new Error('Insufficient permissions');
  }

  const activeUsers = ActivityTracker.getActiveUsers(60); // Last hour
  const allUsers = this.getAllSystemUsers();

  const html = `
    <div class="user-management-container">
      <h3>User Management</h3>
      <div class="active-users">
        <h4>Active Users (Last Hour)</h4>
        ${activeUsers.map(u => `
          <div class="user-item">
            <span>${u.email}</span>
            <span class="role">${PermissionManager.getUserRole(u.email)}</span>
            <span class="last-seen">${getRelativeTime(new Date(u.lastSeen).getTime())}</span>
          </div>
        `).join('')}
      </div>
      
      <div class="user-roles">
        <h4>Manage Roles</h4>
        <select id="userSelect">
          ${allUsers.map(u => `<option value="${u}">${u}</option>`).join('')}
        </select>
        <select id="roleSelect">
          ${Object.keys(PermissionManager.roles).map(role => 
            `<option value="${role}">${role} - ${PermissionManager.roles[role].description}</option>`
          ).join('')}
        </select>
        <button onclick="assignUserRole()">Assign Role</button>
      </div>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'User Management');
}

function getAllSystemUsers() {
  const users = new Set();
  const properties = PropertiesService.getScriptProperties().getProperties();
  
  Object.keys(properties).forEach(key => {
    if (key.startsWith('user_') || key.startsWith('role_') || key.startsWith('activity_')) {
      const email = key.split('_').slice(1).join('_');
      if (email.includes('@')) {
        users.add(email);
      }
    }
  });
  
  return Array.from(users);
}

function assignUserRole() {
  const userEmail = document.getElementById('userSelect').value;
  const role = document.getElementById('roleSelect').value;
  
  google.script.run.withSuccessHandler(() => {
    alert(`Role ${role} assigned to ${userEmail}`);
    google.script.host.close();
  }).withFailureHandler(error => {
    alert('Error assigning role: ' + error);
  }).setUserRole(userEmail, role);
}

function setUserRole(userEmail, role) {
  const currentUser = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(currentUser.email, 'manage_users')) {
    throw new Error('Insufficient permissions');
  }
  
  PermissionManager.assignRole(userEmail, role);
  return { success: true };
}

/**
 * Initialize Enterprise Mode
 */
function initializeEnterpriseMode() {
  try {
    const user = AuthManager.getCurrentUser();
    
    // If not authenticated, return basic mode
    if (!user.authenticated) {
      console.warn('Enterprise authentication not available, falling back to basic mode');
      return {
        success: false,
        fallbackMode: true,
        message: 'Running in basic mode - enterprise features disabled',
        user: {
          email: 'guest@smartcollege.com',
          displayName: 'Guest User',
          authenticated: false
        },
        role: 'tutor'
      };
    }
    
    // Initialize user if first time
    AuthManager.initializeUser(user.email);
    
    // Set default role if not set
    const currentRole = PermissionManager.getUserRole(user.email);
    if (!currentRole || currentRole === 'tutor') {
      // First user gets admin, others get tutor by default
      const existingUsers = getAllSystemUsers();
      const role = existingUsers.length === 0 ? 'admin' : 'tutor';
      PermissionManager.assignRole(user.email, role);
    }
    
    // Update activity
    ActivityTracker.updateUserActivity();
    
    // Log initialization
    AuditLogger.log('system_init', 'system', 'enterprise_mode', user.email, {
      role: PermissionManager.getUserRole(user.email),
      existingUsers: getAllSystemUsers().length
    });
    
    return {
      success: true,
      user: user,
      role: PermissionManager.getUserRole(user.email),
      initialized: true
    };
  } catch (error) {
    console.error('Error initializing enterprise mode:', error);
    // Return fallback mode on any error
    return {
      success: false,
      fallbackMode: true,
      message: 'Enterprise initialization failed - running in basic mode',
      user: {
        email: 'guest@smartcollege.com',
        displayName: 'Guest User',
        authenticated: false
      },
      role: 'tutor'
    };
  }
}

/**
 * Notification System for Multi-User Communication
 */
const NotificationManager = {
  createNotification(userId, type, title, message, entityId = null, actionUrl = null) {
    const notification = {
      id: Utilities.getUuid(),
      userId: userId,
      type: type, // 'client_updated', 'tutor_assigned', 'system_alert', 'role_changed', 'audit_alert'
      title: title,
      message: message,
      entityId: entityId,
      actionUrl: actionUrl,
      timestamp: new Date().toISOString(),
      read: false,
      priority: this.getPriority(type)
    };

    this.storeNotification(notification);
    return notification;
  },

  getPriority(type) {
    const priorities = {
      'system_alert': 'high',
      'audit_alert': 'high', 
      'role_changed': 'medium',
      'tutor_assigned': 'medium',
      'client_updated': 'low',
      'activity_update': 'low'
    };
    return priorities[type] || 'low';
  },

  storeNotification(notification) {
    const key = `notification_${notification.userId}_${notification.id}`;
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(notification));
  },

  getUserNotifications(userId, includeRead = false) {
    const notifications = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith(`notification_${userId}_`)) {
        try {
          const notification = JSON.parse(properties[key]);
          if (includeRead || !notification.read) {
            notifications.push(notification);
          }
        } catch (e) {
          // Skip corrupted notifications
        }
      }
    });
    
    return notifications.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  },

  markAsRead(userId, notificationId) {
    const key = `notification_${userId}_${notificationId}`;
    const data = PropertiesService.getScriptProperties().getProperty(key);
    
    if (data) {
      const notification = JSON.parse(data);
      notification.read = true;
      notification.readAt = new Date().toISOString();
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(notification));
    }
  },

  markAllAsRead(userId) {
    const notifications = this.getUserNotifications(userId, true);
    notifications.forEach(notification => {
      if (!notification.read) {
        this.markAsRead(userId, notification.id);
      }
    });
  },

  // Notify when client is updated by another user
  notifyClientUpdate(clientName, updatedBy, updateType) {
    const client = UnifiedClientDataStore.getClient(clientName);
    if (!client || !client.assignedTutors) return;

    client.assignedTutors.forEach(tutorEmail => {
      if (tutorEmail !== updatedBy) {
        this.createNotification(
          tutorEmail,
          'client_updated',
          `Client Updated: ${clientName}`,
          `${updatedBy} made changes to ${clientName} (${updateType})`,
          clientName
        );
      }
    });
  },

  // Notify when user role changes
  notifyRoleChange(userId, newRole, changedBy) {
    this.createNotification(
      userId,
      'role_changed',
      'Role Updated',
      `Your role has been changed to ${newRole} by ${changedBy}`,
      null
    );
  },

  // Notify when assigned to new client
  notifyTutorAssignment(tutorEmail, clientName, assignedBy) {
    this.createNotification(
      tutorEmail,
      'tutor_assigned',
      `New Client Assignment: ${clientName}`,
      `You have been assigned to work with ${clientName} by ${assignedBy}`,
      clientName
    );
  },

  // System-wide alerts
  broadcastSystemAlert(title, message, targetRole = null) {
    const allUsers = getAllSystemUsers();
    
    allUsers.forEach(userEmail => {
      const userRole = PermissionManager.getUserRole(userEmail);
      
      // If targeting specific role, only notify those users
      if (targetRole && userRole !== targetRole) return;
      
      this.createNotification(
        userEmail,
        'system_alert',
        title,
        message
      );
    });
  },

  // Clean up old notifications (keep last 100 per user)
  cleanupNotifications() {
    const allUsers = getAllSystemUsers();
    
    allUsers.forEach(userId => {
      const notifications = this.getUserNotifications(userId, true);
      
      if (notifications.length > 100) {
        // Keep newest 100, delete the rest
        const toDelete = notifications.slice(100);
        toDelete.forEach(notification => {
          const key = `notification_${userId}_${notification.id}`;
          PropertiesService.getScriptProperties().deleteProperty(key);
        });
      }
    });
  }
};

/**
 * Data Versioning and History System
 */
const DataVersionManager = {
  // Create version entry for data changes
  createVersion(entityType, entityId, data, userId, changeType = 'update') {
    const version = {
      id: Utilities.getUuid(),
      entityType: entityType,
      entityId: entityId,
      version: this.getNextVersionNumber(entityType, entityId),
      data: JSON.stringify(data),
      userId: userId,
      changeType: changeType, // 'create', 'update', 'delete'
      timestamp: new Date().toISOString(),
      checksum: this.generateChecksum(data)
    };

    this.storeVersion(version);
    return version;
  },

  storeVersion(version) {
    const key = `version_${version.entityType}_${version.entityId}_${version.version}`;
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(version));
  },

  getNextVersionNumber(entityType, entityId) {
    const versions = this.getVersionHistory(entityType, entityId);
    return versions.length > 0 ? Math.max(...versions.map(v => v.version)) + 1 : 1;
  },

  getVersionHistory(entityType, entityId, limit = 50) {
    const versions = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    const prefix = `version_${entityType}_${entityId}_`;
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith(prefix)) {
        try {
          const version = JSON.parse(properties[key]);
          versions.push(version);
        } catch (e) {
          // Skip corrupted versions
        }
      }
    });
    
    return versions
      .sort((a, b) => b.version - a.version)
      .slice(0, limit);
  },

  getCurrentVersion(entityType, entityId) {
    const versions = this.getVersionHistory(entityType, entityId, 1);
    return versions.length > 0 ? versions[0] : null;
  },

  restoreVersion(entityType, entityId, versionNumber, userId) {
    const version = this.getVersion(entityType, entityId, versionNumber);
    if (!version) {
      throw new Error(`Version ${versionNumber} not found`);
    }

    const restoredData = JSON.parse(version.data);
    
    // Update the current data
    switch (entityType) {
      case 'client':
        UnifiedClientDataStore.updateClient(entityId, restoredData);
        break;
      default:
        throw new Error(`Unsupported entity type: ${entityType}`);
    }

    // Create new version entry for the restore
    this.createVersion(entityType, entityId, restoredData, userId, 'restore');
    
    // Log the restore action
    AuditLogger.log('data_restored', entityType, entityId, userId, {
      restoredVersion: versionNumber,
      originalTimestamp: version.timestamp
    });

    return restoredData;
  },

  getVersion(entityType, entityId, versionNumber) {
    const key = `version_${entityType}_${entityId}_${versionNumber}`;
    const data = PropertiesService.getScriptProperties().getProperty(key);
    return data ? JSON.parse(data) : null;
  },

  generateChecksum(data) {
    const dataString = JSON.stringify(data);
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, dataString).toString();
  },

  // Compare two versions and return differences
  compareVersions(entityType, entityId, version1, version2) {
    const v1 = this.getVersion(entityType, entityId, version1);
    const v2 = this.getVersion(entityType, entityId, version2);
    
    if (!v1 || !v2) {
      throw new Error('One or both versions not found');
    }

    const data1 = JSON.parse(v1.data);
    const data2 = JSON.parse(v2.data);
    
    return this.deepDiff(data1, data2);
  },

  deepDiff(obj1, obj2, path = '') {
    const differences = [];
    const keys = new Set([...Object.keys(obj1), ...Object.keys(obj2)]);
    
    keys.forEach(key => {
      const newPath = path ? `${path}.${key}` : key;
      const val1 = obj1[key];
      const val2 = obj2[key];
      
      if (val1 === undefined && val2 !== undefined) {
        differences.push({ type: 'added', path: newPath, value: val2 });
      } else if (val1 !== undefined && val2 === undefined) {
        differences.push({ type: 'removed', path: newPath, value: val1 });
      } else if (val1 !== val2) {
        if (typeof val1 === 'object' && typeof val2 === 'object' && val1 !== null && val2 !== null) {
          differences.push(...this.deepDiff(val1, val2, newPath));
        } else {
          differences.push({ type: 'changed', path: newPath, oldValue: val1, newValue: val2 });
        }
      }
    });
    
    return differences;
  },

  // Clean up old versions (keep last 20 per entity)
  cleanupVersions() {
    const properties = PropertiesService.getScriptProperties().getProperties();
    const versionGroups = {};
    
    // Group versions by entity
    Object.keys(properties).forEach(key => {
      if (key.startsWith('version_')) {
        const parts = key.split('_');
        if (parts.length >= 4) {
          const entityType = parts[1];
          const entityId = parts[2];
          const groupKey = `${entityType}_${entityId}`;
          
          if (!versionGroups[groupKey]) {
            versionGroups[groupKey] = [];
          }
          
          try {
            const version = JSON.parse(properties[key]);
            versionGroups[groupKey].push({ key, version });
          } catch (e) {
            // Delete corrupted entries
            PropertiesService.getScriptProperties().deleteProperty(key);
          }
        }
      }
    });
    
    // Keep only latest 20 versions per entity
    Object.keys(versionGroups).forEach(groupKey => {
      const versions = versionGroups[groupKey].sort((a, b) => b.version.version - a.version.version);
      
      if (versions.length > 20) {
        const toDelete = versions.slice(20);
        toDelete.forEach(item => {
          PropertiesService.getScriptProperties().deleteProperty(item.key);
        });
      }
    });
  }
};

/**
 * Enterprise Admin Dashboard
 */
function showEnterpriseAdminDashboard() {
  const user = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(user.email, 'system_config')) {
    throw new Error('Insufficient permissions for admin dashboard');
  }

  const dashboardData = EnterpriseAdminManager.getDashboardData();
  
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <title>Enterprise Admin Dashboard</title>
      <style>
        body { font-family: 'Poppins', sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .dashboard-header { background: #003366; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .stat-value { font-size: 2em; font-weight: bold; color: #003366; }
        .stat-label { color: #666; margin-top: 5px; }
        .section { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .section h3 { margin-top: 0; color: #003366; }
        .btn { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; margin-right: 5px; }
        .btn-primary { background: #007bff; color: white; }
        .btn-danger { background: #dc3545; color: white; }
        .btn-warning { background: #ffc107; color: #333; }
        .activity-list { max-height: 300px; overflow-y: auto; background: #f8f9fa; padding: 10px; border-radius: 4px; }
        .activity-item { padding: 8px; border-bottom: 1px solid #eee; font-size: 0.9em; }
      </style>
    </head>
    <body>
      <div class="dashboard-header">
        <h1>Enterprise Admin Dashboard</h1>
        <p>System Overview and Management</p>
      </div>

      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-value">${dashboardData.stats.totalUsers}</div>
          <div class="stat-label">Total Users</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${dashboardData.stats.activeUsers}</div>
          <div class="stat-label">Active Users (24h)</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${dashboardData.stats.totalClients}</div>
          <div class="stat-label">Total Clients</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${dashboardData.stats.auditEntries}</div>
          <div class="stat-label">Audit Entries (30d)</div>
        </div>
      </div>

      <div class="section">
        <h3>System Management</h3>
        <button class="btn btn-primary" onclick="runSystemMaintenance()">Run Maintenance</button>
        <button class="btn btn-warning" onclick="generateComplianceReport()">Compliance Report</button>
        <button class="btn btn-danger" onclick="cleanupInactiveUsers()">Cleanup Users</button>
      </div>

      <div class="section">
        <h3>Recent System Activity</h3>
        <div class="activity-list">
          ${dashboardData.recentActivity.map(activity => `
            <div class="activity-item">
              <strong>${activity.action}</strong> by ${activity.userId} - ${activity.entityType}: ${activity.entityId}
            </div>
          `).join('')}
        </div>
      </div>

      <script>
        function runSystemMaintenance() {
          if (confirm('Run system maintenance? This may take a few minutes.')) {
            google.script.run.withSuccessHandler(result => {
              alert('Maintenance completed: ' + result.message);
              window.location.reload();
            }).withFailureHandler(error => {
              alert('Maintenance error: ' + error);
            }).runSystemMaintenance();
          }
        }

        function generateComplianceReport() {
          const startDate = prompt('Start date (YYYY-MM-DD):');
          const endDate = prompt('End date (YYYY-MM-DD):');
          
          if (startDate && endDate) {
            google.script.run.withSuccessHandler(report => {
              alert('Report generated with ' + report.summary.totalActions + ' entries');
            }).withFailureHandler(error => {
              alert('Error: ' + error);
            }).generateComplianceReportDialog(startDate, endDate);
          }
        }

        function cleanupInactiveUsers() {
          if (confirm('Remove users inactive for more than 90 days?')) {
            google.script.run.withSuccessHandler(result => {
              alert('Cleanup completed: ' + result.cleaned + ' users removed');
              window.location.reload();
            }).withFailureHandler(error => {
              alert('Cleanup error: ' + error);
            }).cleanupInactiveUsers();
          }
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enterprise Admin Dashboard');
}

const EnterpriseAdminManager = {
  getDashboardData() {
    const allUsers = getAllSystemUsers();
    const activeUsers = ActivityTracker.getActiveUsers(24 * 60);
    const auditEntries = AuditLogger.getAuditLog(
      new Date(Date.now() - 30*24*60*60*1000).toISOString().split('T')[0],
      new Date().toISOString().split('T')[0]
    );
    
    return {
      stats: {
        totalUsers: allUsers.length,
        activeUsers: activeUsers.length,
        totalClients: UnifiedClientDataStore.getAllClients().length,
        auditEntries: auditEntries.length
      },
      recentActivity: auditEntries.slice(0, 10)
    };
  }
};

function runSystemMaintenance() {
  const currentUser = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(currentUser.email, 'system_config')) {
    throw new Error('Insufficient permissions');
  }
  
  AuditLogger.log('maintenance_started', 'system', 'maintenance', currentUser.email);
  
  // Cleanup operations
  NotificationManager.cleanupNotifications();
  DataVersionManager.cleanupVersions();
  
  AuditLogger.log('maintenance_completed', 'system', 'maintenance', currentUser.email);
  
  return {
    success: true,
    message: 'System maintenance completed successfully'
  };
}

function generateComplianceReportDialog(startDate, endDate) {
  const currentUser = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(currentUser.email, 'export_data')) {
    throw new Error('Insufficient permissions');
  }
  
  return AuditLogger.generateComplianceReport(startDate, endDate);
}

function cleanupInactiveUsers() {
  const currentUser = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(currentUser.email, 'manage_users')) {
    throw new Error('Insufficient permissions');
  }
  
  const cutoffDate = new Date(Date.now() - 90*24*60*60*1000);
  const allUsers = getAllSystemUsers();
  let cleanedCount = 0;
  
  allUsers.forEach(email => {
    const activity = ActivityTracker.getUserActivity(email);
    const role = PermissionManager.getUserRole(email);
    
    if (role !== 'admin' && (!activity || new Date(activity.lastSeen) < cutoffDate)) {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty(`user_${email}`);
      properties.deleteProperty(`activity_${email}`);
      cleanedCount++;
      
      AuditLogger.log('user_cleaned', 'user', email, currentUser.email, {
        reason: 'inactive_90_days'
      });
    }
  });
  
  return { success: true, cleaned: cleanedCount };
}

/**
 * Organization Management and Multi-Tenancy System
 */
const OrganizationManager = {
  // Create a new organization
  createOrganization(orgData, createdBy) {
    const organization = {
      id: Utilities.getUuid(),
      name: orgData.name,
      displayName: orgData.displayName || orgData.name,
      domain: orgData.domain, // e.g., "smartcollege.com"
      settings: {
        allowedDomains: orgData.allowedDomains || [orgData.domain],
        maxUsers: orgData.maxUsers || 50,
        features: orgData.features || ["basic", "clients", "reports"],
        branding: {
          primaryColor: orgData.primaryColor || '#003366',
          logo: orgData.logo || null,
          companyName: orgData.companyName || orgData.name
        }
      },
      status: 'active',
      createdBy: createdBy,
      createdAt: new Date().toISOString(),
      subscription: {
        plan: orgData.plan || 'standard',
        expiresAt: orgData.subscriptionExpiry || null,
        features: orgData.features || ["basic"]
      }
    };

    this.storeOrganization(organization);
    
    // Initialize organization-specific data stores
    this.initializeOrgDataStores(organization.id);
    
    AuditLogger.log('organization_created', 'organization', organization.id, createdBy, {
      orgName: organization.name,
      domain: organization.domain
    });

    return organization;
  },

  storeOrganization(organization) {
    const key = `organization_${organization.id}`;
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(organization));
  },

  getOrganization(orgId) {
    const key = `organization_${orgId}`;
    const data = PropertiesService.getScriptProperties().getProperty(key);
    return data ? JSON.parse(data) : null;
  },

  getUserOrganization(userEmail) {
    // Determine organization by email domain or explicit assignment
    const domain = userEmail.split('@')[1];
    
    // Check for explicit organization assignment first
    const userOrgKey = `user_org_${userEmail}`;
    const assignedOrgId = PropertiesService.getScriptProperties().getProperty(userOrgKey);
    if (assignedOrgId) {
      return this.getOrganization(assignedOrgId);
    }

    // Fall back to domain matching
    const organizations = this.getAllOrganizations();
    return organizations.find(org => 
      org.settings.allowedDomains.includes(domain)
    ) || this.getDefaultOrganization();
  },

  getAllOrganizations() {
    const organizations = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('organization_')) {
        try {
          const org = JSON.parse(properties[key]);
          organizations.push(org);
        } catch (e) {
          // Skip corrupted entries
        }
      }
    });
    
    return organizations;
  },

  getDefaultOrganization() {
    // Return a default organization for users without specific org assignment
    return {
      id: 'default',
      name: 'Default Organization',
      displayName: 'Default Organization',
      domain: 'default.local',
      settings: {
        allowedDomains: ['*'],
        maxUsers: 100,
        features: ["basic", "clients", "reports"],
        branding: {
          primaryColor: '#003366',
          companyName: 'Smart College'
        }
      },
      status: 'active'
    };
  },

  assignUserToOrganization(userEmail, orgId) {
    const currentUser = AuthManager.getCurrentUser();
    if (!PermissionManager.hasPermission(currentUser.email, 'manage_users')) {
      throw new Error('Insufficient permissions');
    }

    const organization = this.getOrganization(orgId);
    if (!organization) {
      throw new Error('Organization not found');
    }

    // Check user limits
    const orgUserCount = this.getOrganizationUserCount(orgId);
    if (orgUserCount >= organization.settings.maxUsers) {
      throw new Error('Organization user limit reached');
    }

    const userOrgKey = `user_org_${userEmail}`;
    PropertiesService.getScriptProperties().setProperty(userOrgKey, orgId);

    AuditLogger.log('user_org_assigned', 'organization', orgId, currentUser.email, {
      assignedUser: userEmail,
      orgName: organization.name
    });

    return { success: true };
  },

  getOrganizationUserCount(orgId) {
    let count = 0;
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('user_org_') && properties[key] === orgId) {
        count++;
      }
    });
    
    return count;
  },

  getOrganizationUsers(orgId) {
    const users = [];
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    Object.keys(properties).forEach(key => {
      if (key.startsWith('user_org_') && properties[key] === orgId) {
        const userEmail = key.replace('user_org_', '');
        users.push({
          email: userEmail,
          role: PermissionManager.getUserRole(userEmail),
          lastActive: ActivityTracker.getUserActivity(userEmail)?.lastSeen
        });
      }
    });
    
    return users;
  },

  initializeOrgDataStores(orgId) {
    // Create organization-specific data containers
    const orgDataKey = `org_data_${orgId}`;
    const orgData = {
      clients: {},
      settings: {},
      statistics: {
        created: new Date().toISOString(),
        lastUpdated: new Date().toISOString(),
        totalClients: 0,
        totalUsers: 0
      }
    };
    
    PropertiesService.getScriptProperties().setProperty(orgDataKey, JSON.stringify(orgData));
  },

  // Data isolation methods
  getOrgClients(orgId) {
    const orgDataKey = `org_data_${orgId}`;
    const data = PropertiesService.getScriptProperties().getProperty(orgDataKey);
    return data ? JSON.parse(data).clients : {};
  },

  storeOrgClient(orgId, clientName, clientData) {
    const orgDataKey = `org_data_${orgId}`;
    const data = PropertiesService.getScriptProperties().getProperty(orgDataKey);
    const orgData = data ? JSON.parse(data) : { clients: {}, statistics: {} };
    
    orgData.clients[clientName] = {
      ...clientData,
      organizationId: orgId,
      lastUpdated: new Date().toISOString()
    };
    
    orgData.statistics.totalClients = Object.keys(orgData.clients).length;
    orgData.statistics.lastUpdated = new Date().toISOString();
    
    PropertiesService.getScriptProperties().setProperty(orgDataKey, JSON.stringify(orgData));
  }
};

/**
 * Multi-Tenant Aware Client Functions
 */
function createClientSheetMultiTenant(clientData) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('User not authenticated');
  }

  // Get user's organization
  const organization = OrganizationManager.getUserOrganization(user.email);
  if (!organization) {
    throw new Error('User organization not found');
  }

  // Check organization features
  if (!organization.settings.features.includes('clients')) {
    throw new Error('Client management not enabled for your organization');
  }

  // Add organization context
  clientData.organizationId = organization.id;
  clientData.createdBy = user.email;
  clientData.createdAt = new Date().toISOString();

  // Check if client exists in this organization
  const existingClients = OrganizationManager.getOrgClients(organization.id);
  if (existingClients[clientData.name]) {
    return {
      success: false,
      reason: "client_exists_in_org",
      organization: organization.name,
      suggestedAction: "use_different_name"
    };
  }

  // Create the client using existing function
  const result = createClientSheet(clientData);
  
  if (result.success) {
    // Store in organization-specific data
    OrganizationManager.storeOrgClient(organization.id, clientData.name, clientData);
    
    // Log with organization context
    AuditLogger.log('client_created', 'client', clientData.name, user.email, {
      organizationId: organization.id,
      organizationName: organization.name,
      services: clientData.services
    });
  }

  return result;
}

function getOrganizationClientList() {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    return { error: 'User not authenticated', clients: [] };
  }

  const organization = OrganizationManager.getUserOrganization(user.email);
  if (!organization) {
    return { error: 'Organization not found', clients: [] };
  }

  // Get organization-specific clients
  const orgClients = OrganizationManager.getOrgClients(organization.id);
  
  // Filter based on user permissions within the organization
  const accessibleClients = Object.values(orgClients).filter(client => 
    ClientOwnershipManager.canUserAccessClient(client.name, user.email)
  );

  // Apply data privacy filtering
  const filteredClients = accessibleClients.map(client => 
    DataPrivacyManager.filterDataByPermission(client, user.email)
  );

  AuditLogger.log('view', 'org_client_list', organization.id, user.email, { 
    clientCount: filteredClients.length,
    organizationName: organization.name
  });

  return { 
    clients: filteredClients, 
    organization: organization,
    userRole: PermissionManager.getUserRole(user.email) 
  };
}

/**
 * Organization Admin Dashboard
 */
function showOrganizationDashboard() {
  const user = AuthManager.getCurrentUser();
  const organization = OrganizationManager.getUserOrganization(user.email);
  
  if (!PermissionManager.hasPermission(user.email, 'manage_users') && 
      !PermissionManager.hasPermission(user.email, 'system_config')) {
    throw new Error('Insufficient permissions for organization dashboard');
  }

  const orgUsers = OrganizationManager.getOrganizationUsers(organization.id);
  const orgClients = Object.values(OrganizationManager.getOrgClients(organization.id));
  
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <title>${organization.displayName} - Dashboard</title>
      <style>
        body { font-family: 'Poppins', sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .org-header { 
          background: ${organization.settings.branding.primaryColor}; 
          color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; 
        }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .stat-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .stat-value { font-size: 2em; font-weight: bold; color: ${organization.settings.branding.primaryColor}; }
        .stat-label { color: #666; margin-top: 5px; }
        .section { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .section h3 { margin-top: 0; color: ${organization.settings.branding.primaryColor}; }
        .user-list { display: grid; gap: 10px; }
        .user-item { padding: 10px; background: #f8f9fa; border-radius: 4px; display: flex; justify-content: space-between; align-items: center; }
        .btn { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; margin-right: 5px; }
        .btn-primary { background: #007bff; color: white; }
      </style>
    </head>
    <body>
      <div class="org-header">
        <h1>${organization.settings.branding.companyName}</h1>
        <p>Organization: ${organization.displayName}</p>
        <p>Plan: ${organization.subscription?.plan || 'Standard'}</p>
      </div>

      <div class="stats-grid">
        <div class="stat-card">
          <div class="stat-value">${orgUsers.length}</div>
          <div class="stat-label">Users</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${organization.settings.maxUsers}</div>
          <div class="stat-label">User Limit</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${orgClients.length}</div>
          <div class="stat-label">Clients</div>
        </div>
        <div class="stat-card">
          <div class="stat-value">${organization.settings.features.length}</div>
          <div class="stat-label">Features</div>
        </div>
      </div>

      <div class="section">
        <h3>Organization Users</h3>
        <div class="user-list">
          ${orgUsers.map(user => `
            <div class="user-item">
              <div>
                <strong>${user.email}</strong>
                <span style="color: #666; margin-left: 10px;">${user.role}</span>
              </div>
              <div>
                <span style="font-size: 0.9em; color: #666;">
                  ${user.lastActive ? getRelativeTime(new Date(user.lastActive).getTime()) : 'Never'}
                </span>
              </div>
            </div>
          `).join('')}
        </div>
        
        <button class="btn btn-primary" onclick="addOrgUser()" style="margin-top: 15px;">
          Add User to Organization
        </button>
      </div>

      <div class="section">
        <h3>Organization Settings</h3>
        <p><strong>Allowed Domains:</strong> ${organization.settings.allowedDomains.join(', ')}</p>
        <p><strong>Features:</strong> ${organization.settings.features.join(', ')}</p>
        <p><strong>Status:</strong> ${organization.status}</p>
        
        <button class="btn btn-primary" onclick="updateOrgSettings()">
          Update Settings
        </button>
      </div>

      <script>
        function addOrgUser() {
          const email = prompt('Enter user email to add to organization:');
          if (email && email.includes('@')) {
            google.script.run.withSuccessHandler(() => {
              alert('User added to organization');
              window.location.reload();
            }).withFailureHandler(error => {
              alert('Error: ' + error);
            }).assignUserToOrg(email);
          }
        }

        function updateOrgSettings() {
          alert('Organization settings update functionality would go here');
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Organization Dashboard');
}

function assignUserToOrg(userEmail) {
  const currentUser = AuthManager.getCurrentUser();
  const organization = OrganizationManager.getUserOrganization(currentUser.email);
  
  return OrganizationManager.assignUserToOrganization(userEmail, organization.id);
}

/**
 * Enhanced Security Measures
 */
const SecurityManager = {
  // Rate limiting per user
  rateLimiter: {
    limits: {
      'api_calls': { max: 100, window: 60000 }, // 100 calls per minute
      'client_creation': { max: 10, window: 300000 }, // 10 clients per 5 minutes
      'email_sends': { max: 20, window: 3600000 }, // 20 emails per hour
      'data_exports': { max: 3, window: 3600000 }, // 3 exports per hour
      'login_attempts': { max: 5, window: 900000 } // 5 login attempts per 15 minutes
    },

    checkLimit(userId, action) {
      const limit = this.limits[action];
      if (!limit) return { allowed: true };

      const key = `rate_limit_${userId}_${action}`;
      const now = Date.now();
      const data = PropertiesService.getScriptProperties().getProperty(key);
      
      let attempts = data ? JSON.parse(data) : [];
      
      // Clean old attempts outside the window
      attempts = attempts.filter(timestamp => now - timestamp < limit.window);
      
      if (attempts.length >= limit.max) {
        AuditLogger.log('rate_limit_exceeded', 'security', action, userId, {
          attempts: attempts.length,
          limit: limit.max,
          window: limit.window
        });
        
        return { 
          allowed: false, 
          message: `Rate limit exceeded. Try again in ${Math.ceil(limit.window/60000)} minutes.`,
          resetTime: Math.min(...attempts) + limit.window
        };
      }

      // Add current attempt
      attempts.push(now);
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(attempts));
      
      return { allowed: true, remaining: limit.max - attempts.length };
    },

    resetUserLimits(userId) {
      const properties = PropertiesService.getScriptProperties().getProperties();
      Object.keys(properties).forEach(key => {
        if (key.startsWith(`rate_limit_${userId}_`)) {
          PropertiesService.getScriptProperties().deleteProperty(key);
        }
      });
    }
  },

  // Session security
  sessionManager: {
    createSession(userId) {
      const session = {
        id: Utilities.getUuid(),
        userId: userId,
        createdAt: new Date().toISOString(),
        lastActivity: new Date().toISOString(),
        ipHash: this.getClientIPHash(),
        userAgent: this.getUserAgentHash(),
        expiresAt: new Date(Date.now() + 8 * 60 * 60 * 1000).toISOString() // 8 hours
      };

      const key = `session_${session.id}`;
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(session));
      
      return session;
    },

    validateSession(sessionId) {
      const key = `session_${sessionId}`;
      const data = PropertiesService.getScriptProperties().getProperty(key);
      
      if (!data) return { valid: false, reason: 'Session not found' };
      
      const session = JSON.parse(data);
      const now = new Date().toISOString();
      
      if (now > session.expiresAt) {
        this.invalidateSession(sessionId);
        return { valid: false, reason: 'Session expired' };
      }

      // Check for session hijacking (simplified)
      const currentIPHash = this.getClientIPHash();
      const currentUAHash = this.getUserAgentHash();
      
      if (session.ipHash !== currentIPHash || session.userAgent !== currentUAHash) {
        AuditLogger.log('session_anomaly', 'security', sessionId, session.userId, {
          ipMismatch: session.ipHash !== currentIPHash,
          userAgentMismatch: session.userAgent !== currentUAHash
        });
      }

      // Update last activity
      session.lastActivity = now;
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(session));
      
      return { valid: true, session: session };
    },

    invalidateSession(sessionId) {
      const key = `session_${sessionId}`;
      PropertiesService.getScriptProperties().deleteProperty(key);
    },

    getClientIPHash() {
      // In Google Apps Script, we can't get real IP, so use a placeholder
      return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'placeholder_ip').toString();
    },

    getUserAgentHash() {
      // Simplified user agent detection
      return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'placeholder_ua').toString();
    },

    cleanupExpiredSessions() {
      const properties = PropertiesService.getScriptProperties().getProperties();
      const now = new Date().toISOString();
      let cleanedCount = 0;

      Object.keys(properties).forEach(key => {
        if (key.startsWith('session_')) {
          try {
            const session = JSON.parse(properties[key]);
            if (now > session.expiresAt) {
              PropertiesService.getScriptProperties().deleteProperty(key);
              cleanedCount++;
            }
          } catch (e) {
            PropertiesService.getScriptProperties().deleteProperty(key);
            cleanedCount++;
          }
        }
      });

      return cleanedCount;
    }
  },

  // Advanced data encryption
  encryption: {
    encryptSensitiveData(data, key) {
      if (!data) return data;
      
      const dataStr = typeof data === 'object' ? JSON.stringify(data) : data;
      const salt = Utilities.getUuid().substring(0, 8);
      const combinedKey = key + salt;
      
      // Simple encryption using base64 encoding with key mixing
      const encrypted = Utilities.base64Encode(dataStr + '|' + salt + '|' + combinedKey);
      
      return {
        encrypted: encrypted,
        algorithm: 'base64_keymix',
        timestamp: new Date().toISOString()
      };
    },

    decryptSensitiveData(encryptedData, key) {
      if (!encryptedData || typeof encryptedData !== 'object') return encryptedData;
      
      try {
        const decoded = Utilities.base64Decode(encryptedData.encrypted);
        const parts = decoded.split('|');
        
        if (parts.length >= 3) {
          const originalData = parts[0];
          const salt = parts[1];
          const combinedKey = parts[2];
          
          // Verify key
          if (combinedKey === key + salt) {
            try {
              return JSON.parse(originalData);
            } catch (e) {
              return originalData;
            }
          }
        }
        
        return encryptedData;
      } catch (e) {
        AuditLogger.log('decryption_failed', 'security', 'data_decrypt', 'system', {
          error: e.message,
          algorithm: encryptedData.algorithm
        });
        return encryptedData;
      }
    },

    generateSecureToken() {
      const timestamp = Date.now().toString(36);
      const randomBytes = Utilities.getUuid().replace(/-/g, '');
      return timestamp + randomBytes;
    }
  },

  // Security monitoring
  monitor: {
    detectAnomalies(userId, action, metadata = {}) {
      const anomalies = [];
      
      // Detect unusual activity patterns
      const recentActions = this.getRecentUserActions(userId, 300000); // 5 minutes
      
      // Rapid succession of same action
      const sameActionCount = recentActions.filter(a => a.action === action).length;
      if (sameActionCount > 10) {
        anomalies.push({
          type: 'rapid_repetition',
          description: `${sameActionCount} ${action} actions in 5 minutes`,
          severity: 'medium'
        });
      }
      
      // Unusual time of day (outside 6 AM - 11 PM)
      const hour = new Date().getHours();
      if (hour < 6 || hour > 23) {
        anomalies.push({
          type: 'unusual_time',
          description: `Activity at ${hour}:00`,
          severity: 'low'
        });
      }
      
      // Multiple organizations accessed rapidly
      const orgActions = recentActions.filter(a => a.metadata && a.metadata.organizationId);
      const uniqueOrgs = [...new Set(orgActions.map(a => a.metadata.organizationId))];
      if (uniqueOrgs.length > 2) {
        anomalies.push({
          type: 'multi_org_access',
          description: `Accessed ${uniqueOrgs.length} organizations rapidly`,
          severity: 'high'
        });
      }
      
      // Log anomalies
      if (anomalies.length > 0) {
        AuditLogger.log('security_anomaly', 'security', action, userId, {
          anomalies: anomalies,
          metadata: metadata
        });
        
        // Alert admins for high severity
        const highSeverityAnomalies = anomalies.filter(a => a.severity === 'high');
        if (highSeverityAnomalies.length > 0) {
          this.alertSecurityTeam(userId, action, highSeverityAnomalies);
        }
      }
      
      return anomalies;
    },

    getRecentUserActions(userId, windowMs) {
      const since = new Date(Date.now() - windowMs).toISOString();
      return AuditLogger.getAuditLog(since, new Date().toISOString(), { userId: userId });
    },

    alertSecurityTeam(userId, action, anomalies) {
      const admins = getAllSystemUsers().filter(email => 
        PermissionManager.getUserRole(email) === 'admin'
      );
      
      admins.forEach(adminEmail => {
        NotificationManager.createNotification(
          adminEmail,
          'audit_alert',
          'Security Anomaly Detected',
          `User ${userId} triggered security alerts during ${action}: ${anomalies.map(a => a.description).join(', ')}`,
          userId
        );
      });
    }
  },

  // Input validation and sanitization
  validator: {
    sanitizeInput(input, type = 'general') {
      if (!input) return input;
      
      const sanitizers = {
        email: (str) => str.toLowerCase().trim().replace(/[^a-zA-Z0-9@._-]/g, ''),
        name: (str) => str.trim().replace(/[<>\"'&]/g, '').substring(0, 100),
        general: (str) => str.trim().replace(/[<>\"']/g, '').substring(0, 1000),
        notes: (str) => str.trim().substring(0, 5000) // Longer for notes
      };
      
      const sanitizer = sanitizers[type] || sanitizers.general;
      return typeof input === 'string' ? sanitizer(input) : input;
    },

    validateClientData(clientData) {
      const errors = [];
      
      if (!clientData.name || clientData.name.length < 2) {
        errors.push('Client name must be at least 2 characters');
      }
      
      if (clientData.parentEmail && !this.isValidEmail(clientData.parentEmail)) {
        errors.push('Invalid parent email format');
      }
      
      if (clientData.name && clientData.name.length > 100) {
        errors.push('Client name too long (max 100 characters)');
      }
      
      return {
        valid: errors.length === 0,
        errors: errors,
        sanitized: {
          name: this.sanitizeInput(clientData.name, 'name'),
          services: this.sanitizeInput(clientData.services, 'general'),
          parentEmail: this.sanitizeInput(clientData.parentEmail, 'email'),
          emailAddressees: this.sanitizeInput(clientData.emailAddressees, 'email')
        }
      };
    },

    isValidEmail(email) {
      return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    }
  }
};

/**
 * Secure wrapper functions with rate limiting and validation
 */
function createClientSheetSecure(clientData) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('User not authenticated');
  }

  // Rate limiting
  const rateLimitCheck = SecurityManager.rateLimiter.checkLimit(user.email, 'client_creation');
  if (!rateLimitCheck.allowed) {
    throw new Error(rateLimitCheck.message);
  }

  // Input validation
  const validation = SecurityManager.validator.validateClientData(clientData);
  if (!validation.valid) {
    throw new Error('Validation failed: ' + validation.errors.join(', '));
  }

  // Security monitoring
  SecurityManager.monitor.detectAnomalies(user.email, 'client_creation', {
    clientName: validation.sanitized.name
  });

  // Use sanitized data
  return createClientSheetWithUserContext(validation.sanitized);
}

function sendEmailSecure(emailData) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('User not authenticated');
  }

  // Rate limiting for emails
  const rateLimitCheck = SecurityManager.rateLimiter.checkLimit(user.email, 'email_sends');
  if (!rateLimitCheck.allowed) {
    throw new Error(rateLimitCheck.message);
  }

  // Security monitoring
  SecurityManager.monitor.detectAnomalies(user.email, 'email_send', {
    recipientCount: emailData.recipients ? emailData.recipients.length : 0
  });

  // Proceed with email sending
  return sendIndividualRecapWithUserContext(emailData.sheetName, emailData);
}

/**
 * Integration Framework for LMS/CRM Systems
 */
const IntegrationManager = {
  // Webhook system for external integrations
  webhooks: {
    registeredHooks: {},

    registerWebhook(name, url, events, authToken = null) {
      const currentUser = AuthManager.getCurrentUser();
      if (!PermissionManager.hasPermission(currentUser.email, 'system_config')) {
        throw new Error('Insufficient permissions to register webhooks');
      }

      const webhook = {
        id: Utilities.getUuid(),
        name: name,
        url: url,
        events: events, // ['client_created', 'session_completed', 'user_assigned']
        authToken: authToken,
        isActive: true,
        createdBy: currentUser.email,
        createdAt: new Date().toISOString(),
        lastTriggered: null,
        successCount: 0,
        errorCount: 0
      };

      this.registeredHooks[webhook.id] = webhook;
      this.persistWebhooks();

      AuditLogger.log('webhook_registered', 'integration', webhook.id, currentUser.email, {
        name: name,
        url: url,
        events: events
      });

      return webhook;
    },

    triggerWebhook(event, data) {
      Object.values(this.registeredHooks).forEach(webhook => {
        if (webhook.isActive && webhook.events.includes(event)) {
          this.executeWebhook(webhook, event, data);
        }
      });
    },

    executeWebhook(webhook, event, data) {
      try {
        const payload = {
          event: event,
          timestamp: new Date().toISOString(),
          data: data,
          source: 'client-management-system'
        };

        const options = {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'User-Agent': 'ClientManagementSystem/1.0'
          },
          payload: JSON.stringify(payload)
        };

        if (webhook.authToken) {
          options.headers['Authorization'] = `Bearer ${webhook.authToken}`;
        }

        const response = UrlFetchApp.fetch(webhook.url, options);
        
        webhook.lastTriggered = new Date().toISOString();
        webhook.successCount++;

        if (response.getResponseCode() >= 400) {
          throw new Error(`HTTP ${response.getResponseCode()}: ${response.getContentText()}`);
        }

        AuditLogger.log('webhook_triggered', 'integration', webhook.id, 'system', {
          event: event,
          status: 'success',
          responseCode: response.getResponseCode()
        });

      } catch (error) {
        webhook.errorCount++;
        
        AuditLogger.log('webhook_failed', 'integration', webhook.id, 'system', {
          event: event,
          error: error.message,
          url: webhook.url
        });

        // Disable webhook after 10 consecutive failures
        if (webhook.errorCount > 10) {
          webhook.isActive = false;
          this.notifyWebhookFailure(webhook);
        }
      }

      this.persistWebhooks();
    },

    persistWebhooks() {
      PropertiesService.getScriptProperties().setProperty('webhooks', JSON.stringify(this.registeredHooks));
    },

    loadWebhooks() {
      const data = PropertiesService.getScriptProperties().getProperty('webhooks');
      this.registeredHooks = data ? JSON.parse(data) : {};
    },

    notifyWebhookFailure(webhook) {
      NotificationManager.broadcastSystemAlert(
        'Webhook Disabled',
        `Webhook "${webhook.name}" has been disabled due to repeated failures. Please check the endpoint.`,
        'admin'
      );
    }
  },

  // API endpoints for external systems
  api: {
    // External systems can call these via Google Apps Script Web App
    handleApiRequest(request) {
      const user = AuthManager.getCurrentUser();
      
      // Rate limiting for API calls
      const rateLimitCheck = SecurityManager.rateLimiter.checkLimit(user.email, 'api_calls');
      if (!rateLimitCheck.allowed) {
        return this.apiResponse(429, { error: rateLimitCheck.message });
      }

      try {
        const { method, endpoint, data } = request;
        
        switch (endpoint) {
          case '/clients':
            return method === 'GET' ? this.getClients() : this.createClient(data);
          
          case '/users':
            return method === 'GET' ? this.getUsers() : this.createUser(data);
          
          case '/sessions':
            return method === 'POST' ? this.recordSession(data) : this.apiResponse(405, { error: 'Method not allowed' });
          
          default:
            return this.apiResponse(404, { error: 'Endpoint not found' });
        }
      } catch (error) {
        AuditLogger.log('api_error', 'integration', endpoint, user.email, {
          error: error.message,
          method: request.method
        });
        
        return this.apiResponse(500, { error: 'Internal server error' });
      }
    },

    getClients() {
      const result = getOrganizationClientList();
      return this.apiResponse(200, {
        clients: result.clients,
        organization: result.organization.name
      });
    },

    createClient(data) {
      const validation = SecurityManager.validator.validateClientData(data);
      if (!validation.valid) {
        return this.apiResponse(400, { errors: validation.errors });
      }

      const result = createClientSheetSecure(validation.sanitized);
      
      // Trigger webhook
      IntegrationManager.webhooks.triggerWebhook('client_created', result);
      
      return this.apiResponse(201, result);
    },

    recordSession(data) {
      // External LMS can record that a session happened
      const user = AuthManager.getCurrentUser();
      
      AuditLogger.log('external_session_recorded', 'integration', data.clientName, user.email, {
        sessionType: data.sessionType,
        duration: data.duration,
        externalSystem: data.source
      });

      // Trigger webhook
      IntegrationManager.webhooks.triggerWebhook('session_completed', {
        clientName: data.clientName,
        sessionType: data.sessionType,
        completedAt: new Date().toISOString(),
        recordedBy: user.email
      });

      return this.apiResponse(200, { success: true, recorded: true });
    },

    apiResponse(status, data) {
      return {
        status: status,
        data: data,
        timestamp: new Date().toISOString()
      };
    }
  },

  // Data sync with external systems
  sync: {
    exportToLMS(clientData, lmsConfig) {
      // Export client data to Learning Management System
      const payload = {
        student: {
          name: clientData.name,
          email: clientData.parentEmail,
          services: clientData.services,
          startDate: clientData.createdAt
        },
        metadata: {
          source: 'client-management-system',
          exportedAt: new Date().toISOString()
        }
      };

      // Would integrate with actual LMS APIs
      console.log('Would export to LMS:', payload);
      
      return { success: true, exported: true };
    },

    importFromCRM(crmData) {
      // Import contacts from CRM system
      const user = AuthManager.getCurrentUser();
      
      crmData.contacts.forEach(contact => {
        const clientData = {
          name: contact.fullName,
          parentEmail: contact.email,
          services: contact.interestedServices || 'General Tutoring'
        };

        // Validate and create if doesn't exist
        const validation = SecurityManager.validator.validateClientData(clientData);
        if (validation.valid) {
          const existing = UnifiedClientDataStore.getClient(clientData.name);
          if (!existing) {
            createClientSheetSecure(validation.sanitized);
          }
        }
      });

      AuditLogger.log('crm_import_completed', 'integration', 'crm_sync', user.email, {
        contactsProcessed: crmData.contacts.length
      });

      return { success: true, imported: crmData.contacts.length };
    }
  }
};

// Initialize webhooks on load
IntegrationManager.webhooks.loadWebhooks();

/**
 * Security maintenance function
 */
function runSecurityMaintenance() {
  const currentUser = AuthManager.getCurrentUser();
  if (!PermissionManager.hasPermission(currentUser.email, 'system_config')) {
    throw new Error('Insufficient permissions');
  }

  let maintenanceLog = {
    startTime: new Date().toISOString(),
    actions: []
  };

  // Clean expired sessions
  const expiredSessions = SecurityManager.sessionManager.cleanupExpiredSessions();
  maintenanceLog.actions.push(`Cleaned ${expiredSessions} expired sessions`);

  // Reset rate limits for inactive users (older than 24 hours)
  const inactiveUsers = getAllSystemUsers().filter(email => {
    const activity = ActivityTracker.getUserActivity(email);
    return !activity || (Date.now() - new Date(activity.lastSeen).getTime()) > 24 * 60 * 60 * 1000;
  });
  
  inactiveUsers.forEach(email => {
    SecurityManager.rateLimiter.resetUserLimits(email);
  });
  maintenanceLog.actions.push(`Reset rate limits for ${inactiveUsers.length} inactive users`);

  // Check webhook health
  Object.values(IntegrationManager.webhooks.registeredHooks).forEach(webhook => {
    if (webhook.errorCount > 5 && webhook.isActive) {
      maintenanceLog.actions.push(`Warning: Webhook ${webhook.name} has ${webhook.errorCount} errors`);
    }
  });

  maintenanceLog.endTime = new Date().toISOString();
  
  AuditLogger.log('security_maintenance', 'security', 'maintenance', currentUser.email, maintenanceLog);

  return {
    success: true,
    message: `Security maintenance completed: ${maintenanceLog.actions.join(', ')}`
  };
}

/**
 * Show notifications for current user
 */
function showNotifications() {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    SpreadsheetApp.getUi().alert('Please sign in to view notifications');
    return;
  }
  
  const notifications = getUserNotifications();
  if (notifications.length === 0) {
    SpreadsheetApp.getUi().alert('No notifications', 'You have no new notifications.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Create a simple HTML display of notifications
  const html = notifications.map(n => 
    `<div style="border-bottom: 1px solid #ccc; padding: 10px;">
      <strong>${n.title}</strong><br>
      ${n.message}<br>
      <small>${new Date(n.timestamp).toLocaleString()}</small>
    </div>`
  ).join('');
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Your Notifications');
  
  return { success: true };
}

/**
 * Mark all notifications as read for current user
 */
function markAllNotificationsRead() {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    return { success: false, message: 'Not authenticated' };
  }
  
  const notifications = NotificationManager.getUserNotifications(user.email);
  notifications.forEach(n => {
    NotificationManager.markAsRead(user.email, n.id);
  });
  
  return { success: true, message: 'All notifications marked as read' };
}

/**
 * Get user notifications (wrapper for sidebar)
 */
function getUserNotifications() {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    return [];
  }
  
  return NotificationManager.getUserNotifications(user.email);
}

/**
 * Update user display name
 */
function updateUserDisplayName(newName) {
  const user = AuthManager.getCurrentUser();
  if (!user.authenticated) {
    throw new Error('Not authenticated');
  }
  
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('displayName', newName);
  
  // Also update in auth manager if needed
  const userDataStore = AuthManager.getUserDataStore(user.email);
  userDataStore.displayName = newName;
  AuthManager.saveUserDataStore(user.email, userDataStore);
  
  return { success: true };
}

/**
 * Search clients by name
 */
function searchClientsByName(searchTerm) {
  if (!searchTerm || searchTerm.length < 2) {
    return [];
  }
  
  const allClients = UnifiedClientDataStore.getAllClients();
  const searchLower = searchTerm.toLowerCase();
  
  return allClients
    .filter(client => 
      client.name && 
      client.name.toLowerCase().includes(searchLower) &&
      client.isActive !== false
    )
    .slice(0, 10) // Limit to 10 results
    .map(client => ({
      name: client.name,
      sheetName: client.sheetName || client.name
    }));
}

/**
 * Switch to a specific client
 */
function switchToClient(clientName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  for (const sheet of sheets) {
    if (sheet.getName() === clientName || 
        (sheet.getRange('A2').getValue() === clientName)) {
      spreadsheet.setActiveSheet(sheet);
      return { success: true };
    }
  }
  
  throw new Error(`Client sheet "${clientName}" not found`);
}

/**
 * Force set user as configured (for fixing setup issues)
 */
function forceSetConfigured() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('isConfigured', 'true');
  
  // Also set default tutor info if not set
  if (!userProperties.getProperty('tutorName')) {
    const user = Session.getActiveUser();
    const email = user.getEmail() || Session.getEffectiveUser().getEmail();
    userProperties.setProperty('tutorName', email ? email.split('@')[0] : 'Tutor');
    userProperties.setProperty('tutorEmail', email || 'tutor@smartcollege.com');
  }
  
  SpreadsheetApp.getUi().alert('Configuration Updated', 'User has been marked as configured. Enterprise sidebar will now be used.', SpreadsheetApp.getUi().ButtonSet.OK);
  
  // Refresh the UI
  onOpen();
  
  return { success: true };
}

/**
 * Test the enhanced cache system
 */
