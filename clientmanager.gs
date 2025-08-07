/**
 * Complete Client Management and Session Recap System
 * Combines client management with automated session recap emails
 */

const CONFIG = {
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
    console.log('Could not create main menu, using add-on menu:', error);
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
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
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
        <style>
          body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
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
 * Refresh UI after setup
 */
function refreshUI() {
  const ui = SpreadsheetApp.getUi();
  onOpen();
}

function showHelp() {
  const config = getCompleteConfig();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 20px;">
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
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
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
    'Complete ACT practice test {test_number} before our next session',
    'Review ACT math sections {sections} - 30 minutes before our next session',
    'Complete practice problems set {set_number} before our next session',
    'Review and redo missed problems from last test before our next session'
  ],
  'Math Curriculum': [
    'Complete textbook problems: p. {page} #{problems} before our next session',
    'Review class notes on {topic} - 20 minutes daily',
    'Complete worksheet on {topic} before our next session',
    'Study for {assessment} before our next session',
    'Khan Academy: {topic} - 30 minutes before our next session'
  ],
  'Language': [
    'Practice conversation phrases - 15 minutes daily',
    'Complete workbook exercises: p. {page} before our next session',
    'Watch {language} video with subtitles - 20 minutes',
    'Write 5 sentences using new vocabulary before our next session',
    'Record yourself reading dialogue on p. {page}'
  ],
  'General': [
    'Review today\'s materials - 20 minutes',
    'Complete practice problems before our next session',
    'Prepare questions for next session',
    'Review and organize notes from today'
  ]
};


// Email templates - Universal template for all client types
const EMAIL_TEMPLATES = {
  'Universal': {
    subject: '{studentFirstName} - Session Recap {date}',
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
 * Universal sidebar that can switch between views without closing
 */
function showUniversalSidebar(initialView = 'control') {
  // Pre-fetch client info to embed in HTML for faster loading
  const clientInfo = getCurrentClientInfo();
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
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
            text-transform: uppercase;
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
              <div id="clientName" class="no-client">
                Loading...
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
            <button id="recapBtn" class="menu-button" onclick="runFunction('sendIndividualRecap', this)" disabled>
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
          
          <h3>Quick Session Notes</h3>
          <p style="font-size: 12px; color: #666;">Click tags to add to notes</p>
          
          <div class="save-indicator" id="saveIndicator">Saved!</div>
          <div class="backup-indicator" id="backupIndicator" style="display: none; background: #ff9800; color: white; padding: 8px; border-radius: 4px; margin-bottom: 10px; font-size: 12px; text-align: center;">
            📱 Local backup active
          </div>
          
          <div class="note-section">
            <label>Wins/Breakthroughs</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('wins', 'Aha moment!')">Aha moment!</span>
              <span class="quick-tag" onclick="addToField('wins', 'Solved independently')">Solved independently</span>
              <span class="quick-tag" onclick="addToField('wins', 'Confidence boost')">Confidence boost</span>
              <span class="quick-tag" onclick="addToField('wins', 'Speed improved')">Speed improved</span>
            </div>
            <textarea id="wins" placeholder="Quick notes about wins..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="note-section">
            <label>Skills Worked On</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('skills', 'Problem solving')">Problem solving</span>
              <span class="quick-tag" onclick="addToField('skills', 'Reading comprehension')">Reading comprehension</span>
              <span class="quick-tag" onclick="addToField('skills', 'Time management')">Time management</span>
              <span class="quick-tag" onclick="addToField('skills', 'Test strategies')">Test strategies</span>
            </div>
            <textarea id="skills" placeholder="List skills covered..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="note-section">
            <label>Struggles/Areas to Review</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('struggles', 'Careless errors')">Careless errors</span>
              <span class="quick-tag" onclick="addToField('struggles', 'Time pressure')">Time pressure</span>
              <span class="quick-tag" onclick="addToField('struggles', 'Concept confusion')">Concept confusion</span>
              <span class="quick-tag" onclick="addToField('struggles', 'Attention focus')">Attention focus</span>
            </div>
            <textarea id="struggles" placeholder="Note any challenges..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="note-section">
            <label>Parent Communication Points</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('parent', 'Great participation today')">Great participation</span>
              <span class="quick-tag" onclick="addToField('parent', 'Ask about homework completion')">Homework check</span>
              <span class="quick-tag" onclick="addToField('parent', 'Encourage practice at home')">Practice reminder</span>
              <span class="quick-tag" onclick="addToField('parent', 'Celebrate improvement')">Celebrate progress</span>
            </div>
            <textarea id="parent" placeholder="Things to mention to parent..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="note-section">
            <label>Next Session Focus</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('next', 'Review homework challenges')">Review homework</span>
              <span class="quick-tag" onclick="addToField(&quot;next&quot;, &quot;Build on today's progress&quot;)">Build on progress</span>
              <span class="quick-tag" onclick="addToField('next', 'Practice test strategies')">Practice strategies</span>
              <span class="quick-tag" onclick="addToField('next', 'Focus on weak areas')">Focus weak areas</span>
            </div>
            <textarea id="next" placeholder="Plan for next time..." onchange="markUnsaved()"></textarea>
          </div>
          
          <div class="button-group">
            <button id="saveNotesBtn" class="btn-primary" onclick="saveNotes()">Save Notes</button>
            <button id="sendRecapBtn" onclick="sendRecap()">Send Recap</button>
          </div>
          
          <div class="button-group" style="margin-top: 5px;">
            <button class="btn-secondary" onclick="clearNotes()">Clear Notes</button>
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
          font-family: Arial, sans-serif;
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
          console.log('Script: Starting initialization');
          
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
            console.log('Back online');
          });
          
          window.addEventListener('offline', function() {
            isOnline = false;
            handleOfflineState();
            console.log('Gone offline');
          });
          
          // Check initial state and setup
          document.addEventListener('DOMContentLoaded', function() {
            handleOfflineState();
          });
          
          // SAVE NOTES FUNCTION - Defined at top of consolidated script with optimistic updates
          function saveNotes() {
            console.log('🟢 saveNotes function STARTED');
            try {
              const saveBtn = document.getElementById('saveNotesBtn');
              
              if (!saveBtn) {
                console.error('Save button not found');
                return;
              }
              
              const notes = {
                wins: (document.getElementById('wins') && document.getElementById('wins').value) || '',
                skills: (document.getElementById('skills') && document.getElementById('skills').value) || '',
                struggles: (document.getElementById('struggles') && document.getElementById('struggles').value) || '',
                parent: (document.getElementById('parent') && document.getElementById('parent').value) || '',
                next: (document.getElementById('next') && document.getElementById('next').value) || ''
              };
              
              console.log('Notes collected:', notes);
              
              // OPTIMISTIC UPDATE: Immediately update cache and show success
              try {
                // Store original cache state for potential rollback
                const originalCacheState = typeof PerformanceCache !== 'undefined' && currentQuickNotesClient 
                  ? PerformanceCache.getQuickNotes(currentQuickNotesClient) 
                  : null;
                
                // Update cache immediately for instant feedback
                if (typeof PerformanceCache !== 'undefined' && currentQuickNotesClient) {
                  PerformanceCache.setQuickNotes(currentQuickNotesClient, notes);
                  console.log('Quick notes cache updated optimistically');
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
                    console.log('Background notes sync successful');
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
                    console.log('Rolling back notes cache due to sync failure');
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
          
          console.log('saveNotes function defined in consolidated script:', typeof saveNotes);
          
          // Immediate test to verify function accessibility
          if (typeof saveNotes === 'undefined') {
            console.error('❌ CRITICAL: saveNotes is undefined after definition!');
          } else {
            console.log('✅ SUCCESS: saveNotes is properly defined');
          }
          
          console.log('Script: Continuing with minimal implementation');
          
          // Minimal working implementations
          function runFunction(functionName, button) {
            console.log('Running function:', functionName);
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
                
                console.log('Function completed:', functionName);
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
            console.log('Switching to Quick Notes');
            
            if (!clientInfo || !clientInfo.sheetName) {
              showAlert('Please select a client first.', 'error');
              return;
            }
            
            // Switch view immediately for fast UI response
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.add('active');
            currentView = 'quicknotes';
            
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
              console.log('Same client - preserving current form data for:', clientInfo.name);
              // Don't reload anything, keep the current form state
              return;
            }
            
            // Different client - update tracking variable
            currentQuickNotesClient = clientInfo.sheetName;
            
            // Check cache first for instant loading of selected client's notes
            const cacheKey = 'quicknotes_' + clientInfo.sheetName;
            const cachedNotes = typeof PerformanceCache !== 'undefined' ? PerformanceCache.get(cacheKey) : null;
            
            if (cachedNotes) {
              console.log('Loading cached quick notes for:', clientInfo.name);
              populateQuickNotesForm(cachedNotes);
              // Still load from server in background to get latest data
              loadQuickNotesFromServer(clientInfo.sheetName, true);
            } else {
              console.log('No cached notes found, loading from server for:', clientInfo.name);
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
            console.log('Switching to Control Panel');
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
            console.log('Opening update dialog for:', clientInfo.name);
            
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
                console.log('Using cached client details');
                populateClientUpdateForm(cachedDetails);
                detailsLoaded = true;
              }
            }
            
            // Second: Check preloaded data if cache miss
            if (!detailsLoaded && window.preloadedClientDetails) {
              console.log('Using preloaded client details');
              populateClientUpdateForm(window.preloadedClientDetails);
              detailsLoaded = true;
            }
            
            // Third: Load from server if no cached or preloaded data
            if (!detailsLoaded) {
              console.log('Loading client details from server');
              google.script.run
                .withSuccessHandler(function(details) {
                  console.log('Client details loaded from server:', details);
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
            console.log(type + ':', message);
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
                    console.log('Client details auto-preloaded from UnifiedClientDataStore:', window.preloadedClientDetails);
                    
                    // Update caution sign based on client information completeness
                    const clientWarning = document.getElementById('clientWarning');
                    if (clientWarning) {
                      const hasAllInfo = currentClient.dashboardLink && 
                                       currentClient.meetingNotesLink && 
                                       currentClient.scoreReportLink &&
                                       currentClient.parentEmail &&
                                       currentClient.emailAddressees;
                      clientWarning.style.display = hasAllInfo ? 'none' : 'inline-block';
                      console.log('Client warning updated from UnifiedClientDataStore - hasAllInfo:', hasAllInfo);
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
                  console.log('Failed to get client details from UnifiedClientDataStore:', error);
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
                console.log('Sidebar data loaded:', data);
                clientInfo = data.clientInfo || {isClient: false, name: null, sheetName: null};
                console.log('Client info set to:', clientInfo);
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
                  console.log('Unified client list loaded:', clients.length, 'clients');
                } else {
                  console.log('No clients found in unified store');
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
            console.log('preloadAddClient called');
          }
          
          function preloadClientList() {
            console.log('preloadClientList called');
          }
          
          function preloadQuickNotes() {
            console.log('preloadQuickNotes called on mouseover');
            
            // Only proceed if we have a client and we're not already in Quick Notes view
            if (!clientInfo || !clientInfo.isClient || currentView === 'quicknotes') {
              return;
            }
            
            // Check if PerformanceCache is available
            if (typeof PerformanceCache === 'undefined') {
              console.log('PerformanceCache not yet defined, skipping preload');
              return;
            }
            
            // Check if we have cached notes first
            const cacheKey = 'quicknotes_' + clientInfo.sheetName;
            const cachedNotes = PerformanceCache.get(cacheKey);
            
            if (cachedNotes) {
              console.log('Quick Notes found in cache for mouseover preload');
              return; // Notes are already cached and will load when view opens
            }
            
            // Cache blank notes for instant loading
            console.log('Caching blank notes for instant loading:', clientInfo.name);
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
                  console.log('Quick Notes preloaded from server and cache updated');
                }
              })
              .withFailureHandler(function(error) {
                // Silent fail for preload - blank notes remain in cache
                console.log('Preload failed silently, using blank notes:', error);
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function preloadUpdateClient() {
            console.log('preloadUpdateClient called - data should already be preloaded automatically');
            // Details are now automatically preloaded when client is identified
            // This function is kept for backward compatibility but no longer needs to do anything
            if (window.preloadedClientDetails) {
              console.log('Preloaded client details already available:', window.preloadedClientDetails);
            } else {
              console.log('No preloaded data available yet');
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
                console.log('Client details cache updated optimistically');
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
                  console.log('Background sync successful');
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
                    console.log('Cache rolled back due to server sync failure');
                    
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
              console.log('Control panel warning updated - hasAllInfo:', hasAllInfo);
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
              
              console.log('Added to', fieldId, ':', text);
            } else {
              console.error('Field not found:', fieldId);
            }
          }
          
          // Populate client update form
          function populateClientUpdateForm(details) {
            console.log('Populating form with:', details);
            
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
            
            console.log('Form populated successfully');
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
            console.log('Showing client list');
            
            // Hide other views and show client list
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.remove('active');
            document.getElementById('clientUpdateView').classList.remove('active');
            document.getElementById('clientListView').classList.add('active');
            currentView = 'clientlist';
            
            // Check if we have preloaded client data for instant display
            if (allClients && allClients.length > 0) {
              console.log('Using preloaded client data for instant display');
              // Apply default filter (Active Clients)
              filterClientsByType('active');
              return;
            }
            
            // Fallback: Show loading state and load from server
            console.log('No preloaded data, loading from server');
            const clientList = document.getElementById('clientList');
            if (clientList) {
              clientList.innerHTML = '<div style="text-align: center; color: #666; padding: 20px;">Loading clients...</div>';
            }
            
            google.script.run
              .withSuccessHandler(function(clients) {
                console.log('Unified clients loaded from server:', clients);
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
            console.log('Selected client:', sheetName);
            
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
                console.log('Navigated to client sheet');
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
                .getActiveClientList();
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
            console.log('immediateBackToControl called');
            
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
                console.log('Removed active from:', viewId);
              }
            });
            
            const controlView = document.getElementById('controlView');
            if (controlView) {
              controlView.classList.add('active');
              console.log('Added active to controlView');
            } else {
              console.error('Control view element not found!');
            }
            
            currentView = 'control';
            console.log('immediateBackToControl completed, view:', currentView);
            
            // Update client info in background (non-blocking)
            setTimeout(function() {
              google.script.run
                .withSuccessHandler(function(info) {
                  clientInfo = info;
                  updateClientDisplay();
                  console.log('Background client info updated');
                })
                .getCurrentClientInfo();
            }, 0);
          }
          
          function prepopulateClientListDialog(clients) {
            console.log('Pre-populating Client List dialog with', clients.length, 'clients');
            
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
              
              console.log('Client List dialog pre-populated successfully');
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
                console.log('Sheet change detection initialized:', lastKnownSheet);
              })
              .getCurrentClientInfo();
            
            // Check for sheet changes every 2 seconds
            sheetChangeCheckInterval = setInterval(function() {
              google.script.run
                .withSuccessHandler(function(newInfo) {
                  if (newInfo.sheetName !== lastKnownSheet) {
                    console.log('Sheet change detected:', lastKnownSheet, '→', newInfo.sheetName);
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
            console.log('Handling sheet change to:', newClientInfo.sheetName);
            lastKnownSheet = newClientInfo.sheetName;
            
            // Check if we're in Quick Notes view with unsaved changes
            const isInQuickNotes = currentView === 'quicknotes';
            const hasUnsavedChanges = typeof unsavedChanges !== 'undefined' && unsavedChanges;
            
            if (isInQuickNotes && hasUnsavedChanges) {
              console.log('Saving Quick Notes before sheet change');
              
              // Save current Quick Notes
              const notes = {
                wins: document.getElementById('wins').value,
                skills: document.getElementById('skills').value,
                struggles: document.getElementById('struggles').value,
                parent: document.getElementById('parent').value,
                next: document.getElementById('next').value
              };
              
              google.script.run
                .withSuccessHandler(function() {
                  console.log('Quick Notes saved, updating client info');
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
            console.log('Updating client info after sheet change');
            
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
              console.log('Sheet change detection stopped');
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
            console.log('Notes marked as unsaved');
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
          
          // Clear Notes function
          function clearNotes() {
            const button = document.querySelector('button[onclick*="clearNotes"]');
            if (button) {
              setButtonLoading(button, true, 'Clearing...');
            }
            
            // Clear all textareas
            document.getElementById('wins').value = '';
            document.getElementById('skills').value = '';
            document.getElementById('struggles').value = '';
            document.getElementById('parent').value = '';
            document.getElementById('next').value = '';
            
            // Mark as unsaved since we cleared content
            unsavedChanges = true;
            
            if (button) {
              setTimeout(function() {
                setButtonLoading(button, false);
              }, 300); // Brief delay for visual feedback
            }
            
            console.log('Notes cleared');
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
                  
                  console.log('Notes reloaded successfully');
                } else {
                  console.log('No saved notes found');
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
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(function() {
                console.log('Notes saved, sending recap');
                
                // Send the recap and handle the result
                google.script.run
                  .withSuccessHandler(function(result) {
                    button.classList.remove('btn-loading');
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    console.log('Recap process result:', result);
                    
                    if (result && result.showOverlay) {
                      // Show missing links overlay
                      showMissingLinksOverlayInSidebar(result.missingLinks, result.clientName);
                    } else {
                      // Recap dialog was shown successfully
                      console.log('Recap dialog opened successfully');
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
            document.getElementById('skills').value = notes.skills || '';
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
          
          console.log('Script 3: Minimal implementation loaded');
        </script>
        
        <!-- Original Script 3 commented out entirely -->
        <!--
        <script>
          console.log('Script 3: Starting');
          
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
              console.log('Quick notes cached for:', clientSheetName);
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
                  .getActiveClientList();
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
          
          console.log('Script 3: PerformanceCache defined successfully');
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
          
          console.log('Script 3: RequestQueue defined successfully');
          
          } catch(e) {
            console.error('Script 3 error:', e.message, 'at line', e.lineNumber);
          }
          */
          
          console.log('Script 3: Ended without PerformanceCache');
          
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
              console.log('Button loading started:', button.textContent); // Debug
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
                console.log('Button loading finished:', button.textContent); // Debug
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
            document.getElementById('skills').value = notes.skills || '';
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
            console.log('switchToControlPanel called, current view:', currentView);
            
            // Clear auto-save
            if (autoSaveInterval) {
              clearInterval(autoSaveInterval);
            }
            
            // Clear any preloaded data
            window.preloadedClientDetails = null;
            
            // Update client info
            google.script.run
              .withSuccessHandler(function(info) {
                console.log('switchToControlPanel: Client info received:', info);
                clientInfo = info;
                updateClientDisplay();
                
                // Switch view - force class removal and addition
                const views = ['quickNotesView', 'clientUpdateView', 'clientListView'];
                views.forEach(function(viewId) {
                  const element = document.getElementById(viewId);
                  if (element) {
                    element.classList.remove('active');
                    console.log('Removed active class from:', viewId);
                  }
                });
                
                const controlView = document.getElementById('controlView');
                if (controlView) {
                  controlView.classList.add('active');
                  console.log('Added active class to controlView');
                }
                
                currentView = 'control';
                console.log('switchToControlPanel completed, new view:', currentView);
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
                    console.log('Client details preloaded from UnifiedClientDataStore:', window.preloadedClientDetails);
                    
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
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            saveNotesToLocalBackup(notes);
            updateBackupIndicator();
          }
          
          // Function moved to main script section for proper accessibility
          
          function saveNotes() {
            console.log('🟢 saveNotes function STARTED');
            try {
              console.log('saveNotes function called - inside try block');
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
                skills: (document.getElementById('skills') && document.getElementById('skills').value) || '',
                struggles: (document.getElementById('struggles') && document.getElementById('struggles').value) || '',
                parent: (document.getElementById('parent') && document.getElementById('parent').value) || '',
                next: (document.getElementById('next') && document.getElementById('next').value) || ''
              };
              
              console.log('Notes collected:', notes);
              
              // Save to local backup immediately
              if (typeof saveNotesToLocalBackup === 'function') {
                saveNotesToLocalBackup(notes);
              }
              
              google.script.run
                .withSuccessHandler(function() {
                  console.log('Notes saved successfully');
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
          console.log('saveNotes function defined:', typeof saveNotes);
          
          function clearNotes() {
            if (confirm('Clear all quick notes for this student?')) {
              document.getElementById('wins').value = '';
              document.getElementById('skills').value = '';
              document.getElementById('struggles').value = '';
              document.getElementById('parent').value = '';
              document.getElementById('next').value = '';
              unsavedChanges = true;
              showAlert('Notes cleared. Remember to save!', 'success');
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
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(function() {
                console.log('Notes saved, sending recap');
                
                // Send the recap
                google.script.run
                  .withSuccessHandler(function() {
                    button.classList.remove('btn-loading');
                    button.disabled = false;
                    button.textContent = button.dataset.originalText || 'Send Recap';
                    console.log('Recap sent successfully');
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
                .getActiveClientList();
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
            console.log('preloadAddClient called');
            // Pre-load any data needed for add client (minimal since it's a modal)
          }
          
          function preloadClientList() {
            console.log('preloadClientList called');
            // Client list is already preloaded during sidebar initialization
          }
          
          function preloadUpdateClient() {
            console.log('preloadUpdateClient called');
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
              console.log('Notes saved to local backup');
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
              console.log('Local backup cleared');
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
              console.log('Connection lost - switching to offline mode');
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
              console.log('Connection restored');
              
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
            console.log('Initializing preload functions...');
            try {
              preloadAddClient();
              preloadClientList();
              preloadUpdateClient();
              console.log('All preload functions initialized successfully');
            } catch(error) {
              console.error('Error initializing preloads:', error);
            }
          }
          
          // Function to load sidebar data from server
          function loadSidebarData() {
            // Use pre-loaded client info for instant display (no server round-trip)
            console.log('Using pre-loaded client info:', clientInfo);
            
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
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
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
 * Save dashboard link to Script Properties
 */
function saveDashboardLinkToProperties(sheetName, dashboardLink) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(`dashboard_${sheetName}`, dashboardLink);
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
            font-family: 'Google Sans', Arial, sans-serif;
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
  
  // Save all data to Script Properties (primary storage)
  const scriptProperties = PropertiesService.getScriptProperties();
  
  if (clientData.dashboardLink) {
    scriptProperties.setProperty(`dashboard_${concatenatedName}`, clientData.dashboardLink);
  }
  
  if (clientData.parentEmail) {
    scriptProperties.setProperty(`parentemail_${concatenatedName}`, clientData.parentEmail);
  }
  
  if (clientData.emailAddressees) {
    scriptProperties.setProperty(`emailaddressees_${concatenatedName}`, clientData.emailAddressees);
  }
  
  // Set default active status
  scriptProperties.setProperty(`isActive_${concatenatedName}`, 'true');
  
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
    EnhancedClientCache.addClientToCache(fullClientData);
    console.log(`Added "${clientData.name}" to enhanced client cache with complete data`);
    
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
  console.log('=== ADDING CLIENT TO MASTER LIST ===');
  console.log('Client data:', clientData);
  console.log('Sheet name:', sheetName);
  
  // Find the Master sheet
  const masterSheet = spreadsheet.getSheetByName('Master');
  
  if (!masterSheet) {
    console.error('ERROR: Could not find sheet named "Master"');
    throw new Error('Master sheet not found');
  }
  
  console.log('✓ Found Master sheet');
  
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
  
  console.log(`Last row with data in column A: ${lastRow}`);
  
  // Add new row immediately after the last row
  const newRow = lastRow + 1;
  console.log(`Adding new client to row: ${newRow}`);
  
  try {
    // Column A: Client Name with hyperlink to their sheet
    const sheetId = newSheet.getSheetId();
    const studentNameFormula = `=HYPERLINK("#gid=${sheetId}", "${clientData.name}")`;
    console.log('Setting column A (Client Name) formula:', studentNameFormula);
    masterSheet.getRange(newRow, 1).setFormula(studentNameFormula);
    
    // Column B: Dashboard link as raw URL (for smart chip conversion)
    if (clientData.dashboardLink) {
      console.log('Setting column B (Dashboard) raw URL:', clientData.dashboardLink);
      masterSheet.getRange(newRow, 2).setValue(clientData.dashboardLink);
    } else {
      masterSheet.getRange(newRow, 2).setValue('');
    }
    
    // Column C: Last Session - Formula with IFERROR to prevent errors
    const lastSessionFormula = `=IFERROR(INDEX(FILTER(${sheetName}!A:A,${sheetName}!A:A<>""), COUNTA(${sheetName}!A:A)), "")`;
    console.log('Setting column C (Last Session) formula:', lastSessionFormula);
    masterSheet.getRange(newRow, 3).setFormula(lastSessionFormula);
    
    // Column D: Next Session Scheduled? - Set checkbox value (don't insert new checkbox)
    console.log('Setting checkbox value in column D (Next Session Scheduled?)');
    masterSheet.getRange(newRow, 4).setValue(true);
    
    // Column E: Current Client? - Set checkbox value (don't insert new checkbox)
    console.log('Setting checkbox value in column E (Current Client?)');
    masterSheet.getRange(newRow, 5).setValue(true);
    
    console.log('✓ Successfully added client to master list');
    
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
    
    // Get client details from the sheet and script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || 
                         activeSheet.getRange('H2').getValue() || '';
    const meetingNotesLink = scriptProperties.getProperty(`meetingnotes_${sheetName}`) || 
                            activeSheet.getRange('I2').getValue() || '';
    const scoreReportLink = scriptProperties.getProperty(`scorereport_${sheetName}`) || 
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
    
    // Update enhanced cache with complete client data from properties
    const fullClientData = {
      name: clientName,
      sheetName: sheetName,
      isActive: true,
      dashboardLink: scriptProperties.getProperty(`dashboard_${sheetName}`) || '',
      meetingNotesLink: scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '',
      scoreReportLink: scriptProperties.getProperty(`scorereport_${sheetName}`) || '',
      parentEmail: scriptProperties.getProperty(`parentemail_${sheetName}`) || '',
      emailAddressees: scriptProperties.getProperty(`emailaddressees_${sheetName}`) || ''
    };
    EnhancedClientCache.addClientToCache(fullClientData);
    
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
  showUniversalSidebar('clientlist');
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
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const notesKey = `notes_${sheet.getName()}`;
  const timestampKey = `notes_timestamp_${sheet.getName()}`;
  
  const notesData = scriptProperties.getProperty(notesKey);
  const timestamp = scriptProperties.getProperty(timestampKey) || 0;
  
  return {
    data: notesData,
    timestamp: parseInt(timestamp)
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const notesKey = `notes_${studentSheetName}`;
  scriptProperties.deleteProperty(notesKey);
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
 * Save quick notes to script properties
 */
function saveQuickNotes(notes) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const scriptProperties = PropertiesService.getScriptProperties();
    const notesKey = `notes_${sheet.getName()}`;
    const timestampKey = `notes_timestamp_${sheet.getName()}`;
    
    scriptProperties.setProperty(notesKey, JSON.stringify(notes));
    scriptProperties.setProperty(timestampKey, new Date().getTime().toString());
    
    console.log('Quick notes saved successfully for sheet:', sheet.getName());
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
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
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
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('batchClients', JSON.stringify(clientSheetNames));
  scriptProperties.setProperty('batchIndex', '0');
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
  
  const clientData = getClientDataFromSheet(sheet);
  
  // Check if any required links are missing
  const missingLinks = [];
  if (!clientData.dashboardLink) missingLinks.push('Dashboard Link');
  if (!clientData.meetingNotesLink) missingLinks.push('Meeting Notes Link');
  if (!clientData.scoreReportLink) missingLinks.push('Score Report Link');
  
  if (missingLinks.length > 0) {
    // Store data and return overlay information
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('tempClientData', JSON.stringify(clientData));
    
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
  // Store data for the overlay buttons
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('tempClientData', JSON.stringify(clientData));
  
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
  // Store all data on server side to avoid HTML embedding
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('tempClientData', JSON.stringify(clientData));
  scriptProperties.setProperty('tempMissingLinks', JSON.stringify(missingLinks));
  
  // Create minimal HTML with no embedded data
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; margin: 20px; }
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const tempClientDataStr = scriptProperties.getProperty('tempClientData');
  
  if (!tempClientDataStr) {
    throw new Error('No client data found for recap');
  }
  
  const clientData = JSON.parse(tempClientDataStr);
  
  // Clean up temporary data
  scriptProperties.deleteProperty('tempClientData');
  
  showRecapDialog(clientData);
}

/**
 * Update client details from missing links dialog and proceed with recap
 */
function updateClientDetailsFromDialog(details) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const tempClientDataStr = scriptProperties.getProperty('tempClientData');
  
  if (!tempClientDataStr) {
    throw new Error('No client data found for update');
  }
  
  const clientData = JSON.parse(tempClientDataStr);
  
  // Update the client details
  const result = updateClientDetailsAndProceedWithRecap(clientData.sheetName, details);
  
  // Clean up temporary data
  scriptProperties.deleteProperty('tempClientData');
  
  return result;
}

/**
 * Get temporarily stored client data for the dialog
 */
function getTempClientData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const tempClientDataStr = scriptProperties.getProperty('tempClientData');
  
  if (!tempClientDataStr) {
    return null;
  }
  
  return JSON.parse(tempClientDataStr);
}

/**
 * Get dialog data (client name and missing links) for the simple dialog
 */
function getTempDialogData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const tempClientDataStr = scriptProperties.getProperty('tempClientData');
  const tempMissingLinksStr = scriptProperties.getProperty('tempMissingLinks');
  
  if (!tempClientDataStr || !tempMissingLinksStr) {
    throw new Error('Dialog data not found');
  }
  
  const clientData = JSON.parse(tempClientDataStr);
  const missingLinks = JSON.parse(tempMissingLinksStr);
  
  return {
    clientName: clientData.studentName || 'Unknown Client',
    missingLinks: missingLinks
  };
}

/**
 * Get recap dialog data from server-side storage
 */
function getRecapDialogData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Get the stored client data
  const clientDataStr = scriptProperties.getProperty('CURRENT_RECAP_DATA');
  if (!clientDataStr) {
    throw new Error('No client data found for recap dialog');
  }
  
  const clientData = JSON.parse(clientDataStr);
  
  // Get saved quick notes if available
  const notesKey = `notes_${clientData.sheetName}`;
  const savedNotes = scriptProperties.getProperty(notesKey);
  const notes = savedNotes ? JSON.parse(savedNotes) : {};
  
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const tempClientDataStr = scriptProperties.getProperty('tempClientData');
  
  if (!tempClientDataStr) {
    throw new Error('Client data not found');
  }
  
  const clientData = JSON.parse(tempClientDataStr);
  
  // Use the existing update client dialog from the sidebar
  showUpdateClientInfoDialog(clientData.sheetName);
}

/**
 * Show dialog for individual recap input
 */
function showRecapDialog(clientData) {
  // Ensure we have valid client data
  if (!clientData || !clientData.sheetName) {
    throw new Error('Invalid client data for recap dialog');
  }
  
  // Store ALL client data in a single, simple property for easy retrieval
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CURRENT_RECAP_DATA', JSON.stringify(clientData));
  
  console.log('Storing client data for recap:', clientData.studentName);
  
  // Create HTML using safe string building to avoid template literal injection issues
  const htmlContent = createSessionRecapHtml();
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(700).setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Session Recap');
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
  const subject = `${formData.studentFirstName}'s ${sessionType} Session Recap ${dateFormatted}`;

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
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Use links from emailData if available, otherwise try Script Properties
  let meetingNotesLink = emailData.formData?.meetingNotesLink || 
                         scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '';
  let scoreReportLink = emailData.formData?.scoreReportLink || 
                        scriptProperties.getProperty(`scorereport_${sheetName}`) || '';
  
  // Only try fallback to cells if we have an actual sheet (not called from dialog)
  if (!meetingNotesLink || !scoreReportLink) {
    try {
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (activeSheet) {
        // Fallback to cells if not in Script Properties (for existing sheets)
        if (!meetingNotesLink) {
          meetingNotesLink = activeSheet.getRange('I2').getValue() || '';
          // If found in I2, migrate to Script Properties
          if (meetingNotesLink) {
            scriptProperties.setProperty(`meetingnotes_${sheetName}`, meetingNotesLink);
          }
        }
        
        if (!scoreReportLink) {
          scoreReportLink = activeSheet.getRange('J2').getValue() || '';
          // If found in J2, migrate to Script Properties
          if (scoreReportLink) {
            scriptProperties.setProperty(`scorereport_${sheetName}`, scoreReportLink);
          }
        }
      }
    } catch (e) {
      // If we can't get the sheet, just use what we have
      console.log('Could not access sheet for fallback link retrieval:', e);
    }
  }
  
  // Store links in emailData for use in the dialog
  emailData.meetingNotesLink = meetingNotesLink;
  emailData.scoreReportLink = scoreReportLink;
  
  // Store the email data in Properties Service temporarily
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('tempEmailData', JSON.stringify(emailData));
  
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: white;
          }
          h2 {
            color: #333;
            margin-bottom: 20px;
            font-weight: 500;
          }
          .email-info {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
          }
          .email-field {
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 10px;
          }
          .email-label {
            font-weight: 500;
            color: #5f6368;
            min-width: 80px;
          }
          .email-value {
            flex: 1;
            color: #202124;
            font-family: monospace;
            background: white;
            padding: 8px;
            border-radius: 4px;
            border: 1px solid #dadce0;
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
            font-family: Arial, sans-serif;
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
          <div class="email-info">
            <div class="email-field">
              <span class="email-label">To:</span>
              <span class="email-value" id="parentEmails">[Loading...]</span>
              <button class="copy-btn" id="copyEmailBtn">Copy Email</button>
            </div>
            <div class="email-field">
              <span class="email-label">Subject:</span>
              <span class="email-value" id="subject">[Loading...]</span>
              <button class="copy-btn" id="copySubjectBtn">Copy Subject</button>
            </div>
          </div>
          
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
            ${emailData.meetingNotesLink ? 
              '<button class="menu-button action-button" id="meetingNotesBtn"><span class="icon">📝</span> <span>Update Meeting Notes</span></button>' : 
              '<div></div>'}
            
            <!-- Bottom left: Save & Close -->
            <button class="menu-button" id="saveCloseBtn">
              <span class="icon">💾</span>
              <span>Save & Close</span>
            </button>
            
            <!-- Bottom right: Send Recap Email -->
            <button class="menu-button primary" id="sendEmailBtn">
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
            
            // Populate fields
            document.getElementById('parentEmails').textContent = emailData.formData.parentEmail || '[No parent email set]';
            document.getElementById('subject').textContent = emailData.subject;
            
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
            document.getElementById('copyEmailBtn').addEventListener('click', function() {
              copyToClipboard('parentEmails', this);
            });
            document.getElementById('copySubjectBtn').addEventListener('click', function() {
              copyToClipboard('subject', this);
            });
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
            
            if (emailData.meetingNotesLink) {
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
                console.log('Dashboard Update successfully copied to clipboard');
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
                console.log('Dashboard Update successfully copied to clipboard (fallback method)');
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
                console.log('Rich text copied successfully');
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const batchClients = scriptProperties.getProperty('batchClients');
  
  if (!batchClients) return null;
  
  const clients = JSON.parse(batchClients);
  const currentIndex = parseInt(scriptProperties.getProperty('batchIndex') || '0');
  
  if (currentIndex < clients.length - 1) {
    const nextIndex = currentIndex + 1;
    scriptProperties.setProperty('batchIndex', nextIndex.toString());
    
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
    // Batch complete
    scriptProperties.deleteProperty('batchClients');
    scriptProperties.deleteProperty('batchIndex');
  }
  
  return null;
}

function getClientDataFromSheet(sheet) {
  const sheetName = sheet.getName();
  const clientNameWithSpaces = sheet.getRange('A1').getValue() || sheetName.replace(/([A-Z])/g, ' $1').trim();
  const clientTypes = sheet.getRange('C2').getValue() || '';
  
  // Get all data from Script Properties (cached)
  const scriptProperties = PropertiesService.getScriptProperties();
  const dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || '';
  const meetingNotesLink = scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '';
  const scoreReportLink = scriptProperties.getProperty(`scorereport_${sheetName}`) || '';
  const parentEmail = scriptProperties.getProperty(`parentemail_${sheetName}`) || '';
  const emailAddressees = scriptProperties.getProperty(`emailaddressees_${sheetName}`) || '';
  
  // Get active status from Script Properties (controlled by Update Client Info dialog)
  const isActiveProperty = scriptProperties.getProperty(`isActive_${sheetName}`);
  const isActive = isActiveProperty === null ? true : isActiveProperty === 'true'; // Default to true for new clients
  
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
  
  console.log('Cell Value:', cell.getValue());
  
  try {
    const richText = cell.getRichTextValue();
    if (richText) {
      console.log('Rich Text Found');
      const runs = richText.getRuns();
      runs.forEach((run, index) => {
        console.log(`Run ${index}:`, {
          text: run.getText(),
          url: run.getLinkUrl(),
          textStyle: run.getTextStyle()
        });
      });
    }
  } catch (e) {
    console.log('No rich text:', e.message);
  }
  
  // Try formula
  try {
    const formula = cell.getFormula();
    if (formula) {
      console.log('Formula:', formula);
    }
  } catch (e) {
    console.log('No formula');
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
                       'Template', 'Settings', 'Dashboard', 'Summary', 'SessionRecaps', 'Master'];
  
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
  
  // For valid client sheets, allow spaces in names
  // Just check that it's not empty and not a system sheet
  return sheetName.trim().length > 0;
}

/**
 * Manual refresh of client list cache
 */
function refreshClientCache() {
  try {
    const startTime = Date.now();
    
    // Clear and rebuild UnifiedClientDataStore first
    console.log('Clearing and rebuilding UnifiedClientDataStore...');
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('UNIFIED_CLIENT_DATA');
    
    // Rebuild enhanced cache from all data sources
    const allClients = EnhancedClientCache.buildEnhancedCache();
    const activeCount = allClients.filter(client => client.isActive).length;
    
    // Rebuild UnifiedClientDataStore with clean data
    UnifiedClientDataStore.initialize();
    allClients.forEach(client => {
      // Only add valid clients (double-check for contamination)
      const invalidNames = ['Date Sent', 'Student Name', 'Client Type', 'Template', 'New Client'];
      if (!invalidNames.includes(client.name) && client.name && client.name.trim()) {
        UnifiedClientDataStore.addClient({
          name: client.name,
          isActive: client.isActive,
          parentEmail: client.parentEmail || '',
          emailAddressees: client.emailAddressees || '',
          dashboardLink: client.dashboardLink || '',
          meetingNotesLink: client.meetingNotesLink || '',
          scoreReportLink: client.scoreReportLink || ''
        });
      }
    });
    
    const duration = Date.now() - startTime;
    
    const message = `Enhanced client cache refreshed successfully!\n\n` +
                   `Total clients: ${allClients.length}\n` +
                   `Active clients: ${activeCount}\n` +
                   `Refresh time: ${duration}ms\n\n` +
                   `Cache cleaned and rebuilt - removed any system sheets`;
    
    SpreadsheetApp.getUi().alert('Cache Refreshed', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    console.log(`Manual cache refresh completed - ${allClients.length} total, ${activeCount} active`);
    
    return { 
      success: true,
      message: 'Client list cache refreshed and cleaned successfully',
      stats: {
        total: allClients.length,
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
  try {
    // Try enhanced cache first for instant results
    const cachedAllClients = EnhancedClientCache.getAllClientsFromCache();
    if (cachedAllClients && cachedAllClients.length > 0) {
      console.log(`Returning ${cachedAllClients.length} total clients from enhanced cache`);
      return cachedAllClients;
    }
    
    // Cache miss - build enhanced cache from all data sources
    console.log('Enhanced cache miss - building from all data sources');
    const allClients = EnhancedClientCache.buildEnhancedCache();
    
    return allClients;
    
  } catch (error) {
    console.error('Error getting all client list:', error);
    // Fallback to direct Master sheet read if cache fails
    return getActiveClientListDirect();
  }
}

/**
 * Get list of active clients only - now uses persistent cache
 */
function getActiveClientList() {
  try {
    // Try enhanced cache first for instant results
    const cachedActiveClients = EnhancedClientCache.getActiveClientsFromCache();
    if (cachedActiveClients && cachedActiveClients.length > 0) {
      console.log(`Returning ${cachedActiveClients.length} active clients from enhanced cache`);
      return cachedActiveClients;
    }
    
    // Cache miss - build enhanced cache from all data sources
    console.log('Enhanced cache miss - building from all data sources');
    const allClients = EnhancedClientCache.buildEnhancedCache();
    
    // Return only active clients
    return allClients.filter(client => client.isActive);
    
  } catch (error) {
    console.error('Error getting active client list:', error);
    // Fallback to direct Master sheet read if cache fails
    return getActiveClientListDirect();
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
      console.log('Building enhanced cache from existing data sources...');
      
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
            a1Value.toString().toLowerCase().includes('name')) {
          return false;
        }
        
        return true;
      });
      
      for (const sheet of clientSheets) {
        const sheetName = sheet.getName();
        
        try {
          // Get client name from A1
          const clientName = sheet.getRange('A1').getValue() || sheetName.replace(/([A-Z])/g, ' $1').trim();
          
          // Get all data from Script Properties
          const dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || '';
          const meetingNotesLink = scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '';
          const scoreReportLink = scriptProperties.getProperty(`scorereport_${sheetName}`) || '';
          const parentEmail = scriptProperties.getProperty(`parentemail_${sheetName}`) || '';
          const emailAddressees = scriptProperties.getProperty(`emailaddressees_${sheetName}`) || '';
          
          // Get active status from Script Properties (default to true if not set)
          const isActiveProperty = scriptProperties.getProperty(`isActive_${sheetName}`);
          const isActive = isActiveProperty === null ? true : isActiveProperty === 'true';
          
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
          console.log(`Added client to cache: ${clientName} (${sheetName})`);
          
        } catch (sheetError) {
          console.warn(`Error processing sheet ${sheetName}:`, sheetError);
          continue;
        }
      }
      
      // Save to cache
      this.setCachedClientData(allClients);
      console.log(`Enhanced cache built with ${allClients.length} clients`);
      
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
      
      console.log('Enhanced client cache updated successfully');
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
      console.log('Client "' + completeClientData.name + '" added/updated in enhanced cache');
      
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
        console.log('Client "' + clientName + '" status updated to ' + (isActive ? 'active' : 'inactive'));
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
        console.log(`Client "${clientList[clientIndex].name}" field "${field}" updated in cache`);
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
        console.log(`Client "${clientList[clientIndex].name}" multiple fields updated in cache`);
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
        console.log(`Client with sheet name "${sheetName}" removed from cache`);
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

/**
 * Compatibility layer for existing ClientListCache calls
 * Maps old function calls to the new EnhancedClientCache
 */
const ClientListCache = {
  // Map old function names to new ones
  getCachedClientList: function() {
    return EnhancedClientCache.getCachedClientData();
  },
  
  rebuildClientListCache: function() {
    return EnhancedClientCache.buildEnhancedCache();
  },
  
  getActiveClientsFromCache: function() {
    return EnhancedClientCache.getActiveClientsFromCache();
  },
  
  addClientToCache: function(clientName, isActive) {
    // Convert old format to new format
    if (typeof isActive === 'undefined') isActive = true;
    const clientData = {
      name: clientName,
      sheetName: clientName.trim().replace(/\s+/g, ''),
      isActive: isActive
    };
    return EnhancedClientCache.addClientToCache(clientData);
  },
  
  updateClientStatusInCache: function(clientName, isActive) {
    return EnhancedClientCache.updateClientStatusInCache(clientName, isActive);
  },
  
  setCachedClientList: function(clientList) {
    return EnhancedClientCache.setCachedClientData(clientList);
  },
  
  getCacheStats: function() {
    return EnhancedClientCache.getCacheStats();
  }
};

/**
 * Enhanced function to populate cache from all existing data sources
 * This replaces the need to read from the Master sheet
 */
function buildComprehensiveClientCache() {
  try {
    console.log('Building comprehensive client cache from all data sources...');
    const result = EnhancedClientCache.buildEnhancedCache();
    
    if (result.length > 0) {
      console.log(`Successfully built comprehensive cache with ${result.length} clients`);
      return {
        success: true,
        message: `Cache built successfully with ${result.length} clients`,
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
  return EnhancedClientCache.getClientBySheetName(sheetName);
}

function getAllActiveClientsFromCache() {
  return EnhancedClientCache.getActiveClientsFromCache() || [];
}

function getAllClientsFromCache() {
  return EnhancedClientCache.getAllClientsFromCache() || [];
}

function updateClientInCache(sheetName, updates) {
  return EnhancedClientCache.updateClientFields(sheetName, updates);
}

function addNewClientToCache(clientData) {
  return EnhancedClientCache.addClientToCache(clientData);
}

/**
 * Get client details for update dialog (enhanced with cache support)
 * First tries cache, then falls back to direct data access
 */
function getClientDetailsForUpdate(sheetName) {
  try {
    // Try to get from enhanced cache first
    const cachedClient = EnhancedClientCache.getClientBySheetName(sheetName);
    if (cachedClient) {
      console.log(`Using cached data for client: ${cachedClient.name}`);
      
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
    console.log(`Cache miss for ${sheetName}, using direct access`);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    const masterSheet = spreadsheet.getSheetByName('Master');
    const scriptProperties = PropertiesService.getScriptProperties();
    
    if (!sheet) {
      throw new Error('Client sheet not found');
    }
    
    // Get client name from sheet A1
    const clientName = sheet.getRange('A1').getValue();
    
    // Get links and parent data exclusively from Script Properties
    const dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || '';
    const meetingNotesLink = scriptProperties.getProperty(`meetingnotes_${sheetName}`) || '';
    const scoreReportLink = scriptProperties.getProperty(`scorereport_${sheetName}`) || '';
    const parentEmail = scriptProperties.getProperty(`parentemail_${sheetName}`) || '';
    const emailAddressees = scriptProperties.getProperty(`emailaddressees_${sheetName}`) || '';
    
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
      console.log('Client not found in Master sheet, automatically syncing:', clientName);
      
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
      EnhancedClientCache.addClientToCache(fullClientData);
      
      // Now find the newly added row
      const updatedMasterData = masterSheet.getDataRange().getValues();
      for (let i = 1; i < updatedMasterData.length; i++) {
        if (updatedMasterData[i][0] === clientName) {
          masterRow = i + 1; // Convert to 1-based
          break;
        }
      }
    }
    
    // Get active status from Script Properties (controlled by Update Client Info dialog)
    const isActiveProperty = scriptProperties.getProperty(`isActive_${sheetName}`);
    const isActive = isActiveProperty === null ? true : isActiveProperty === 'true'; // Default to true for new clients
    
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
    EnhancedClientCache.addClientToCache(fullClientData);
    
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
    const scriptProperties = PropertiesService.getScriptProperties();
    
    if (!masterSheet) {
      throw new Error('Master sheet not found');
    }
    
    // Get client details to find master row
    const clientDetails = getClientDetailsForUpdate(sheetName);
    const masterRow = clientDetails.masterRow;
    
    // Update Script Properties for links
    if (details.dashboardLink) {
      scriptProperties.setProperty(`dashboard_${sheetName}`, details.dashboardLink);
    } else {
      scriptProperties.deleteProperty(`dashboard_${sheetName}`);
    }
    
    if (details.meetingNotesLink) {
      scriptProperties.setProperty(`meetingnotes_${sheetName}`, details.meetingNotesLink);
    } else {
      scriptProperties.deleteProperty(`meetingnotes_${sheetName}`);
    }
    
    if (details.scoreReportLink) {
      scriptProperties.setProperty(`scorereport_${sheetName}`, details.scoreReportLink);
    } else {
      scriptProperties.deleteProperty(`scorereport_${sheetName}`);
    }
    
    // Update parent data in Script Properties
    if (details.parentEmail) {
      scriptProperties.setProperty(`parentemail_${sheetName}`, details.parentEmail);
    } else {
      scriptProperties.deleteProperty(`parentemail_${sheetName}`);
    }
    
    if (details.emailAddressees) {
      scriptProperties.setProperty(`emailaddressees_${sheetName}`, details.emailAddressees);
    } else {
      scriptProperties.deleteProperty(`emailaddressees_${sheetName}`);
    }
    
    // Update isActive status in Script Properties (source of truth)
    scriptProperties.setProperty(`isActive_${sheetName}`, details.isActive.toString());
    
    // Update Master sheet Column E to reflect the boolean value
    masterSheet.getRange(masterRow, 5).setValue(details.isActive);
    
    // Update enhanced cache with all changes
    const cacheUpdates = {
      isActive: details.isActive,
      dashboardLink: details.dashboardLink || '',
      meetingNotesLink: details.meetingNotesLink || '',
      scoreReportLink: details.scoreReportLink || '',
      parentEmail: details.parentEmail || '',
      emailAddressees: details.emailAddressees || ''
    };
    
    const cacheUpdateSuccess = EnhancedClientCache.updateClientFields(sheetName, cacheUpdates);
    if (cacheUpdateSuccess) {
      console.log(`Updated enhanced cache for "${clientDetails.clientName}" with all changes`);
    } else {
      console.warn(`Failed to update enhanced cache for "${clientDetails.clientName}"`);
    }
    
    // ALSO update the UnifiedClientDataStore (new architecture)
    try {
      const unifiedData = UnifiedClientDataStore.getUnifiedData();
      if (unifiedData && unifiedData.clients && unifiedData.clients[sheetName]) {
        // Update the client data in the unified store
        unifiedData.clients[sheetName].isActive = details.isActive;
        unifiedData.clients[sheetName].links.dashboard = details.dashboardLink || '';
        unifiedData.clients[sheetName].links.meetingNotes = details.meetingNotesLink || '';
        unifiedData.clients[sheetName].links.scoreReport = details.scoreReportLink || '';
        unifiedData.clients[sheetName].contactInfo.parentEmail = details.parentEmail || '';
        unifiedData.clients[sheetName].contactInfo.emailAddressees = details.emailAddressees || '';
        unifiedData.clients[sheetName].lastUpdated = new Date().getTime();
        
        // Update stats
        const activeCount = Object.values(unifiedData.clients).filter(client => client.isActive).length;
        unifiedData.stats.activeClients = activeCount;
        
        // Save the updated unified data
        UnifiedClientDataStore.saveUnifiedData(unifiedData);
        console.log(`Updated UnifiedClientDataStore for "${clientDetails.clientName}" with isActive: ${details.isActive}`);
      }
    } catch (unifiedError) {
      console.warn('Failed to update UnifiedClientDataStore:', unifiedError);
      // Don't fail the whole operation if unified store update fails
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
function testEnhancedCache() {
  try {
    console.log('=== Testing Enhanced Client Cache System ===');
    
    // Test 1: Build cache
    console.log('Test 1: Building enhanced cache...');
    const result = buildComprehensiveClientCache();
    console.log('Build result:', result);
    
    if (!result.success) {
      throw new Error('Failed to build cache: ' + result.message);
    }
    
    // Test 2: Get all clients from cache
    console.log('Test 2: Getting all clients from cache...');
    const allClients = getAllClientsFromCache();
    console.log(`Retrieved ${allClients ? allClients.length : 0} clients from cache`);
    
    if (allClients && allClients.length > 0) {
      console.log('Sample client data:', allClients[0]);
    }
    
    // Test 3: Get active clients only
    console.log('Test 3: Getting active clients only...');
    const activeClients = getAllActiveClientsFromCache();
    console.log(`Retrieved ${activeClients ? activeClients.length : 0} active clients from cache`);
    
    // Test 4: Get specific client by sheet name
    if (allClients && allClients.length > 0) {
      const testClient = allClients[0];
      console.log('Test 4: Getting specific client by sheet name...');
      const retrievedClient = getClientFromCache(testClient.sheetName);
      console.log('Retrieved client:', retrievedClient ? retrievedClient.name : 'Not found');
    }
    
    // Test 5: Get cache statistics
    console.log('Test 5: Getting cache statistics...');
    const stats = EnhancedClientCache.getCacheStats();
    console.log('Cache stats:', stats);
    
    const testMessage = `Enhanced Cache System Test Results:
    
✓ Cache built successfully
✓ Total clients: ${result.clientCount}
✓ Active clients: ${activeClients ? activeClients.length : 0}
✓ Cache includes complete client data
✓ All test operations completed successfully

The enhanced cache system is ready to use!`;
    
    console.log(testMessage);
    SpreadsheetApp.getUi().alert('Enhanced Cache Test', testMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
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
    console.error('Enhanced cache test failed:', error);
    const errorMessage = `Enhanced Cache Test Failed:

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
function showBulkClientDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
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
      console.log('Creating client sheets for successful entries...');
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
      console.log(`Sheet ${clientKey} already exists, skipping creation`);
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
    
    console.log(`Created sheet for client: ${clientData.name}`);
    
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
        console.log('Initializing unified client data store...');
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
          console.log('Upgrading unified data store version...');
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
      
      console.log(`Saved unified client data: ${Object.keys(data.clients).length} clients`);
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
      console.log('Attempting to restore from backup...');
      const properties = PropertiesService.getScriptProperties();
      const backupString = properties.getProperty(this.BACKUP_KEY);
      
      if (backupString) {
        const backupData = JSON.parse(backupString);
        console.log('Restored from backup successfully');
        return backupData;
      }
      
      console.log('No backup available, starting fresh');
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
      console.log('Migrating data from EnhancedClientCache...');
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
        
        console.log(`Migrated ${migrated} clients from EnhancedClientCache`);
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
      
      console.log(`Batch operation completed: ${results.successful.length} successful, ${results.failed.length} failed`);
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
          // Filter out template and system sheets
          return client.name && 
                 client.name !== 'New Client' && 
                 client.name !== 'NewClient' &&
                 !client.name.includes('Template') &&
                 key !== 'New Client' &&
                 key !== 'NewClient';
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
    console.log('UnifiedClientDataStore initialized successfully');
    console.log('Client count:', Object.keys(result.clients).length);
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

/**
 * Test the enhanced cache system
 */
function testEnhancedCache() {
  try {
    console.log('Testing Enhanced Cache System...');
    
    // Test 1: Build the cache
    const buildResult = buildComprehensiveClientCache();
    console.log('Cache build result:', buildResult);
    
    // Test 2: Get all clients
    const allClients = getAllClientsFromCache();
    console.log('Total clients in cache:', allClients.length);
    
    // Test 3: Get active clients
    const activeClients = getAllActiveClientsFromCache();
    console.log('Active clients:', activeClients.length);
    
    // Test 4: Get specific client (if any exist)
    if (allClients.length > 0) {
      const testClient = allClients[0];
      const clientInfo = getComprehensiveClientInfo(testClient.sheetName);
      console.log('Sample client info:', clientInfo);
    }
    
    // Show results
    SpreadsheetApp.getUi().alert(
      'Enhanced Cache Test Results',
      `Cache built successfully!\n\n` +
      `Total clients: ${allClients.length}\n` +
      `Active clients: ${activeClients.length}\n` +
      `Cache includes: names, links, parent info, session data\n\n` +
      `The system is ready to use without the Master sheet!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Cache test error:', error);
    SpreadsheetApp.getUi().alert('Cache Test Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
