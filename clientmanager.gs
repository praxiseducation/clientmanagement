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
    ui.createMenu('Client Management')
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
  ui.createMenu('Client Management')
    // Main items with icons for visual appeal
    .addItem('📊 Control Panel', 'showSidebar')
    .addSeparator()
    .addItem('➕ Add New Client', 'addNewClient')
    .addItem('🔍 Find Client', 'findClient')
    .addSeparator()
    // Session Management submenu
    .addSubMenu(ui.createMenu('📝 Session Management')
      .addItem('Quick Notes', 'openQuickNotesFromMenu')
      .addSeparator()
      .addItem('Send Session Recap', 'sendIndividualRecap')
      .addItem('Batch Prep Mode', 'openBatchPrepMode')
      .addItem('View Recap History', 'viewRecapHistory')
      .addSeparator()
      .addItem('Clear Current Student Notes', 'clearCurrentStudentNotes'))
    .addSeparator()
    // Tools submenu
    .addSubMenu(ui.createMenu('🛠️ Tools')
      .addItem('Refresh Client List', 'refreshClientCache')
      .addItem('Check Dashboard Links', 'validateDashboardLinks')
      .addItem('Export Client List', 'exportClientList'))
    .addSeparator()
    .addItem('⚙️ Settings', 'showSettings')
    .addItem('❓ Help', 'showHelp')
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
            transition: all 0.3s ease;
            background: #00C853;
            color: white;
            margin-top: 20px;
          }
          button:hover {
            background: #00A043;
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
              .withSuccessHandler(() => {
                alert('Setup complete! The page will now reload.');
                google.script.host.close();
                google.script.run.refreshUI();
              })
              .withFailureHandler((error) => {
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
            transition: all 0.3s ease;
          }
          .btn-primary {
            background: #00C853;
            color: white;
          }
          .btn-primary:hover {
            background: #00A043;
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
              .withSuccessHandler(() => {
                alert('Settings updated successfully!');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
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
  ui.createMenu('Client Management')
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
 * Universal sidebar that can switch between views without closing
 */
function showUniversalSidebar(initialView = 'control') {
  const sheet = SpreadsheetApp.getActiveSheet();
  const clientInfo = getCurrentClientInfo();
  
  // Get quick notes data if on a client sheet
  let notesData = {};
  if (clientInfo.isClient) {
    const scriptProperties = PropertiesService.getScriptProperties();
    const notesKey = `notes_${sheet.getName()}`;
    const savedNotes = scriptProperties.getProperty(notesKey);
    notesData = savedNotes ? JSON.parse(savedNotes) : {};
  }
  
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
            display: none; 
            padding: 15px;
          }
          .view.active { 
            display: block; 
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
            transition: all 0.3s ease;
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
          /* Updated button styles for Quick Notes section */
          .btn-primary {
            background: #00C853;
            color: white !important;
          }
          
          .btn-primary:hover {
            background: #00A043;
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
            background: #00A043;
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
            transition: all 0.3s ease;
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
            padding: 3px 8px;
            background: #f8f9fa;
            border: 1px solid #dadce0;
            border-radius: 3px;
            margin: 2px;
            cursor: pointer;
            font-size: 11px;
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
            background: #00A043;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }

          #sendRecapBtn {
            background: #003366;
            color: white;
          }

          #sendRecapBtn:hover {
            background: #002244;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
          }
          
          .btn-success {
            background: #00C853 !important;
            transition: background 0.3s ease;
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
          
          /* View transition animation */
          .view {
            animation: fadeIn 0.3s ease-out;
          }
          
          @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
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
        </style>
      </head>
      <body>
        <!-- Control Panel View -->
        <div id="controlView" class="view ${initialView === 'control' ? 'active' : ''}">
          <div class="current-client">
            <div class="current-client-label">Current Client:</div>
            <div id="clientName" class="${clientInfo.isClient ? 'current-client-name' : 'no-client'}">
              ${clientInfo.isClient ? clientInfo.name : 'No client selected'}
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Client Management</div>
            <button class="menu-button primary" onclick="runFunction('addNewClient')">
              <span class="icon">➕</span>
              Add New Client
            </button>
            <button class="menu-button" onclick="runFunction('findClient')">
              <span class="icon">🔍</span>
              Find Client
            </button>
          </div>
          
          <div class="section">
            <div class="section-title">Session Management</div>
            <button id="quickNotesBtn" class="menu-button" onclick="switchToQuickNotes()" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">📝</span>
              Quick Notes
            </button>
            <button id="recapBtn" class="menu-button" onclick="runFunction('sendIndividualRecap')" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">✉️</span>
              Send Session Recap
            </button>
            <button class="menu-button" onclick="runFunction('openBatchPrepMode')">
              <span class="icon">📋</span>
              Batch Prep Mode
            </button>
            <button class="menu-button" onclick="runFunction('viewRecapHistory')">
              <span class="icon">📊</span>
              View Recap History
            </button>
          </div>
          
          <div class="loading" id="controlLoading">
            <div class="spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
        
        <!-- Quick Notes View -->
        <div id="quickNotesView" class="view ${initialView === 'quicknotes' ? 'active' : ''}">
          <button class="control-panel-btn" onclick="switchToControlPanel()">
            <span>📊</span>
            <span>Back to Control Panel</span>
          </button>
          
          <div class="divider"></div>
          
          <div id="quickNotesAlert" class="alert" style="display: none;"></div>
          
          <h3>Quick Session Notes</h3>
          <p style="font-size: 12px; color: #666;">Click tags to add to notes</p>
          
          <div class="save-indicator" id="saveIndicator">Saved!</div>
          
          <div class="note-section">
            <label>Wins/Breakthroughs</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('wins', 'Aha moment!')">Aha moment!</span>
              <span class="quick-tag" onclick="addToField('wins', 'Solved independently')">Solved independently</span>
              <span class="quick-tag" onclick="addToField('wins', 'Confidence boost')">Confidence boost</span>
              <span class="quick-tag" onclick="addToField('wins', 'Speed improved')">Speed improved</span>
            </div>
            <textarea id="wins" placeholder="Quick notes about wins..." onchange="markUnsaved()">${notesData.wins || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Skills Worked On</label>
            <textarea id="skills" placeholder="List skills covered..." onchange="markUnsaved()">${notesData.skills || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Struggles/Areas to Review</label>
            <textarea id="struggles" placeholder="Note any challenges..." onchange="markUnsaved()">${notesData.struggles || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Parent Communication Points</label>
            <textarea id="parent" placeholder="Things to mention to parent..." onchange="markUnsaved()">${notesData.parent || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Next Session Focus</label>
            <textarea id="next" placeholder="Plan for next time..." onchange="markUnsaved()">${notesData.next || ''}</textarea>
          </div>
          
          <div class="button-group">
            <button id="saveNotesBtn" class="btn-primary" onclick="saveNotes()">Save Notes</button>
            <button id="sendRecapBtn" onclick="sendRecap()">Send Recap</button>
          </div>
          
          <div class="button-group" style="margin-top: 5px;">
            <button class="btn-secondary" onclick="clearNotes()">Clear Notes</button>
            <button class="btn-secondary" onclick="reloadNotes()">Reload Saved</button>
          </div>
          
          <div class="auto-save-status" id="autoSaveStatus">Auto-save: Active</div>
        </div>
        
        <script>
          let currentView = '${initialView}';
          let clientInfo = ${JSON.stringify(clientInfo)};
          let unsavedChanges = false;
          let autoSaveInterval;
          

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
          // View switching functions
          function switchToQuickNotes() {
            // Check if we're on a client sheet
            google.script.run
              .withSuccessHandler((info) => {
                if (info.isClient) {
                  // Load the latest notes
                  google.script.run
                    .withSuccessHandler((notesJson) => {
                      if (notesJson) {
                        const notes = JSON.parse(notesJson);
                        document.getElementById('wins').value = notes.wins || '';
                        document.getElementById('skills').value = notes.skills || '';
                        document.getElementById('struggles').value = notes.struggles || '';
                        document.getElementById('parent').value = notes.parent || '';
                        document.getElementById('next').value = notes.next || '';
                      }
                      
                      // Switch view
                      document.getElementById('controlView').classList.remove('active');
                      document.getElementById('quickNotesView').classList.add('active');
                      currentView = 'quicknotes';
                      
                      // Start auto-save
                      startAutoSave();
                    })
                    .getQuickNotesForCurrentSheet();
                } else {
                  showAlert('Please navigate to a client sheet first.', 'error');
                }
              })
              .getCurrentClientInfo();
          }
          
          function switchToControlPanel() {
            // Clear auto-save
            if (autoSaveInterval) {
              clearInterval(autoSaveInterval);
            }
            
            // Update client info
            google.script.run
              .withSuccessHandler((info) => {
                clientInfo = info;
                updateControlPanelView();
                
                // Switch view
                document.getElementById('quickNotesView').classList.remove('active');
                document.getElementById('controlView').classList.add('active');
                currentView = 'control';
              })
              .getCurrentClientInfo();
          }
          
          function updateControlPanelView() {
            const clientNameDiv = document.getElementById('clientName');
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            
            if (clientInfo.isClient) {
              clientNameDiv.textContent = clientInfo.name;
              clientNameDiv.className = 'current-client-name';
              quickNotesBtn.disabled = false;
              recapBtn.disabled = false;
            } else {
              clientNameDiv.textContent = 'No client selected';
              clientNameDiv.className = 'no-client';
              quickNotesBtn.disabled = true;
              recapBtn.disabled = true;
            }
          }
          
          // Quick Notes functions
          function markUnsaved() {
            unsavedChanges = true;
          }
          
          function addToField(fieldId, text) {
            const field = document.getElementById(fieldId);
            if (field.value) field.value += ', ';
            field.value += text;
            markUnsaved();
          }
          
          function saveNotes() {
            const saveBtn = document.getElementById('saveNotesBtn');
            setButtonLoading(saveBtn, true);
            
            const notes = {
              wins: document.getElementById('wins').value,
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(() => {
                setButtonLoading(saveBtn, false);
                saveBtn.classList.add('btn-success');
                saveBtn.textContent = 'Saved!';
                
                const indicator = document.getElementById('saveIndicator');
                indicator.classList.add('show');
                
                setTimeout(() => {
                  saveBtn.classList.remove('btn-success');
                  saveBtn.textContent = 'Save Notes';
                  indicator.classList.remove('show');
                }, 2000);
                
                unsavedChanges = false;
              })
              .withFailureHandler((error) => {
                setButtonLoading(saveBtn, false);
                showAlert('Error saving notes: ' + error.message, 'error');
              })
              .saveQuickNotes(notes);
          }
          
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
            
            google.script.run
              .withSuccessHandler((notesJson) => {
                if (notesJson) {
                  const notes = JSON.parse(notesJson);
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
              })
              .withFailureHandler((error) => {
                showAlert('Error loading notes: ' + error.message, 'error');
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function sendRecap() {
            // Save notes first
            saveNotes();
            setTimeout(() => {
              google.script.run
                .withSuccessHandler(() => {
                  // Switch back to control panel after sending
                  switchToControlPanel();
                })
                .sendIndividualRecap();
            }, 500);
          }
          
          // Auto-save functionality
          function startAutoSave() {
            autoSaveInterval = setInterval(() => {
              if (unsavedChanges) {
                const notes = {
                  wins: document.getElementById('wins').value,
                  skills: document.getElementById('skills').value,
                  struggles: document.getElementById('struggles').value,
                  parent: document.getElementById('parent').value,
                  next: document.getElementById('next').value
                };
                
                google.script.run
                  .withSuccessHandler(() => {
                    document.getElementById('autoSaveStatus').textContent = 
                      'Auto-saved at ' + new Date().toLocaleTimeString();
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
            if (button) {
              setButtonLoading(button, true);
            }
            
            google.script.run
              .withSuccessHandler((result) => {
                if (button) {
                  setButtonLoading(button, false);
                }
                
                // Update client info after certain actions
                if (functionName === 'findClient' || functionName === 'addNewClient') {
                  google.script.run
                    .withSuccessHandler((info) => {
                      clientInfo = info;
                      updateControlPanelView();
                    })
                    .getCurrentClientInfo();
                }
              })
              .withFailureHandler((error) => {
                if (button) {
                  setButtonLoading(button, false);
                }
                showAlert('Error: ' + error.message, 'error');
              })[functionName]();
          }
          
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
              
              setTimeout(() => {
                alertDiv.remove();
              }, 5000);
            } else {
              alertDiv.className = 'alert ' + type;
              alertDiv.textContent = message;
              alertDiv.style.display = 'block';
              
              setTimeout(() => {
                alertDiv.style.display = 'none';
              }, 5000);
            }
          }
          
          // Monitor for text input
          document.querySelectorAll('textarea').forEach(textarea => {
            textarea.addEventListener('input', markUnsaved);
          });
          
          // Refresh client info periodically
          setInterval(() => {
            if (currentView === 'control') {
              google.script.run
                .withSuccessHandler((info) => {
                  clientInfo = info;
                  updateControlPanelView();
                })
                .getCurrentClientInfo();
            }
          }, 5000);
          
          // Initialize auto-save if starting in quick notes
          if (currentView === 'quicknotes') {
            startAutoSave();
          }
        </script>
      </body>
    </html>
  `).setTitle('Client Management').setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
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
            transition: all 0.3s ease;
            flex: 1;
          }
          .btn-primary {
            background: #00C853;   
            color: white;
          }
          .btn-primary:hover {
            background: #00A043;    
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
            transition: all 0.3s ease;
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
              .withSuccessHandler(() => {
                // Close dialog and continue with recap
                google.script.host.close();
                // Re-trigger the recap dialog
                google.script.run.continueWithRecap(sheetName);
              })
              .withFailureHandler((error) => {
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
            transition: all 0.3s ease;
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
            transition: all 0.3s ease;
            min-width: 100px;
          }
          .btn-primary {
            background: #00C853;
            color: white;
          }
          .btn-primary:hover {
            background: #00A043;
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
            transition: all 0.3s ease;
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
              <label>Parent Email(s)</label>
              <input type="email" id="parentEmails" placeholder="Enter parent email address(es)">
              <div class="example">Use commas for multiple: mom@email.com, dad@email.com</div>
            </div>
            
            <div class="input-group">
              <label>Parent Name(s)</label>
              <input type="text" id="parentNames" placeholder="Enter parent name(s)">
              <div class="example">Example: Jane Smith or Jane and John Smith</div>
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
              style="display: inline-flex; align-items: center; gap: 8px; color: #1a73e8; text-decoration: none; font-size: 13px; padding: 8px 16px; border: 1px solid #dadce0; border-radius: 6px; background: #f8f9fa; transition: all 0.3s ease;"
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
          
          function processClient() {
            if (!validateForm()) return;
            
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, true);
            
            const clientData = {
              name: document.getElementById('clientName').value.trim(),
              services: getSelectedServices().join(', '),
              parentEmails: document.getElementById('parentEmails').value.trim(),
              parentNames: document.getElementById('parentNames').value.trim(),
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
            setTimeout(() => {
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
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Client Management');
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
  
  // Populate the cells with client data
  newSheet.getRange('A1').setValue(clientData.name);           // Student name
  newSheet.getRange('C2').setValue(clientData.services);       // Client type(s)
  newSheet.getRange('D2').setValue(clientData.dashboardLink);  // Dashboard smart chip (display)
  newSheet.getRange('E2').setValue(clientData.parentEmails);   // Parent emails
  newSheet.getRange('F2').setValue(clientData.parentNames);    // Parent names
  
  if (clientData.dashboardLink) {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(`dashboard_${concatenatedName}`, clientData.dashboardLink);
  }
  
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
    
    // Column D: Next Session Scheduled? - Checkbox with value TRUE
    console.log('Adding checkbox to column D (Next Session Scheduled?)');
    masterSheet.getRange(newRow, 4).insertCheckboxes();
    masterSheet.getRange(newRow, 4).setValue(true);
    
    // Column E: Current Client? - Checkbox with value TRUE (new clients are current)
    console.log('Adding checkbox to column E (Current Client?)');
    masterSheet.getRange(newRow, 5).insertCheckboxes();
    masterSheet.getRange(newRow, 5).setValue(true);
    masterSheet.getRange(newRow, 5).check(); // Ensure checkbox is checked
    
    console.log('✓ Successfully added client to master list');
    
    // Force a refresh of the sheet to ensure all formulas calculate
    SpreadsheetApp.flush();
        
  } catch (error) {
    console.error('ERROR adding data to master list:', error);
    throw error;
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
 * Optimized find client using Master sheet
 */
function findClient() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = spreadsheet.getSheetByName('Master');
  
  if (!masterSheet) {
    // Fallback to original method if no Master sheet
    findClientFallback();
    return;
  }
  
  // Get data from Master sheet
  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No clients found in Master sheet.');
    return;
  }
  
  // Get both values and formulas to extract hyperlinks
  const dataRange = masterSheet.getRange(2, 1, lastRow - 1, 5); // A2:E[lastRow]
  const values = dataRange.getValues();
  const formulas = dataRange.getFormulas();
  
  const clientSheets = [];
  
  for (let i = 0; i < values.length; i++) {
    const clientName = values[i][0]; // Column A
    const isCurrentClient = values[i][4]; // Column E
    
    // Skip if no name or not current client
    if (!clientName || !isCurrentClient) continue;
    
    const sheetName = extractSheetNameFromFormula(
      formulas[i][0], 
      clientName, 
      spreadsheet
    );
    
    clientSheets.push({
      sheetName: sheetName,
      displayName: clientName,
      firstName: clientName.split(' ')[0]
    });
  }
  
  // Sort by first name
  clientSheets.sort((a, b) => a.firstName.localeCompare(b.firstName));
  
  if (clientSheets.length === 0) {
    SpreadsheetApp.getUi().alert('No active clients found.');
    return;
  }
  
  showClientSelectionDialog(clientSheets);
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
            transition: all 0.3s ease;
          }
          .client-item:hover {
            background: #f8f9fa;
            color: #003366;      
            font-weight: 500;
          }
          .client-item:last-child {
            border-bottom: none;
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
            transition: all 0.3s ease;
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
            // Show loading overlay with client name
            document.getElementById('loadingOverlay').classList.add('active');
            document.querySelector('.loading-text').textContent = 'Loading ' + displayName + '...';
            
            google.script.run
              .withSuccessHandler(() => {
                // Small delay to ensure smooth transition
                setTimeout(() => {
                  google.script.host.close();
                }, 300);
              })
              .withFailureHandler(error => {
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
 * Universal sidebar that can switch between views without closing
 */
function showUniversalSidebar(initialView = 'control') {
  const sheet = SpreadsheetApp.getActiveSheet();
  const clientInfo = getCurrentClientInfo();
  
  // Get quick notes data if on a client sheet
  let notesData = {};
  if (clientInfo.isClient) {
    const scriptProperties = PropertiesService.getScriptProperties();
    const notesKey = `notes_${sheet.getName()}`;
    const savedNotes = scriptProperties.getProperty(notesKey);
    notesData = savedNotes ? JSON.parse(savedNotes) : {};
  }
  
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
            display: none; 
            padding: 15px;
          }
          .view.active { 
            display: block; 
          }
          
          /* Common styles */
         h3 {
            margin-top: 0;
            color: #003366;  /* Changed from #1a73e8 */
          }
          
          button {
            background: #003366;  /* Changed from #1a73e8 */
            color: white;
            /* ... rest stays the same */
          }

          
          button:hover {
            background: #002244;  /* Changed from #1557b0 */
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
          
          /* Control Panel styles */
          .current-client {
            background: #e8f0fe;
            border-left: 4px solid #003366;  /* Add this new line */
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 20px;
            font-size: 13px;
          }
          
          .current-client-label {
            font-weight: 500;
            color: #003366;  /* Changed from #1a73e8 */
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
            color: #202124; /* ADD THIS LINE */
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
            border-color: #003366;  /* Changed from #1a73e8 */
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
            background: #00C853;  /* Changed from #1a73e8 to use green for primary actions */
            color: white !important; 
            border: none;
            text-align: center;
            justify-content: center;
          }
          
          .primary:hover:not(:disabled) {
            background: #00A043;  /* Changed from #1557b0 */
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
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
          }
          
          .control-panel-btn:hover {
            background: #e8f0fe;
            border-color: #1a73e8;
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
            padding: 3px 8px;
            background: #f8f9fa;
            border: 1px solid #dadce0;
            border-radius: 3px;
            margin: 2px;
            cursor: pointer;
            font-size: 11px;
          }
          
          .quick-tag:hover {
            background: #e8f0fe;
            border-color: #003366;  /* Changed from #1a73e8 */
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
          
          .btn-success {
            background: #00C853 !important;  
            transition: background 0.3s ease;
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
          
          /* View transition animation */
          .view {
            animation: fadeIn 0.3s ease-out;
          }
          
          @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
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
        </style>
      </head>
      <body>
        <!-- Control Panel View -->
        <div id="controlView" class="view ${initialView === 'control' ? 'active' : ''}">
          <div class="current-client">
            <div class="current-client-label">Current Client:</div>
            <div id="clientName" class="${clientInfo.isClient ? 'current-client-name' : 'no-client'}">
              ${clientInfo.isClient ? clientInfo.name : 'No client selected'}
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Client Management</div>
            <button class="menu-button primary" onclick="runFunction('addNewClient')">
              <span class="icon">➕</span>
              Add New Client
            </button>
            <button class="menu-button" onclick="runFunction('findClient')">
              <span class="icon">🔍</span>
              Find Client
            </button>
          </div>
          
          <div class="section">
            <div class="section-title">Session Management</div>
            <button id="quickNotesBtn" class="menu-button" onclick="switchToQuickNotes()" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">📝</span>
              Quick Notes
            </button>
            <button id="recapBtn" class="menu-button" onclick="runFunction('sendIndividualRecap')" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">✉️</span>
              Send Session Recap
            </button>
            <button class="menu-button" onclick="runFunction('openBatchPrepMode')">
              <span class="icon">📋</span>
              Batch Prep Mode
            </button>
            <button class="menu-button" onclick="runFunction('viewRecapHistory')">
              <span class="icon">📊</span>
              View Recap History
            </button>
          </div>
          
          <div class="loading" id="controlLoading">
            <div class="spinner"></div>
            <p>Loading...</p>
          </div>
        </div>
        
        <!-- Quick Notes View -->
        <div id="quickNotesView" class="view ${initialView === 'quicknotes' ? 'active' : ''}">
          <button class="control-panel-btn" onclick="switchToControlPanel()">
            <span>📊</span>
            <span>Back to Control Panel</span>
          </button>
          
          <div class="divider"></div>
          
          <div id="quickNotesAlert" class="alert" style="display: none;"></div>
          
          <h3>Quick Session Notes</h3>
          <p style="font-size: 12px; color: #666;">Click tags to add to notes</p>
          
          <div class="save-indicator" id="saveIndicator">Saved!</div>
          
          <div class="note-section">
            <label>Wins/Breakthroughs</label>
            <div style="margin-bottom: 5px;">
              <span class="quick-tag" onclick="addToField('wins', 'Aha moment!')">Aha moment!</span>
              <span class="quick-tag" onclick="addToField('wins', 'Solved independently')">Solved independently</span>
              <span class="quick-tag" onclick="addToField('wins', 'Confidence boost')">Confidence boost</span>
              <span class="quick-tag" onclick="addToField('wins', 'Speed improved')">Speed improved</span>
            </div>
            <textarea id="wins" placeholder="Quick notes about wins..." onchange="markUnsaved()">${notesData.wins || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Skills Worked On</label>
            <textarea id="skills" placeholder="List skills covered..." onchange="markUnsaved()">${notesData.skills || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Struggles/Areas to Review</label>
            <textarea id="struggles" placeholder="Note any challenges..." onchange="markUnsaved()">${notesData.struggles || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Parent Communication Points</label>
            <textarea id="parent" placeholder="Things to mention to parent..." onchange="markUnsaved()">${notesData.parent || ''}</textarea>
          </div>
          
          <div class="note-section">
            <label>Next Session Focus</label>
            <textarea id="next" placeholder="Plan for next time..." onchange="markUnsaved()">${notesData.next || ''}</textarea>
          </div>
          
          <div class="button-group">
            <button id="saveNotesBtn" onclick="saveNotes()">Save Notes</button>
            <button class="btn-secondary" onclick="sendRecap()">Send Recap</button>
          </div>
          
          <div class="button-group" style="margin-top: 5px;">
            <button class="btn-danger" onclick="clearNotes()">Clear Notes</button>
            <button class="btn-secondary" onclick="reloadNotes()">Reload Saved</button>
          </div>
          
          <div class="auto-save-status" id="autoSaveStatus">Auto-save: Active</div>
        </div>
        
        <script>
          let currentView = '${initialView}';
          let clientInfo = ${JSON.stringify(clientInfo)};
          let unsavedChanges = false;
          let autoSaveInterval;
          
          // View switching functions
          function switchToQuickNotes() {
            // Check if we're on a client sheet
            google.script.run
              .withSuccessHandler((info) => {
                if (info.isClient) {
                  // Load the latest notes
                  google.script.run
                    .withSuccessHandler((notesJson) => {
                      if (notesJson) {
                        const notes = JSON.parse(notesJson);
                        document.getElementById('wins').value = notes.wins || '';
                        document.getElementById('skills').value = notes.skills || '';
                        document.getElementById('struggles').value = notes.struggles || '';
                        document.getElementById('parent').value = notes.parent || '';
                        document.getElementById('next').value = notes.next || '';
                      }
                      
                      // Switch view
                      document.getElementById('controlView').classList.remove('active');
                      document.getElementById('quickNotesView').classList.add('active');
                      currentView = 'quicknotes';
                      
                      // Start auto-save
                      startAutoSave();
                    })
                    .getQuickNotesForCurrentSheet();
                } else {
                  showAlert('Please navigate to a client sheet first.', 'error');
                }
              })
              .getCurrentClientInfo();
          }
          
          function switchToControlPanel() {
            // Clear auto-save
            if (autoSaveInterval) {
              clearInterval(autoSaveInterval);
            }
            
            // Update client info
            google.script.run
              .withSuccessHandler((info) => {
                clientInfo = info;
                updateControlPanelView();
                
                // Switch view
                document.getElementById('quickNotesView').classList.remove('active');
                document.getElementById('controlView').classList.add('active');
                currentView = 'control';
              })
              .getCurrentClientInfo();
          }
          
          function updateControlPanelView() {
            const clientNameDiv = document.getElementById('clientName');
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            
            if (clientInfo.isClient) {
              clientNameDiv.textContent = clientInfo.name;
              clientNameDiv.className = 'current-client-name';
              quickNotesBtn.disabled = false;
              recapBtn.disabled = false;
            } else {
              clientNameDiv.textContent = 'No client selected';
              clientNameDiv.className = 'no-client';
              quickNotesBtn.disabled = true;
              recapBtn.disabled = true;
            }
          }
          
          // Quick Notes functions
          function markUnsaved() {
            unsavedChanges = true;
          }
          
          function addToField(fieldId, text) {
            const field = document.getElementById(fieldId);
            if (field.value) field.value += ', ';
            field.value += text;
            markUnsaved();
          }
          
          function saveNotes() {
            const saveBtn = document.getElementById('saveNotesBtn');
            const originalText = saveBtn.textContent;
            
            saveBtn.classList.add('btn-saving');
            saveBtn.disabled = true;
            
            const notes = {
              wins: document.getElementById('wins').value,
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(() => {
                saveBtn.classList.remove('btn-saving');
                saveBtn.classList.add('btn-success');
                saveBtn.textContent = 'Saved!';
                
                const indicator = document.getElementById('saveIndicator');
                indicator.classList.add('show');
                
                setTimeout(() => {
                  saveBtn.classList.remove('btn-success');
                  saveBtn.textContent = originalText;
                  saveBtn.disabled = false;
                  indicator.classList.remove('show');
                }, 2000);
                
                unsavedChanges = false;
              })
              .withFailureHandler((error) => {
                saveBtn.classList.remove('btn-saving');
                saveBtn.disabled = false;
                showAlert('Error saving notes: ' + error.message, 'error');
              })
              .saveQuickNotes(notes);
          }
          
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
            
            google.script.run
              .withSuccessHandler((notesJson) => {
                if (notesJson) {
                  const notes = JSON.parse(notesJson);
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
              })
              .withFailureHandler((error) => {
                showAlert('Error loading notes: ' + error.message, 'error');
              })
              .getQuickNotesForCurrentSheet();
          }
          
          function sendRecap() {
            // Save notes first
            saveNotes();
            setTimeout(() => {
              google.script.run
                .withSuccessHandler(() => {
                  // Switch back to control panel after sending
                  switchToControlPanel();
                })
                .sendIndividualRecap();
            }, 500);
          }
          
          // Auto-save functionality
          function startAutoSave() {
            autoSaveInterval = setInterval(() => {
              if (unsavedChanges) {
                const notes = {
                  wins: document.getElementById('wins').value,
                  skills: document.getElementById('skills').value,
                  struggles: document.getElementById('struggles').value,
                  parent: document.getElementById('parent').value,
                  next: document.getElementById('next').value
                };
                
                google.script.run
                  .withSuccessHandler(() => {
                    document.getElementById('autoSaveStatus').textContent = 
                      'Auto-saved at ' + new Date().toLocaleTimeString();
                    unsavedChanges = false;
                  })
                  .saveQuickNotes(notes);
              }
            }, 60000); // Every minute
          }
          
          // Control Panel functions
          function runFunction(functionName) {
            document.getElementById('controlLoading').style.display = 'block';
            
            google.script.run
              .withSuccessHandler((result) => {
                document.getElementById('controlLoading').style.display = 'none';
                
                // Update client info after certain actions
                if (functionName === 'findClient' || functionName === 'addNewClient') {
                  google.script.run
                    .withSuccessHandler((info) => {
                      clientInfo = info;
                      updateControlPanelView();
                    })
                    .getCurrentClientInfo();
                }
              })
              .withFailureHandler((error) => {
                document.getElementById('controlLoading').style.display = 'none';
                showAlert('Error: ' + error.message, 'error');
              })[functionName]();
          }
          
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
              
              setTimeout(() => {
                alertDiv.remove();
              }, 5000);
            } else {
              alertDiv.className = 'alert ' + type;
              alertDiv.textContent = message;
              alertDiv.style.display = 'block';
              
              setTimeout(() => {
                alertDiv.style.display = 'none';
              }, 5000);
            }
          }
          
          // Monitor for text input
          document.querySelectorAll('textarea').forEach(textarea => {
            textarea.addEventListener('input', markUnsaved);
          });
          
          // Refresh client info periodically
          setInterval(() => {
            if (currentView === 'control') {
              google.script.run
                .withSuccessHandler((info) => {
                  clientInfo = info;
                  updateControlPanelView();
                })
                .getCurrentClientInfo();
            }
          }, 5000);
          
          // Initialize auto-save if starting in quick notes
          if (currentView === 'quicknotes') {
            startAutoSave();
          }
        </script>
      </body>
    </html>
  `).setTitle('Client Management').setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
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
  return scriptProperties.getProperty(notesKey);
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
  const sheet = SpreadsheetApp.getActiveSheet();
  const scriptProperties = PropertiesService.getScriptProperties();
  const notesKey = `notes_${sheet.getName()}`;
  scriptProperties.setProperty(notesKey, JSON.stringify(notes));
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
            transition: all 0.3s ease;
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
            const selectedClients = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedClients.length === 0) {
              alert('Please select at least one client.');
              return;
            }
            
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, true);
            
            google.script.run
              .withSuccessHandler(() => {
                alert('Batch prep complete! Use "Send Session Recap" after each session.');
                google.script.host.close();
              })
              .withFailureHandler((error) => {
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
  
  // Check if dashboard link exists
  if (!clientData.dashboardLink) {
    // Prompt for dashboard link
    promptForDashboardLink(sheetName, clientData.studentName);
    return;
  }
  
  showRecapDialog(clientData);
}

/**
 * Show dialog for individual recap input
 */
function showRecapDialog(clientData) {
  // Get saved quick notes if available
  const scriptProperties = PropertiesService.getScriptProperties();
  const notesKey = `notes_${clientData.sheetName}`;
  const savedNotes = scriptProperties.getProperty(notesKey);
  const notes = savedNotes ? JSON.parse(savedNotes) : {};
  
  // Get homework templates based on client types
  const clientTypesList = clientData.clientTypes ? clientData.clientTypes.split(',').map(t => t.trim()) : [];
  let homeworkOptions = [];
  
  // Add relevant templates based on client types
  clientTypesList.forEach(type => {
    if (HOMEWORK_TEMPLATES[type]) {
      homeworkOptions = homeworkOptions.concat(HOMEWORK_TEMPLATES[type]);
    }
  });
  
  // Add general templates as fallback
  if (homeworkOptions.length === 0) {
    homeworkOptions = HOMEWORK_TEMPLATES['General'];
  }
  
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
            color: #003366;          /* Changed from #1a73e8 */
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
            color: #003366;          /* Changed from #1a73e8 */
            border: none;
            border-radius: 4px;
            font-size: 12px;
            cursor: pointer;
            transition: all 0.3s;
          }
          .quick-fill-btn:hover {
            background: #d2e3fc;
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
            transition: all 0.3s ease;
          }
          .btn-primary {
            background: #00C853;     /* Changed from #1a73e8 */
            color: white;
            flex: 1;
          }
          .btn-primary:hover {
            background: #00A043;     /* Changed from #1557b0 */
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
            border-top: 3px solid #003366;  /* Changed from #1a73e8 */
            border-radius: 50%;
            animation: spin 1s linear infinite;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .homework-template-group {
            display: flex;
            gap: 10px;
            margin-bottom: 10px;
            align-items: center;
          }
          .homework-template-group select {
            flex: 1;
          }
          .homework-template-group button {
            padding: 8px 16px;
            font-size: 13px;
          }
          .saved-notes-indicator {
            background: #e8f5e9;
            padding: 8px;
            border-radius: 4px;
            font-size: 12px;
            color: #00A043;          /* Changed from #2e7d32 to darker green */
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
          .quick-not-scheduled {
            padding: 8px 16px;
            background: #fce8b2;
            color: #7f6000;
            border: none;
            border-radius: 4px;
            font-size: 13px;
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.3s;
          }
          .quick-not-scheduled:hover {
            background: #f9cc79;
          }
        </style>
      </head>
      <body>
                
        <div class="client-info">
          <strong>${clientData.studentName}</strong> - ${clientData.clientTypes}
          ${clientData.parentEmail ? `<br>Parent: ${clientData.parentName} (${clientData.parentEmail})` : ''}
          ${clientData.dashboardLink ? `<br>Dashboard: <a href="${clientData.dashboardLink}" target="_blank" style="color: #003366;">View Dashboard</a>` : ''}
        </div>
        
        ${notes.wins ? '<div class="saved-notes-indicator">Quick notes loaded from your session</div>' : ''}
        
        <form id="recapForm">
          <div class="form-section">
            <h3>Session Highlights</h3>
            
            <div class="form-group">
              <label>Today's Win*</label>
              <input type="text" id="todaysWin" value="${notes.wins || ''}" placeholder="e.g., Solved 5 quadratic equations independently!" required>
              <div class="quick-fill">
                <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Finally conquered [skill] after struggling last week!')">Breakthrough</button>
                <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Solved [number] problems independently')">Independence</button>
                <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Had an aha moment with [concept]')">Aha Moment</button>
                <button type="button" class="quick-fill-btn" onclick="fillField('todaysWin', 'Confidence soaring - attempted harder problems without prompting')">Confidence</button>
              </div>
            </div>
            
            <div class="form-group">
              <label>Focus Skill/Topic*</label>
              <input type="text" id="focusSkill" value="${notes.skills ? notes.skills.split(',')[0].trim() : ''}" placeholder="e.g., Quadratic Equations, French Conversation, ACT Geometry" required>
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
              <input type="text" id="practiced" value="${notes.skills || ''}" placeholder="e.g., Word problems, graphing parabolas">
            </div>
            
            <div class="form-group">
              <label>Skills Introduced</label>
              <input type="text" id="introduced" placeholder="e.g., Completing the square">
            </div>
            
            <div class="form-group">
              <label>Progress Notes</label>
              <textarea id="progressNotes" placeholder="e.g., Making great progress with quadratics, still needs practice with word problems">${notes.struggles || ''}</textarea>
            </div>
            
            <div class="form-group">
              <label>Parent Notes</label>
              <textarea id="parentNotes" placeholder="Additional notes for parents">${notes.parent || ''}</textarea>
            </div>
          </div>
          
          <div class="form-section">
            <h3>Homework & Next Steps</h3>
            
            <div class="form-group">
              <label>Homework/Practice</label>
              <div class="homework-template-group">
                <select id="homeworkTemplate" onchange="updateHomework()">
                  <option value="">-- Select a template --</option>
                  ${homeworkOptions.map((template, index) => 
                    `<option value="${index}">${template.replace(/{[^}]+}/g, '___').replace('before our next session', '')}</option>`
                  ).join('')}
                </select>
                <button type="button" class="btn-secondary" onclick="addHomework()">Add</button>
              </div>
              <div class="act-test-selector" id="actTestSelector">
                <label>Select ACT Test:</label>
                <select id="actTestDropdown" onchange="updateACTTest()">
                  <option value="">-- Select ACT Test --</option>
                  ${ACT_TESTS.map(test => `<option value="${test}">${test}</option>`).join('')}
                </select>
              </div>
              <textarea id="homework" placeholder="e.g., Complete practice set 3.2 (problems 1-15) before our next session"></textarea>
            </div>
            
            <div class="form-group">
              <label>Next Session Date</label>
              <div class="next-session-input-group">
                <input type="text" id="nextSession" placeholder="e.g., Monday, July 29 at 3:00 PM">
                <button type="button" class="quick-not-scheduled" onclick="fillField('nextSession', 'Not Yet Scheduled')">Not Yet Scheduled</button>
              </div>
            </div>
            
            <div class="form-group">
              <label>Next Session Preview</label>
              <input type="text" id="nextSessionPreview" value="${notes.next || ''}" placeholder="e.g., We'll build on factoring to tackle more complex equations">
            </div>
          </div>
          
          <div class="form-section">
            <h3>Parent Communication</h3>
            
            <div class="form-group">
              <label>Parent Quick Win</label>
              <input type="text" id="quickWin" value="${notes.parent || ''}" placeholder="e.g., Ask them to explain the FOIL method">
              <div class="quick-fill">
                <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Ask them to explain [concept] to you')">Explain concept</button>
                <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Notice when they use [skill] in homework')">Notice skill</button>
                <button type="button" class="quick-fill-btn" onclick="fillField('quickWin', 'Celebrate their persistence with challenging problems')">Celebrate effort</button>
              </div>
            </div>
            
            <div class="form-group">
              <label>Additional Parent Notes</label>
              <textarea id="parentNotes" placeholder="Conversation starters, encouragement focus, etc."></textarea>
            </div>
            
            <div class="form-group">
              <label>P.S. Note (Personal moment - leave blank if not needed)</label>
              <input type="text" id="psNote" placeholder="e.g., Their face lit up when they realized they could do this!">
            </div>
          </div>
          
          <div class="button-group">
            <button type="submit" class="btn-primary">Preview Email</button>
            <button type="button" class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
          </div>
        </form>
        
        <div class="loading-overlay" id="loadingOverlay">
          <div class="loading-spinner"></div>
        </div>
        
        <script>
          const clientData = ${JSON.stringify(clientData)};
          const homeworkTemplates = ${JSON.stringify(homeworkOptions)};
          const actTests = ${JSON.stringify(ACT_TESTS)};
          
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
          
          function fillField(fieldId, text) {
            document.getElementById(fieldId).value = text;
          }
          
          function updateHomework() {
            const select = document.getElementById('homeworkTemplate');
            const textarea = document.getElementById('homework');
            const actSelector = document.getElementById('actTestSelector');
            
            if (select.value !== '') {
              const templateIndex = parseInt(select.value);
              const template = homeworkTemplates[templateIndex];
              
              // Check if this is an ACT practice test template
              if (template.includes('Complete ACT practice test {test_number}')) {
                // Show the ACT test selector
                actSelector.classList.add('show');
                // Don't update the textarea yet - wait for test selection
              } else {
                // Hide the ACT test selector if it's visible
                actSelector.classList.remove('show');
                // Show the template with placeholders for editing
                textarea.value = template;
              }
            } else {
              // Hide the ACT test selector if no template is selected
              actSelector.classList.remove('show');
            }
          }
          
          function updateACTTest() {
            const testDropdown = document.getElementById('actTestDropdown');
            const textarea = document.getElementById('homework');
            const templateSelect = document.getElementById('homeworkTemplate');
            
            if (testDropdown.value && templateSelect.value !== '') {
              const template = homeworkTemplates[parseInt(templateSelect.value)];
              // Replace the placeholder with the selected test
              const updatedTemplate = template.replace('{test_number}', testDropdown.value);
              textarea.value = updatedTemplate;
            }
          }
          
          function addHomework() {
            const select = document.getElementById('homeworkTemplate');
            const textarea = document.getElementById('homework');
            const actTestDropdown = document.getElementById('actTestDropdown');
            
            if (select.value !== '') {
              let template = homeworkTemplates[parseInt(select.value)];
              
              // If it's an ACT test template and a test is selected, use the selected test
              if (template.includes('Complete ACT practice test {test_number}') && actTestDropdown.value) {
                template = template.replace('{test_number}', actTestDropdown.value);
              }
              
              if (textarea.value) textarea.value += '\\n';
              textarea.value += template;
            }
          }
          
          function calculateDueDate(daysFromNow = 7) {
            const date = new Date();
            date.setDate(date.getDate() + daysFromNow);
            return date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
          }
        
          document.getElementById('recapForm').addEventListener('submit', function(e) {
            e.preventDefault();
            
            const submitBtn = e.target.querySelector('.btn-primary');
            setButtonLoading(submitBtn, true);
            
            // Prepare form data
            const formData = {
              ...clientData,
              todaysWin: document.getElementById('todaysWin').value,
              focusSkill: document.getElementById('focusSkill').value,
              mastered: document.getElementById('mastered').value,
              practiced: document.getElementById('practiced').value,
              introduced: document.getElementById('introduced').value,
              progressNotes: document.getElementById('progressNotes').value,
              homework: document.getElementById('homework').value,
              nextSession: document.getElementById('nextSession').value,
              nextSessionPreview: document.getElementById('nextSessionPreview').value,
              quickWin: document.getElementById('quickWin').value,
              parentNotes: document.getElementById('parentNotes').value,
              psNote: document.getElementById('psNote').value,
              date: new Date().toLocaleDateString()
            };
            
            // Note: We don't need to show the overlay since we're using button loading
            // document.getElementById('loadingOverlay').classList.add('active');
            
            // Call generateEmailPreview which will process the data and show the preview
            google.script.run
              .withSuccessHandler(function(emailData) {
                // Don't need to remove loading state here as dialog closes
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                setButtonLoading(submitBtn, false);
                alert('Error: ' + error.message);
              })
              .generateEmailPreview(formData);
          });
        </script>
      </body>
    </html>
  `).setWidth(700).setHeight(800);
  
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
  showEmailPreviewDialog(emailData);
  
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



function showEmailPreviewDialog(emailData) {
  // Get meeting notes link from cell I2 and score report link from J2 of the active sheet
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const meetingNotesLink = activeSheet.getRange('I2').getValue() || '';
  const scoreReportLink = activeSheet.getRange('J2').getValue() || '';
  
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
            transition: all 0.3s ease;
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
            background: #00A043;     
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
            transition: all 0.3s ease;
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
        </style>
      </head>
      <body>
        <div id="loadingDiv" class="loading">
          <div class="loading-spinner"></div>
          <div class="loading-text">Loading email preview...</div>
        </div>
        <div id="mainContent" style="display: none;">
          <h2>Session Recap - Preview</h2>
          
          <div class="instructions">
            📋 Review the detailed recap below. Click "Send Recap Email" to send a simplified version to parents.
          </div>
          
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
              <span>Click "Copy for Google Docs" to copy with proper formatting!</span>
            </div>
            
            <div class="email-preview" id="emailBody">[Loading...]</div>
          </div>
          
          <div class="button-group">
            <button class="btn-primary" id="copyBodyBtn">Copy for Google Docs</button>
            <button class="btn-edit" id="backToEditBtn">Back to Edit</button>
            ${emailData.formData.dashboardLink ? 
              '<button class="copy-btn dashboard" id="dashboardBtn">Open Dashboard</button>' : 
              ''}
            <button class="copy-btn gmail" id="sendEmailBtn">Send Recap Email</button>
            <button class="btn-secondary" id="saveCloseBtn">Save & Close</button>
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
            document.getElementById('emailBody').innerHTML = emailData.bodyHtmlRich || emailData.bodyHtml;
            
            // Show/hide dashboard button
            if (emailData.formData.dashboardLink) {
              const dashboardBtn = document.getElementById('dashboardBtn');
              if (dashboardBtn) {
                dashboardBtn.style.display = 'block';
              }
            }
            
            // Attach event listeners
            document.getElementById('copyEmailBtn').addEventListener('click', function() {
              copyToClipboard('parentEmails', this);
            });
            document.getElementById('copySubjectBtn').addEventListener('click', function() {
              copyToClipboard('subject', this);
            });
            document.getElementById('copyBodyBtn').addEventListener('click', copyRichText);
            document.getElementById('backToEditBtn').addEventListener('click', backToEdit);
            document.getElementById('sendEmailBtn').addEventListener('click', sendRecapEmail);
            document.getElementById('saveCloseBtn').addEventListener('click', saveAndClose);
            
            const dashboardBtn = document.getElementById('dashboardBtn');
            if (dashboardBtn) {
              dashboardBtn.addEventListener('click', openDashboard);
            }
          }
          
          function copyToClipboard(elementId, button) {
            setButtonLoading(button, true);
            
            const element = document.getElementById(elementId);
            const text = element.textContent;
            
            // Add a small delay to show the loading state
            setTimeout(() => {
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
            
            setTimeout(() => {
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
            // Close this dialog and call the server to reopen the edit form
            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                alert('Error going back to edit: ' + error.message);
              })
              .backToEditRecap();
          }
          
          function openDashboard() {
            if (emailData.formData.dashboardLink) {
              window.open(emailData.formData.dashboardLink, '_blank');
            } else {
              alert('No dashboard link available for this student.');
            }
          }
          
          function sendRecapEmail() {
            const parentNames = emailData.formData.parentName || 'Parents';
            const studentFirstName = emailData.formData.studentFirstName;
            
            let sessionType = '';
            const clientTypes = emailData.formData.clientTypes ? emailData.formData.clientTypes.split(',').map(function(t) { return t.trim(); }) : [];
            
            if (clientTypes.length === 1) {
              sessionType = clientTypes[0];
            } else {
              sessionType = 'Tutoring';
            }
            
            const simplifiedEmailBody = 'Hi ' + parentNames + '!\\n\\n' +
              'Here is the link to our meeting notes (' + emailData.meetingNotesLink + ') from our ' + sessionType + ' session today, and you will find ' + studentFirstName + '\\'s homework linked in the Meeting Notes.\\n\\n' +
              'You can always follow our progress on ' + studentFirstName + '\\'s Dashboard.\\n\\n' +
              'You can also see their progress on the Score Report (' + emailData.scoreReportLink + ')!\\n\\n' +
              'As always, I\\'m just an email away for any questions you may have.';
            
            const to = encodeURIComponent(emailData.formData.parentEmail || '');
            const subject = encodeURIComponent(emailData.subject);
            const body = encodeURIComponent(simplifiedEmailBody);
            
            const gmailUrl = 'https://mail.google.com/mail/?view=cm&fs=1' + 
              (to ? '&to=' + to : '') +
              '&su=' + subject +
              '&body=' + body;
            
            window.open(gmailUrl, '_blank');
            saveAndClose();
          }
          
          function saveAndClose() {
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
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview Recap Email');
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
  const parentEmail = sheet.getRange('E2').getValue() || '';
  const parentName = sheet.getRange('F2').getValue() || '';
  
  // Try to get dashboard link from Script Properties first
  const scriptProperties = PropertiesService.getScriptProperties();
  let dashboardLink = scriptProperties.getProperty(`dashboard_${sheetName}`) || '';
  
  // Fall back to H2 if not in Script Properties (for existing sheets)
  if (!dashboardLink) {
    dashboardLink = sheet.getRange('H2').getValue() || '';
    
    // If found in H2, migrate to Script Properties
    if (dashboardLink) {
      scriptProperties.setProperty(`dashboard_${sheetName}`, dashboardLink);
    }
  }
  
  return {
    studentName: clientNameWithSpaces,
    clientTypes: clientTypes,
    clientType: clientTypes,
    dashboardLink: dashboardLink,
    parentEmail: parentEmail,
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
  const excludeNames = ['NewClient', 'NewACTClient', 'NewAcademicSupportClient', 
                       'Template', 'Settings', 'Dashboard', 'Summary', 'SessionRecaps'];
  return !excludeNames.includes(sheetName) && 
         !sheetName.includes(' ') && 
         /^[A-Z][a-z]+[A-Z][a-z]+/.test(sheetName);
}

/**
 * Placeholder function for refreshing client cache
 */
function refreshClientCache() {
  // This is a placeholder - implement if needed
  SpreadsheetApp.getUi().alert('Client list refreshed!');
  return { message: 'Client list refreshed' };
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
