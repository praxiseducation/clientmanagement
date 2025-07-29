/**
 * Complete Client Management and Session Recap System
 * Combines client management with automated session recap emails
 */

// Global configuration
const CONFIG = {
  tutorName: 'Nicholas',
  tutorEmail: 'leekenick@gmail.com', // Update this with your email
  company: 'Smart College'
};

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
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Client Management')
    .addItem('Add New Client', 'addNewClient')
    .addSeparator()
    .addItem('Find Client', 'findClient')
    .addSeparator()
    .addSubMenu(ui.createMenu('Session Management')
      .addItem('Quick Notes', 'openQuickNotes')
      .addSeparator()
      .addItem('Send Session Recap', 'sendIndividualRecap')
      .addItem('Batch Prep Mode', 'openBatchPrepMode')
      .addItem('View Recap History', 'viewRecapHistory'))
    .addToUi();
}

/**
 * Main function to add a new client sheet
 */
function addNewClient() {
  showClientTypeDialog();
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
            border-radius: 8px;
            font-size: 16px;
            font-family: inherit;
            transition: border-color 0.3s ease;
            box-sizing: border-box;
          }
          input[type="text"]:focus, input[type="email"]:focus, input[type="url"]:focus {
            outline: none;
            border-color: #1a73e8;
            box-shadow: 0 0 0 3px rgba(26, 115, 232, 0.1);
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
            border-radius: 8px;
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
            background-color: rgba(138, 180, 138, 0.1);
            border-color: #8ab48a;
          }
          .service-checkbox.math:hover {
            background-color: rgba(217, 131, 122, 0.1);
            border-color: #d9837a;
          }
          .service-checkbox.language:hover {
            background-color: rgba(159, 124, 172, 0.1);
            border-color: #9f7cac;
          }
          .service-checkbox.selected {
            border-color: #1a73e8;
            background-color: #f8f9fa;
          }
          .service-checkbox.act.selected {
            background-color: rgba(138, 180, 138, 0.15);
            border-color: #8ab48a;
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
            background: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background: #1557b0;
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
            border-top: 4px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            transition: all 0.3s ease;
          }
          
          .loading-spinner.success {
            animation: none;
            border: 4px solid #34a853;
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
              <div class="example">Use semicolons for multiple: mom@email.com; dad@email.com</div>
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
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `).setWidth(550).setHeight(700);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Client Management');
}

/**
 * Creates a new client sheet based on the template and adds to master list
 * @param {Object} clientData - Object containing all client information
 */
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
  newSheet.getRange('H2').setValue(clientData.dashboardLink);  // Dashboard actual link
  
  // Move the new sheet to position 3 (index 2, since it's 0-based)
  const totalSheets = spreadsheet.getSheets().length;
  const targetPosition = Math.min(3, totalSheets);
  
  // Activate the new sheet first, then move it
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(targetPosition);
  
  // Add client to master list
  try {
    addClientToMasterList(clientData, concatenatedName, spreadsheet);
  } catch (e) {
    console.error('Error adding client to master list:', e);
    // Continue even if master list update fails
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
    
    // Column C: Last Session - Formula to get last non-empty cell from student's sheet
    // This will show the most recent entry from their sheet
    const lastSessionFormula = `=INDEX(FILTER(${sheetName}!A:A,${sheetName}!A:A<>""), COUNTA(${sheetName}!A:A)`;
    console.log('Setting column C (Last Session) formula:', lastSessionFormula);
    masterSheet.getRange(newRow, 3).setFormula(lastSessionFormula);
    
    // Column D: Next Session Scheduled? - Checkbox with value TRUE (assuming they're scheduled)
    console.log('Adding checkbox to column D (Next Session Scheduled?)');
    masterSheet.getRange(newRow, 4).insertCheckboxes();
    masterSheet.getRange(newRow, 4).setValue(true);
    
    // Column E: Current Client? - Checkbox with value TRUE (new clients are current)
    console.log('Adding checkbox to column E (Current Client?)');
    masterSheet.getRange(newRow, 5).insertCheckboxes();
    masterSheet.getRange(newRow, 5).setValue(true);
    
    console.log('✓ Successfully added client to master list');
        
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
 * Find and navigate to a client sheet
 */
function findClient() {
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
          }
          .client-item:hover {
            background: #f8f9fa;
            color: #1a73e8;
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
            padding: 10px 20px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: #f8f9fa;
            color: #5f6368;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
          }
          button:hover {
            background: #f1f3f4;
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
            border-top: 4px solid #1a73e8;
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
 * Updated openQuickNotes function with custom clear confirmation
 */
function openQuickNotes() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!isClientSheet(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Please navigate to a client sheet first.');
    return;
  }
  
  // Get student name for the confirmation dialog
  const studentName = sheet.getRange('A1').getValue();
  
  // Get any existing notes
  const scriptProperties = PropertiesService.getScriptProperties();
  const notesKey = `notes_${sheet.getName()}`;
  const savedNotes = scriptProperties.getProperty(notesKey);
  const notes = savedNotes ? JSON.parse(savedNotes) : {};
  
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 15px;
            font-size: 14px;
          }
          h3 {
            margin-top: 0;
            color: #1a73e8;
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
          }
          label {
            display: block;
            margin-bottom: 4px;
            font-weight: 500;
            color: #5f6368;
            font-size: 12px;
          }
          button {
            background: #1a73e8;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
            margin-top: 10px;
            width: 48%;
          }
          button:hover {
            background: #1557b0;
          }
          .button-group {
            display: flex;
            gap: 4%;
            width: 100%;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
            border-color: #d2d6dc;
          }
          .btn-danger {
            background: #ea4335;
            color: white;
          }
          .btn-danger:hover {
            background: #d33b2c;
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
            border-color: #1a73e8;
          }
          .save-indicator {
            position: fixed;
            top: 10px;
            right: 10px;
            padding: 5px 10px;
            background: #34a853;
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
          
          /* Custom confirmation overlay styles */
          .confirmation-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            animation: fadeIn 0.3s ease-out;
          }
          
          .confirmation-overlay.show {
            display: flex;
          }
          
          @keyframes fadeIn {
            from {
              opacity: 0;
            }
            to {
              opacity: 1;
            }
          }
          
          /* Loading overlay styles */
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(10px);
            -webkit-backdrop-filter: blur(10px);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 9999;
            animation: fadeIn 0.3s ease-out;
          }
          
          .loading-overlay.show {
            display: flex;
          }
          
          .loading-content {
            text-align: center;
          }
          
          .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1a73e8;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          .loading-text {
            font-size: 14px;
            color: #5f6368;
            font-weight: 500;
          }
          
          .confirmation-dialog {
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
            max-width: 400px;
            text-align: center;
            animation: slideUp 0.3s ease-out;
          }
          
          @keyframes slideUp {
            from {
              transform: translateY(20px);
              opacity: 0;
            }
            to {
              transform: translateY(0);
              opacity: 1;
            }
          }
          
          @keyframes slideDown {
            from {
              transform: translateY(0);
              opacity: 1;
            }
            to {
              transform: translateY(20px);
              opacity: 0;
            }
          }
          
          .confirmation-overlay.hiding {
            animation: fadeOut 0.3s ease-out forwards;
          }
          
          .confirmation-overlay.hiding .confirmation-dialog {
            animation: slideDown 0.3s ease-out forwards;
          }
          
          @keyframes fadeOut {
            from {
              opacity: 1;
            }
            to {
              opacity: 0;
            }
          }
          
          .confirmation-icon {
            font-size: 48px;
            margin-bottom: 20px;
          }
          
          .confirmation-title {
            font-size: 18px;
            font-weight: 500;
            color: #202124;
            margin-bottom: 10px;
          }
          
          .confirmation-message {
            font-size: 14px;
            color: #5f6368;
            margin-bottom: 25px;
            line-height: 1.5;
          }
          
          .confirmation-buttons {
            display: flex;
            gap: 12px;
            justify-content: center;
          }
          
          .confirm-btn {
            padding: 10px 24px;
            border: none;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }
          
          .confirm-btn-danger {
            background: #ea4335;
            color: white;
          }
          
          .confirm-btn-danger:hover {
            background: #d33b2c;
            transform: translateY(-1px);
          }
          
          .confirm-btn-cancel {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          
          .confirm-btn-cancel:hover {
            background: #f1f3f4;
            border-color: #d2d6dc;
          }
        </style>
      </head>
      <body>
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
          <textarea id="wins" placeholder="Quick notes about wins..." onchange="markUnsaved()">${notes.wins || ''}</textarea>
        </div>
        
        <div class="note-section">
          <label>Skills Worked On</label>
          <textarea id="skills" placeholder="List skills covered..." onchange="markUnsaved()">${notes.skills || ''}</textarea>
        </div>
        
        <div class="note-section">
          <label>Struggles/Areas to Review</label>
          <textarea id="struggles" placeholder="Note any challenges..." onchange="markUnsaved()">${notes.struggles || ''}</textarea>
        </div>
        
        <div class="note-section">
          <label>Parent Communication Points</label>
          <textarea id="parent" placeholder="Things to mention to parent..." onchange="markUnsaved()">${notes.parent || ''}</textarea>
        </div>
        
        <div class="note-section">
          <label>Next Session Focus</label>
          <textarea id="next" placeholder="Plan for next time..." onchange="markUnsaved()">${notes.next || ''}</textarea>
        </div>
        
        <div class="button-group">
          <button onclick="saveNotes()">Save Notes</button>
          <button class="btn-secondary" onclick="sendRecap()">Send Recap</button>
        </div>
        
        <div class="button-group" style="margin-top: 5px;">
          <button class="btn-danger" onclick="showClearConfirmation()">Clear Notes</button>
          <button class="btn-secondary" onclick="reloadNotes()">Reload Saved</button>
        </div>
        
        <div class="auto-save-status" id="autoSaveStatus">Auto-save: Active</div>
        
        <!-- Custom confirmation overlay for Clear Notes -->
        <div class="confirmation-overlay" id="confirmationOverlay">
          <div class="confirmation-dialog">
            <div class="confirmation-icon">⚠️</div>
            <div class="confirmation-title">Clear Quick Notes?</div>
            <div class="confirmation-message">
              This will clear all quick notes for ${studentName}. Continue?
            </div>
            <div class="confirmation-buttons">
              <button class="confirm-btn confirm-btn-danger" onclick="confirmClearNotes()">Yes, Clear Notes</button>
              <button class="confirm-btn confirm-btn-cancel" onclick="hideClearConfirmation()">Cancel</button>
            </div>
          </div>
        </div>
        
        <!-- Custom confirmation overlay for Reload -->
        <div class="confirmation-overlay" id="reloadConfirmationOverlay">
          <div class="confirmation-dialog">
            <div class="confirmation-icon">🔄</div>
            <div class="confirmation-title">Reload Saved Notes?</div>
            <div class="confirmation-message">
              You have unsaved changes. Reload anyway?
            </div>
            <div class="confirmation-buttons">
              <button class="confirm-btn confirm-btn-danger" onclick="performReload()">Yes, Reload</button>
              <button class="confirm-btn confirm-btn-cancel" onclick="hideReloadConfirmation()">Cancel</button>
            </div>
          </div>
        </div>
        
        <!-- Loading overlay for clearing notes -->
        <div class="loading-overlay" id="clearLoadingOverlay">
          <div class="loading-content">
            <div class="loading-spinner"></div>
            <div class="loading-text">Clearing notes...</div>
          </div>
        </div>
        
        <!-- Loading overlay for reloading notes -->
        <div class="loading-overlay" id="reloadLoadingOverlay">
          <div class="loading-content">
            <div class="loading-spinner"></div>
            <div class="loading-text">Reloading last saved notes...</div>
          </div>
        </div>
        
        <!-- Success message overlay -->
        <div class="confirmation-overlay" id="successOverlay">
          <div class="confirmation-dialog" style="background: #e8f5e9;">
            <div class="confirmation-icon" style="color: #2e7d32;">✓</div>
            <div class="confirmation-title" style="color: #2e7d32;">Success!</div>
            <div class="confirmation-message" id="successMessage" style="color: #2e7d32;">
              Notes reloaded from saved version.
            </div>
          </div>
        </div>
        
        <!-- Error message overlay -->
        <div class="confirmation-overlay" id="errorOverlay">
          <div class="confirmation-dialog">
            <div class="confirmation-icon" style="color: #d93025;">❌</div>
            <div class="confirmation-title" style="color: #d93025;">Error</div>
            <div class="confirmation-message" id="errorMessage" style="color: #d93025;">
              No saved notes found for this student.
            </div>
            <div class="confirmation-buttons">
              <button class="confirm-btn confirm-btn-cancel" onclick="hideErrorOverlay()">OK</button>
            </div>
          </div>
        </div>
        
        <script>
          let unsavedChanges = false;
          let autoSaveInterval;
          
          // Mark that we have unsaved changes
          function markUnsaved() {
            unsavedChanges = true;
          }
          
          // Add to field function
          function addToField(fieldId, text) {
            const field = document.getElementById(fieldId);
            if (field.value) field.value += ', ';
            field.value += text;
            markUnsaved();
          }
          
          // Save notes function
          function saveNotes() {
            const notes = {
              wins: document.getElementById('wins').value,
              skills: document.getElementById('skills').value,
              struggles: document.getElementById('struggles').value,
              parent: document.getElementById('parent').value,
              next: document.getElementById('next').value
            };
            
            google.script.run
              .withSuccessHandler(() => {
                // Show save indicator
                const indicator = document.getElementById('saveIndicator');
                indicator.classList.add('show');
                setTimeout(() => {
                  indicator.classList.remove('show');
                }, 2000);
                
                unsavedChanges = false;
              })
              .withFailureHandler((error) => {
                alert('Error saving notes: ' + error.message);
              })
              .saveQuickNotes(notes);
          }
          
          // Show custom confirmation overlay
          function showClearConfirmation() {
            document.getElementById('confirmationOverlay').classList.add('show');
          }
          
          // Hide custom confirmation overlay
          function hideClearConfirmation() {
            document.getElementById('confirmationOverlay').classList.remove('show');
          }
          
          // Confirm clear notes
          function confirmClearNotes() {
            // Hide the confirmation dialog
            hideClearConfirmation();
            
            // Show loading overlay
            document.getElementById('clearLoadingOverlay').classList.add('show');
            
            // Simulate clearing process with a small delay for visual effect
            setTimeout(() => {
              document.getElementById('wins').value = '';
              document.getElementById('skills').value = '';
              document.getElementById('struggles').value = '';
              document.getElementById('parent').value = '';
              document.getElementById('next').value = '';
              
              // Mark as having unsaved changes (empty state)
              unsavedChanges = true;
              
              // Hide loading overlay
              document.getElementById('clearLoadingOverlay').classList.remove('show');
              
              // Show success message
              showSuccessMessage('Cleared!');
            }, 800);
          }
          
          // Reload saved notes
          function reloadNotes() {
            if (unsavedChanges) {
              showReloadConfirmation();
            } else {
              performReload();
            }
          }
          
          // Show reload confirmation overlay
          function showReloadConfirmation() {
            document.getElementById('reloadConfirmationOverlay').classList.add('show');
          }
          
          // Hide reload confirmation overlay
          function hideReloadConfirmation() {
            document.getElementById('reloadConfirmationOverlay').classList.remove('show');
          }
          
          // Show success message overlay
          function showSuccessMessage(message) {
            document.getElementById('successMessage').textContent = message;
            const overlay = document.getElementById('successOverlay');
            overlay.classList.remove('hiding');
            overlay.classList.add('show');
            
            setTimeout(() => {
              overlay.classList.add('hiding');
              setTimeout(() => {
                overlay.classList.remove('show', 'hiding');
              }, 300); // Wait for animation to complete
            }, 2000);
          }
          
          // Show error message overlay
          function showErrorMessage(message) {
            document.getElementById('errorMessage').textContent = message;
            document.getElementById('errorOverlay').classList.add('show');
          }
          
          // Hide error overlay
          function hideErrorOverlay() {
            document.getElementById('errorOverlay').classList.remove('show');
          }
          
          // Perform the actual reload
          function performReload() {
            // Hide confirmation dialog if it's showing
            hideReloadConfirmation();
            
            // Show loading overlay
            document.getElementById('reloadLoadingOverlay').classList.add('show');
            
            google.script.run
              .withSuccessHandler((savedNotes) => {
                // Hide loading overlay
                document.getElementById('reloadLoadingOverlay').classList.remove('show');
                
                if (savedNotes) {
                  const notes = JSON.parse(savedNotes);
                  document.getElementById('wins').value = notes.wins || '';
                  document.getElementById('skills').value = notes.skills || '';
                  document.getElementById('struggles').value = notes.struggles || '';
                  document.getElementById('parent').value = notes.parent || '';
                  document.getElementById('next').value = notes.next || '';
                  
                  unsavedChanges = false;
                  showSuccessMessage('Notes reloaded from saved version.');
                } else {
                  showErrorMessage('No saved notes found for this student.');
                }
              })
              .withFailureHandler((error) => {
                // Hide loading overlay
                document.getElementById('reloadLoadingOverlay').classList.remove('show');
                showErrorMessage('Error loading notes: ' + error.message);
              })
              .getQuickNotesForCurrentSheet();
          }
          
          // Send recap function
          function sendRecap() {
            // Save notes first, then close sidebar and open recap
            saveNotes();
            setTimeout(() => {
              google.script.run
                .withSuccessHandler(() => {
                  google.script.host.close();
                })
                .sendIndividualRecap();
            }, 500);
          }
          
          // Auto-save every 60 seconds (1 minute)
          autoSaveInterval = setInterval(() => {
            if (unsavedChanges) {
              saveNotes();
              document.getElementById('autoSaveStatus').textContent = 'Auto-saved at ' + 
                new Date().toLocaleTimeString();
            }
          }, 60000);
          
          // Save on window close
          window.addEventListener('beforeunload', (event) => {
            if (unsavedChanges) {
              saveNotes();
              event.returnValue = 'You have unsaved changes.';
            }
          });
          
          // Monitor for text input (not just change events)
          document.querySelectorAll('textarea').forEach(textarea => {
            textarea.addEventListener('input', markUnsaved);
          });
          
          // Close overlay when clicking outside
          document.getElementById('confirmationOverlay').addEventListener('click', function(e) {
            if (e.target === this) {
              hideClearConfirmation();
            }
          });
          
          document.getElementById('reloadConfirmationOverlay').addEventListener('click', function(e) {
            if (e.target === this) {
              hideReloadConfirmation();
            }
          });
          
          document.getElementById('successOverlay').addEventListener('click', function(e) {
            if (e.target === this && !this.classList.contains('hiding')) {
              const overlay = document.getElementById('successOverlay');
              overlay.classList.add('hiding');
              setTimeout(() => {
                overlay.classList.remove('show', 'hiding');
              }, 300);
            }
          });
          
          document.getElementById('errorOverlay').addEventListener('click', function(e) {
            if (e.target === this) {
              hideErrorOverlay();
            }
          });
        </script>
      </body>
    </html>
  `).setTitle('Quick Notes').setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
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
 * Add option to the menu for managing quick notes
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Client Management')
    .addItem('Add New Client', 'addNewClient')
    .addSeparator()
    .addItem('Find Client', 'findClient')
    .addSeparator()
    .addSubMenu(ui.createMenu('Session Management')
      .addItem('Quick Notes', 'openQuickNotes')
      .addSeparator()
      .addItem('Send Session Recap', 'sendIndividualRecap')
      .addItem('Batch Prep Mode', 'openBatchPrepMode')
      .addItem('View Recap History', 'viewRecapHistory')
      .addSeparator()
      .addItem('Clear Current Student Notes', 'clearCurrentStudentNotes'))
    .addToUi();
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
            padding: 10px 20px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 500;
          }
          .btn-primary {
            background: #1a73e8;
            color: white;
          }
          .btn-primary:hover {
            background: #1557b0;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
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
          function prepareBatch() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:checked');
            const selectedClients = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedClients.length === 0) {
              alert('Please select at least one client.');
              return;
            }
            
            google.script.run
              .withSuccessHandler(() => {
                alert('Batch prep complete! Use "Send Session Recap" after each session.');
                google.script.host.close();
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
            color: #1a73e8;
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
            color: #1a73e8;
            border: none;
            border-radius: 4px;
            font-size: 12px;
            cursor: pointer;
            transition: all 0.2s;
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
            background: #1a73e8;
            color: white;
            flex: 1;
          }
          .btn-primary:hover {
            background: #1557b0;
          }
          .btn-secondary {
            background: #f8f9fa;
            color: #5f6368;
            border: 1px solid #dadce0;
          }
          .btn-secondary:hover {
            background: #f1f3f4;
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
            border-top: 3px solid #1a73e8;
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
            color: #2e7d32;
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
            transition: all 0.2s;
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
          ${clientData.dashboardLink ? `<br>Dashboard: <a href="${clientData.dashboardLink}" target="_blank" style="color: #1a73e8;">View Dashboard</a>` : ''}
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
            
            document.getElementById('loadingOverlay').classList.add('active');
            
            // Call generateEmailPreview which will process the data and show the preview
            google.script.run
              .withSuccessHandler(function(emailData) {
                document.getElementById('loadingOverlay').classList.remove('active');
                // Close this dialog - the server has already shown the preview
                google.script.host.close();
              })
              .withFailureHandler(function(error) {
                document.getElementById('loadingOverlay').classList.remove('active');
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
  formData.studentFirstName = formData.studentName.split(' ')[0];
  formData.tutorName = CONFIG.tutorName;
  
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
  sections.push(data.tutorName || CONFIG.tutorName);
  
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




/**
 * Fixed showEmailPreviewDialog function with Back to Edit button
 */
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
            color: #1a73e8;
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
            background: #34a853;
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
            background: #34a853;
            color: white;
          }
          .copy-btn.dashboard:hover {
            background: #2e7d32;
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
            color: #1a73e8;
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
            background: #1a73e8;
            color: white;
            flex: 1;
          }
          .btn-primary:hover {
            background: #1557b0;
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
            color: #5f6368;
          }
        </style>
      </head>
      <body>
        <div id="loadingDiv" class="loading">Loading email preview...</div>
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
            const element = document.getElementById(elementId);
            const text = element.textContent;
            
            const textArea = document.createElement('textarea');
            textArea.value = text;
            textArea.style.position = 'fixed';
            textArea.style.left = '-999999px';
            document.body.appendChild(textArea);
            textArea.select();
            
            try {
              document.execCommand('copy');
              const originalText = button.textContent;
              button.textContent = 'Copied!';
              button.classList.add('copied');
              setTimeout(function() {
                button.textContent = originalText;
                button.classList.remove('copied');
              }, 2000);
            } catch (err) {
              console.error('Copy failed:', err);
            }
            
            document.body.removeChild(textArea);
          }
          
          function copyRichText() {
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
              
              // Show success message
              showSuccessMessage('Copied with formatting! Paste into Google Docs.');
            } catch (err) {
              console.error('Copy failed:', err);
              // Fallback to plain text copy
              copyPlainTextFallback();
            }
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

/**
 * Enhanced get client data from sheet
 */
function getClientDataFromSheet(sheet) {
  const sheetName = sheet.getName();
  const clientNameWithSpaces = sheet.getRange('A1').getValue() || sheetName.replace(/([A-Z])/g, ' $1').trim();
  const clientTypes = sheet.getRange('C2').getValue() || '';
  const dashboardLink = sheet.getRange('H2').getValue() || '';  // Changed to H2
  const parentEmail = sheet.getRange('E2').getValue() || '';
  const parentName = sheet.getRange('F2').getValue() || '';
  
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
