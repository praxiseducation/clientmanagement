/**
 * Client Creation Backend Module - Enterprise Architecture
 * Features: Sheet creation, data storage, cache integration, error handling
 */

/**
 * Main function to show the new client dialog
 * Called from the enterprise sidebar
 */
function addNewClient() {
  showNewClientDialog();
}

/**
 * Show the new client dialog with modern UI
 */
function showNewClientDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          /* Loading skeleton styles */
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

          /* Base styles */
          body {
            font-family: 'Poppins', 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 30px;
            background: white;
            color: #202124;
          }
          .container {
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #202124;
            margin-bottom: 8px;
            font-weight: 600;
            text-align: center;
            font-size: 24px;
          }
          .subtitle {
            color: #5f6368;
            font-size: 14px;
            margin-bottom: 25px;
            text-align: center;
          }

          /* Form styles */
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
            border: 2px solid #dadce0;
            border-radius: 6px;
            font-size: 16px;
            font-family: inherit;
            transition: border-color 0.2s ease;
            box-sizing: border-box;
          }
          input[type="text"]:focus, input[type="email"]:focus, input[type="url"]:focus {
            outline: none;
            border-color: #003366;
            box-shadow: 0 0 0 3px rgba(0, 51, 102, 0.1);
          }
          .example {
            font-style: italic;
            color: #5f6368;
            font-size: 13px;
            margin-top: 8px;
          }
          .error-message {
            color: #d93025;
            font-size: 13px;
            margin-top: 4px;
            display: none;
          }

          /* Service selection */
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
            border: 2px solid #dadce0;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
            background: white;
            flex: 1;
            min-width: 120px;
          }
          .service-checkbox:hover {
            border-color: #5f6368;
            transform: translateY(-1px);
          }
          .service-checkbox.act:hover {
            background-color: rgba(0, 51, 102, 0.1);
            border-color: #003366;
          }
          .service-checkbox.math:hover {
            background-color: rgba(0, 200, 83, 0.1);
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
            background-color: rgba(0, 200, 83, 0.15);
            border-color: #00C853;
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

          /* Section divider */
          .section-divider {
            border-top: 1px solid #dadce0;
            margin: 30px 0;
          }

          /* Buttons */
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

          /* Button loading state */
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

          /* Footer link */
          .footer-link {
            text-align: center;
            margin-top: 20px;
            padding-top: 20px;
            border-top: 1px solid #dadce0;
          }
          .drive-link {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            color: #1a73e8;
            text-decoration: none;
            font-size: 13px;
            padding: 8px 16px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: #f8f9fa;
            transition: all 0.2s ease;
          }
          .drive-link:hover {
            background: #e8f0fe;
            border-color: #1a73e8;
          }

          /* Loading overlay */
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
                <div class="error-message" id="name-error">Please enter a valid student name</div>
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

            <div class="footer-link">
              <a href="https://drive.google.com/drive/u/0/folders/1UYJa--G9hE23a5WfxblqrtdnKdp3r91_"
                target="_blank"
                class="drive-link">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
                  <path d="M10 4H4c-1.11 0-2 .89-2 2v12c0 1.11.89 2 2 2h16c1.11 0 2-.89 2-2V8c0-1.11-.89-2-2-2h-8l-2-2z"/>
                </svg>
                Open Dashboards Drive Folder
              </a>
            </div>
          </div>
        </div>

        <!-- Loading overlay -->
        <div class="loading-overlay" id="loadingOverlay">
          <div class="loading-spinner"></div>
          <div class="loading-text">Adding client...</div>
        </div>

        <script>
          // Dialog JavaScript
          function toggleService(element, service) {
            const checkbox = element.querySelector('input[type="checkbox"]');
            checkbox.checked = !checkbox.checked;
            element.classList.toggle('selected', checkbox.checked);
            // Clear error when user selects a service
            if (checkbox.checked) {
              document.getElementById('service-error').style.display = 'none';
            }
          }

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
            let isValid = true;

            // Clear previous errors
            document.getElementById('name-error').style.display = 'none';
            document.getElementById('service-error').style.display = 'none';

            if (!clientName) {
              document.getElementById('name-error').style.display = 'block';
              isValid = false;
            }

            if (services.length === 0) {
              document.getElementById('service-error').style.display = 'block';
              isValid = false;
            }

            return isValid;
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

            // Show loading overlay
            document.getElementById('loadingOverlay').classList.add('active');
            document.querySelector('.loading-text').textContent = 'Adding ' + clientData.name + '...';

            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .createNewClient(clientData);
          }

          function onSuccess(message) {
            // Transition to success state
            const overlay = document.getElementById('loadingOverlay');
            const spinner = document.querySelector('.loading-spinner');
            const loadingText = document.querySelector('.loading-text');

            overlay.classList.add('success');
            spinner.classList.add('success');
            loadingText.classList.add('success');
            loadingText.textContent = 'Success!';

            // Close dialog after success animation
            setTimeout(function() {
              google.script.host.close();
            }, 1500);
          }

          function onFailure(error) {
            // Hide loading overlay on error
            document.getElementById('loadingOverlay').classList.remove('active');
            const button = document.querySelector('.btn-primary');
            setButtonLoading(button, false);
            alert('Error: ' + (error.message || error));
          }

          // Initialize form after loading
          setTimeout(function() {
            document.getElementById('formSkeleton').style.display = 'none';
            document.getElementById('actualForm').style.display = 'block';

            // Focus first input
            const firstInput = document.getElementById('clientName');
            if (firstInput) {
              firstInput.focus();
            }
          }, 300);

          // Allow Enter key to submit
          document.addEventListener('keypress', function(e) {
            if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA') {
              processClient();
            }
          });

          // Clear name error on input
          document.getElementById('clientName').addEventListener('input', function() {
            if (this.value.trim()) {
              document.getElementById('name-error').style.display = 'none';
            }
          });
        </script>
      </body>
    </html>
  `).setWidth(550).setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Smart College');
}

/**
 * Create a new client sheet with enterprise architecture
 * @param {Object} clientData - Client information from the dialog
 */
function createNewClient(clientData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Validate client data
    if (!clientData || !clientData.name || !clientData.services) {
      throw new Error('Missing required client information');
    }

    // Create concatenated sheet name (remove spaces)
    const sheetName = clientData.name.replace(/\s+/g, '');

    // Check if sheet already exists
    if (spreadsheet.getSheetByName(sheetName)) {
      throw new Error(`A client named "${clientData.name}" already exists.`);
    }

    // Find template sheet
    const templateSheet = spreadsheet.getSheetByName('NewClient');
    if (!templateSheet) {
      throw new Error('Template sheet "NewClient" not found. Please ensure it exists.');
    }

    // Create new sheet from template
    const newSheet = templateSheet.copyTo(spreadsheet);
    newSheet.setName(sheetName);

    // Populate sheet with client data
    populateClientSheet(newSheet, clientData);

    // Store client data in Script Properties
    storeClientData(sheetName, clientData);

    // Add to master client list
    addToMasterList(clientData, sheetName, spreadsheet);

    // Update enterprise caches
    updateEnterpriseCaches(clientData, sheetName);

    // Position sheet and activate
    positionAndActivateSheet(newSheet, spreadsheet);

    // Force sheet creation to complete
    SpreadsheetApp.flush();

    // Verify sheet creation
    if (!spreadsheet.getSheetByName(sheetName)) {
      throw new Error('Sheet creation failed - verification failed');
    }

    console.log(`Successfully created client: ${clientData.name} (${sheetName})`);
    return `Client "${clientData.name}" created successfully!`;

  } catch (error) {
    console.error('Error creating new client:', error);
    throw new Error(error.message || 'Failed to create client');
  }
}

/**
 * Populate the new client sheet with data
 */
function populateClientSheet(sheet, clientData) {
  try {
    // Basic client information
    sheet.getRange('A1').setValue(clientData.name);           // Student name
    sheet.getRange('C2').setValue(clientData.services);       // Services
    sheet.getRange('D2').setValue(clientData.dashboardLink);  // Dashboard link
    sheet.getRange('E2').setValue(clientData.parentEmail);    // Parent email
    sheet.getRange('F2').setValue(clientData.emailAddressees); // Email addressees

    // Add timestamp
    sheet.getRange('B1').setValue(new Date());

  } catch (error) {
    console.error('Error populating client sheet:', error);
    throw new Error('Failed to populate client sheet');
  }
}

/**
 * Store client data in Script Properties
 */
function storeClientData(sheetName, clientData) {
  try {
    const properties = PropertiesService.getScriptProperties();

    // Store individual properties for backward compatibility
    if (clientData.dashboardLink) {
      properties.setProperty(`dashboard_${sheetName}`, clientData.dashboardLink);
    }
    if (clientData.parentEmail) {
      properties.setProperty(`parentemail_${sheetName}`, clientData.parentEmail);
    }
    if (clientData.emailAddressees) {
      properties.setProperty(`emailaddressees_${sheetName}`, clientData.emailAddressees);
    }
    if (clientData.services) {
      properties.setProperty(`services_${sheetName}`, clientData.services);
    }

    // Set as active by default
    properties.setProperty(`isActive_${sheetName}`, 'true');

    // Store creation timestamp
    properties.setProperty(`created_${sheetName}`, new Date().toISOString());

  } catch (error) {
    console.error('Error storing client data:', error);
    throw new Error('Failed to store client data');
  }
}

/**
 * Add client to master list if it exists
 */
function addToMasterList(clientData, sheetName, spreadsheet) {
  try {
    const masterSheet = spreadsheet.getSheetByName('ClientList');
    if (!masterSheet) {
      console.log('No master ClientList sheet found, skipping master list update');
      return;
    }

    const newSheet = spreadsheet.getSheetByName(sheetName);
    if (!newSheet) {
      throw new Error('New client sheet not found for master list addition');
    }

    // Find next available row
    const values = masterSheet.getRange('A:A').getValues();
    let lastRow = 0;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] !== '') {
        lastRow = i + 1;
      }
    }

    const newRow = lastRow + 1;
    const sheetId = newSheet.getSheetId();

    // Create hyperlink formula
    const nameFormula = `=HYPERLINK("#gid=${sheetId}", "${clientData.name}")`;

    // Add to master list
    masterSheet.getRange(newRow, 1).setFormula(nameFormula);

    if (clientData.dashboardLink) {
      masterSheet.getRange(newRow, 2).setValue(clientData.dashboardLink);
    }

    console.log(`Added client to master list at row ${newRow}`);

  } catch (error) {
    console.error('Error adding to master list:', error);
    // Don't throw here - master list is optional
  }
}

/**
 * Update enterprise caches
 */
function updateEnterpriseCaches(clientData, sheetName) {
  try {
    // Refresh client list cache to include new client
    if (typeof ClientList !== 'undefined' && ClientList.onNewClient) {
      // This will be called by the frontend when the dialog closes
      console.log('Enterprise cache will be refreshed by frontend');
    }

    // Clear relevant caches to force refresh
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('clientListCache');
    properties.deleteProperty('ENHANCED_CLIENT_CACHE');

    console.log('Enterprise caches cleared for refresh');

  } catch (error) {
    console.error('Error updating enterprise caches:', error);
    // Don't throw - cache update is not critical for client creation
  }
}

/**
 * Position and activate the new sheet
 */
function positionAndActivateSheet(sheet, spreadsheet) {
  try {
    // Move to position 3 (after template sheets)
    const totalSheets = spreadsheet.getSheets().length;
    const targetPosition = Math.min(3, totalSheets);

    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(targetPosition);

  } catch (error) {
    console.error('Error positioning sheet:', error);
    // Don't throw - positioning is not critical
  }
}

/**
 * Helper function to check if a sheet name is valid for a client
 */
function isValidClientSheetName(sheetName) {
  const invalidNames = [
    'NewClient', 'New Client', 'NewACTClient', 'NewAcademicSupportClient',
    'Template', 'Settings', 'Dashboard', 'Summary', 'SessionRecaps',
    'Master', 'ClientList', 'Date Sent', 'Student Name', 'Client Type'
  ];

  return !invalidNames.includes(sheetName) &&
         !sheetName.includes('Template') &&
         !sheetName.startsWith('Sheet') &&
         sheetName.trim().length > 0;
}

/**
 * Get all existing client names for validation
 */
function getExistingClientNames() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();

    return sheets
      .map(sheet => sheet.getName())
      .filter(name => isValidClientSheetName(name));

  } catch (error) {
    console.error('Error getting existing client names:', error);
    return [];
  }
}