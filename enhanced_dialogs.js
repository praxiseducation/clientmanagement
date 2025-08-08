/**
 * Enhanced Dialogs Module - Second Beta Iteration
 * Features: Skeleton screens, improved UX, advanced functionality
 */

// ============================================================================
// SKELETON SCREEN COMPONENTS
// ============================================================================

const SkeletonScreens = {
  /**
   * Session Recap Dialog Skeleton
   */
  sessionRecap: () => `
    <div class="skeleton-container">
      <div class="skeleton-header">
        <div class="skeleton-text skeleton-title"></div>
        <div class="skeleton-text skeleton-subtitle"></div>
      </div>
      <div class="skeleton-body">
        <div class="skeleton-section">
          <div class="skeleton-text skeleton-label"></div>
          <div class="skeleton-textarea"></div>
        </div>
        <div class="skeleton-section">
          <div class="skeleton-text skeleton-label"></div>
          <div class="skeleton-textarea"></div>
        </div>
        <div class="skeleton-section">
          <div class="skeleton-text skeleton-label"></div>
          <div class="skeleton-textarea"></div>
        </div>
      </div>
      <div class="skeleton-actions">
        <div class="skeleton-button"></div>
        <div class="skeleton-button skeleton-button-primary"></div>
      </div>
    </div>
  `,

  /**
   * Bulk Client Dialog Skeleton
   */
  bulkClients: () => `
    <div class="skeleton-container">
      <div class="skeleton-header">
        <div class="skeleton-text skeleton-title"></div>
        <div class="skeleton-text skeleton-subtitle"></div>
      </div>
      <div class="skeleton-toolbar">
        <div class="skeleton-button"></div>
        <div class="skeleton-button"></div>
        <div class="skeleton-search"></div>
      </div>
      <div class="skeleton-table">
        ${Array(5).fill().map(() => `
          <div class="skeleton-row">
            <div class="skeleton-cell"></div>
            <div class="skeleton-cell"></div>
            <div class="skeleton-cell"></div>
            <div class="skeleton-cell"></div>
          </div>
        `).join('')}
      </div>
    </div>
  `,

  /**
   * Skeleton Screen CSS
   */
  css: () => `
    .skeleton-container {
      padding: 20px;
      background: white;
      border-radius: 8px;
    }
    
    .skeleton-text, .skeleton-button, .skeleton-textarea, .skeleton-search, .skeleton-cell {
      background: linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%);
      background-size: 200% 100%;
      animation: skeleton-loading 1.5s infinite;
      border-radius: 4px;
    }
    
    @keyframes skeleton-loading {
      0% { background-position: 200% 0; }
      100% { background-position: -200% 0; }
    }
    
    .skeleton-header {
      margin-bottom: 24px;
      text-align: center;
    }
    
    .skeleton-title {
      height: 28px;
      width: 60%;
      margin: 0 auto 8px;
    }
    
    .skeleton-subtitle {
      height: 16px;
      width: 40%;
      margin: 0 auto;
    }
    
    .skeleton-toolbar {
      display: flex;
      gap: 12px;
      margin-bottom: 20px;
      align-items: center;
    }
    
    .skeleton-button {
      height: 36px;
      width: 100px;
    }
    
    .skeleton-button-primary {
      width: 120px;
    }
    
    .skeleton-search {
      height: 36px;
      flex: 1;
    }
    
    .skeleton-section {
      margin-bottom: 20px;
    }
    
    .skeleton-label {
      height: 16px;
      width: 100px;
      margin-bottom: 8px;
    }
    
    .skeleton-textarea {
      height: 80px;
      width: 100%;
    }
    
    .skeleton-table {
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      overflow: hidden;
    }
    
    .skeleton-row {
      display: flex;
      gap: 1px;
      background: #e0e0e0;
      margin-bottom: 1px;
    }
    
    .skeleton-row:last-child {
      margin-bottom: 0;
    }
    
    .skeleton-cell {
      height: 40px;
      flex: 1;
      background: #f5f5f5;
    }
    
    .skeleton-actions {
      display: flex;
      justify-content: flex-end;
      gap: 12px;
      margin-top: 24px;
      padding-top: 20px;
      border-top: 1px solid #e0e0e0;
    }
    
    .fade-in {
      animation: fadeIn 0.3s ease-in;
    }
    
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }
  `
};

// ============================================================================
// ENHANCED BULK CLIENT DIALOG
// ============================================================================

/**
 * Modern Bulk Client Add/Edit Dialog with improved UX
 */
function showEnhancedBulkClientDialog() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" rel="stylesheet">
        <style>
          ${SkeletonScreens.css()}
          
          :root {
            --primary-color: #1976d2;
            --primary-light: #e3f2fd;
            --success-color: #4caf50;
            --warning-color: #ff9800;
            --error-color: #f44336;
            --text-primary: #212121;
            --text-secondary: #757575;
            --border-color: #e0e0e0;
            --background: #fafafa;
          }
          
          * {
            box-sizing: border-box;
          }
          
          body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            margin: 0;
            padding: 0;
            background: var(--background);
            color: var(--text-primary);
            line-height: 1.5;
          }
          
          .main-container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            max-height: 800px;
          }
          
          /* Header */
          .header {
            background: white;
            border-bottom: 1px solid var(--border-color);
            padding: 20px 24px;
            text-align: center;
          }
          
          .header h1 {
            margin: 0 0 4px;
            font-size: 24px;
            font-weight: 600;
            color: var(--text-primary);
          }
          
          .header .subtitle {
            margin: 0;
            font-size: 14px;
            color: var(--text-secondary);
          }
          
          /* Mode Toggle */
          .mode-toggle {
            display: flex;
            background: var(--primary-light);
            border-radius: 8px;
            padding: 4px;
            margin: 16px auto 0;
            width: fit-content;
          }
          
          .mode-btn {
            padding: 8px 16px;
            border: none;
            background: transparent;
            border-radius: 6px;
            cursor: pointer;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            transition: all 0.2s ease;
            color: var(--text-secondary);
          }
          
          .mode-btn.active {
            background: white;
            color: var(--primary-color);
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          /* Toolbar */
          .toolbar {
            background: white;
            border-bottom: 1px solid var(--border-color);
            padding: 16px 24px;
            display: flex;
            gap: 12px;
            align-items: center;
            flex-wrap: wrap;
          }
          
          .search-container {
            flex: 1;
            min-width: 250px;
            position: relative;
          }
          
          .search-input {
            width: 100%;
            padding: 8px 12px 8px 40px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            font-family: inherit;
            font-size: 14px;
            transition: all 0.2s ease;
          }
          
          .search-input:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 2px var(--primary-light);
          }
          
          .search-icon {
            position: absolute;
            left: 12px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-secondary);
            font-size: 20px;
          }
          
          .btn {
            padding: 8px 16px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            background: white;
            cursor: pointer;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
            text-decoration: none;
            white-space: nowrap;
          }
          
          .btn:hover {
            background: var(--background);
            transform: translateY(-1px);
          }
          
          .btn-primary {
            background: var(--primary-color);
            border-color: var(--primary-color);
            color: white;
          }
          
          .btn-primary:hover {
            background: #1565c0;
            border-color: #1565c0;
          }
          
          .btn-success {
            background: var(--success-color);
            border-color: var(--success-color);
            color: white;
          }
          
          /* Main Content */
          .content {
            flex: 1;
            overflow: hidden;
            display: flex;
            flex-direction: column;
          }
          
          /* Wizard Steps */
          .wizard-container {
            display: none;
            flex-direction: column;
            height: 100%;
          }
          
          .wizard-container.active {
            display: flex;
          }
          
          .step-indicator {
            background: white;
            border-bottom: 1px solid var(--border-color);
            padding: 16px 24px;
          }
          
          .steps {
            display: flex;
            justify-content: center;
            gap: 40px;
          }
          
          .step {
            display: flex;
            align-items: center;
            gap: 8px;
            color: var(--text-secondary);
            font-size: 14px;
          }
          
          .step.active {
            color: var(--primary-color);
          }
          
          .step.completed {
            color: var(--success-color);
          }
          
          .step-number {
            width: 24px;
            height: 24px;
            border-radius: 50%;
            border: 2px solid currentColor;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
            font-weight: 600;
          }
          
          .step.active .step-number,
          .step.completed .step-number {
            background: currentColor;
            color: white;
          }
          
          /* Step Content */
          .step-content {
            flex: 1;
            overflow: auto;
            padding: 24px;
          }
          
          /* Table View */
          .table-view {
            background: white;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          
          .table-container {
            overflow: auto;
            max-height: 500px;
          }
          
          table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
          }
          
          th, td {
            padding: 12px 16px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
          }
          
          th {
            background: var(--background);
            font-weight: 600;
            color: var(--text-secondary);
            position: sticky;
            top: 0;
          }
          
          tr:hover {
            background: var(--primary-light);
          }
          
          /* Form Styles */
          .form-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 24px;
          }
          
          .form-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
          }
          
          .form-label {
            font-size: 14px;
            font-weight: 500;
            color: var(--text-primary);
          }
          
          .form-input, .form-select, .form-textarea {
            padding: 8px 12px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            font-family: inherit;
            font-size: 14px;
            transition: all 0.2s ease;
          }
          
          .form-input:focus, .form-select:focus, .form-textarea:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 2px var(--primary-light);
          }
          
          .form-textarea {
            min-height: 80px;
            resize: vertical;
          }
          
          /* Actions */
          .actions {
            background: white;
            border-top: 1px solid var(--border-color);
            padding: 16px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .actions-left {
            display: flex;
            gap: 8px;
          }
          
          .actions-right {
            display: flex;
            gap: 12px;
          }
          
          /* Loading States */
          .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(255, 255, 255, 0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
          }
          
          .loading-spinner {
            width: 40px;
            height: 40px;
            border: 4px solid var(--border-color);
            border-left-color: var(--primary-color);
            border-radius: 50%;
            animation: spin 1s linear infinite;
          }
          
          @keyframes spin {
            to { transform: rotate(360deg); }
          }
          
          /* Success/Error States */
          .notification {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 12px 20px;
            border-radius: 6px;
            color: white;
            font-weight: 500;
            transform: translateX(100%);
            transition: transform 0.3s ease;
            z-index: 1001;
          }
          
          .notification.show {
            transform: translateX(0);
          }
          
          .notification.success {
            background: var(--success-color);
          }
          
          .notification.error {
            background: var(--error-color);
          }
          
          /* Responsive */
          @media (max-width: 768px) {
            .main-container {
              height: 100vh;
            }
            
            .header, .toolbar, .actions {
              padding: 16px;
            }
            
            .steps {
              gap: 20px;
            }
            
            .step-content {
              padding: 16px;
            }
            
            .form-grid {
              grid-template-columns: 1fr;
            }
          }
        </style>
      </head>
      <body>
        <!-- Loading Screen -->
        <div id="loadingScreen" class="loading-overlay">
          ${SkeletonScreens.bulkClients()}
        </div>
        
        <!-- Main Interface -->
        <div id="mainInterface" class="main-container" style="display: none;">
          <!-- Header -->
          <div class="header">
            <h1>Client Management</h1>
            <p class="subtitle">Add new clients or edit existing ones</p>
            
            <div class="mode-toggle">
              <button class="mode-btn active" data-mode="add">
                <span class="material-symbols-outlined">person_add</span>
                Add Clients
              </button>
              <button class="mode-btn" data-mode="edit">
                <span class="material-symbols-outlined">edit</span>
                Edit Clients
              </button>
            </div>
          </div>
          
          <!-- Add Mode -->
          <div id="addMode" class="wizard-container active">
            <!-- Step Indicator -->
            <div class="step-indicator">
              <div class="steps">
                <div class="step active" data-step="1">
                  <div class="step-number">1</div>
                  <span>Client Info</span>
                </div>
                <div class="step" data-step="2">
                  <div class="step-number">2</div>
                  <span>Services</span>
                </div>
                <div class="step" data-step="3">
                  <div class="step-number">3</div>
                  <span>Review</span>
                </div>
              </div>
            </div>
            
            <!-- Step 1: Basic Info -->
            <div class="step-content" id="step1" style="display: block;">
              <div class="form-grid">
                <div class="form-group">
                  <label class="form-label">Student Name *</label>
                  <input type="text" class="form-input" id="studentName" placeholder="Enter student name" required>
                </div>
                <div class="form-group">
                  <label class="form-label">Parent Email *</label>
                  <input type="email" class="form-input" id="parentEmail" placeholder="parent@email.com" required>
                </div>
                <div class="form-group">
                  <label class="form-label">Additional Email Recipients</label>
                  <input type="text" class="form-input" id="emailAddressees" placeholder="email1@domain.com, email2@domain.com">
                </div>
              </div>
            </div>
            
            <!-- Step 2: Services -->
            <div class="step-content" id="step2" style="display: none;">
              <div class="form-grid">
                <div class="form-group">
                  <label class="form-label">Service Type *</label>
                  <select class="form-select" id="serviceType" required>
                    <option value="">Select service type</option>
                    <option value="ACT Prep">ACT Prep</option>
                    <option value="SAT Prep">SAT Prep</option>
                    <option value="Academic Support">Academic Support</option>
                    <option value="College Counseling">College Counseling</option>
                    <option value="Test Prep + Academic">Test Prep + Academic</option>
                  </select>
                </div>
                <div class="form-group">
                  <label class="form-label">Dashboard Link</label>
                  <input type="url" class="form-input" id="dashboardLink" placeholder="https://dashboard.example.com">
                </div>
                <div class="form-group">
                  <label class="form-label">Meeting Notes Link</label>
                  <input type="url" class="form-input" id="meetingNotesLink" placeholder="https://notes.example.com">
                </div>
                <div class="form-group">
                  <label class="form-label">Score Report Link</label>
                  <input type="url" class="form-input" id="scoreReportLink" placeholder="https://scores.example.com">
                </div>
              </div>
            </div>
            
            <!-- Step 3: Review -->
            <div class="step-content" id="step3" style="display: none;">
              <div id="reviewContent" class="table-view">
                <div class="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th>Student Name</th>
                        <th>Parent Email</th>
                        <th>Service Type</th>
                        <th>Links</th>
                      </tr>
                    </thead>
                    <tbody id="reviewTableBody">
                      <!-- Review content will be inserted here -->
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          
          <!-- Edit Mode -->
          <div id="editMode" class="wizard-container">
            <!-- Toolbar -->
            <div class="toolbar">
              <div class="search-container">
                <span class="material-symbols-outlined search-icon">search</span>
                <input type="text" class="search-input" id="clientSearch" placeholder="Search clients by name, email, or service type...">
              </div>
              <button class="btn" id="selectAllBtn">
                <span class="material-symbols-outlined">select_all</span>
                Select All
              </button>
              <button class="btn" id="bulkEditBtn" disabled>
                <span class="material-symbols-outlined">edit</span>
                Edit Selected
              </button>
            </div>
            
            <!-- Client List -->
            <div class="step-content">
              <div class="table-view">
                <div class="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th width="40">
                          <input type="checkbox" id="selectAllCheckbox">
                        </th>
                        <th>Student Name</th>
                        <th>Parent Email</th>
                        <th>Service Type</th>
                        <th>Status</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody id="clientsTableBody">
                      <!-- Client rows will be inserted here -->
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          
          <!-- Actions -->
          <div class="actions">
            <div class="actions-left">
              <span id="statusText" class="text-secondary"></span>
            </div>
            <div class="actions-right">
              <button class="btn" id="cancelBtn">Cancel</button>
              <button class="btn" id="prevBtn" style="display: none;">Previous</button>
              <button class="btn btn-primary" id="nextBtn">Next</button>
              <button class="btn btn-success" id="saveBtn" style="display: none;">Save Clients</button>
            </div>
          </div>
        </div>
        
        <!-- Notification -->
        <div id="notification" class="notification"></div>
        
        <script>
          // Application state
          let currentMode = 'add';
          let currentStep = 1;
          let clients = [];
          let allClients = [];
          let selectedClients = new Set();
          
          // Initialize application
          window.addEventListener('load', async () => {
            await initializeApp();
          });
          
          async function initializeApp() {
            try {
              // Load existing clients for edit mode
              await loadExistingClients();
              
              // Hide loading screen and show main interface
              document.getElementById('loadingScreen').style.display = 'none';
              document.getElementById('mainInterface').style.display = 'flex';
              document.getElementById('mainInterface').classList.add('fade-in');
              
              // Setup event listeners
              setupEventListeners();
              
              // Initialize fuzzy search
              setupFuzzySearch();
              
            } catch (error) {
              showNotification('Failed to initialize application: ' + error.message, 'error');
            }
          }
          
          async function loadExistingClients() {
            try {
              const response = await new Promise((resolve, reject) => {
                google.script.run
                  .withSuccessHandler(resolve)
                  .withFailureHandler(reject)
                  .getUnifiedClientList();
              });
              
              allClients = response || [];
              renderClientTable();
              
            } catch (error) {
              console.error('Error loading clients:', error);
              throw error;
            }
          }
          
          function setupEventListeners() {
            // Mode toggle
            document.querySelectorAll('.mode-btn').forEach(btn => {
              btn.addEventListener('click', () => switchMode(btn.dataset.mode));
            });
            
            // Navigation buttons
            document.getElementById('nextBtn').addEventListener('click', handleNext);
            document.getElementById('prevBtn').addEventListener('click', handlePrevious);
            document.getElementById('saveBtn').addEventListener('click', handleSave);
            document.getElementById('cancelBtn').addEventListener('click', handleCancel);
            
            // Form inputs
            document.getElementById('studentName').addEventListener('input', validateStep1);
            document.getElementById('parentEmail').addEventListener('input', validateStep1);
            document.getElementById('serviceType').addEventListener('change', validateStep2);
            
            // Search functionality
            document.getElementById('clientSearch').addEventListener('input', handleSearch);
            
            // Bulk actions
            document.getElementById('selectAllBtn').addEventListener('click', toggleSelectAll);
            document.getElementById('selectAllCheckbox').addEventListener('change', toggleSelectAll);
            document.getElementById('bulkEditBtn').addEventListener('click', handleBulkEdit);
          }
          
          function switchMode(mode) {
            currentMode = mode;
            
            // Update mode buttons
            document.querySelectorAll('.mode-btn').forEach(btn => {
              btn.classList.toggle('active', btn.dataset.mode === mode);
            });
            
            // Show/hide mode containers
            document.getElementById('addMode').classList.toggle('active', mode === 'add');
            document.getElementById('editMode').classList.toggle('active', mode === 'edit');
            
            // Update actions
            updateActions();
          }
          
          function handleNext() {
            if (currentMode === 'add') {
              if (currentStep < 3) {
                currentStep++;
                showStep(currentStep);
              }
            }
          }
          
          function handlePrevious() {
            if (currentStep > 1) {
              currentStep--;
              showStep(currentStep);
            }
          }
          
          function showStep(step) {
            // Hide all steps
            document.querySelectorAll('.step-content').forEach(el => {
              el.style.display = 'none';
            });
            
            // Show current step
            document.getElementById(\`step\${step}\`).style.display = 'block';
            
            // Update step indicator
            document.querySelectorAll('.step').forEach((el, index) => {
              el.classList.toggle('active', index + 1 === step);
              el.classList.toggle('completed', index + 1 < step);
            });
            
            // Update review content if on step 3
            if (step === 3) {
              updateReviewContent();
            }
            
            updateActions();
          }
          
          function updateReviewContent() {
            const reviewBody = document.getElementById('reviewTableBody');
            const formData = getFormData();
            
            reviewBody.innerHTML = \`
              <tr>
                <td>\${formData.studentName}</td>
                <td>\${formData.parentEmail}</td>
                <td>\${formData.serviceType}</td>
                <td>
                  <div style="display: flex; gap: 8px;">
                    \${formData.dashboardLink ? '<span class="badge">Dashboard</span>' : ''}
                    \${formData.meetingNotesLink ? '<span class="badge">Notes</span>' : ''}
                    \${formData.scoreReportLink ? '<span class="badge">Scores</span>' : ''}
                  </div>
                </td>
              </tr>
            \`;
          }
          
          function getFormData() {
            return {
              studentName: document.getElementById('studentName').value,
              parentEmail: document.getElementById('parentEmail').value,
              emailAddressees: document.getElementById('emailAddressees').value,
              serviceType: document.getElementById('serviceType').value,
              dashboardLink: document.getElementById('dashboardLink').value,
              meetingNotesLink: document.getElementById('meetingNotesLink').value,
              scoreReportLink: document.getElementById('scoreReportLink').value
            };
          }
          
          function validateStep1() {
            const name = document.getElementById('studentName').value.trim();
            const email = document.getElementById('parentEmail').value.trim();
            
            const isValid = name && email && isValidEmail(email);
            updateStepValidation(1, isValid);
            return isValid;
          }
          
          function validateStep2() {
            const serviceType = document.getElementById('serviceType').value;
            const isValid = !!serviceType;
            updateStepValidation(2, isValid);
            return isValid;
          }
          
          function updateStepValidation(step, isValid) {
            const stepEl = document.querySelector(\`.step[data-step="\${step}"]\`);
            stepEl.classList.toggle('completed', isValid);
          }
          
          function updateActions() {
            const prevBtn = document.getElementById('prevBtn');
            const nextBtn = document.getElementById('nextBtn');
            const saveBtn = document.getElementById('saveBtn');
            const statusText = document.getElementById('statusText');
            
            if (currentMode === 'add') {
              prevBtn.style.display = currentStep > 1 ? 'inline-flex' : 'none';
              nextBtn.style.display = currentStep < 3 ? 'inline-flex' : 'none';
              saveBtn.style.display = currentStep === 3 ? 'inline-flex' : 'none';
              
              statusText.textContent = \`Step \${currentStep} of 3\`;
              
              // Disable next button if current step is invalid
              if (currentStep === 1) {
                nextBtn.disabled = !validateStep1();
              } else if (currentStep === 2) {
                nextBtn.disabled = !validateStep2();
              } else {
                nextBtn.disabled = false;
              }
            } else {
              prevBtn.style.display = 'none';
              nextBtn.style.display = 'none';
              saveBtn.style.display = selectedClients.size > 0 ? 'inline-flex' : 'none';
              
              statusText.textContent = \`\${selectedClients.size} client(s) selected\`;
            }
          }
          
          async function handleSave() {
            if (currentMode === 'add') {
              await saveNewClient();
            } else {
              await saveBulkEdits();
            }
          }
          
          async function saveNewClient() {
            const formData = getFormData();
            
            if (!validateStep1() || !validateStep2()) {
              showNotification('Please fill in all required fields', 'error');
              return;
            }
            
            try {
              showLoading('Creating client...');
              
              const response = await new Promise((resolve, reject) => {
                google.script.run
                  .withSuccessHandler(resolve)
                  .withFailureHandler(reject)
                  .createClientSheet(formData);
              });
              
              hideLoading();
              showNotification('Client created successfully!', 'success');
              
              // Reset form
              resetForm();
              
            } catch (error) {
              hideLoading();
              showNotification('Error creating client: ' + error.message, 'error');
            }
          }
          
          function resetForm() {
            document.getElementById('studentName').value = '';
            document.getElementById('parentEmail').value = '';
            document.getElementById('emailAddressees').value = '';
            document.getElementById('serviceType').value = '';
            document.getElementById('dashboardLink').value = '';
            document.getElementById('meetingNotesLink').value = '';
            document.getElementById('scoreReportLink').value = '';
            
            currentStep = 1;
            showStep(1);
          }
          
          function renderClientTable() {
            const tbody = document.getElementById('clientsTableBody');
            
            if (allClients.length === 0) {
              tbody.innerHTML = \`
                <tr>
                  <td colspan="6" style="text-align: center; padding: 40px; color: var(--text-secondary);">
                    No clients found. Switch to Add mode to create your first client.
                  </td>
                </tr>
              \`;
              return;
            }
            
            tbody.innerHTML = allClients.map(client => \`
              <tr>
                <td>
                  <input type="checkbox" class="client-checkbox" value="\${client.name}" 
                         onchange="handleClientSelect(this)">
                </td>
                <td>\${client.name || 'N/A'}</td>
                <td>\${client.parentEmail || 'N/A'}</td>
                <td>\${client.services || 'N/A'}</td>
                <td>
                  <span class="badge \${client.isActive ? 'success' : 'warning'}">
                    \${client.isActive ? 'Active' : 'Inactive'}
                  </span>
                </td>
                <td>
                  <button class="btn btn-sm" onclick="editClient('\${client.name}')">
                    <span class="material-symbols-outlined">edit</span>
                  </button>
                </td>
              </tr>
            \`).join('');
          }
          
          function handleClientSelect(checkbox) {
            const clientName = checkbox.value;
            
            if (checkbox.checked) {
              selectedClients.add(clientName);
            } else {
              selectedClients.delete(clientName);
            }
            
            updateBulkActions();
          }
          
          function updateBulkActions() {
            const bulkEditBtn = document.getElementById('bulkEditBtn');
            const selectAllCheckbox = document.getElementById('selectAllCheckbox');
            
            bulkEditBtn.disabled = selectedClients.size === 0;
            
            // Update select all checkbox state
            const allCheckboxes = document.querySelectorAll('.client-checkbox');
            const checkedCount = document.querySelectorAll('.client-checkbox:checked').length;
            
            selectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < allCheckboxes.length;
            selectAllCheckbox.checked = checkedCount === allCheckboxes.length && allCheckboxes.length > 0;
            
            updateActions();
          }
          
          function toggleSelectAll() {
            const selectAllCheckbox = document.getElementById('selectAllCheckbox');
            const clientCheckboxes = document.querySelectorAll('.client-checkbox');
            
            clientCheckboxes.forEach(checkbox => {
              checkbox.checked = selectAllCheckbox.checked;
              handleClientSelect(checkbox);
            });
          }
          
          function handleSearch() {
            // Implement fuzzy search functionality
            // This will be connected to the advanced search system
          }
          
          function setupFuzzySearch() {
            // Initialize fuzzy search with Fuse.js equivalent
            // Implementation will be added in the advanced search feature
          }
          
          function handleBulkEdit() {
            // Open bulk edit dialog
            showBulkEditDialog(Array.from(selectedClients));
          }
          
          function handleCancel() {
            google.script.host.close();
          }
          
          // Utility functions
          function isValidEmail(email) {
            return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
          }
          
          function showLoading(message = 'Loading...') {
            // Show loading overlay with message
          }
          
          function hideLoading() {
            // Hide loading overlay
          }
          
          function showNotification(message, type = 'success') {
            const notification = document.getElementById('notification');
            notification.textContent = message;
            notification.className = \`notification \${type} show\`;
            
            setTimeout(() => {
              notification.classList.remove('show');
            }, 3000);
          }
          
          // Additional utility styles
          const additionalStyles = \`
            .badge {
              padding: 4px 8px;
              border-radius: 12px;
              font-size: 12px;
              font-weight: 500;
              text-transform: uppercase;
            }
            
            .badge.success {
              background: var(--success-color);
              color: white;
            }
            
            .badge.warning {
              background: var(--warning-color);
              color: white;
            }
            
            .btn-sm {
              padding: 4px 8px;
              font-size: 12px;
            }
          \`;
          
          // Inject additional styles
          const styleSheet = document.createElement('style');
          styleSheet.textContent = additionalStyles;
          document.head.appendChild(styleSheet);
        </script>
      </body>
    </html>
  `).setWidth(1200).setHeight(800);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Client Management');
}

// ============================================================================
// ENHANCED SESSION RECAP DIALOG
// ============================================================================

/**
 * Enhanced Session Recap Dialog with skeleton loading and improved UX
 */
function showEnhancedSessionRecapDialog(clientData) {
  // Store client data for the dialog
  UnifiedClientDataStore.setTempData('CURRENT_RECAP_DATA', clientData);
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
        <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" rel="stylesheet">
        <style>
          ${SkeletonScreens.css()}
          
          :root {
            --primary-color: #1976d2;
            --primary-light: #e3f2fd;
            --success-color: #4caf50;
            --warning-color: #ff9800;
            --error-color: #f44336;
            --text-primary: #212121;
            --text-secondary: #757575;
            --border-color: #e0e0e0;
            --background: #fafafa;
          }
          
          * {
            box-sizing: border-box;
          }
          
          body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            margin: 0;
            padding: 0;
            background: white;
            color: var(--text-primary);
            line-height: 1.5;
          }
          
          .main-container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            max-height: 800px;
          }
          
          /* Header */
          .header {
            background: linear-gradient(135deg, var(--primary-color), #1565c0);
            color: white;
            padding: 20px 24px;
            text-align: center;
          }
          
          .header h1 {
            margin: 0 0 4px;
            font-size: 24px;
            font-weight: 600;
          }
          
          .header .subtitle {
            margin: 0;
            font-size: 14px;
            opacity: 0.9;
          }
          
          /* Content */
          .content {
            flex: 1;
            overflow: auto;
            padding: 24px;
            background: var(--background);
          }
          
          .form-section {
            background: white;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            transition: all 0.2s ease;
          }
          
          .form-section:hover {
            box-shadow: 0 4px 16px rgba(0,0,0,0.12);
          }
          
          .section-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 16px;
            padding-bottom: 12px;
            border-bottom: 1px solid var(--border-color);
          }
          
          .section-icon {
            color: var(--primary-color);
            font-size: 24px;
          }
          
          .section-title {
            font-size: 18px;
            font-weight: 600;
            margin: 0;
          }
          
          .section-description {
            color: var(--text-secondary);
            font-size: 14px;
            margin: 0;
            margin-left: auto;
          }
          
          /* Quick Pills */
          .quick-pills {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-bottom: 16px;
          }
          
          .quick-pill {
            padding: 6px 12px;
            background: var(--primary-light);
            color: var(--primary-color);
            border: none;
            border-radius: 20px;
            cursor: pointer;
            font-size: 12px;
            font-weight: 500;
            transition: all 0.2s ease;
            user-select: none;
          }
          
          .quick-pill:hover {
            background: var(--primary-color);
            color: white;
            transform: translateY(-1px);
          }
          
          .quick-pill.active {
            background: var(--success-color);
            color: white;
          }
          
          /* Form Elements */
          .form-textarea {
            width: 100%;
            min-height: 100px;
            padding: 12px;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            font-family: inherit;
            font-size: 14px;
            resize: vertical;
            transition: all 0.2s ease;
          }
          
          .form-textarea:focus {
            outline: none;
            border-color: var(--primary-color);
            box-shadow: 0 0 0 2px var(--primary-light);
          }
          
          .form-textarea.expanded {
            min-height: 150px;
          }
          
          /* Character Counter */
          .character-counter {
            text-align: right;
            font-size: 12px;
            color: var(--text-secondary);
            margin-top: 4px;
          }
          
          .character-counter.warning {
            color: var(--warning-color);
          }
          
          .character-counter.error {
            color: var(--error-color);
          }
          
          /* Auto-save Status */
          .auto-save-status {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 12px;
            color: var(--text-secondary);
            margin-top: 8px;
          }
          
          .auto-save-status.saving {
            color: var(--warning-color);
          }
          
          .auto-save-status.saved {
            color: var(--success-color);
          }
          
          .auto-save-status.error {
            color: var(--error-color);
          }
          
          /* Actions */
          .actions {
            background: white;
            border-top: 1px solid var(--border-color);
            padding: 20px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
          }
          
          .actions-left {
            display: flex;
            align-items: center;
            gap: 16px;
          }
          
          .actions-right {
            display: flex;
            gap: 12px;
          }
          
          .btn {
            padding: 10px 20px;
            border: 1px solid var(--border-color);
            border-radius: 8px;
            background: white;
            cursor: pointer;
            font-family: inherit;
            font-size: 14px;
            font-weight: 500;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
            text-decoration: none;
          }
          
          .btn:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          }
          
          .btn-primary {
            background: var(--primary-color);
            border-color: var(--primary-color);
            color: white;
          }
          
          .btn-success {
            background: var(--success-color);
            border-color: var(--success-color);
            color: white;
          }
          
          .btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
          }
          
          /* Progress Indicator */
          .progress-indicator {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 14px;
            color: var(--text-secondary);
          }
          
          .progress-bar {
            width: 100px;
            height: 4px;
            background: var(--border-color);
            border-radius: 2px;
            overflow: hidden;
          }
          
          .progress-fill {
            height: 100%;
            background: var(--success-color);
            transition: width 0.3s ease;
          }
          
          /* Offline Indicator */
          .offline-indicator {
            background: var(--warning-color);
            color: white;
            padding: 8px 16px;
            text-align: center;
            font-size: 14px;
            font-weight: 500;
            display: none;
          }
          
          .offline-indicator.show {
            display: block;
          }
          
          /* Animations */
          .fade-in {
            animation: fadeIn 0.3s ease-in;
          }
          
          @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
          }
          
          .pulse {
            animation: pulse 1s ease-in-out infinite;
          }
          
          @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
          }
          
          /* Responsive */
          @media (max-width: 768px) {
            .header, .content, .actions {
              padding: 16px;
            }
            
            .form-section {
              margin-bottom: 16px;
              padding: 20px;
            }
            
            .actions {
              flex-direction: column;
              gap: 16px;
            }
            
            .actions-left, .actions-right {
              width: 100%;
              justify-content: center;
            }
          }
        </style>
      </head>
      <body>
        <!-- Loading Screen -->
        <div id="loadingScreen" class="loading-overlay">
          ${SkeletonScreens.sessionRecap()}
        </div>
        
        <!-- Offline Indicator -->
        <div id="offlineIndicator" class="offline-indicator">
          <span class="material-symbols-outlined">wifi_off</span>
          You're offline. Notes are being saved locally and will sync when you reconnect.
        </div>
        
        <!-- Main Interface -->
        <div id="mainInterface" class="main-container" style="display: none;">
          <!-- Header -->
          <div class="header">
            <h1 id="clientName">Session Recap</h1>
            <p class="subtitle" id="sessionDate">Loading session details...</p>
          </div>
          
          <!-- Content -->
          <div class="content">
            <!-- Wins Section -->
            <div class="form-section">
              <div class="section-header">
                <span class="material-symbols-outlined section-icon">celebration</span>
                <h3 class="section-title">Wins & Achievements</h3>
                <p class="section-description">What went well today?</p>
              </div>
              
              <div class="quick-pills" id="winsPills">
                <!-- Quick pills will be populated by user preferences -->
              </div>
              
              <textarea 
                class="form-textarea" 
                id="wins" 
                placeholder="Share the positive moments and breakthroughs from today's session..."
                data-section="wins"></textarea>
              
              <div class="character-counter" id="winsCounter">0 / 500</div>
              <div class="auto-save-status" id="winsAutoSave">
                <span class="material-symbols-outlined">save</span>
                <span>Auto-saved</span>
              </div>
            </div>
            
            <!-- Skills Section -->
            <div class="form-section">
              <div class="section-header">
                <span class="material-symbols-outlined section-icon">psychology</span>
                <h3 class="section-title">Skills Development</h3>
                <p class="section-description">Track learning progress</p>
              </div>
              
              <!-- Skills Tabs -->
              <div class="skills-tabs">
                <button class="skill-tab active" data-skill="mastered">
                  <span class="material-symbols-outlined">check_circle</span>
                  Mastered
                </button>
                <button class="skill-tab" data-skill="practiced">
                  <span class="material-symbols-outlined">repeat</span>
                  Practiced
                </button>
                <button class="skill-tab" data-skill="introduced">
                  <span class="material-symbols-outlined">lightbulb</span>
                  New Topics
                </button>
              </div>
              
              <div class="skills-content">
                <div class="skill-panel active" id="masteredPanel">
                  <div class="quick-pills" id="masteredPills"></div>
                  <textarea 
                    class="form-textarea" 
                    id="skillsMastered" 
                    placeholder="What skills has the student fully mastered today?"
                    data-section="mastered"></textarea>
                  <div class="character-counter" id="masteredCounter">0 / 300</div>
                </div>
                
                <div class="skill-panel" id="practicedPanel">
                  <div class="quick-pills" id="practicedPills"></div>
                  <textarea 
                    class="form-textarea" 
                    id="skillsPracticed" 
                    placeholder="What did you practice together?"
                    data-section="practiced"></textarea>
                  <div class="character-counter" id="practicedCounter">0 / 300</div>
                </div>
                
                <div class="skill-panel" id="introducedPanel">
                  <div class="quick-pills" id="introducedPills"></div>
                  <textarea 
                    class="form-textarea" 
                    id="skillsIntroduced" 
                    placeholder="What new concepts or skills were introduced?"
                    data-section="introduced"></textarea>
                  <div class="character-counter" id="introducedCounter">0 / 300</div>
                </div>
              </div>
            </div>
            
            <!-- Additional Sections -->
            <div class="form-section">
              <div class="section-header">
                <span class="material-symbols-outlined section-icon">psychology_alt</span>
                <h3 class="section-title">Challenges & Areas for Growth</h3>
                <p class="section-description">What needs attention?</p>
              </div>
              
              <div class="quick-pills" id="strugglesPills"></div>
              <textarea 
                class="form-textarea" 
                id="struggles" 
                placeholder="Note any challenges or areas that need more focus..."
                data-section="struggles"></textarea>
              <div class="character-counter" id="strugglesCounter">0 / 400</div>
            </div>
            
            <div class="form-section">
              <div class="section-header">
                <span class="material-symbols-outlined section-icon">family_restroom</span>
                <h3 class="section-title">Parent Communication</h3>
                <p class="section-description">Key points to share</p>
              </div>
              
              <div class="quick-pills" id="parentPills"></div>
              <textarea 
                class="form-textarea" 
                id="parent" 
                placeholder="Important information to share with parents..."
                data-section="parent"></textarea>
              <div class="character-counter" id="parentCounter">0 / 400</div>
            </div>
            
            <div class="form-section">
              <div class="section-header">
                <span class="material-symbols-outlined section-icon">schedule</span>
                <h3 class="section-title">Next Session Planning</h3>
                <p class="section-description">Prepare for next time</p>
              </div>
              
              <div class="quick-pills" id="nextPills"></div>
              <textarea 
                class="form-textarea" 
                id="next" 
                placeholder="What should we focus on in the next session?"
                data-section="next"></textarea>
              <div class="character-counter" id="nextCounter">0 / 400</div>
            </div>
          </div>
          
          <!-- Actions -->
          <div class="actions">
            <div class="actions-left">
              <div class="progress-indicator">
                <span>Completion</span>
                <div class="progress-bar">
                  <div class="progress-fill" id="progressFill" style="width: 0%;"></div>
                </div>
                <span id="progressPercent">0%</span>
              </div>
            </div>
            <div class="actions-right">
              <button class="btn" id="clearBtn">
                <span class="material-symbols-outlined">refresh</span>
                Clear All
              </button>
              <button class="btn" id="previewBtn">
                <span class="material-symbols-outlined">preview</span>
                Preview
              </button>
              <button class="btn btn-success" id="sendBtn" disabled>
                <span class="material-symbols-outlined">send</span>
                Send Recap
              </button>
            </div>
          </div>
        </div>
        
        <script>
          // Application state
          let clientData = null;
          let userPreferences = null;
          let autoSaveTimeouts = {};
          let isOnline = navigator.onLine;
          let unsavedChanges = {};
          
          // Initialize application
          window.addEventListener('load', async () => {
            await initializeApp();
          });
          
          // Monitor online/offline status
          window.addEventListener('online', () => {
            isOnline = true;
            document.getElementById('offlineIndicator').classList.remove('show');
            syncOfflineChanges();
          });
          
          window.addEventListener('offline', () => {
            isOnline = false;
            document.getElementById('offlineIndicator').classList.add('show');
          });
          
          async function initializeApp() {
            try {
              // Load client data and preferences
              await Promise.all([
                loadClientData(),
                loadUserPreferences(),
                loadExistingNotes()
              ]);
              
              // Hide loading screen and show main interface
              document.getElementById('loadingScreen').style.display = 'none';
              document.getElementById('mainInterface').style.display = 'flex';
              document.getElementById('mainInterface').classList.add('fade-in');
              
              // Setup interface
              setupInterface();
              setupEventListeners();
              setupAutoSave();
              
              // Initial progress calculation
              calculateProgress();
              
            } catch (error) {
              console.error('Error initializing session recap:', error);
              // Fallback to basic mode without advanced features
              initializeBasicMode();
            }
          }
          
          async function loadClientData() {
            const response = await new Promise((resolve, reject) => {
              google.script.run
                .withSuccessHandler(resolve)
                .withFailureHandler(reject)
                .getRecapDialogData();
            });
            
            clientData = response;
            
            // Update header
            document.getElementById('clientName').textContent = \`Session Recap - \${clientData.studentName}\`;
            document.getElementById('sessionDate').textContent = new Date().toLocaleDateString('en-US', {
              weekday: 'long',
              year: 'numeric',
              month: 'long',
              day: 'numeric'
            });
          }
          
          async function loadUserPreferences() {
            try {
              const response = await new Promise((resolve, reject) => {
                google.script.run
                  .withSuccessHandler(resolve)
                  .withFailureHandler(reject)
                  .getUserPreferences();
              });
              
              userPreferences = response;
              populateQuickPills();
              
            } catch (error) {
              // Use default preferences
              userPreferences = getDefaultPreferences();
              populateQuickPills();
            }
          }
          
          function getDefaultPreferences() {
            return {
              quickPills: {
                wins: ['Great participation', 'Improved focus', 'Asked good questions', 'Completed homework'],
                mastered: ['Algebra basics', 'Reading comprehension', 'Essay structure', 'Math formulas'],
                practiced: ['Problem solving', 'Time management', 'Note taking', 'Critical thinking'],
                introduced: ['New concept', 'Advanced technique', 'Study strategy', 'Test strategy'],
                struggles: ['Time pressure', 'Concept confusion', 'Motivation', 'Organization'],
                parent: ['Homework assigned', 'Test preparation', 'Progress update', 'Practice needed'],
                next: ['Review material', 'Practice problems', 'Prepare for test', 'Continue topic']
              }
            };
          }
          
          function populateQuickPills() {
            const sections = ['wins', 'mastered', 'practiced', 'introduced', 'struggles', 'parent', 'next'];
            
            sections.forEach(section => {
              const pillsContainer = document.getElementById(\`\${section}Pills\`);
              const pills = userPreferences.quickPills[section] || [];
              
              pillsContainer.innerHTML = pills.map(pill => 
                \`<button class="quick-pill" onclick="addQuickPill('\${section}', '\${pill}')">\${pill}</button>\`
              ).join('');
            });
          }
          
          function addQuickPill(section, text) {
            const textarea = document.getElementById(section === 'mastered' || section === 'practiced' || section === 'introduced' ? 'skills' + section.charAt(0).toUpperCase() + section.slice(1) : section);
            
            if (textarea.value) {
              textarea.value += ', ' + text;
            } else {
              textarea.value = text;
            }
            
            // Trigger auto-save
            handleTextareaInput(textarea);
            
            // Visual feedback
            event.target.classList.add('active');
            setTimeout(() => event.target.classList.remove('active'), 200);
          }
          
          function setupInterface() {
            // Setup skills tabs
            setupSkillsTabs();
            
            // Setup character counters
            setupCharacterCounters();
          }
          
          function setupSkillsTabs() {
            const tabs = document.querySelectorAll('.skill-tab');
            const panels = document.querySelectorAll('.skill-panel');
            
            tabs.forEach(tab => {
              tab.addEventListener('click', () => {
                // Remove active class from all tabs and panels
                tabs.forEach(t => t.classList.remove('active'));
                panels.forEach(p => p.classList.remove('active'));
                
                // Add active class to clicked tab and corresponding panel
                tab.classList.add('active');
                document.getElementById(\`\${tab.dataset.skill}Panel\`).classList.add('active');
              });
            });
          }
          
          function setupCharacterCounters() {
            const textareas = document.querySelectorAll('.form-textarea');
            
            textareas.forEach(textarea => {
              const counterId = textarea.id + 'Counter';
              const counter = document.getElementById(counterId);
              
              if (counter) {
                const maxLength = parseInt(counter.textContent.split(' / ')[1]);
                
                textarea.addEventListener('input', () => {
                  const length = textarea.value.length;
                  counter.textContent = \`\${length} / \${maxLength}\`;
                  
                  counter.classList.remove('warning', 'error');
                  if (length > maxLength * 0.9) {
                    counter.classList.add('warning');
                  }
                  if (length > maxLength) {
                    counter.classList.add('error');
                  }
                });
              }
            });
          }
          
          function setupEventListeners() {
            // Text areas
            document.querySelectorAll('.form-textarea').forEach(textarea => {
              textarea.addEventListener('input', () => handleTextareaInput(textarea));
              textarea.addEventListener('focus', () => textarea.classList.add('expanded'));
              textarea.addEventListener('blur', () => {
                if (!textarea.value) {
                  textarea.classList.remove('expanded');
                }
              });
            });
            
            // Action buttons
            document.getElementById('clearBtn').addEventListener('click', clearAllNotes);
            document.getElementById('previewBtn').addEventListener('click', previewRecap);
            document.getElementById('sendBtn').addEventListener('click', sendRecap);
          }
          
          function setupAutoSave() {
            // Auto-save setup is handled in handleTextareaInput
          }
          
          function handleTextareaInput(textarea) {
            const section = textarea.dataset.section || textarea.id;
            
            // Clear existing timeout
            if (autoSaveTimeouts[section]) {
              clearTimeout(autoSaveTimeouts[section]);
            }
            
            // Mark as unsaved
            unsavedChanges[section] = textarea.value;
            
            // Update auto-save status to "saving"
            updateAutoSaveStatus(section, 'saving');
            
            // Set new timeout for auto-save
            autoSaveTimeouts[section] = setTimeout(() => {
              autoSaveNotes(section, textarea.value);
            }, 1000); // Auto-save after 1 second of inactivity
            
            // Update progress
            calculateProgress();
          }
          
          async function autoSaveNotes(section, value) {
            try {
              const notes = {};
              notes[section] = value;
              
              if (isOnline) {
                await new Promise((resolve, reject) => {
                  google.script.run
                    .withSuccessHandler(resolve)
                    .withFailureHandler(reject)
                    .saveQuickNotes(notes);
                });
                
                updateAutoSaveStatus(section, 'saved');
                delete unsavedChanges[section];
              } else {
                // Save to local storage when offline
                localStorage.setItem(\`recap_\${section}\`, value);
                updateAutoSaveStatus(section, 'saved');
              }
              
            } catch (error) {
              console.error('Auto-save error:', error);
              updateAutoSaveStatus(section, 'error');
            }
          }
          
          function updateAutoSaveStatus(section, status) {
            const statusEl = document.getElementById(\`\${section}AutoSave\`);
            if (!statusEl) return;
            
            statusEl.classList.remove('saving', 'saved', 'error');
            statusEl.classList.add(status);
            
            const icon = statusEl.querySelector('.material-symbols-outlined');
            const text = statusEl.querySelector('span:last-child');
            
            switch (status) {
              case 'saving':
                icon.textContent = 'sync';
                icon.classList.add('pulse');
                text.textContent = 'Saving...';
                break;
              case 'saved':
                icon.textContent = 'check_circle';
                icon.classList.remove('pulse');
                text.textContent = 'Auto-saved';
                break;
              case 'error':
                icon.textContent = 'error';
                icon.classList.remove('pulse');
                text.textContent = 'Save failed';
                break;
            }
          }
          
          function calculateProgress() {
            const textareas = document.querySelectorAll('.form-textarea');
            const filled = Array.from(textareas).filter(ta => ta.value.trim().length > 0).length;
            const total = textareas.length;
            const percentage = Math.round((filled / total) * 100);
            
            document.getElementById('progressFill').style.width = percentage + '%';
            document.getElementById('progressPercent').textContent = percentage + '%';
            
            // Enable send button if at least 50% complete
            document.getElementById('sendBtn').disabled = percentage < 50;
          }
          
          async function syncOfflineChanges() {
            const sections = Object.keys(unsavedChanges);
            
            for (const section of sections) {
              try {
                await autoSaveNotes(section, unsavedChanges[section]);
              } catch (error) {
                console.error('Failed to sync offline changes for', section, error);
              }
            }
          }
          
          function clearAllNotes() {
            if (confirm('Are you sure you want to clear all notes? This action cannot be undone.')) {
              document.querySelectorAll('.form-textarea').forEach(textarea => {
                textarea.value = '';
                handleTextareaInput(textarea);
              });
            }
          }
          
          function previewRecap() {
            // Open preview dialog with formatted recap
            const notes = gatherAllNotes();
            showPreviewDialog(notes);
          }
          
          async function sendRecap() {
            try {
              const notes = gatherAllNotes();
              
              // Show loading state
              const sendBtn = document.getElementById('sendBtn');
              const originalText = sendBtn.innerHTML;
              sendBtn.innerHTML = '<span class="material-symbols-outlined pulse">send</span> Sending...';
              sendBtn.disabled = true;
              
              await new Promise((resolve, reject) => {
                google.script.run
                  .withSuccessHandler(resolve)
                  .withFailureHandler(reject)
                  .sendIndividualRecap(clientData.sheetName, notes);
              });
              
              // Success feedback
              sendBtn.innerHTML = '<span class="material-symbols-outlined">check_circle</span> Sent!';
              sendBtn.classList.add('btn-success');
              
              setTimeout(() => {
                google.script.host.close();
              }, 1500);
              
            } catch (error) {
              console.error('Error sending recap:', error);
              
              // Error feedback
              const sendBtn = document.getElementById('sendBtn');
              sendBtn.innerHTML = '<span class="material-symbols-outlined">error</span> Failed to Send';
              sendBtn.classList.add('btn-error');
              
              setTimeout(() => {
                sendBtn.innerHTML = originalText;
                sendBtn.classList.remove('btn-error');
                sendBtn.disabled = false;
              }, 3000);
            }
          }
          
          function gatherAllNotes() {
            const textareas = document.querySelectorAll('.form-textarea');
            const notes = {};
            
            textareas.forEach(textarea => {
              const section = textarea.dataset.section || textarea.id;
              notes[section] = textarea.value.trim();
            });
            
            return notes;
          }
          
          // Additional styles for skills tabs
          const skillsTabsStyles = \`
            .skills-tabs {
              display: flex;
              gap: 4px;
              margin-bottom: 16px;
              background: var(--background);
              padding: 4px;
              border-radius: 8px;
            }
            
            .skill-tab {
              flex: 1;
              padding: 8px 12px;
              border: none;
              background: transparent;
              border-radius: 6px;
              cursor: pointer;
              font-family: inherit;
              font-size: 14px;
              font-weight: 500;
              color: var(--text-secondary);
              display: flex;
              align-items: center;
              justify-content: center;
              gap: 6px;
              transition: all 0.2s ease;
            }
            
            .skill-tab:hover {
              background: white;
            }
            
            .skill-tab.active {
              background: white;
              color: var(--primary-color);
              box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }
            
            .skills-content {
              position: relative;
            }
            
            .skill-panel {
              display: none;
            }
            
            .skill-panel.active {
              display: block;
            }
            
            .btn-error {
              background: var(--error-color) !important;
              border-color: var(--error-color) !important;
              color: white !important;
            }
          \`;
          
          // Inject skills tabs styles
          const styleSheet = document.createElement('style');
          styleSheet.textContent = skillsTabsStyles;
          document.head.appendChild(styleSheet);
        </script>
      </body>
    </html>
  `).setWidth(800).setHeight(900);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Session Recap');
}