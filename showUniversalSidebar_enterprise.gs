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
          
          /* User info section (Enterprise) */
          .user-info {
            background: linear-gradient(135deg, #003366 0%, #1a73e8 100%);
            color: white;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 20px;
            font-size: 13px;
            display: ${isEnterprise && userInfo ? 'block' : 'none'};
          }
          
          .user-name {
            font-weight: 600;
            margin-bottom: 4px;
            cursor: pointer;
          }
          
          .user-name:hover {
            text-decoration: underline;
          }
          
          .user-role {
            opacity: 0.9;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
          }
          
          /* Search dropdown */
          .search-container {
            position: relative;
            width: 100%;
            margin-bottom: 15px;
          }
          
          .search-bar {
            width: 100%;
            padding: 10px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            font-size: 14px;
            box-sizing: border-box;
          }
          
          .search-bar:focus {
            outline: none;
            border-color: #1a73e8;
            box-shadow: 0 0 0 2px rgba(26,115,232,0.1);
          }
          
          .search-dropdown {
            display: none;
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            background: white;
            border: 1px solid #dadce0;
            border-top: none;
            border-radius: 0 0 6px 6px;
            max-height: 200px;
            overflow-y: auto;
            z-index: 100;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          }
          
          .search-result {
            padding: 10px;
            cursor: pointer;
            border-bottom: 1px solid #f1f3f4;
          }
          
          .search-result:hover {
            background: #f8f9fa;
          }
          
          /* Current client section */
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
          
          /* Section styles */
          .section {
            margin-bottom: 25px;
            background: white;
            border-radius: 8px;
            padding: 15px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          .section-title {
            font-weight: 600;
            color: #202124;
            margin-bottom: 12px;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          
          .section-icon {
            font-size: 16px;
          }
          
          /* Button styles */
          .menu-button {
            width: 100%;
            padding: 12px;
            margin-bottom: 8px;
            border: 1px solid #dadce0;
            border-radius: 6px;
            background: white;
            color: #202124;
            cursor: pointer;
            font-family: 'Poppins', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
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
          
          .menu-button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
          }
          
          .primary {
            background: linear-gradient(135deg, #1a73e8 0%, #1557b0 100%);
            color: white !important;
            border: none;
            text-align: center;
            justify-content: center;
          }
          
          .primary:hover:not(:disabled) {
            background: linear-gradient(135deg, #1557b0 0%, #0d47a1 100%);
          }
          
          .admin-only {
            background: linear-gradient(135deg, #d32f2f 0%, #b71c1c 100%);
            color: white !important;
            border: none;
          }
          
          .admin-only:hover:not(:disabled) {
            background: linear-gradient(135deg, #b71c1c 0%, #8e24aa 100%);
          }
          
          .icon {
            font-size: 18px;
            width: 24px;
            text-align: center;
          }
          
          /* Loading spinner */
          .btn-loading {
            position: relative;
            pointer-events: none;
            opacity: 0.8;
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
          }
          
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          
          /* Messages */
          .message {
            font-size: 13px;
            margin-top: 10px;
            padding: 12px;
            border-radius: 6px;
            display: none;
            border-left: 4px solid;
          }
          
          .error {
            color: #d93025;
            background: #fce8e6;
            border-left-color: #d93025;
          }
          
          .success {
            color: #188038;
            background: #e6f4ea;
            border-left-color: #188038;
          }
          
          /* Version info */
          .version-info {
            margin-top: 20px;
            padding: 10px;
            background: #f8f9fa;
            border-radius: 6px;
            font-size: 11px;
            color: #5f6368;
            text-align: center;
            border-top: 1px solid #e1e5e9;
          }
        </style>
      </head>
      <body>
        <!-- Control Panel View -->
        <div id="controlView" class="view active">
          ${isEnterprise && userInfo ? `
          <!-- User Information (Enterprise) -->
          <div class="user-info">
            <div class="user-name" ondblclick="editUserName()" title="Double-click to edit">
              ${userInfo.displayName || userInfo.email.split('@')[0]}
            </div>
            <div class="user-role">${userRole}</div>
          </div>
          ` : ''}
          
          <!-- Current Client -->
          <div class="current-client">
            <div class="current-client-label">Current Client:</div>
            <div id="clientName" class="${clientInfo.isClient ? 'current-client-name' : 'no-client'}">
              ${clientInfo.isClient ? clientInfo.name : 'No client selected'}
            </div>
          </div>
          
          <!-- Client Management Section -->
          <div class="section">
            <div class="section-title">
              <span class="section-icon">👥</span>
              Client Management
            </div>
            
            <!-- Enhanced Search Bar with Dropdown -->
            <div class="search-container">
              <input type="text" class="search-bar" id="clientSearchBar" 
                placeholder="Search for a client..." 
                onkeyup="searchClients(this.value)" 
                onfocus="showSearchDropdown()" 
                onblur="hideSearchDropdown()">
              <div id="searchDropdown" class="search-dropdown"></div>
            </div>
            
            <button class="menu-button primary" onclick="runFunction('${isEnterprise ? 'createClientSheetSecure' : 'addNewClient'}', this)">
              <span class="icon">➕</span>
              Add New Client
            </button>
            <button class="menu-button" onclick="showClientList()">
              <span class="icon">📋</span>
              View All Clients
            </button>
            ${clientInfo.isClient ? `
            <button class="menu-button" onclick="openUpdateClientDialog()">
              <span class="icon">⚙️</span>
              Update Client Info
            </button>
            ` : ''}
          </div>
          
          <!-- Session Management Section -->
          <div class="section">
            <div class="section-title">
              <span class="section-icon">📝</span>
              Session Management
            </div>
            <button id="quickNotesBtn" class="menu-button" onclick="switchToQuickNotes()" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">📝</span>
              Quick Notes
            </button>
            <button id="recapBtn" class="menu-button" onclick="runFunction('${isEnterprise ? 'sendEmailSecure' : 'sendIndividualRecap'}', this)" ${!clientInfo.isClient ? 'disabled' : ''}>
              <span class="icon">✉️</span>
              Send Session Recap
            </button>
          </div>
          
          <!-- External Tools Section -->
          <div class="section">
            <div class="section-title">
              <span class="section-icon">🔗</span>
              External Tools
            </div>
            <button class="menu-button" onclick="window.open('https://secure.acuityscheduling.com/admin/clients', '_blank')">
              <span class="icon">📅</span>
              Open Acuity
            </button>
          </div>
          
          ${isEnterprise && (userRole === 'admin' || userRole === 'supervisor') ? `
          <!-- Admin Section (Enterprise) -->
          <div class="section">
            <div class="section-title">
              <span class="section-icon">⚙️</span>
              Administration
            </div>
            <button class="menu-button admin-only" onclick="runFunction('showEnterpriseAdminDashboard', this)">
              <span class="icon">🏢</span>
              Admin Dashboard
            </button>
            <button class="menu-button admin-only" onclick="runFunction('showOrganizationDashboard', this)">
              <span class="icon">🏗️</span>
              Organization Dashboard
            </button>
            <button class="menu-button" onclick="runFunction('runSystemMaintenance', this)">
              <span class="icon">🔧</span>
              System Maintenance
            </button>
            <button class="menu-button" onclick="runFunction('generateComplianceReportDialog', this)">
              <span class="icon">📋</span>
              Compliance Report
            </button>
          </div>
          ` : ''}
          
          <!-- Messages -->
          <div class="message error" id="error"></div>
          <div class="message success" id="success"></div>
          
          <!-- Version Info -->
          <div class="version-info">
            <div>Universal Sidebar v${CONFIG.version} (${CONFIG.build})</div>
          </div>
        </div>
        
        <!-- Quick Notes View (existing implementation) -->
        <div id="quickNotesView" class="view">
          <!-- Existing Quick Notes implementation remains unchanged -->
        </div>
        
        <!-- Client List View (existing implementation) -->
        <div id="clientListView" class="view">
          <!-- Existing Client List implementation remains unchanged -->
        </div>
        
        <script>
          let currentView = 'control';
          let searchTimeout;
          let searchResults = [];
          
          // Initialize on load
          window.onload = function() {
            checkClientStatus();
            ${isEnterprise ? 'setupEnterpriseFeatures();' : ''}
          };
          
          // Check client status
          function checkClientStatus() {
            const clientInfo = ${JSON.stringify(clientInfo)};
            updateButtonStates(clientInfo.isClient);
          }
          
          // Update button states based on client selection
          function updateButtonStates(hasClient) {
            const quickNotesBtn = document.getElementById('quickNotesBtn');
            const recapBtn = document.getElementById('recapBtn');
            if (quickNotesBtn) quickNotesBtn.disabled = !hasClient;
            if (recapBtn) recapBtn.disabled = !hasClient;
          }
          
          // Enhanced client search with dropdown
          function searchClients(searchTerm) {
            clearTimeout(searchTimeout);
            const dropdown = document.getElementById('searchDropdown');
            
            if (searchTerm.length < 2) {
              dropdown.style.display = 'none';
              dropdown.innerHTML = '';
              return;
            }
            
            // Debounce search
            searchTimeout = setTimeout(() => {
              google.script.run
                .withSuccessHandler((clients) => {
                  searchResults = clients || [];
                  displaySearchResults(searchResults);
                })
                .withFailureHandler((error) => {
                  console.error('Search failed:', error);
                  dropdown.style.display = 'none';
                })
                .searchClientsByName(searchTerm);
            }, 300);
          }
          
          function displaySearchResults(clients) {
            const dropdown = document.getElementById('searchDropdown');
            
            if (!clients || clients.length === 0) {
              dropdown.innerHTML = '<div style="padding: 10px; color: #5f6368;">No clients found</div>';
              dropdown.style.display = 'block';
              return;
            }
            
            dropdown.innerHTML = clients.map(client => 
              '<div class="search-result" onmousedown="selectClientFromSearch(\\'' + client.name + '\\')">' +
                '<strong>' + client.name + '</strong>' +
                (client.sheetName !== client.name ? '<small style="color: #5f6368;"> (' + client.sheetName + ')</small>' : '') +
              '</div>'
            ).join('');
            
            dropdown.style.display = 'block';
          }
          
          function selectClientFromSearch(clientName) {
            document.getElementById('clientSearchBar').value = '';
            document.getElementById('searchDropdown').style.display = 'none';
            
            google.script.run
              .withSuccessHandler(() => {
                updateCurrentClient();
                showMessage('success', 'Switched to ' + clientName);
              })
              .withFailureHandler((error) => {
                showMessage('error', 'Failed to switch client');
              })
              .switchToClient(clientName);
          }
          
          function showSearchDropdown() {
            const searchTerm = document.getElementById('clientSearchBar').value;
            if (searchTerm.length >= 2 && searchResults.length > 0) {
              displaySearchResults(searchResults);
            }
          }
          
          function hideSearchDropdown() {
            setTimeout(() => {
              document.getElementById('searchDropdown').style.display = 'none';
            }, 200);
          }
          
          // Update current client display
          function updateCurrentClient() {
            google.script.run
              .withSuccessHandler((clientInfo) => {
                const clientNameDiv = document.getElementById('clientName');
                if (clientInfo && clientInfo.isClient) {
                  clientNameDiv.textContent = clientInfo.name;
                  clientNameDiv.className = 'current-client-name';
                  updateButtonStates(true);
                } else {
                  clientNameDiv.textContent = 'No client selected';
                  clientNameDiv.className = 'no-client';
                  updateButtonStates(false);
                }
              })
              .withFailureHandler(() => {
                document.getElementById('clientName').textContent = 'Error loading client';
                document.getElementById('clientName').className = 'no-client';
              })
              .getCurrentClientInfo();
          }
          
          // Show client list
          function showClientList() {
            google.script.run
              .withSuccessHandler(() => {
                // Client list dialog will be shown by the server
              })
              .withFailureHandler((error) => {
                showMessage('error', 'Failed to load client list');
              })
              .showClientSelectionDialogEnterprise();
          }
          
          // Switch to Quick Notes view
          function switchToQuickNotes() {
            document.getElementById('controlView').classList.remove('active');
            document.getElementById('quickNotesView').classList.add('active');
            currentView = 'quickNotes';
          }
          
          // Switch back to Control Panel
          function switchToControlPanel() {
            document.getElementById('quickNotesView').classList.remove('active');
            document.getElementById('controlView').classList.add('active');
            currentView = 'control';
          }
          
          // Run server function with loading state
          function runFunction(functionName, button) {
            if (button) {
              button.classList.add('btn-loading');
              button.disabled = true;
            }
            
            google.script.run
              .withSuccessHandler((result) => {
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                }
                if (result && result.message) {
                  showMessage('success', result.message);
                }
                if (functionName.includes('Client')) {
                  updateCurrentClient();
                }
              })
              .withFailureHandler((error) => {
                if (button) {
                  button.classList.remove('btn-loading');
                  button.disabled = false;
                }
                showMessage('error', error.message || 'An error occurred');
              })[functionName]();
          }
          
          // Show messages
          function showMessage(type, message) {
            const messageDiv = document.getElementById(type);
            if (messageDiv) {
              messageDiv.textContent = message;
              messageDiv.style.display = 'block';
              setTimeout(() => {
                messageDiv.style.display = 'none';
              }, 5000);
            }
          }
          
          ${isEnterprise ? `
          // Enterprise-specific features
          function setupEnterpriseFeatures() {
            // Add any enterprise-specific initialization here
          }
          
          // Edit user name (Enterprise)
          function editUserName() {
            const currentName = document.querySelector('.user-name').textContent.trim();
            const newName = prompt('Enter your name:', currentName);
            if (newName && newName !== currentName) {
              google.script.run
                .withSuccessHandler(() => {
                  document.querySelector('.user-name').textContent = newName;
                  showMessage('success', 'Name updated successfully');
                })
                .withFailureHandler((error) => {
                  showMessage('error', 'Failed to update name');
                })
                .updateUserDisplayName(newName);
            }
          }
          ` : ''}
          
          // Refresh data periodically
          setInterval(() => {
            updateCurrentClient();
          }, 30000);
        </script>
      </body>
    </html>
  `).setWidth(350).setHeight(600);
  
  SpreadsheetApp.getUi().showSidebar(html);
}