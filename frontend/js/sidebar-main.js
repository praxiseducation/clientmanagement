// Sidebar Main JavaScript
  let currentView = 'control';
  let searchTimeout;
  let searchResults = [];
  let isOnline = true;
  let disconnectTimeout = null;
  
  // Initialize on load
  window.onload = function() {
    checkClientStatus();
    if (typeof setupEnterpriseFeatures === 'function') {
      setupEnterpriseFeatures();
    }
    // Start connection monitoring
    setInterval(checkConnection, 10000);
    
    // Preload client list for instant search
    google.script.run
      .withSuccessHandler((clients) => {
        window.cachedClientList = clients;
      })
      .withFailureHandler((error) => {
        console.log('Could not preload client list:', error);
      })
      .getCachedClientData();
  };
  
  // Check client status
  function checkClientStatus() {
    google.script.run
      .withSuccessHandler((clientInfo) => {
        updateButtonStates(clientInfo && clientInfo.isClient);
      })
      .withFailureHandler(() => {
        updateButtonStates(false);
      })
      .getCurrentClientInfo();
  }
  
  // Update button states based on client selection
  function updateButtonStates(hasClient) {
    const quickNotesBtn = document.getElementById('quickNotesBtn');
    const recapBtn = document.getElementById('recapBtn');
    if (quickNotesBtn) quickNotesBtn.disabled = !hasClient;
    if (recapBtn) recapBtn.disabled = !hasClient;
  }

  // Switch between views in the sidebar - Enhanced for smooth transitions
  function switchView(viewId) {
    // Hide all views
    const views = document.querySelectorAll('.view');
    views.forEach(view => view.classList.remove('active'));

    // Show the selected view with proper timing
    const targetView = document.getElementById(viewId);
    if (targetView) {
      // Small delay to ensure clean transition
      requestAnimationFrame(() => {
        targetView.classList.add('active');
        currentView = viewId;
      });
    }
  }
  
  // Switch back to control panel
  function switchToControlPanel() {
    switchView('controlView');
  }
  
  // Enhanced client search with dropdown - INSTANT
  function searchClients(searchTerm) {
    const dropdown = document.getElementById('searchDropdown');
    
    if (searchTerm.length < 2) {
      dropdown.style.display = 'none';
      dropdown.innerHTML = '';
      return;
    }
    
    // Try to use cached data first for instant results
    if (window.cachedClientList && window.cachedClientList.length > 0) {
      const searchLower = searchTerm.toLowerCase();
      const filtered = window.cachedClientList.filter(client => 
        client.name.toLowerCase().includes(searchLower)
      ).slice(0, 10);
      searchResults = filtered;
      displaySearchResults(filtered);
    } else {
      // Fall back to server call if no cache
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        google.script.run
          .withSuccessHandler((clients) => {
            window.cachedClientList = clients; // Cache for next time
            searchResults = clients || [];
            displaySearchResults(searchResults);
          })
          .withFailureHandler((error) => {
            console.error('Search failed:', error);
            dropdown.style.display = 'none';
          })
          .getCachedClientData(); // Use cached data function
      }, 100); // Reduced delay
    }
  }
  
  function displaySearchResults(clients) {
    const dropdown = document.getElementById('searchDropdown');
    
    if (!clients || clients.length === 0) {
      dropdown.innerHTML = '<div style="padding: 10px; color: #5f6368;">No clients found</div>';
      dropdown.style.display = 'block';
      return;
    }
    
    dropdown.innerHTML = clients.map(client => 
      '<div class="search-result" onmousedown="selectClientFromSearch(\'' + (client.sheetName || client.name) + '\', \'' + client.name + '\')">' +
        '<strong>' + client.name + '</strong>' +
        (client.email ? '<div style="font-size: 11px; color: #5f6368;">' + client.email + '</div>' : '') +
      '</div>'
    ).join('');
    
    dropdown.style.display = 'block';
  }
  
  function selectClientFromSearch(sheetName, displayName) {
    document.getElementById('clientSearchBar').value = '';
    document.getElementById('searchDropdown').style.display = 'none';
    
    google.script.run
      .withSuccessHandler(() => {
        updateCurrentClient();
        showMessage('success', 'Switched to ' + (displayName || sheetName));
      })
      .withFailureHandler((error) => {
        showMessage('error', 'Failed to switch client: ' + error.message);
      })
      .navigateToSheet(sheetName);
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
  
  // Switch between views in the sidebar
  function switchView(viewId) {
    // Hide all views
    const views = document.querySelectorAll('.view');
    views.forEach(view => view.classList.remove('active'));
    
    // Show the selected view
    const targetView = document.getElementById(viewId);
    if (targetView) {
      targetView.classList.add('active');
      currentView = viewId;
    }
  }
  
  // Switch back to control panel
  function switchToControlPanel() {
    switchView('controlView');
  }
  
  // Show client list
  function showClientList() {
    switchView('clientListView');
    
    // Initialize client list if it hasn't been loaded yet
    if (typeof initializeClientList === 'function') {
      setTimeout(() => {
        initializeClientList();
      }, 100);
    }
  }
  
  // Switch to Quick Notes view
  function switchToQuickNotes() {
    document.getElementById('controlView').classList.remove('active');
    document.getElementById('quickNotesView').classList.add('active');
    currentView = 'quickNotes';
    
    // Initialize Quick Notes if function exists
    if (typeof initializeQuickNotes === 'function') {
      initializeQuickNotes();
    }
  }
  
  // Switch back to Control Panel
  function switchToControlPanel() {
    document.getElementById('quickNotesView').classList.remove('active');
    document.getElementById('controlView').classList.add('active');
    currentView = 'control';
  }
  
  // Open new client dialog
  function openNewClientDialog(button) {
    if (button) {
      button.classList.add('btn-loading');
      button.disabled = true;
    }

    google.script.run
      .withSuccessHandler(() => {
        if (button) {
          button.classList.remove('btn-loading');
          button.disabled = false;
        }
        // Refresh client list cache after dialog closes
        if (typeof ClientList !== 'undefined' && ClientList.onNewClient) {
          ClientList.onNewClient();
        }
        // Refresh cached client list for search
        refreshClientCache();
        // Update current client display
        updateCurrentClient();
      })
      .withFailureHandler((error) => {
        if (button) {
          button.classList.remove('btn-loading');
          button.disabled = false;
        }
        showMessage('error', error.message || 'Failed to open new client dialog');
      })
      .addNewClient();
  }

  // Refresh cached client data
  function refreshClientCache() {
    google.script.run
      .withSuccessHandler((clients) => {
        window.cachedClientList = clients;
        console.log('Client cache refreshed with', clients ? clients.length : 0, 'clients');
      })
      .withFailureHandler((error) => {
        console.error('Failed to refresh client cache:', error);
      })
      .getCachedClientData();
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
  
  // Connection monitoring
  function checkConnection() {
    google.script.run
      .withSuccessHandler(() => {
        handleOnline();
        if (disconnectTimeout) {
          clearTimeout(disconnectTimeout);
          disconnectTimeout = null;
        }
      })
      .withFailureHandler(() => {
        if (isOnline) {
          if (!disconnectTimeout) {
            disconnectTimeout = setTimeout(() => {
              handleOffline();
              disconnectTimeout = null;
            }, 2000);
          }
        }
      })
      .testConnectivity();
  }
  
  function handleOnline() {
    isOnline = true;
    const banner = document.getElementById('connectionStatus');
    if (banner) {
      banner.classList.remove('show');
    }
  }
  
  function handleOffline() {
    isOnline = false;
    const banner = document.getElementById('connectionStatus');
    if (banner) {
      banner.classList.add('show');
    }
  }
  
  // Browser offline/online events
  window.addEventListener('online', () => {
    handleOnline();
    checkConnection();
  });
  
  window.addEventListener('offline', () => {
    if (disconnectTimeout) {
      clearTimeout(disconnectTimeout);
      disconnectTimeout = null;
    }
    handleOffline();
  });
  
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
  
  // Open update client dialog
  function openUpdateClientDialog() {
    google.script.run
      .withSuccessHandler(() => {
        // Dialog will be opened by server
      })
      .withFailureHandler((error) => {
        showMessage('error', 'Failed to open update dialog');
      })
      .showFullUpdateClientDialog();
  }
  
  // Refresh data periodically
  setInterval(() => {
    updateCurrentClient();
  }, 30000);