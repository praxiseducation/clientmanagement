<script>
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
  }
  
  // Switch to Quick Notes view
  function switchToQuickNotes() {
    switchView('quickNotesView');
    
    // Initialize Quick Notes after a brief delay to ensure view is active
    setTimeout(() => {
      if (typeof initializeQuickNotes === 'function') {
        initializeQuickNotes();
      }
    }, 100);
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
      dropdown.innerHTML = '<div style="padding: 10px; color: #999;">No clients found</div>';
      dropdown.style.display = 'block';
      return;
    }
    
    let html = '';
    clients.slice(0, 5).forEach(client => {
      html += `<div class="search-result-item" onclick="selectClient('${client.name}')">${client.name}</div>`;
    });
    
    dropdown.innerHTML = html;
    dropdown.style.display = 'block';
  }

  function showSearchDropdown() {
    const dropdown = document.getElementById('searchDropdown');
    const searchBar = document.getElementById('clientSearchBar');
    if (searchBar.value.length >= 2 && searchResults.length > 0) {
      dropdown.style.display = 'block';
    }
  }

  function hideSearchDropdown() {
    setTimeout(() => {
      document.getElementById('searchDropdown').style.display = 'none';
    }, 200);
  }

  function selectClient(clientName) {
    document.getElementById('clientSearchBar').value = clientName;
    hideSearchDropdown();
    
    google.script.run
      .withSuccessHandler(() => {
        checkClientStatus();
        showMessage('success', `Switched to ${clientName}`);
      })
      .withFailureHandler(() => {
        showMessage('error', 'Failed to switch to client');
      })
      .navigateToClient(clientName);
  }

  // Show/hide message
  function showMessage(type, message) {
    // Create or get message element
    let messageEl = document.getElementById('sidebarMessage');
    if (!messageEl) {
      messageEl = document.createElement('div');
      messageEl.id = 'sidebarMessage';
      messageEl.style.cssText = `
        position: fixed;
        top: 10px;
        right: 10px;
        padding: 10px 15px;
        border-radius: 4px;
        font-size: 14px;
        z-index: 9999;
        display: none;
      `;
      document.body.appendChild(messageEl);
    }
    
    // Set message style and content
    if (type === 'success') {
      messageEl.style.background = '#d4edda';
      messageEl.style.color = '#155724';
      messageEl.style.border = '1px solid #c3e6cb';
    } else {
      messageEl.style.background = '#f8d7da';
      messageEl.style.color = '#721c24';
      messageEl.style.border = '1px solid #f5c6cb';
    }
    
    messageEl.textContent = message;
    messageEl.style.display = 'block';
    
    // Auto hide after 3 seconds
    setTimeout(() => {
      messageEl.style.display = 'none';
    }, 3000);
  }

  // Connection monitoring
  function checkConnection() {
    google.script.run
      .withSuccessHandler(() => {
        if (!isOnline) {
          isOnline = true;
          hideConnectionBanner();
        }
      })
      .withFailureHandler(() => {
        if (isOnline) {
          isOnline = false;
          showConnectionBanner();
        }
      })
      .checkConnection();
  }

  function showConnectionBanner() {
    const banner = document.getElementById('connectionStatus');
    if (banner) {
      banner.style.display = 'flex';
    }
  }

  function hideConnectionBanner() {
    const banner = document.getElementById('connectionStatus');
    if (banner) {
      banner.style.display = 'none';
    }
  }

  // Function runner with loading states
  function runFunction(functionName, button) {
    if (button) {
      button.disabled = true;
      const originalText = button.textContent;
      button.textContent = 'Loading...';
      
      setTimeout(() => {
        button.disabled = false;
        button.textContent = originalText;
      }, 2000);
    }
    
    google.script.run[functionName]();
  }
</script>