/**
 * Simple Client List Module - Minimal design with caching and filtering
 * Features: Instant cached loading, simple filtering, alphabetical sorting
 */

(function() {
  'use strict';

  // Configuration
  const CONFIG = {
    CACHE_DURATION: 300000, // 5 minutes
    CACHE_KEY: 'clientListCache'
  };

  // State
  const state = {
    allClients: [],
    filteredClients: [],
    currentFilter: 'active',
    isLoading: false,
    selectedClientElement: null,
    cache: {
      data: null,
      timestamp: 0
    }
  };

  // Cache Manager
  const CacheManager = {
    get() {
      // Check in-memory cache first
      if (state.cache.data && (Date.now() - state.cache.timestamp < CONFIG.CACHE_DURATION)) {
        return state.cache.data;
      }

      // Check localStorage
      try {
        const stored = localStorage.getItem(CONFIG.CACHE_KEY);
        if (stored) {
          const cache = JSON.parse(stored);
          if (Date.now() - cache.timestamp < CONFIG.CACHE_DURATION) {
            state.cache = cache;
            return cache.data;
          }
        }
      } catch (e) {
        console.warn('Failed to load cache:', e);
      }
      return null;
    },

    set(data) {
      state.cache = {
        data: data,
        timestamp: Date.now()
      };

      // Store in localStorage
      try {
        localStorage.setItem(CONFIG.CACHE_KEY, JSON.stringify(state.cache));
      } catch (e) {
        console.warn('Failed to cache data:', e);
      }
    },

    clear() {
      state.cache = { data: null, timestamp: 0 };
      try {
        localStorage.removeItem(CONFIG.CACHE_KEY);
      } catch (e) {
        console.warn('Failed to clear cache:', e);
      }
    }
  };

  // UI Manager
  const UIManager = {
    showLoading() {
      const container = document.getElementById('clientListContainer');
      if (container) {
        container.innerHTML = `
          <div class="loading-state">
            <div class="loading-spinner"></div>
            <span>Loading clients...</span>
          </div>
        `;
      }
    },

    showEmpty() {
      const container = document.getElementById('clientListContainer');
      if (container) {
        container.innerHTML = '<div class="empty-state">No clients found</div>';
      }
    },

    renderClients(clients) {
      const container = document.getElementById('clientListContainer');
      if (!container) return;

      if (!clients || clients.length === 0) {
        this.showEmpty();
        return;
      }

      const html = clients.map(client => this.renderClientCard(client)).join('');
      container.innerHTML = html;
    },

    renderClientCard(client) {
      const sheetName = client.sheetName || client.name.replace(/\s+/g, '');
      const statusClass = client.isActive ? 'active' : 'inactive';

      return `
        <div class="client-card ${statusClass}"
             onclick="ClientList.selectClient('${sheetName}', '${client.name}')"
             data-sheet="${sheetName}"
             data-name="${client.name}">
          <span class="client-name">${client.name}</span>
          <div class="client-loading" style="display: none;">
            <div class="client-spinner"></div>
          </div>
        </div>
      `;
    },

    showClientLoading(clientElement) {
      if (clientElement) {
        clientElement.classList.add('loading');
        const spinner = clientElement.querySelector('.client-loading');
        if (spinner) {
          spinner.style.display = 'flex';
        }
        state.selectedClientElement = clientElement;
      }
    },

    hideClientLoading() {
      if (state.selectedClientElement) {
        state.selectedClientElement.classList.remove('loading');
        const spinner = state.selectedClientElement.querySelector('.client-loading');
        if (spinner) {
          spinner.style.display = 'none';
        }
        state.selectedClientElement = null;
      }
    },

    updateFilterButton(filter) {
      const button = document.getElementById('filterButton');
      const dropdown = document.getElementById('filterDropdown');

      if (button) {
        const labels = {
          'all': 'All Clients',
          'active': 'Active',
          'inactive': 'Inactive'
        };
        button.textContent = labels[filter] || 'Active';
      }

      if (dropdown) {
        dropdown.style.display = 'none';
      }

      // Update active state in dropdown
      document.querySelectorAll('.filter-option').forEach(option => {
        option.classList.toggle('active', option.dataset.filter === filter);
      });
    }
  };

  // Data Manager
  const DataManager = {
    async loadClients(forceRefresh = false) {
      // Try cache first unless forced refresh
      if (!forceRefresh) {
        const cached = CacheManager.get();
        if (cached) {
          state.allClients = cached;
          this.applyFilter();
          return;
        }
      }

      // Load from server
      state.isLoading = true;
      UIManager.showLoading();

      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(clients => {
            state.allClients = clients || [];
            CacheManager.set(state.allClients);
            this.applyFilter();
            state.isLoading = false;
            resolve(clients);
          })
          .withFailureHandler(error => {
            console.error('Failed to load clients:', error);
            UIManager.showEmpty();
            state.isLoading = false;
            reject(error);
          })
          .getAllClientsForDisplay();
      });
    },

    applyFilter() {
      let filtered = [...state.allClients];

      // Apply status filter
      if (state.currentFilter === 'active') {
        filtered = filtered.filter(c => c.isActive);
      } else if (state.currentFilter === 'inactive') {
        filtered = filtered.filter(c => !c.isActive);
      }

      // Sort alphabetically
      filtered.sort((a, b) => {
        const nameA = (a.name || '').toLowerCase();
        const nameB = (b.name || '').toLowerCase();
        return nameA.localeCompare(nameB);
      });

      state.filteredClients = filtered;
      UIManager.renderClients(filtered);
    },

    setFilter(filter) {
      state.currentFilter = filter;
      UIManager.updateFilterButton(filter);
      this.applyFilter();
    }
  };

  // Filter Manager
  const FilterManager = {
    init() {
      // Toggle dropdown
      const filterButton = document.getElementById('filterButton');
      if (filterButton) {
        filterButton.addEventListener('click', this.toggleDropdown.bind(this));
      }

      // Handle filter selection
      document.querySelectorAll('.filter-option').forEach(option => {
        option.addEventListener('click', this.selectFilter.bind(this));
      });

      // Close dropdown on outside click
      document.addEventListener('click', this.handleOutsideClick.bind(this));
    },

    toggleDropdown(event) {
      event.stopPropagation();
      const dropdown = document.getElementById('filterDropdown');
      if (dropdown) {
        const isVisible = dropdown.style.display === 'block';
        dropdown.style.display = isVisible ? 'none' : 'block';
      }
    },

    selectFilter(event) {
      event.stopPropagation();
      const filter = event.currentTarget.dataset.filter;
      DataManager.setFilter(filter);
    },

    handleOutsideClick(event) {
      if (!event.target.closest('.filter-container')) {
        const dropdown = document.getElementById('filterDropdown');
        if (dropdown) {
          dropdown.style.display = 'none';
        }
      }
    }
  };

  // Public API
  const ClientListAPI = {
    init() {
      // Initialize filter manager
      FilterManager.init();

      // Load clients with cache
      DataManager.loadClients();

      // Set default filter
      UIManager.updateFilterButton(state.currentFilter);
    },

    selectClient(sheetName, displayName) {
      const clientElement = window.event ? window.event.currentTarget : null;
      UIManager.showClientLoading(clientElement);

      google.script.run
        .withSuccessHandler(() => {
          UIManager.hideClientLoading();
          // Navigate back to control panel
          if (typeof switchToControlPanel === 'function') {
            switchToControlPanel();
          }
          // Update current client display
          if (typeof updateCurrentClient === 'function') {
            updateCurrentClient();
          }
        })
        .withFailureHandler(error => {
          UIManager.hideClientLoading();
          console.error('Failed to navigate to client:', error);
        })
        .navigateToSheet(sheetName);
    },

    refresh() {
      CacheManager.clear();
      DataManager.loadClients(true);
    },

    // Called when "New Client" is clicked to refresh cache
    onNewClient() {
      this.refresh();
    }
  };

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => ClientListAPI.init());
  } else {
    ClientListAPI.init();
  }

  // Export to global scope
  window.ClientList = ClientListAPI;

})();