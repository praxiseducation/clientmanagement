/**
 * Quick Notes Module - Complete functionality for Quick Notes feature
 * Includes: Auto-save, Quick Buttons, Keyboard Shortcuts, Local Caching
 */

(function() {
  'use strict';

  // Module configuration
  const CONFIG = {
    AUTO_SAVE_DELAY: 2000,
    MESSAGE_DISPLAY_TIME: 3000,
    CACHE_KEY_PREFIX: 'quickNotes_cache_',
    MAX_QUICK_BUTTONS: 4,
    SECTIONS: ['wins', 'skillsMastered', 'skillsPracticed', 'skillsIntroduced', 'struggles', 'parent', 'next']
  };

  // State management
  const state = {
    currentEditingSection: '',
    autoSaveTimeout: null,
    isDirty: false,
    cache: {},
    clientInfo: null,
    quickButtonSettings: {},
    keyboardShortcutsEnabled: true
  };

  // Cache Management
  const CacheManager = {
    get(key) {
      try {
        const cached = localStorage.getItem(CONFIG.CACHE_KEY_PREFIX + key);
        if (cached) {
          const data = JSON.parse(cached);
          // Check if cache is still valid (24 hours)
          if (Date.now() - data.timestamp < 86400000) {
            return data.value;
          }
          localStorage.removeItem(CONFIG.CACHE_KEY_PREFIX + key);
        }
      } catch (e) {
        console.error('Cache read error:', e);
      }
      return null;
    },

    set(key, value) {
      try {
        localStorage.setItem(CONFIG.CACHE_KEY_PREFIX + key, JSON.stringify({
          value: value,
          timestamp: Date.now()
        }));
      } catch (e) {
        console.error('Cache write error:', e);
        // Clear old cache if storage is full
        this.clearOldCache();
      }
    },

    clearOldCache() {
      const keys = Object.keys(localStorage).filter(k => k.startsWith(CONFIG.CACHE_KEY_PREFIX));
      if (keys.length > 10) {
        keys.slice(0, 5).forEach(key => localStorage.removeItem(key));
      }
    }
  };

  // UI Manager
  const UIManager = {
    showMessage(message, type = 'info') {
      let messageDiv = document.getElementById('quickNotesMessage');
      if (!messageDiv) {
        messageDiv = document.createElement('div');
        messageDiv.id = 'quickNotesMessage';
        messageDiv.style.cssText = `
          position: fixed;
          top: 20px;
          right: 20px;
          padding: 12px 20px;
          border-radius: 6px;
          font-size: 14px;
          z-index: 1000;
          animation: slideIn 0.3s ease;
        `;
        document.body.appendChild(messageDiv);
      }

      const styles = {
        success: { background: '#e6f4ea', color: '#188038', border: '1px solid #188038' },
        error: { background: '#fce8e6', color: '#d93025', border: '1px solid #d93025' },
        info: { background: '#e8f0fe', color: '#1a73e8', border: '1px solid #1a73e8' }
      };

      Object.assign(messageDiv.style, styles[type] || styles.info);
      messageDiv.textContent = message;
      messageDiv.style.display = 'block';

      setTimeout(() => {
        messageDiv.style.display = 'none';
      }, CONFIG.MESSAGE_DISPLAY_TIME);
    },

    updateSaveButton(status = 'default') {
      const saveBtn = document.getElementById('saveQuickNotesBtn');
      const saveText = document.getElementById('saveNotesText');

      if (!saveBtn || !saveText) return;

      const states = {
        default: { className: 'save-button', text: 'Save' },
        loading: { className: 'save-button loading', text: 'Saving...' },
        success: { className: 'save-button success', text: 'Saved' },
        error: { className: 'save-button error', text: 'Error' }
      };

      const newState = states[status] || states.default;
      saveBtn.className = newState.className;
      saveText.textContent = newState.text;

      if (status === 'success' || status === 'error') {
        setTimeout(() => this.updateSaveButton('default'), 2000);
      }
    },

    renderPillButtons(container, buttons, section) {
      if (!container) return;

      container.innerHTML = '';
      if (!buttons || buttons.length === 0) {
        container.innerHTML = '<span style="color: #999; font-size: 12px;">No quick buttons configured</span>';
        return;
      }

      buttons.forEach((buttonData, index) => {
        if (buttonData && buttonData.text) {
          const button = document.createElement('button');
          button.className = 'pill-button';
          button.textContent = buttonData.text;
          button.setAttribute('data-content', buttonData.content || buttonData.text);
          button.setAttribute('data-index', index);
          button.title = `Click to insert • Ctrl+${index + 1} for keyboard shortcut`;
          button.onclick = () => QuickButtons.insert(button, section);
          container.appendChild(button);
        }
      });
    }
  };

  // Data Manager
  const DataManager = {
    async loadClientInfo() {
      // Try cache first
      const cached = CacheManager.get('clientInfo');
      if (cached && state.clientInfo === null) {
        state.clientInfo = cached;
        this.updateClientDisplay(cached);
      }

      // Then fetch fresh data
      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(clientInfo => {
            state.clientInfo = clientInfo;
            CacheManager.set('clientInfo', clientInfo);
            this.updateClientDisplay(clientInfo);
            resolve(clientInfo);
          })
          .withFailureHandler(error => {
            console.error('Failed to load client info:', error);
            reject(error);
          })
          .getCurrentClientInfo();
      });
    },

    updateClientDisplay(clientInfo) {
      const nameElement = document.getElementById('quickNotesClientName');
      if (nameElement && clientInfo) {
        nameElement.textContent = clientInfo.name || 'No client selected';
      }
    },

    async loadSavedNotes() {
      // Try cache first for immediate display
      const cacheKey = state.clientInfo ? `notes_${state.clientInfo.name}` : 'notes_default';
      const cached = CacheManager.get(cacheKey);

      if (cached) {
        this.populateNotes(cached);
      }

      // Fetch fresh data
      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(notesData => {
            if (notesData && notesData.data) {
              try {
                const notes = JSON.parse(notesData.data);
                this.populateNotes(notes);
                CacheManager.set(cacheKey, notes);
                resolve(notes);
              } catch (e) {
                console.error('Error parsing notes:', e);
                reject(e);
              }
            }
          })
          .withFailureHandler(error => {
            console.error('Failed to load notes:', error);
            reject(error);
          })
          .getQuickNotes();
      });
    },

    populateNotes(notes) {
      CONFIG.SECTIONS.forEach(field => {
        const element = document.getElementById(field);
        if (element && notes[field] !== undefined) {
          element.value = notes[field];
        }
      });
    },

    async saveNotes(isAutoSave = false) {
      const notes = {};
      CONFIG.SECTIONS.forEach(field => {
        const element = document.getElementById(field);
        if (element) {
          notes[field] = element.value;
        }
      });

      // Update cache immediately for offline capability
      const cacheKey = state.clientInfo ? `notes_${state.clientInfo.name}` : 'notes_default';
      CacheManager.set(cacheKey, notes);

      if (!isAutoSave) {
        UIManager.updateSaveButton('loading');
      }

      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(() => {
            state.isDirty = false;
            if (!isAutoSave) {
              UIManager.updateSaveButton('success');
              UIManager.showMessage('Notes saved successfully!', 'success');
            }
            resolve();
          })
          .withFailureHandler(error => {
            if (!isAutoSave) {
              UIManager.updateSaveButton('error');
              UIManager.showMessage('Failed to save notes', 'error');
            }
            console.error(error);
            reject(error);
          })
          .saveQuickNotes(notes);
      });
    }
  };

  // Quick Buttons Manager
  const QuickButtons = {
    async loadAll() {
      const promises = CONFIG.SECTIONS.map(section => this.loadForSection(section));
      await Promise.all(promises);
    },

    async loadForSection(section) {
      // Try cache first
      const cacheKey = `buttons_${section}`;
      const cached = CacheManager.get(cacheKey);

      if (cached) {
        state.quickButtonSettings[section] = cached;
        const container = document.getElementById(section + 'Buttons');
        if (container) {
          UIManager.renderPillButtons(container, cached, section);
        }
      }

      // Fetch fresh data
      return new Promise(resolve => {
        google.script.run
          .withSuccessHandler(buttons => {
            state.quickButtonSettings[section] = buttons;
            CacheManager.set(cacheKey, buttons);
            const container = document.getElementById(section + 'Buttons');
            if (container) {
              UIManager.renderPillButtons(container, buttons, section);
            }
            resolve(buttons);
          })
          .withFailureHandler(error => {
            console.error(`Failed to load buttons for ${section}:`, error);
            resolve([]);
          })
          .getQuickButtonSettings(section);
      });
    },

    insert(button, section) {
      const textarea = document.getElementById(section);
      if (!textarea) return;

      let content = button.getAttribute('data-content') || button.textContent;
      content = this.processPlaceholders(content);

      const start = textarea.selectionStart;
      const end = textarea.selectionEnd;
      const currentValue = textarea.value;

      // Smart spacing
      let newContent = content;
      if (start > 0 && currentValue[start - 1] !== '\n' && currentValue[start - 1] !== ' ') {
        newContent = ' ' + newContent;
      }
      if (end < currentValue.length && currentValue[end] !== '\n' && currentValue[end] !== ' ') {
        newContent = newContent + ' ';
      }

      textarea.value = currentValue.substring(0, start) + newContent + currentValue.substring(end);

      const newPosition = start + newContent.length;
      textarea.focus();
      textarea.setSelectionRange(newPosition, newPosition);

      // Mark as dirty and trigger auto-save
      state.isDirty = true;
      AutoSave.trigger();

      // Visual feedback
      button.classList.add('selected');
      setTimeout(() => button.classList.remove('selected'), 300);
    },

    processPlaceholders(content) {
      const replacements = {
        '[userFirstName]': () => {
          const clientName = document.getElementById('quickNotesClientName')?.textContent;
          return clientName && clientName !== 'Loading...' && clientName !== 'No client selected'
            ? clientName.split(' ')[0]
            : '[Student Name]';
        },
        '[date]': () => new Date().toLocaleDateString(),
        '[time]': () => new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
        '[day]': () => new Date().toLocaleDateString('en-US', { weekday: 'long' })
      };

      Object.entries(replacements).forEach(([placeholder, replacer]) => {
        content = content.replace(new RegExp(placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), replacer());
      });

      return content;
    }
  };

  // Auto-save Manager
  const AutoSave = {
    setup() {
      const textareas = document.querySelectorAll('.notes-textarea');
      textareas.forEach(textarea => {
        textarea.addEventListener('input', () => {
          state.isDirty = true;
          this.trigger();
        });

        // Save on blur for better data persistence
        textarea.addEventListener('blur', () => {
          if (state.isDirty) {
            DataManager.saveNotes(true);
          }
        });
      });
    },

    trigger() {
      if (state.autoSaveTimeout) {
        clearTimeout(state.autoSaveTimeout);
      }

      state.autoSaveTimeout = setTimeout(() => {
        if (state.isDirty) {
          DataManager.saveNotes(true);
        }
      }, CONFIG.AUTO_SAVE_DELAY);
    }
  };

  // Keyboard Shortcuts Manager
  const KeyboardShortcuts = {
    setup() {
      document.addEventListener('keydown', this.handleKeyPress.bind(this));
    },

    handleKeyPress(event) {
      if (!state.keyboardShortcutsEnabled) return;

      // Escape key - close modal if open
      if (event.key === 'Escape') {
        const modal = document.getElementById('quickButtonModal');
        if (modal && modal.style.display === 'flex') {
          Modal.close();
          event.preventDefault();
        }
      }

      // Ctrl/Cmd + S - Save notes
      if ((event.ctrlKey || event.metaKey) && event.key === 's') {
        event.preventDefault();
        DataManager.saveNotes(false);
      }

      // Ctrl/Cmd + 1-4 - Insert quick buttons for focused section
      if ((event.ctrlKey || event.metaKey) && event.key >= '1' && event.key <= '4') {
        const activeElement = document.activeElement;
        if (activeElement && activeElement.classList.contains('notes-textarea')) {
          const section = activeElement.id;
          const buttons = state.quickButtonSettings[section];
          const buttonIndex = parseInt(event.key) - 1;

          if (buttons && buttons[buttonIndex]) {
            event.preventDefault();
            const fakeButton = {
              getAttribute: (attr) => attr === 'data-content' ? buttons[buttonIndex].content : buttons[buttonIndex].text,
              textContent: buttons[buttonIndex].text,
              classList: { add: () => {}, remove: () => {} }
            };
            QuickButtons.insert(fakeButton, section);
            UIManager.showMessage(`Inserted: ${buttons[buttonIndex].text}`, 'info');
          }
        }
      }

      // Alt + Q - Quick save and close
      if (event.altKey && event.key === 'q') {
        event.preventDefault();
        DataManager.saveNotes(false).then(() => {
          UIManager.showMessage('Notes saved! Press Alt+Q again to switch view', 'success');
        });
      }
    }
  };

  // Modal Manager
  const Modal = {
    open(section) {
      state.currentEditingSection = section;
      const modal = document.getElementById('quickButtonModal');
      const modalTitle = document.getElementById('modalTitle');

      const sectionTitles = {
        wins: "Today's Wins",
        skillsMastered: "Skills Mastered",
        skillsPracticed: "Skills Practiced",
        skillsIntroduced: "Skills Introduced",
        struggles: "Struggles",
        parent: "Parent Communication",
        next: "Next Session"
      };

      modalTitle.textContent = `Customize Quick Buttons - ${sectionTitles[section] || section}`;
      this.loadSettings(section);
      modal.style.display = 'flex';

      // Focus first input
      setTimeout(() => {
        document.getElementById('buttonText1')?.focus();
      }, 100);
    },

    close() {
      const modal = document.getElementById('quickButtonModal');
      modal.style.display = 'none';
      state.currentEditingSection = '';

      // Clear form
      for (let i = 1; i <= CONFIG.MAX_QUICK_BUTTONS; i++) {
        const textInput = document.getElementById(`buttonText${i}`);
        const contentInput = document.getElementById(`buttonContent${i}`);
        if (textInput) textInput.value = '';
        if (contentInput) contentInput.value = '';
      }
    },

    loadSettings(section) {
      google.script.run
        .withSuccessHandler(buttons => {
          if (buttons && Array.isArray(buttons)) {
            for (let i = 0; i < Math.min(buttons.length, CONFIG.MAX_QUICK_BUTTONS); i++) {
              if (buttons[i]) {
                const textInput = document.getElementById(`buttonText${i + 1}`);
                const contentInput = document.getElementById(`buttonContent${i + 1}`);
                if (textInput) textInput.value = buttons[i].text || '';
                if (contentInput) contentInput.value = buttons[i].content || '';
              }
            }
          }
        })
        .withFailureHandler(error => {
          console.error('Failed to load settings:', error);
        })
        .getQuickButtonSettings(section);
    },

    save() {
      const buttons = [];

      for (let i = 1; i <= CONFIG.MAX_QUICK_BUTTONS; i++) {
        const text = document.getElementById(`buttonText${i}`)?.value.trim();
        const content = document.getElementById(`buttonContent${i}`)?.value.trim();

        if (text && content) {
          buttons.push({ text, content });
        }
      }

      // Update cache immediately
      const cacheKey = `buttons_${state.currentEditingSection}`;
      CacheManager.set(cacheKey, buttons);
      state.quickButtonSettings[state.currentEditingSection] = buttons;

      google.script.run
        .withSuccessHandler(() => {
          // Update UI immediately
          const container = document.getElementById(state.currentEditingSection + 'Buttons');
          if (container) {
            UIManager.renderPillButtons(container, buttons, state.currentEditingSection);
          }

          this.close();
          UIManager.showMessage('Quick buttons saved successfully!', 'success');
        })
        .withFailureHandler(error => {
          UIManager.showMessage('Failed to save quick buttons', 'error');
          console.error(error);
        })
        .saveQuickButtonSettings(state.currentEditingSection, buttons);
    }
  };

  // Public API
  const QuickNotesAPI = {
    init() {
      // Load initial data
      Promise.all([
        DataManager.loadClientInfo(),
        DataManager.loadSavedNotes(),
        QuickButtons.loadAll()
      ]).then(() => {
        console.log('Quick Notes initialized successfully');
      }).catch(error => {
        console.error('Quick Notes initialization error:', error);
      });

      // Setup features
      AutoSave.setup();
      KeyboardShortcuts.setup();

      // Setup modal close on outside click
      document.addEventListener('click', event => {
        const modal = document.getElementById('quickButtonModal');
        if (event.target === modal) {
          Modal.close();
        }
      });
    },

    // Public methods for HTML onclick handlers
    saveQuickNotes: (isAutoSave = false) => DataManager.saveNotes(isAutoSave),

    clearQuickNotes() {
      if (confirm('Are you sure you want to clear all notes?')) {
        CONFIG.SECTIONS.forEach(field => {
          const element = document.getElementById(field);
          if (element) element.value = '';
        });
        DataManager.saveNotes(false);
      }
    },

    loadSavedNotes: () => DataManager.loadSavedNotes(),

    openSessionRecap() {
      DataManager.saveNotes(false).then(() => {
        google.script.run
          .withSuccessHandler(() => {
            UIManager.showMessage('Session recap dialog opened', 'success');
          })
          .withFailureHandler(() => {
            UIManager.showMessage('Failed to open recap dialog', 'error');
          })
          .showRecapDialog();
      });
    },

    openQuickNotesSettings() {
      google.script.run
        .withSuccessHandler(() => {
          UIManager.showMessage('Settings dialog opened', 'success');
        })
        .withFailureHandler(() => {
          UIManager.showMessage('Failed to open settings', 'error');
        })
        .showQuickNotesSettings();
    },

    openQuickButtonSettings: (section) => Modal.open(section),
    closeQuickButtonModal: () => Modal.close(),
    saveQuickButtons: () => Modal.save(),

    // Utility method to refresh all data
    refresh() {
      state.cache = {};
      this.init();
      UIManager.showMessage('Quick Notes refreshed', 'info');
    }
  };

  // Initialize when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => QuickNotesAPI.init());
  } else {
    QuickNotesAPI.init();
  }

  // Export to global scope for HTML onclick handlers
  window.QuickNotes = QuickNotesAPI;
  window.saveQuickNotes = QuickNotesAPI.saveQuickNotes;
  window.clearQuickNotes = QuickNotesAPI.clearQuickNotes;
  window.loadSavedNotes = QuickNotesAPI.loadSavedNotes;
  window.openSessionRecap = QuickNotesAPI.openSessionRecap;
  window.openQuickNotesSettings = QuickNotesAPI.openQuickNotesSettings;
  window.openQuickButtonSettings = QuickNotesAPI.openQuickButtonSettings;
  window.closeQuickButtonModal = QuickNotesAPI.closeQuickButtonModal;
  window.saveQuickButtons = QuickNotesAPI.saveQuickButtons;

})();