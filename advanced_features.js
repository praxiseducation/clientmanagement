/**
 * Advanced Features Module - Second Beta Iteration
 * Core system enhancements for production-ready application
 */

// ============================================================================
// DATA LOADING STRATEGIES
// ============================================================================

/**
 * Advanced data loading with pagination, lazy loading, and caching
 */
const DataLoadingManager = {
  // Cache management
  cache: new Map(),
  cacheExpiry: new Map(),
  CACHE_DURATION: 5 * 60 * 1000, // 5 minutes
  
  // Pagination state
  pageSize: 50,
  currentPage: 1,
  totalPages: 1,
  totalItems: 0,
  
  // Loading states
  isLoading: false,
  loadedData: new Set(),
  
  /**
   * Get paginated client data with caching
   */
  async getPaginatedClients(page = 1, pageSize = 50, filters = {}) {
    const cacheKey = `clients_${page}_${pageSize}_${JSON.stringify(filters)}`;
    
    // Check cache first
    if (this.isCacheValid(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    try {
      this.isLoading = true;
      
      const response = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .getPaginatedClientList(page, pageSize, filters);
      });
      
      // Cache the response
      this.setCache(cacheKey, response);
      
      // Update pagination state
      this.currentPage = page;
      this.pageSize = pageSize;
      this.totalPages = response.totalPages || 1;
      this.totalItems = response.totalItems || 0;
      
      return response;
      
    } finally {
      this.isLoading = false;
    }
  },
  
  /**
   * Lazy load client details on demand
   */
  async getClientDetails(clientName) {
    const cacheKey = `client_details_${clientName}`;
    
    if (this.isCacheValid(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    // Only load if not already loaded
    if (this.loadedData.has(clientName)) {
      return this.cache.get(cacheKey);
    }
    
    try {
      const response = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .getEnhancedClientDetails(clientName);
      });
      
      this.setCache(cacheKey, response);
      this.loadedData.add(clientName);
      
      return response;
      
    } catch (error) {
      console.error(`Error loading client details for ${clientName}:`, error);
      throw error;
    }
  },
  
  /**
   * Debounced search with intelligent caching
   */
  async searchClients(query, options = {}) {
    const { 
      debounceMs = 300, 
      includeInactive = false,
      fuzzySearch = true,
      maxResults = 100 
    } = options;
    
    // Clear previous search timeout
    if (this.searchTimeout) {
      clearTimeout(this.searchTimeout);
    }
    
    return new Promise((resolve) => {
      this.searchTimeout = setTimeout(async () => {
        const cacheKey = `search_${query}_${JSON.stringify(options)}`;
        
        if (this.isCacheValid(cacheKey)) {
          resolve(this.cache.get(cacheKey));
          return;
        }
        
        try {
          const response = await new Promise((resolveSearch, reject) => {
            google.script.run
              .withSuccessHandler(resolveSearch)
              .withFailureHandler(reject)
              .searchClientsAdvanced(query, options);
          });
          
          this.setCache(cacheKey, response);
          resolve(response);
          
        } catch (error) {
          console.error('Search error:', error);
          resolve({ clients: [], error: error.message });
        }
      }, debounceMs);
    });
  },
  
  /**
   * Preload likely-to-be-accessed data
   */
  async preloadCriticalData() {
    const preloadTasks = [
      this.getPaginatedClients(1, this.pageSize), // First page
      this.getActiveClientsCount(),
      this.getRecentSessions(),
      this.getUserPreferences()
    ];
    
    try {
      const results = await Promise.allSettled(preloadTasks);
      
      // Log any failed preload tasks
      results.forEach((result, index) => {
        if (result.status === 'rejected') {
          console.warn(`Preload task ${index} failed:`, result.reason);
        }
      });
      
      return results.filter(r => r.status === 'fulfilled').map(r => r.value);
      
    } catch (error) {
      console.error('Error during critical data preload:', error);
      return [];
    }
  },
  
  // Cache management methods
  isCacheValid(key) {
    if (!this.cache.has(key)) return false;
    
    const expiry = this.cacheExpiry.get(key);
    if (Date.now() > expiry) {
      this.cache.delete(key);
      this.cacheExpiry.delete(key);
      return false;
    }
    
    return true;
  },
  
  setCache(key, data) {
    this.cache.set(key, data);
    this.cacheExpiry.set(key, Date.now() + this.CACHE_DURATION);
  },
  
  clearCache() {
    this.cache.clear();
    this.cacheExpiry.clear();
    this.loadedData.clear();
  },
  
  // Helper methods for server-side functions
  async getActiveClientsCount() {
    const cacheKey = 'active_clients_count';
    
    if (this.isCacheValid(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    const response = await new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .getActiveClientsCount();
    });
    
    this.setCache(cacheKey, response);
    return response;
  },
  
  async getRecentSessions() {
    const cacheKey = 'recent_sessions';
    
    if (this.isCacheValid(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    const response = await new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .getRecentSessionsData();
    });
    
    this.setCache(cacheKey, response);
    return response;
  },
  
  async getUserPreferences() {
    const cacheKey = 'user_preferences';
    
    if (this.isCacheValid(cacheKey)) {
      return this.cache.get(cacheKey);
    }
    
    const response = await new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .getUserPreferences();
    });
    
    this.setCache(cacheKey, response);
    return response;
  }
};

// ============================================================================
// BACKGROUND OPERATIONS
// ============================================================================

/**
 * Background operations manager for auto-sync and offline support
 */
const BackgroundOperations = {
  // Auto-sync state
  autoSyncEnabled: true,
  syncInterval: 30000, // 30 seconds
  syncTimeout: null,
  
  // Offline support
  isOnline: navigator.onLine,
  offlineQueue: [],
  offlineData: new Map(),
  
  // Smart caching
  preloadQueue: [],
  isPreloading: false,
  
  /**
   * Initialize background operations
   */
  init() {
    // Monitor online/offline status
    window.addEventListener('online', () => this.handleOnline());
    window.addEventListener('offline', () => this.handleOffline());
    
    // Start auto-sync if enabled
    if (this.autoSyncEnabled) {
      this.startAutoSync();
    }
    
    // Start preloading
    this.startSmartPreloading();
    
    // Setup periodic cleanup
    setInterval(() => this.cleanup(), 300000); // 5 minutes
  },
  
  /**
   * Auto-sync background operations
   */
  startAutoSync() {
    if (this.syncTimeout) {
      clearTimeout(this.syncTimeout);
    }
    
    this.syncTimeout = setTimeout(async () => {
      if (this.isOnline) {
        await this.performAutoSync();
      }
      
      // Schedule next sync
      this.startAutoSync();
    }, this.syncInterval);
  },
  
  async performAutoSync() {
    try {
      // Sync any pending offline changes
      if (this.offlineQueue.length > 0) {
        await this.syncOfflineChanges();
      }
      
      // Refresh cache for active data
      await this.refreshActiveCache();
      
      // Background validation
      await this.validateDataIntegrity();
      
    } catch (error) {
      console.warn('Auto-sync error:', error);
    }
  },
  
  /**
   * Smart caching and preloading
   */
  startSmartPreloading() {
    // Preload data based on user patterns
    this.schedulePreload('recent_clients', () => DataLoadingManager.getPaginatedClients(1));
    this.schedulePreload('user_preferences', () => DataLoadingManager.getUserPreferences());
    
    // Process preload queue
    this.processPreloadQueue();
  },
  
  schedulePreload(key, loadFunction) {
    this.preloadQueue.push({ key, loadFunction, priority: 1 });
  },
  
  async processPreloadQueue() {
    if (this.isPreloading || this.preloadQueue.length === 0) {
      return;
    }
    
    this.isPreloading = true;
    
    try {
      // Sort by priority
      this.preloadQueue.sort((a, b) => b.priority - a.priority);
      
      // Process items with exponential backoff
      for (const item of this.preloadQueue) {
        if (!this.isOnline) break;
        
        try {
          await item.loadFunction();
          await this.delay(100); // Small delay between preloads
        } catch (error) {
          console.warn(`Preload failed for ${item.key}:`, error);
        }
      }
      
      this.preloadQueue = [];
      
    } finally {
      this.isPreloading = false;
    }
  },
  
  /**
   * Offline support
   */
  handleOffline() {
    this.isOnline = false;
    
    // Notify user
    this.dispatchEvent('offline', { 
      message: 'You are now offline. Changes will be saved locally.' 
    });
    
    // Switch to offline mode
    this.enableOfflineMode();
  },
  
  handleOnline() {
    this.isOnline = true;
    
    // Notify user
    this.dispatchEvent('online', { 
      message: 'You are back online. Syncing changes...' 
    });
    
    // Sync offline changes
    this.syncOfflineChanges();
  },
  
  enableOfflineMode() {
    // Load offline data from localStorage
    this.loadOfflineData();
  },
  
  saveOfflineChange(key, data) {
    this.offlineQueue.push({
      key,
      data,
      timestamp: Date.now(),
      type: 'update'
    });
    
    // Save to localStorage as backup
    localStorage.setItem(`offline_${key}`, JSON.stringify({
      data,
      timestamp: Date.now()
    }));
  },
  
  loadOfflineData() {
    // Load any offline changes from localStorage
    const keys = Object.keys(localStorage).filter(k => k.startsWith('offline_'));
    
    keys.forEach(key => {
      try {
        const item = JSON.parse(localStorage.getItem(key));
        this.offlineData.set(key.replace('offline_', ''), item);
      } catch (error) {
        console.warn('Error loading offline data:', error);
      }
    });
  },
  
  async syncOfflineChanges() {
    if (this.offlineQueue.length === 0) return;
    
    const changes = [...this.offlineQueue];
    this.offlineQueue = [];
    
    try {
      for (const change of changes) {
        await this.syncSingleChange(change);
        
        // Remove from localStorage after successful sync
        localStorage.removeItem(`offline_${change.key}`);
      }
      
      this.dispatchEvent('sync_complete', {
        message: 'All offline changes synced successfully',
        changesCount: changes.length
      });
      
    } catch (error) {
      // Re-add failed changes to queue
      this.offlineQueue.unshift(...changes);
      
      this.dispatchEvent('sync_error', {
        message: 'Some changes failed to sync',
        error: error.message
      });
    }
  },
  
  async syncSingleChange(change) {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .syncOfflineChange(change);
    });
  },
  
  // Utility methods
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  },
  
  dispatchEvent(type, data) {
    window.dispatchEvent(new CustomEvent(`background_${type}`, { detail: data }));
  },
  
  async refreshActiveCache() {
    // Refresh cache for frequently accessed data
    const activeKeys = ['active_clients_count', 'recent_sessions', 'user_preferences'];
    
    for (const key of activeKeys) {
      try {
        DataLoadingManager.cache.delete(key);
        DataLoadingManager.cacheExpiry.delete(key);
      } catch (error) {
        console.warn(`Error refreshing cache for ${key}:`, error);
      }
    }
  },
  
  async validateDataIntegrity() {
    // Basic data validation
    try {
      const response = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .validateDataIntegrity();
      });
      
      if (!response.isValid) {
        console.warn('Data integrity issues detected:', response.issues);
      }
      
    } catch (error) {
      console.warn('Data validation error:', error);
    }
  },
  
  cleanup() {
    // Clean up expired data
    const now = Date.now();
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours
    
    // Clean localStorage
    Object.keys(localStorage).forEach(key => {
      if (key.startsWith('offline_')) {
        try {
          const item = JSON.parse(localStorage.getItem(key));
          if (now - item.timestamp > maxAge) {
            localStorage.removeItem(key);
          }
        } catch (error) {
          localStorage.removeItem(key); // Remove corrupted items
        }
      }
    });
    
    // Clean offline data map
    for (const [key, item] of this.offlineData.entries()) {
      if (now - item.timestamp > maxAge) {
        this.offlineData.delete(key);
      }
    }
  }
};

// ============================================================================
// GRACEFUL ERROR RECOVERY
// ============================================================================

/**
 * Comprehensive error handling and recovery system
 */
const ErrorRecoveryManager = {
  // Retry configuration
  maxRetries: 3,
  baseDelay: 1000,
  maxDelay: 10000,
  
  // Circuit breaker state
  circuitBreakers: new Map(),
  
  // Error tracking
  errorHistory: [],
  maxErrorHistory: 100,
  
  // Fallback handlers
  fallbackHandlers: new Map(),
  
  /**
   * Execute function with retry and circuit breaker logic
   */
  async executeWithRetry(fn, options = {}) {
    const {
      maxRetries = this.maxRetries,
      baseDelay = this.baseDelay,
      exponential = true,
      circuitBreakerKey = null,
      fallbackFn = null
    } = options;
    
    // Check circuit breaker
    if (circuitBreakerKey && this.isCircuitBreakerOpen(circuitBreakerKey)) {
      if (fallbackFn) {
        return await fallbackFn();
      }
      throw new Error(`Circuit breaker is open for ${circuitBreakerKey}`);
    }
    
    let lastError;
    
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const result = await fn();
        
        // Reset circuit breaker on success
        if (circuitBreakerKey) {
          this.resetCircuitBreaker(circuitBreakerKey);
        }
        
        return result;
        
      } catch (error) {
        lastError = error;
        
        // Track error
        this.recordError(error, circuitBreakerKey);
        
        // Don't retry on the last attempt
        if (attempt === maxRetries) {
          break;
        }
        
        // Calculate delay
        const delay = exponential 
          ? Math.min(baseDelay * Math.pow(2, attempt), this.maxDelay)
          : baseDelay;
        
        console.warn(`Attempt ${attempt + 1} failed, retrying in ${delay}ms:`, error);
        
        await this.delay(delay);
      }
    }
    
    // Update circuit breaker on final failure
    if (circuitBreakerKey) {
      this.recordCircuitBreakerFailure(circuitBreakerKey);
    }
    
    // Try fallback
    if (fallbackFn) {
      try {
        return await fallbackFn();
      } catch (fallbackError) {
        console.error('Fallback also failed:', fallbackError);
      }
    }
    
    throw lastError;
  },
  
  /**
   * Circuit breaker implementation
   */
  isCircuitBreakerOpen(key) {
    const breaker = this.circuitBreakers.get(key);
    if (!breaker) return false;
    
    const now = Date.now();
    
    // Check if circuit breaker should reset
    if (breaker.state === 'open' && now > breaker.nextAttempt) {
      breaker.state = 'half-open';
      breaker.failures = 0;
    }
    
    return breaker.state === 'open';
  },
  
  recordCircuitBreakerFailure(key) {
    let breaker = this.circuitBreakers.get(key);
    
    if (!breaker) {
      breaker = {
        failures: 0,
        state: 'closed',
        nextAttempt: 0,
        threshold: 5,
        timeout: 60000 // 1 minute
      };
      this.circuitBreakers.set(key, breaker);
    }
    
    breaker.failures++;
    
    if (breaker.failures >= breaker.threshold) {
      breaker.state = 'open';
      breaker.nextAttempt = Date.now() + breaker.timeout;
    }
  },
  
  resetCircuitBreaker(key) {
    const breaker = this.circuitBreakers.get(key);
    if (breaker) {
      breaker.failures = 0;
      breaker.state = 'closed';
      breaker.nextAttempt = 0;
    }
  },
  
  /**
   * Error tracking and analysis
   */
  recordError(error, context = null) {
    const errorRecord = {
      message: error.message,
      stack: error.stack,
      context,
      timestamp: Date.now(),
      userAgent: navigator.userAgent,
      url: window.location.href
    };
    
    this.errorHistory.push(errorRecord);
    
    // Maintain history size
    if (this.errorHistory.length > this.maxErrorHistory) {
      this.errorHistory.shift();
    }
    
    // Send to error reporting service
    this.sendErrorReport(errorRecord);
  },
  
  async sendErrorReport(errorRecord) {
    // Only send critical errors or patterns
    if (this.shouldReportError(errorRecord)) {
      try {
        await new Promise((resolve, reject) => {
          google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(reject)
            .reportError(errorRecord);
        });
      } catch (reportError) {
        // Don't let error reporting break the app
        console.warn('Failed to report error:', reportError);
      }
    }
  },
  
  shouldReportError(errorRecord) {
    // Report if:
    // 1. It's a new type of error
    // 2. Error frequency exceeds threshold
    // 3. Critical error type
    
    const recentSimilar = this.errorHistory.filter(e => 
      e.message === errorRecord.message && 
      Date.now() - e.timestamp < 300000 // 5 minutes
    ).length;
    
    return recentSimilar === 1 || // First occurrence
           recentSimilar % 10 === 0 || // Every 10th occurrence
           this.isCriticalError(errorRecord);
  },
  
  isCriticalError(errorRecord) {
    const criticalKeywords = [
      'permission denied',
      'quota exceeded',
      'service unavailable',
      'network error',
      'data corruption'
    ];
    
    return criticalKeywords.some(keyword => 
      errorRecord.message.toLowerCase().includes(keyword)
    );
  },
  
  /**
   * Undo functionality
   */
  undoStack: [],
  maxUndoStackSize: 10,
  
  async executeWithUndo(operation, undoOperation, description) {
    try {
      const result = await operation();
      
      // Add to undo stack
      this.undoStack.push({
        undo: undoOperation,
        description,
        timestamp: Date.now()
      });
      
      // Maintain stack size
      if (this.undoStack.length > this.maxUndoStackSize) {
        this.undoStack.shift();
      }
      
      return result;
      
    } catch (error) {
      console.error('Operation failed:', error);
      throw error;
    }
  },
  
  async undoLastOperation() {
    if (this.undoStack.length === 0) {
      throw new Error('No operations to undo');
    }
    
    const operation = this.undoStack.pop();
    
    try {
      await operation.undo();
      return operation.description;
    } catch (error) {
      // Put it back on the stack if undo fails
      this.undoStack.push(operation);
      throw error;
    }
  },
  
  getUndoStackSize() {
    return this.undoStack.length;
  },
  
  getLastOperationDescription() {
    if (this.undoStack.length === 0) return null;
    return this.undoStack[this.undoStack.length - 1].description;
  },
  
  // Utility methods
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  },
  
  /**
   * Register fallback handlers for specific operations
   */
  registerFallbackHandler(operationKey, handler) {
    this.fallbackHandlers.set(operationKey, handler);
  },
  
  getFallbackHandler(operationKey) {
    return this.fallbackHandlers.get(operationKey);
  }
};

// ============================================================================
// CONFLICT RESOLUTION SYSTEM
// ============================================================================

/**
 * Comprehensive conflict detection and resolution
 */
const ConflictResolver = {
  // Active conflicts
  activeConflicts: new Map(),
  
  // Resolution strategies
  resolutionStrategies: {
    'client_update': 'merge_with_priority',
    'notes_update': 'timestamp_priority',
    'preferences': 'user_choice',
    'system_settings': 'server_priority'
  },
  
  /**
   * Detect conflicts before applying changes
   */
  async detectConflicts(operation, data) {
    const conflictChecks = [
      this.checkTimestampConflicts(operation, data),
      this.checkVersionConflicts(operation, data),
      this.checkConcurrentEdits(operation, data),
      this.checkDataIntegrity(operation, data)
    ];
    
    const results = await Promise.allSettled(conflictChecks);
    
    const conflicts = results
      .filter(result => result.status === 'fulfilled' && result.value)
      .map(result => result.value);
    
    if (conflicts.length > 0) {
      return {
        hasConflicts: true,
        conflicts: conflicts,
        resolution: await this.getResolutionStrategy(operation, conflicts)
      };
    }
    
    return { hasConflicts: false };
  },
  
  async checkTimestampConflicts(operation, data) {
    if (!data.lastModified) return null;
    
    try {
      const serverData = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .getServerTimestamp(operation.entityType, operation.entityId);
      });
      
      if (serverData.lastModified > data.lastModified) {
        return {
          type: 'timestamp_conflict',
          serverTimestamp: serverData.lastModified,
          clientTimestamp: data.lastModified,
          message: 'Server data is newer than client data'
        };
      }
      
    } catch (error) {
      console.warn('Could not check timestamp conflicts:', error);
    }
    
    return null;
  },
  
  async checkVersionConflicts(operation, data) {
    if (!data.version) return null;
    
    try {
      const serverData = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .getServerVersion(operation.entityType, operation.entityId);
      });
      
      if (serverData.version !== data.version) {
        return {
          type: 'version_conflict',
          serverVersion: serverData.version,
          clientVersion: data.version,
          message: 'Version mismatch detected'
        };
      }
      
    } catch (error) {
      console.warn('Could not check version conflicts:', error);
    }
    
    return null;
  },
  
  async checkConcurrentEdits(operation, data) {
    // Check for other active sessions editing the same entity
    try {
      const activeEditors = await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .getActiveEditors(operation.entityType, operation.entityId);
      });
      
      if (activeEditors.length > 1) {
        return {
          type: 'concurrent_edit',
          activeEditors: activeEditors,
          message: 'Multiple users are editing this item'
        };
      }
      
    } catch (error) {
      console.warn('Could not check concurrent edits:', error);
    }
    
    return null;
  },
  
  async checkDataIntegrity(operation, data) {
    // Validate data structure and required fields
    const validationRules = this.getValidationRules(operation.entityType);
    
    for (const rule of validationRules) {
      if (!rule.validate(data)) {
        return {
          type: 'data_integrity',
          rule: rule.name,
          message: rule.message
        };
      }
    }
    
    return null;
  },
  
  getValidationRules(entityType) {
    const rules = {
      'client': [
        {
          name: 'required_name',
          validate: (data) => data.name && data.name.trim().length > 0,
          message: 'Client name is required'
        },
        {
          name: 'valid_email',
          validate: (data) => !data.email || /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email),
          message: 'Invalid email format'
        }
      ],
      'notes': [
        {
          name: 'max_length',
          validate: (data) => !data.content || data.content.length <= 10000,
          message: 'Notes content too long'
        }
      ]
    };
    
    return rules[entityType] || [];
  },
  
  /**
   * Get resolution strategy for conflicts
   */
  async getResolutionStrategy(operation, conflicts) {
    const strategy = this.resolutionStrategies[operation.type] || 'user_choice';
    
    switch (strategy) {
      case 'merge_with_priority':
        return await this.createMergeStrategy(conflicts);
      
      case 'timestamp_priority':
        return {
          strategy: 'timestamp',
          message: 'Use the most recent version'
        };
      
      case 'server_priority':
        return {
          strategy: 'server',
          message: 'Server version takes priority'
        };
      
      case 'user_choice':
      default:
        return {
          strategy: 'user_choice',
          message: 'User must choose resolution',
          options: await this.getResolutionOptions(conflicts)
        };
    }
  },
  
  async createMergeStrategy(conflicts) {
    // Intelligent merge strategy based on conflict types
    const mergeRules = [];
    
    for (const conflict of conflicts) {
      switch (conflict.type) {
        case 'timestamp_conflict':
          mergeRules.push({
            field: 'lastModified',
            rule: 'use_server'
          });
          break;
        
        case 'concurrent_edit':
          mergeRules.push({
            field: 'all',
            rule: 'user_review_required'
          });
          break;
      }
    }
    
    return {
      strategy: 'merge',
      rules: mergeRules,
      message: 'Intelligent merge will be applied'
    };
  },
  
  async getResolutionOptions(conflicts) {
    return [
      {
        id: 'keep_local',
        label: 'Keep My Changes',
        description: 'Overwrite server data with your local changes'
      },
      {
        id: 'keep_server',
        label: 'Use Server Version',
        description: 'Discard local changes and use server data'
      },
      {
        id: 'merge_manual',
        label: 'Merge Manually',
        description: 'Review and merge changes manually'
      }
    ];
  },
  
  /**
   * Apply conflict resolution
   */
  async resolveConflict(conflictId, resolution) {
    const conflict = this.activeConflicts.get(conflictId);
    if (!conflict) {
      throw new Error('Conflict not found');
    }
    
    try {
      let result;
      
      switch (resolution.strategy) {
        case 'keep_local':
          result = await this.applyLocalChanges(conflict);
          break;
        
        case 'keep_server':
          result = await this.fetchServerVersion(conflict);
          break;
        
        case 'merge_manual':
          result = await this.performManualMerge(conflict, resolution.mergedData);
          break;
        
        case 'merge':
          result = await this.performAutomaticMerge(conflict, resolution.rules);
          break;
        
        default:
          throw new Error('Unknown resolution strategy');
      }
      
      // Remove from active conflicts
      this.activeConflicts.delete(conflictId);
      
      // Notify success
      this.dispatchEvent('conflict_resolved', {
        conflictId,
        strategy: resolution.strategy,
        result
      });
      
      return result;
      
    } catch (error) {
      this.dispatchEvent('conflict_resolution_failed', {
        conflictId,
        error: error.message
      });
      
      throw error;
    }
  },
  
  async applyLocalChanges(conflict) {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .forceUpdateWithLocalData(conflict.operation, conflict.data);
    });
  },
  
  async fetchServerVersion(conflict) {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .getLatestServerData(conflict.operation.entityType, conflict.operation.entityId);
    });
  },
  
  async performManualMerge(conflict, mergedData) {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        .saveMergedData(conflict.operation, mergedData);
    });
  },
  
  async performAutomaticMerge(conflict, rules) {
    const serverData = await this.fetchServerVersion(conflict);
    const localData = conflict.data;
    
    const mergedData = { ...serverData };
    
    for (const rule of rules) {
      switch (rule.rule) {
        case 'use_server':
          // Already in mergedData
          break;
        
        case 'use_local':
          if (rule.field === 'all') {
            Object.assign(mergedData, localData);
          } else {
            mergedData[rule.field] = localData[rule.field];
          }
          break;
        
        case 'merge_arrays':
          if (Array.isArray(mergedData[rule.field]) && Array.isArray(localData[rule.field])) {
            mergedData[rule.field] = [...new Set([...mergedData[rule.field], ...localData[rule.field]])];
          }
          break;
      }
    }
    
    return await this.performManualMerge(conflict, mergedData);
  },
  
  /**
   * Create conflict resolution UI
   */
  showConflictResolutionDialog(conflicts, resolutionOptions) {
    // This would be implemented as a dialog showing the conflicts
    // and allowing the user to choose resolution strategy
    
    return new Promise((resolve) => {
      // Implementation would show a modal dialog
      // For now, return a default resolution
      resolve({
        strategy: 'user_choice',
        option: 'keep_local'
      });
    });
  },
  
  // Utility methods
  dispatchEvent(type, data) {
    window.dispatchEvent(new CustomEvent(`conflict_${type}`, { detail: data }));
  }
};

// ============================================================================
// STRUCTURED ERROR REPORTING
// ============================================================================

/**
 * Comprehensive error reporting and analytics system
 */
const ErrorReportingSystem = {
  // Configuration
  config: {
    enabled: true,
    batchSize: 10,
    batchTimeout: 30000, // 30 seconds
    maxErrorsPerSession: 100,
    enableUserFeedback: true
  },
  
  // Error collection
  errorBatch: [],
  errorCount: 0,
  batchTimer: null,
  
  // User context
  userSession: {
    sessionId: this.generateSessionId(),
    startTime: Date.now(),
    userAgent: navigator.userAgent,
    viewport: {
      width: window.innerWidth,
      height: window.innerHeight
    },
    features: {
      localStorage: typeof Storage !== 'undefined',
      webWorkers: typeof Worker !== 'undefined',
      offlineSupport: 'serviceWorker' in navigator
    }
  },
  
  /**
   * Initialize error reporting
   */
  init() {
    if (!this.config.enabled) return;
    
    // Global error handlers
    window.addEventListener('error', (event) => {
      this.captureError(event.error, {
        type: 'javascript_error',
        filename: event.filename,
        lineno: event.lineno,
        colno: event.colno
      });
    });
    
    window.addEventListener('unhandledrejection', (event) => {
      this.captureError(event.reason, {
        type: 'unhandled_promise_rejection'
      });
    });
    
    // Custom error boundary for React-like components
    this.setupErrorBoundary();
    
    // Start batch timer
    this.startBatchTimer();
  },
  
  /**
   * Capture and process errors
   */
  captureError(error, context = {}) {
    if (!this.config.enabled || this.errorCount >= this.config.maxErrorsPerSession) {
      return;
    }
    
    const errorReport = this.createErrorReport(error, context);
    
    // Add to batch
    this.errorBatch.push(errorReport);
    this.errorCount++;
    
    // Send immediately for critical errors
    if (this.isCriticalError(error)) {
      this.sendBatch();
    }
    
    // Send if batch is full
    if (this.errorBatch.length >= this.config.batchSize) {
      this.sendBatch();
    }
    
    // Show user feedback for user-facing errors
    if (this.config.enableUserFeedback && this.isUserFacingError(error)) {
      this.showErrorFeedback(errorReport);
    }
  },
  
  createErrorReport(error, context = {}) {
    return {
      // Error details
      message: error.message || String(error),
      stack: error.stack,
      name: error.name,
      
      // Context
      context: {
        ...context,
        timestamp: Date.now(),
        url: window.location.href,
        userAgent: navigator.userAgent,
        viewport: {
          width: window.innerWidth,
          height: window.innerHeight
        }
      },
      
      // User session
      session: {
        ...this.userSession,
        timeInSession: Date.now() - this.userSession.startTime
      },
      
      // Application state
      appState: this.captureApplicationState(),
      
      // Performance data
      performance: this.capturePerformanceData(),
      
      // Severity
      severity: this.calculateSeverity(error, context),
      
      // Unique ID
      errorId: this.generateErrorId()
    };
  },
  
  captureApplicationState() {
    return {
      // Data loading states
      isLoading: DataLoadingManager.isLoading,
      cacheSize: DataLoadingManager.cache.size,
      
      // Background operations
      isOnline: BackgroundOperations.isOnline,
      offlineQueueSize: BackgroundOperations.offlineQueue.length,
      
      // Active conflicts
      activeConflictsCount: ConflictResolver.activeConflicts.size,
      
      // Memory usage (if available)
      memoryUsage: this.getMemoryUsage()
    };
  },
  
  capturePerformanceData() {
    const perfData = {};
    
    // Performance timing
    if (performance.timing) {
      const timing = performance.timing;
      perfData.pageLoadTime = timing.loadEventEnd - timing.navigationStart;
      perfData.domReadyTime = timing.domContentLoadedEventEnd - timing.navigationStart;
      perfData.networkTime = timing.responseEnd - timing.requestStart;
    }
    
    // Performance entries
    if (performance.getEntriesByType) {
      const navigationEntries = performance.getEntriesByType('navigation');
      if (navigationEntries.length > 0) {
        const nav = navigationEntries[0];
        perfData.dnsLookupTime = nav.domainLookupEnd - nav.domainLookupStart;
        perfData.tcpConnectTime = nav.connectEnd - nav.connectStart;
        perfData.serverResponseTime = nav.responseStart - nav.requestStart;
      }
    }
    
    return perfData;
  },
  
  getMemoryUsage() {
    if (performance.memory) {
      return {
        usedJSHeapSize: performance.memory.usedJSHeapSize,
        totalJSHeapSize: performance.memory.totalJSHeapSize,
        jsHeapSizeLimit: performance.memory.jsHeapSizeLimit
      };
    }
    return null;
  },
  
  calculateSeverity(error, context) {
    // Critical errors
    if (this.isCriticalError(error)) return 'critical';
    
    // High severity
    if (context.type === 'javascript_error' || 
        error.message.includes('Permission denied') ||
        error.message.includes('Quota exceeded')) {
      return 'high';
    }
    
    // Medium severity
    if (context.type === 'unhandled_promise_rejection' ||
        error.message.includes('Network error')) {
      return 'medium';
    }
    
    // Low severity (default)
    return 'low';
  },
  
  isCriticalError(error) {
    const criticalPatterns = [
      /corruption/i,
      /fatal/i,
      /security/i,
      /unauthorized/i,
      /service.*unavailable/i
    ];
    
    return criticalPatterns.some(pattern => 
      pattern.test(error.message || String(error))
    );
  },
  
  isUserFacingError(error) {
    const userFacingPatterns = [
      /network/i,
      /connection/i,
      /timeout/i,
      /failed to (save|load|sync)/i
    ];
    
    return userFacingPatterns.some(pattern => 
      pattern.test(error.message || String(error))
    );
  },
  
  /**
   * Batch processing and sending
   */
  startBatchTimer() {
    if (this.batchTimer) {
      clearTimeout(this.batchTimer);
    }
    
    this.batchTimer = setTimeout(() => {
      if (this.errorBatch.length > 0) {
        this.sendBatch();
      }
      this.startBatchTimer();
    }, this.config.batchTimeout);
  },
  
  async sendBatch() {
    if (this.errorBatch.length === 0) return;
    
    const batch = [...this.errorBatch];
    this.errorBatch = [];
    
    try {
      await new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          .submitErrorBatch({
            errors: batch,
            session: this.userSession,
            timestamp: Date.now()
          });
      });
      
    } catch (error) {
      // Don't let error reporting break the app
      console.warn('Failed to send error batch:', error);
      
      // Re-add to batch if it was a temporary failure
      if (batch.length < 50) { // Prevent infinite growth
        this.errorBatch.unshift(...batch.slice(0, 10)); // Only re-add first 10
      }
    }
  },
  
  /**
   * User feedback system
   */
  showErrorFeedback(errorReport) {
    // Create a non-intrusive notification
    const notification = this.createErrorNotification(errorReport);
    document.body.appendChild(notification);
    
    // Auto-remove after delay
    setTimeout(() => {
      if (notification.parentNode) {
        notification.parentNode.removeChild(notification);
      }
    }, 10000);
  },
  
  createErrorNotification(errorReport) {
    const div = document.createElement('div');
    div.className = 'error-notification';
    div.innerHTML = \`
      <div class="error-notification-content">
        <div class="error-icon">⚠️</div>
        <div class="error-message">
          <strong>Something went wrong</strong>
          <p>We've automatically reported this issue. You can continue using the app.</p>
        </div>
        <div class="error-actions">
          <button onclick="this.parentElement.parentElement.parentElement.remove()">Dismiss</button>
          <button onclick="window.ErrorReportingSystem.showFeedbackDialog('\${errorReport.errorId}')">Provide Feedback</button>
        </div>
      </div>
    \`;
    
    // Add styles
    div.style.cssText = \`
      position: fixed;
      top: 20px;
      right: 20px;
      background: #fff;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.15);
      padding: 16px;
      max-width: 400px;
      z-index: 10000;
      font-family: -apple-system, BlinkMacSystemFont, sans-serif;
    \`;
    
    return div;
  },
  
  showFeedbackDialog(errorId) {
    // Implementation would show a feedback dialog
    // For now, just log the request
    console.log('Feedback requested for error:', errorId);
  },
  
  // Utility methods
  generateSessionId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  },
  
  generateErrorId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
  },
  
  setupErrorBoundary() {
    // Set up error boundary for catching React-like component errors
    const originalConsoleError = console.error;
    console.error = (...args) => {
      // Check if this looks like a React error boundary error
      const errorMessage = args[0];
      if (typeof errorMessage === 'string' && 
          errorMessage.includes('React') || 
          errorMessage.includes('component')) {
        
        this.captureError(new Error(errorMessage), {
          type: 'component_error',
          args: args
        });
      }
      
      // Call original console.error
      originalConsoleError.apply(console, args);
    };
  }
};

// ============================================================================
// INITIALIZATION AND EXPORT
// ============================================================================

/**
 * Initialize all advanced features
 */
function initializeAdvancedFeatures() {
  // Initialize in order
  ErrorReportingSystem.init();
  BackgroundOperations.init();
  
  // Preload critical data
  DataLoadingManager.preloadCriticalData();
  
  // Set up global error handling
  window.addEventListener('unhandledrejection', (event) => {
    ErrorRecoveryManager.recordError(event.reason, 'unhandled_promise');
  });
  
  // Export to global scope for use in dialogs
  window.DataLoadingManager = DataLoadingManager;
  window.BackgroundOperations = BackgroundOperations;
  window.ErrorRecoveryManager = ErrorRecoveryManager;
  window.ConflictResolver = ConflictResolver;
  window.ErrorReportingSystem = ErrorReportingSystem;
  
  console.log('Advanced features initialized');
}

// Auto-initialize if in browser environment
if (typeof window !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeAdvancedFeatures);
  } else {
    initializeAdvancedFeatures();
  }
}