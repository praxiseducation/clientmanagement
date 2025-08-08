# API Documentation

## Core Functions

### Client Management

#### `createClientSheet(clientData)`
Creates a new client sheet and adds to the master list.

**Parameters:**
- `clientData` (Object): Client information
  - `name` (String): Client name *required*
  - `services` (String): Service type
  - `parentEmail` (String): Parent email address
  - `emailAddressees` (String): Additional email recipients
  - `dashboardLink` (String): Dashboard URL
  - `meetingNotesLink` (String): Meeting notes URL
  - `scoreReportLink` (String): Score report URL

**Returns:** Object with success status and client data

**Example:**
```javascript
createClientSheet({
  name: "John Doe",
  services: "SAT Prep",
  parentEmail: "parent@email.com"
});
```

#### `updateClientDetails(sheetName, details)`
Updates client information in the unified data store.

**Parameters:**
- `sheetName` (String): Client sheet name
- `details` (Object): Updated client details
  - `isActive` (Boolean): Active status
  - `dashboardLink` (String): Dashboard URL
  - `meetingNotesLink` (String): Meeting notes URL
  - `scoreReportLink` (String): Score report URL
  - `parentEmail` (String): Parent email
  - `emailAddressees` (String): Additional emails

**Returns:** Object with success status

---

### Data Retrieval

#### `getPaginatedClientList(page, pageSize, filters)`
Retrieves paginated client list with optional filtering.

**Parameters:**
- `page` (Number): Page number (default: 1)
- `pageSize` (Number): Items per page (default: 50)
- `filters` (Object): Optional filters
  - `search` (String): Search term
  - `isActive` (Boolean): Active status filter
  - `serviceType` (String): Service type filter

**Returns:** Object with clients array and pagination metadata

#### `searchClientsAdvanced(query, options)`
Advanced search with fuzzy matching support.

**Parameters:**
- `query` (String): Search query
- `options` (Object): Search options
  - `includeInactive` (Boolean): Include inactive clients
  - `fuzzySearch` (Boolean): Enable fuzzy matching
  - `maxResults` (Number): Maximum results to return
  - `threshold` (Number): Fuzzy match threshold (0-1)

**Returns:** Object with matched clients and metadata

---

### Session Management

#### `saveQuickNotes(notes)`
Saves quick notes for the current session.

**Parameters:**
- `notes` (Object): Notes data with section keys
  - `wins` (String): Wins and achievements
  - `struggles` (String): Challenges faced
  - `parent` (String): Parent communication
  - `next` (String): Next session planning

**Returns:** Object with success status

#### `sendIndividualRecap(sheetName, emailData)`
Sends session recap email for a specific client.

**Parameters:**
- `sheetName` (String): Client sheet name
- `emailData` (Object): Email content data

**Returns:** Object with send status and timestamp

---

### User Preferences

#### `getUserPreferences()`
Retrieves current user preferences.

**Returns:** Object with user preferences
- `quickPills` (Object): Quick pill configurations
- `theme` (String): Theme preference
- `autoSave` (Boolean): Auto-save setting
- `notifications` (Boolean): Notification setting

#### `saveUserPreferences(preferences)`
Saves user preferences.

**Parameters:**
- `preferences` (Object): Updated preferences object

**Returns:** Object with success status

---

## UnifiedClientDataStore

### Client Operations

#### `UnifiedClientDataStore.getClient(clientName)`
Retrieves a specific client's data.

**Parameters:**
- `clientName` (String): Client name

**Returns:** Client object or null

#### `UnifiedClientDataStore.getAllClients()`
Retrieves all clients from the data store.

**Returns:** Array of client objects

#### `UnifiedClientDataStore.getActiveClients()`
Retrieves only active clients.

**Returns:** Array of active client objects

#### `UnifiedClientDataStore.addClient(clientData)`
Adds a new client to the data store.

**Parameters:**
- `clientData` (Object): Client data object

**Returns:** Success status

#### `UnifiedClientDataStore.updateClient(clientName, updates)`
Updates an existing client's data.

**Parameters:**
- `clientName` (String): Client name
- `updates` (Object): Fields to update

**Returns:** Updated client object

#### `UnifiedClientDataStore.upsertClient(clientData)`
Updates existing or creates new client.

**Parameters:**
- `clientData` (Object): Client data with name field

**Returns:** Client object

---

### Temporary Data Management

#### `UnifiedClientDataStore.setTempData(key, data)`
Stores temporary data for dialogs and operations.

**Parameters:**
- `key` (String): Storage key
- `data` (Any): Data to store

**Returns:** void

#### `UnifiedClientDataStore.getTempData(key)`
Retrieves temporary data by key.

**Parameters:**
- `key` (String): Storage key

**Returns:** Stored data or null

#### `UnifiedClientDataStore.deleteTempData(key)`
Removes temporary data.

**Parameters:**
- `key` (String): Storage key

**Returns:** void

---

## Background Operations

### Sync Functions

#### `syncOfflineChange(change)`
Synchronizes offline changes when connection restored.

**Parameters:**
- `change` (Object): Change object
  - `type` (String): Change type
  - `key` (String): Entity key
  - `data` (Object): Change data
  - `timestamp` (Number): Change timestamp

**Returns:** Sync result

#### `validateDataIntegrity()`
Validates data integrity across all stores.

**Returns:** Object with validation results
- `isValid` (Boolean): Overall validity
- `issues` (Array): List of issues found
- `timestamp` (String): Validation timestamp

---

## Error Handling

### Error Reporting

#### `submitErrorBatch(batch)`
Submits batch of errors for logging.

**Parameters:**
- `batch` (Object): Error batch
  - `errors` (Array): Error objects
  - `session` (Object): Session information
  - `timestamp` (Number): Batch timestamp

**Returns:** Submission result

---

## Conflict Resolution

### Conflict Detection

#### `getServerTimestamp(entityType, entityId)`
Gets server timestamp for conflict detection.

**Parameters:**
- `entityType` (String): Entity type (e.g., 'client')
- `entityId` (String): Entity identifier

**Returns:** Object with lastModified timestamp

#### `getServerVersion(entityType, entityId)`
Gets server version for conflict detection.

**Parameters:**
- `entityType` (String): Entity type
- `entityId` (String): Entity identifier

**Returns:** Object with version string

---

## Utility Functions

### Data Helpers

#### `truncateText(text, maxLength)`
Truncates text to specified length.

**Parameters:**
- `text` (String): Text to truncate
- `maxLength` (Number): Maximum length

**Returns:** Truncated string with ellipsis

#### `getRelativeTime(timestamp)`
Converts timestamp to relative time string.

**Parameters:**
- `timestamp` (Number): Unix timestamp

**Returns:** Relative time string (e.g., "2 hours ago")

#### `isValidEmail(email)`
Validates email address format.

**Parameters:**
- `email` (String): Email address

**Returns:** Boolean validation result

---

## Event Triggers

### Custom Events

The system dispatches custom events for various operations:

- `background_online`: Connection restored
- `background_offline`: Connection lost
- `background_sync_complete`: Sync completed
- `background_sync_error`: Sync failed
- `conflict_resolved`: Conflict resolved
- `conflict_resolution_failed`: Resolution failed

**Listening to Events:**
```javascript
window.addEventListener('background_sync_complete', (event) => {
  console.log('Sync completed:', event.detail);
});
```

---

## Constants and Configuration

### System Constants

```javascript
const CONFIG = {
  tutorName: '',          // Set during setup
  tutorEmail: '',         // Set during setup  
  company: 'Smart College',
  companyUrl: 'https://www.smartcollege.com',
  primaryColor: '#003366',
  accentColor: '#00C853',
  useMenuBar: true,
  isConfigured: false
};
```

### Data Loading Configuration

```javascript
DataLoadingManager.config = {
  pageSize: 50,           // Default page size
  cacheExpiry: 300000,    // 5 minutes
  debounceMs: 300,        // Search debounce
  maxRetries: 3           // Retry attempts
};
```

### Background Operations Configuration

```javascript
BackgroundOperations.config = {
  syncInterval: 30000,    // 30 seconds
  offlineQueueMax: 100,   // Max offline operations
  preloadPriority: 1      // Preload priority level
};
```

---

## Migration Functions

### Data Migration

#### `migrateFromLegacyCache()`
Migrates data from legacy cache systems.

**Returns:** Migration result with statistics

#### `cleanupLegacyData()`
Removes deprecated data structures.

**Returns:** Cleanup result

---

## Testing Functions

#### `testUnifiedDataStore()`
Tests the unified data store functionality.

**Returns:** Test results object

---

*For implementation examples and advanced usage, refer to the source code documentation.*