/**
 * Feature Flag System
 * Allows toggling between CardService UI and HTML UI
 * Provides safe rollback mechanism
 *
 * DEBUG_ID: FEATURE_FLAGS_20250109
 */

// ============================================================================
// FEATURE FLAGS
// ============================================================================

const FEATURE_FLAGS = {
  USE_CARDSERVICE: 'use_cardservice_ui',
  CARDSERVICE_VERSION: 'cardservice_version',
  ENABLE_PERFORMANCE_LOGGING: 'enable_performance_logging'
};

/**
 * Gets a feature flag value
 * @param {string} flagName - Name of the flag
 * @param {boolean} defaultValue - Default value if not set
 * @return {boolean} Flag value
 */
function getFeatureFlag(flagName, defaultValue = false) {
  try {
    const props = PropertiesService.getScriptProperties();
    const value = props.getProperty(flagName);

    if (value === null || value === undefined) {
      return defaultValue;
    }

    return value === 'true';
  } catch (error) {
    console.error('Error getting feature flag:', flagName, error);
    return defaultValue;
  }
}

/**
 * Sets a feature flag value
 * @param {string} flagName - Name of the flag
 * @param {boolean} value - Value to set
 */
function setFeatureFlag(flagName, value) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty(flagName, value.toString());
    console.log(`Feature flag set: ${flagName} = ${value}`);
  } catch (error) {
    console.error('Error setting feature flag:', flagName, error);
    throw error;
  }
}

/**
 * Gets all feature flags
 * @return {Object} All feature flags and their values
 */
function getAllFeatureFlags() {
  const flags = {};
  for (const key in FEATURE_FLAGS) {
    flags[FEATURE_FLAGS[key]] = getFeatureFlag(FEATURE_FLAGS[key]);
  }
  return flags;
}

// ============================================================================
// FEATURE FLAG MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Enables CardService UI
 */
function enableCardServiceUI() {
  setFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, true);
  setFeatureFlag(FEATURE_FLAGS.CARDSERVICE_VERSION, true);
  SpreadsheetApp.getUi().alert(
    'CardService UI Enabled',
    'The new CardService UI is now active. Refresh your sidebar to see changes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Disables CardService UI (rollback to HTML)
 */
function disableCardServiceUI() {
  setFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);
  SpreadsheetApp.getUi().alert(
    'HTML UI Restored',
    'Rolled back to classic HTML UI. Refresh your sidebar to see changes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Toggles CardService UI
 */
function toggleCardServiceUI() {
  const current = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE);
  setFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, !current);

  const message = !current
    ? 'CardService UI is now ENABLED'
    : 'CardService UI is now DISABLED (using HTML)';

  SpreadsheetApp.getUi().alert(
    'UI Mode Toggled',
    message + '\n\nRefresh your sidebar to see changes.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows current feature flag status
 */
function showFeatureFlagStatus() {
  const flags = getAllFeatureFlags();
  const useCardService = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE);

  let message = 'Current Feature Flags:\n\n';
  message += `UI Mode: ${useCardService ? 'CardService (Native)' : 'HTML (Classic)'}\n`;
  message += `Performance Logging: ${getFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING) ? 'ON' : 'OFF'}\n`;

  SpreadsheetApp.getUi().alert(
    'Feature Flag Status',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Enables performance logging
 */
function enablePerformanceLogging() {
  setFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING, true);
  SpreadsheetApp.getUi().alert('Performance logging enabled');
}

/**
 * Disables performance logging
 */
function disablePerformanceLogging() {
  setFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING, false);
  SpreadsheetApp.getUi().alert('Performance logging disabled');
}

// ============================================================================
// PERFORMANCE LOGGING
// ============================================================================

/**
 * Logs performance metrics if enabled
 * @param {string} action - Action being measured
 * @param {number} duration - Duration in milliseconds
 */
function logPerformance(action, duration) {
  if (!getFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING)) {
    return;
  }

  const timestamp = new Date().toISOString();
  const uiMode = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE) ? 'CardService' : 'HTML';

  console.log(`[PERFORMANCE] ${timestamp} | ${uiMode} | ${action} | ${duration}ms`);

  // Optionally log to sheet for analysis
  try {
    logPerformanceToSheet(timestamp, uiMode, action, duration);
  } catch (error) {
    // Silent fail - performance logging shouldn't break functionality
    console.error('Error logging performance to sheet:', error);
  }
}

/**
 * Logs performance data to a sheet
 */
function logPerformanceToSheet(timestamp, uiMode, action, duration) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let perfSheet = ss.getSheetByName('Performance_Log');

  // Create sheet if it doesn't exist
  if (!perfSheet) {
    perfSheet = ss.insertSheet('Performance_Log');
    perfSheet.appendRow(['Timestamp', 'UI Mode', 'Action', 'Duration (ms)']);
    perfSheet.getRange('A1:D1').setFontWeight('bold');
  }

  perfSheet.appendRow([timestamp, uiMode, action, duration]);
}

/**
 * Clears performance log
 */
function clearPerformanceLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perfSheet = ss.getSheetByName('Performance_Log');

  if (perfSheet) {
    perfSheet.clearContents();
    perfSheet.appendRow(['Timestamp', 'UI Mode', 'Action', 'Duration (ms)']);
    perfSheet.getRange('A1:D1').setFontWeight('bold');
    SpreadsheetApp.getUi().alert('Performance log cleared');
  } else {
    SpreadsheetApp.getUi().alert('No performance log found');
  }
}

/**
 * Analyzes performance data
 */
function analyzePerformance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perfSheet = ss.getSheetByName('Performance_Log');

  if (!perfSheet) {
    SpreadsheetApp.getUi().alert('No performance data available');
    return;
  }

  const data = perfSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No performance data available');
    return;
  }

  // Analyze by UI mode
  const cardServiceTimes = [];
  const htmlTimes = [];

  for (let i = 1; i < data.length; i++) {
    const uiMode = data[i][1];
    const duration = data[i][3];

    if (uiMode === 'CardService') {
      cardServiceTimes.push(duration);
    } else if (uiMode === 'HTML') {
      htmlTimes.push(duration);
    }
  }

  const avgCardService = cardServiceTimes.length > 0
    ? cardServiceTimes.reduce((a, b) => a + b, 0) / cardServiceTimes.length
    : 0;

  const avgHTML = htmlTimes.length > 0
    ? htmlTimes.reduce((a, b) => a + b, 0) / htmlTimes.length
    : 0;

  const improvement = avgHTML > 0
    ? ((avgHTML - avgCardService) / avgHTML * 100).toFixed(1)
    : 0;

  let message = 'Performance Analysis:\n\n';
  message += `CardService Average: ${avgCardService.toFixed(0)}ms\n`;
  message += `HTML Average: ${avgHTML.toFixed(0)}ms\n`;
  message += `Improvement: ${improvement}%\n\n`;
  message += `Total Samples: ${data.length - 1}`;

  SpreadsheetApp.getUi().alert('Performance Analysis', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// GRADUAL ROLLOUT FUNCTIONS
// ============================================================================

/**
 * Enables CardService for specific users (gradual rollout)
 * @param {Array<string>} userEmails - Array of user emails
 */
function enableCardServiceForUsers(userEmails) {
  const props = PropertiesService.getScriptProperties();
  const allowedUsers = JSON.stringify(userEmails);
  props.setProperty('cardservice_allowed_users', allowedUsers);
  console.log('CardService enabled for users:', userEmails);
}

/**
 * Checks if current user can use CardService
 * @return {boolean} True if user is allowed
 */
function isCardServiceEnabledForUser() {
  // If globally enabled, return true
  if (getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE)) {
    return true;
  }

  // Check if user is in allowed list
  try {
    const props = PropertiesService.getScriptProperties();
    const allowedUsersJson = props.getProperty('cardservice_allowed_users');

    if (!allowedUsersJson) {
      return false;
    }

    const allowedUsers = JSON.parse(allowedUsersJson);
    const currentUser = Session.getActiveUser().getEmail();

    return allowedUsers.includes(currentUser);
  } catch (error) {
    console.error('Error checking user allowlist:', error);
    return false;
  }
}

/**
 * Enables CardService for a percentage of users (A/B testing)
 * @param {number} percentage - Percentage of users (0-100)
 */
function enableCardServiceForPercentage(percentage) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('cardservice_rollout_percentage', percentage.toString());
  console.log(`CardService enabled for ${percentage}% of users`);
}

/**
 * Checks if current user is in the rollout percentage
 * @return {boolean} True if user is in rollout
 */
function isInCardServiceRollout() {
  try {
    const props = PropertiesService.getScriptProperties();
    const percentage = parseInt(props.getProperty('cardservice_rollout_percentage') || '0');

    if (percentage >= 100) {
      return true;
    }

    if (percentage <= 0) {
      return false;
    }

    // Use user email hash to deterministically assign users
    const userEmail = Session.getActiveUser().getEmail();
    const hash = hashString(userEmail);
    const userPercentile = hash % 100;

    return userPercentile < percentage;
  } catch (error) {
    console.error('Error checking rollout percentage:', error);
    return false;
  }
}

/**
 * Simple hash function for user email
 */
function hashString(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash);
}

// ============================================================================
// ADMIN UI FOR FEATURE FLAGS
// ============================================================================

/**
 * Shows feature flag management dialog
 */
function showFeatureFlagDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            margin: 0;
          }
          h2 {
            color: #1a73e8;
            margin-top: 0;
          }
          .flag-section {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
          }
          .flag-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 10px;
          }
          button {
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
          }
          .primary {
            background: #1a73e8;
            color: white;
          }
          .secondary {
            background: #e8f0fe;
            color: #1a73e8;
          }
          .danger {
            background: #d93025;
            color: white;
          }
          .status {
            font-weight: bold;
            padding: 4px 8px;
            border-radius: 4px;
          }
          .enabled {
            background: #e6f4ea;
            color: #188038;
          }
          .disabled {
            background: #fce8e6;
            color: #d93025;
          }
        </style>
      </head>
      <body>
        <h2>⚙️ Feature Flag Management</h2>

        <div class="flag-section">
          <h3>UI Mode</h3>
          <div class="flag-row">
            <span>CardService UI:</span>
            <span class="status ${getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE) ? 'enabled' : 'disabled'}">
              ${getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE) ? 'ENABLED' : 'DISABLED'}
            </span>
          </div>
          <button class="primary" onclick="enableCardService()">Enable CardService</button>
          <button class="danger" onclick="disableCardService()">Disable (Rollback to HTML)</button>
          <button class="secondary" onclick="toggleCardService()">Toggle</button>
        </div>

        <div class="flag-section">
          <h3>Performance Tracking</h3>
          <div class="flag-row">
            <span>Performance Logging:</span>
            <span class="status ${getFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING) ? 'enabled' : 'disabled'}">
              ${getFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING) ? 'ON' : 'OFF'}
            </span>
          </div>
          <button class="secondary" onclick="togglePerformanceLogging()">Toggle Logging</button>
          <button class="secondary" onclick="viewPerformanceAnalysis()">View Analysis</button>
          <button class="secondary" onclick="clearLog()">Clear Log</button>
        </div>

        <div class="flag-section">
          <h3>Gradual Rollout (Advanced)</h3>
          <p style="font-size: 13px; color: #5f6368;">
            Enable CardService for a percentage of users for A/B testing.
          </p>
          <label>Rollout Percentage:</label>
          <input type="number" id="rolloutPercentage" min="0" max="100" value="0" style="width: 80px; margin: 0 10px;">
          <button class="secondary" onclick="setRolloutPercentage()">Set Percentage</button>
        </div>

        <script>
          function enableCardService() {
            google.script.run.withSuccessHandler(() => {
              alert('CardService UI enabled! Refresh your sidebar.');
              google.script.host.close();
            }).enableCardServiceUI();
          }

          function disableCardService() {
            google.script.run.withSuccessHandler(() => {
              alert('Rolled back to HTML UI. Refresh your sidebar.');
              google.script.host.close();
            }).disableCardServiceUI();
          }

          function toggleCardService() {
            google.script.run.withSuccessHandler(() => {
              alert('UI mode toggled! Refresh your sidebar.');
              google.script.host.close();
            }).toggleCardServiceUI();
          }

          function togglePerformanceLogging() {
            google.script.run.withSuccessHandler(() => {
              alert('Performance logging toggled!');
              google.script.host.close();
            }).withFailureHandler((err) => {
              alert('Error: ' + err);
            }).togglePerformanceLogging();
          }

          function viewPerformanceAnalysis() {
            google.script.run.analyzePerformance();
          }

          function clearLog() {
            google.script.run.clearPerformanceLog();
          }

          function setRolloutPercentage() {
            const percentage = document.getElementById('rolloutPercentage').value;
            google.script.run.withSuccessHandler(() => {
              alert(\`CardService enabled for \${percentage}% of users\`);
            }).enableCardServiceForPercentage(parseInt(percentage));
          }

          function togglePerformanceLogging() {
            const isEnabled = ${getFeatureFlag(FEATURE_FLAGS.ENABLE_PERFORMANCE_LOGGING)};
            if (isEnabled) {
              google.script.run.disablePerformanceLogging();
            } else {
              google.script.run.enablePerformanceLogging();
            }
            setTimeout(() => google.script.host.close(), 500);
          }
        </script>
      </body>
    </html>
  `)
    .setWidth(600)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(html, 'Feature Flag Management');
}
