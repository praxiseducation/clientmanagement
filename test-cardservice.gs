/**
 * Diagnostic Tests for CardService Setup
 * Run these functions manually to diagnose issues
 */

/**
 * Test 1: Check if all required files/functions are loaded
 */
function test1_CheckFilesLoaded() {
  const results = [];

  try {
    results.push('getFeatureFlag: ' + (typeof getFeatureFlag !== 'undefined' ? '✅ FOUND' : '❌ MISSING'));
  } catch (e) {
    results.push('getFeatureFlag: ❌ ERROR - ' + e.message);
  }

  try {
    results.push('showCardServiceSidebar: ' + (typeof showCardServiceSidebar !== 'undefined' ? '✅ FOUND' : '❌ MISSING'));
  } catch (e) {
    results.push('showCardServiceSidebar: ❌ ERROR - ' + e.message);
  }

  try {
    results.push('FEATURE_FLAGS: ' + (typeof FEATURE_FLAGS !== 'undefined' ? '✅ FOUND' : '❌ MISSING'));
  } catch (e) {
    results.push('FEATURE_FLAGS: ❌ ERROR - ' + e.message);
  }

  try {
    results.push('getAllClients: ' + (typeof getAllClients !== 'undefined' ? '✅ FOUND' : '❌ MISSING'));
  } catch (e) {
    results.push('getAllClients: ❌ ERROR - ' + e.message);
  }

  Logger.log('=== FILE LOAD TEST ===');
  results.forEach(r => Logger.log(r));

  SpreadsheetApp.getUi().alert('File Load Test', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);

  return results;
}

/**
 * Test 2: Check if CardService is available in this context
 */
function test2_CheckCardServiceAvailable() {
  const results = [];

  try {
    results.push('CardService object: ' + (typeof CardService !== 'undefined' ? '✅ AVAILABLE' : '❌ NOT AVAILABLE'));

    if (typeof CardService !== 'undefined') {
      const card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle('Test'))
        .build();
      results.push('CardService.newCardBuilder(): ✅ WORKS');
      results.push('Card object type: ' + typeof card);
    }
  } catch (e) {
    results.push('CardService ERROR: ❌ ' + e.message);
    results.push('This means CardService is not available in this context');
    results.push('You are likely using a container-bound script');
    results.push('CardService only works in Workspace Add-ons');
  }

  Logger.log('=== CARDSERVICE AVAILABILITY TEST ===');
  results.forEach(r => Logger.log(r));

  SpreadsheetApp.getUi().alert('CardService Test', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);

  return results;
}

/**
 * Test 3: Check feature flag status
 */
function test3_CheckFeatureFlags() {
  const results = [];

  try {
    const props = PropertiesService.getScriptProperties();
    const useCardService = props.getProperty('use_cardservice_ui');

    results.push('Feature Flag Value: ' + (useCardService || 'NOT SET (defaults to false)'));
    results.push('CardService Enabled: ' + (useCardService === 'true' ? '✅ YES' : '❌ NO'));

    if (typeof getFeatureFlag !== 'undefined') {
      const flagValue = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);
      results.push('getFeatureFlag() returns: ' + flagValue);
    }

  } catch (e) {
    results.push('ERROR reading feature flags: ' + e.message);
  }

  Logger.log('=== FEATURE FLAG TEST ===');
  results.forEach(r => Logger.log(r));

  SpreadsheetApp.getUi().alert('Feature Flag Test', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);

  return results;
}

/**
 * Test 4: Test HTML sidebar (baseline)
 */
function test4_TestHtmlSidebar() {
  try {
    const html = HtmlService.createHtmlOutput(`
      <html>
        <head>
          <style>
            body { font-family: Arial; padding: 20px; }
            h2 { color: #1a73e8; }
          </style>
        </head>
        <body>
          <h2>✅ HTML Sidebar Works!</h2>
          <p>This confirms that HtmlService is working correctly.</p>
          <p>If you see this, the issue is specifically with CardService, not with sidebars in general.</p>
        </body>
      </html>
    `).setWidth(300);

    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log('HTML sidebar test: SUCCESS');

  } catch (e) {
    Logger.log('HTML sidebar test: FAILED - ' + e.message);
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}

/**
 * Test 5: Try to build a simple CardService card
 */
function test5_TryBuildSimpleCard() {
  const results = [];

  try {
    results.push('Attempting to build a simple CardService card...');

    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('Test Card'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('This is a test card')))
      .build();

    results.push('✅ Card built successfully!');
    results.push('Card object type: ' + typeof card);
    results.push('');
    results.push('IMPORTANT: Card built, but cannot be shown in sidebar.');
    results.push('In container-bound scripts, sidebars use HtmlService, not CardService.');
    results.push('CardService cards are for Workspace Add-ons only.');

  } catch (e) {
    results.push('❌ Card build FAILED');
    results.push('Error: ' + e.message);
    results.push('');
    results.push('This confirms CardService is not available.');
    results.push('You are using a container-bound script.');
    results.push('Recommendation: Stick with HTML UI.');
  }

  Logger.log('=== CARD BUILD TEST ===');
  results.forEach(r => Logger.log(r));

  SpreadsheetApp.getUi().alert('Card Build Test', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);

  return results;
}

/**
 * Test 6: Full integration test
 */
function test6_FullIntegrationTest() {
  const results = [];
  results.push('=== FULL INTEGRATION TEST ===\n');

  // Test 1
  results.push('1. Files loaded...');
  try {
    if (typeof getFeatureFlag === 'undefined') throw new Error('getFeatureFlag not found');
    if (typeof showCardServiceSidebar === 'undefined') throw new Error('showCardServiceSidebar not found');
    results.push('   ✅ All required functions found\n');
  } catch (e) {
    results.push('   ❌ FAILED: ' + e.message);
    results.push('   FIX: Upload missing .gs files to Apps Script\n');
  }

  // Test 2
  results.push('2. CardService available...');
  try {
    if (typeof CardService === 'undefined') throw new Error('CardService not defined');
    CardService.newCardBuilder();
    results.push('   ✅ CardService is available\n');
  } catch (e) {
    results.push('   ❌ FAILED: ' + e.message);
    results.push('   FIX: Cannot use CardService in container-bound scripts');
    results.push('   RECOMMENDATION: Disable CardService, use HTML UI\n');
  }

  // Test 3
  results.push('3. Feature flags...');
  try {
    const flagValue = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);
    results.push('   ✅ Feature flag readable: ' + flagValue + '\n');
  } catch (e) {
    results.push('   ❌ FAILED: ' + e.message + '\n');
  }

  // Test 4
  results.push('4. showSidebar() function...');
  try {
    if (typeof showSidebar === 'undefined') throw new Error('showSidebar not found');
    results.push('   ✅ showSidebar function exists\n');
  } catch (e) {
    results.push('   ❌ FAILED: ' + e.message + '\n');
  }

  results.push('\n=== SUMMARY ===');
  results.push('Check the results above for any ❌ failures.');
  results.push('See Execution Log for details.');

  Logger.log(results.join('\n'));

  const ui = SpreadsheetApp.getUi();
  const alert = HtmlService.createHtmlOutput(
    '<html><body style="font-family:monospace;padding:20px;"><pre>' +
    results.join('\n') +
    '</pre></body></html>'
  ).setWidth(600).setHeight(400);

  ui.showModalDialog(alert, 'Integration Test Results');
}

/**
 * Quick Fix: Disable CardService
 */
function quickFix_DisableCardService() {
  try {
    PropertiesService.getScriptProperties().setProperty('use_cardservice_ui', 'false');
    Logger.log('CardService DISABLED');
    SpreadsheetApp.getUi().alert(
      'CardService Disabled',
      'CardService has been disabled. Your sidebar will now use the HTML UI.\n\nClose and reopen the sidebar to see the change.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Error disabling CardService: ' + e.message);
  }
}

/**
 * Quick Fix: Enable CardService (for testing)
 */
function quickFix_EnableCardService() {
  try {
    PropertiesService.getScriptProperties().setProperty('use_cardservice_ui', 'true');
    Logger.log('CardService ENABLED');
    SpreadsheetApp.getUi().alert(
      'CardService Enabled',
      'CardService has been enabled (for testing).\n\nClose and reopen the sidebar.\n\nNote: This may not work in container-bound scripts.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Error enabling CardService: ' + e.message);
  }
}

/**
 * Run ALL tests in sequence
 */
function runAllTests() {
  Logger.log('\n\n=== RUNNING ALL DIAGNOSTIC TESTS ===\n');

  test1_CheckFilesLoaded();
  Utilities.sleep(1000);

  test2_CheckCardServiceAvailable();
  Utilities.sleep(1000);

  test3_CheckFeatureFlags();
  Utilities.sleep(1000);

  test6_FullIntegrationTest();

  Logger.log('\n=== ALL TESTS COMPLETE ===');
  Logger.log('Check the results above and the modal dialogs.');
}
