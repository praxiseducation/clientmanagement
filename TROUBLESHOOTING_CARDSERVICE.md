# Troubleshooting: CardService Sidebar Not Loading

## Symptom
After enabling CardService, the sidebar doesn't load or shows an error.

---

## Quick Diagnostics

### Step 1: Check if files were uploaded to Apps Script

1. Open your Google Sheet
2. Extensions > Apps Script
3. **Verify these 3 files exist** (left sidebar):
   - ✅ cardservice-ui
   - ✅ feature-flags
   - ✅ cardservice-helpers

**If files are missing:**
- You need to upload them to Apps Script (not just save locally)
- See "Upload Files" section below

---

### Step 2: Check Apps Script Console for Errors

1. In Apps Script editor, click **Execution log** (bottom)
2. Or go to View > Logs
3. Look for error messages

**Common Errors:**

**Error: "getFeatureFlag is not defined"**
- **Fix:** feature-flags.gs not uploaded or not saved

**Error: "showCardServiceSidebar is not defined"**
- **Fix:** cardservice-ui.gs not uploaded or not saved

**Error: "FEATURE_FLAGS is not defined"**
- **Fix:** feature-flags.gs not loaded, see fix below

**Error: "CardService is not available"**
- **Fix:** Wrong add-on type, see fix below

---

## Common Issues & Fixes

### Issue 1: Files Not Uploaded to Apps Script

**Problem:** Files are saved locally but not in Apps Script project

**Fix:**
1. Go to Extensions > Apps Script
2. Click **+** next to "Files"
3. Select "Script"
4. Name it exactly: `cardservice-ui` (no .gs extension)
5. Copy entire contents from local file cardservice-ui.gs
6. **Ctrl+S to save**
7. Repeat for `feature-flags` and `cardservice-helpers`

---

### Issue 2: FEATURE_FLAGS Not Defined

**Problem:** The constant `FEATURE_FLAGS` is referenced but not accessible

**Fix:** Add this to the TOP of feature-flags.gs (if not there):

```javascript
const FEATURE_FLAGS = {
  USE_CARDSERVICE: 'use_cardservice_ui',
  CARDSERVICE_VERSION: 'cardservice_version',
  ENABLE_PERFORMANCE_LOGGING: 'enable_performance_logging'
};
```

Then save and refresh.

---

### Issue 3: CardService Not Available

**Problem:** CardService only works in specific contexts

**Cause:** You might be trying to use CardService in:
- A standalone script (not an add-on)
- Web app mode
- Editor add-on (Docs/Sheets/Slides)

**CardService ONLY works in:**
- Google Workspace Add-ons (Gmail, Calendar, Drive, etc.)
- Google Chat apps

**Fix:** This is a limitation. If your add-on is a "Container-bound script" (bound to the Sheet), CardService won't work for sidebar. You need to:

**Option A: Stick with HTML UI (Recommended for now)**
1. Smart College > Settings > Feature Flags
2. Click "Disable (Rollback to HTML)"
3. HTML UI will work fine

**Option B: Convert to Workspace Add-on (Advanced)**
- Requires publishing as a Workspace Add-on
- Requires manifest changes
- See "Convert to Workspace Add-on" below

---

### Issue 4: Function Not Found After Upload

**Problem:** Uploaded files but functions still not found

**Fix:**
1. **Save all files** in Apps Script (Ctrl+S on each)
2. **Close Apps Script tab**
3. **Refresh Google Sheet**
4. Try again

Sometimes Apps Script needs a hard refresh.

---

### Issue 5: Sidebar Shows Blank/White Screen

**Problem:** Sidebar opens but shows nothing

**Fix:**

**Test 1: Check if HTML UI works**
```javascript
// In Apps Script, run this manually:
function testHtmlSidebar() {
  const html = HtmlService.createHtmlOutput('<h1>Test</h1>').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
```

If this works, issue is with CardService.

**Test 2: Check if CardService is available**
```javascript
// Run this in Apps Script console:
function testCardService() {
  try {
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('Test'))
      .build();
    Logger.log('CardService available: ' + typeof card);
    return card;
  } catch (error) {
    Logger.log('CardService ERROR: ' + error.message);
    throw error;
  }
}
```

If error: "CardService is not defined" → CardService not available (see Issue 3)

---

## Step-by-Step Fix Guide

### Fix Attempt 1: Verify Files (Most Common)

1. Open Apps Script (Extensions > Apps Script)
2. Check left sidebar - should see:
   ```
   Files
   ├── clientmanager
   ├── backend/sidebar-functions
   ├── cardservice-ui          ← MUST BE HERE
   ├── feature-flags           ← MUST BE HERE
   ├── cardservice-helpers     ← MUST BE HERE
   └── ... other files
   ```

3. If any missing, upload them:
   - Click + next to Files
   - Select "Script"
   - Name exactly as shown (no .gs)
   - Copy contents from local file
   - Save (Ctrl+S)

4. **Save ALL files** (very important)
5. Close Apps Script tab
6. Refresh Google Sheet
7. Try opening sidebar

---

### Fix Attempt 2: Manual Test

Run this in Apps Script console to test CardService:

```javascript
function testCardServiceSidebar() {
  try {
    // Test 1: Check feature flags exist
    Logger.log('Testing feature flags...');
    const useCS = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);
    Logger.log('Feature flag value: ' + useCS);

    // Test 2: Try to build a simple card
    Logger.log('Testing CardService...');
    const card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader()
        .setTitle('TEST - It Works!'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText('CardService is working!')))
      .build();

    Logger.log('Card built successfully');
    return card;

  } catch (error) {
    Logger.log('ERROR: ' + error.message);
    Logger.log('Stack: ' + error.stack);
    throw error;
  }
}
```

**Run this function:**
1. Select `testCardServiceSidebar` from dropdown
2. Click Run (▶️)
3. Check Execution Log for results

**If successful:** CardService works, issue is in your code
**If error:** See what error says

---

### Fix Attempt 3: Disable and Test HTML First

Before debugging CardService, verify HTML works:

1. **Disable CardService:**
   ```javascript
   function disableCardServiceManually() {
     PropertiesService.getScriptProperties()
       .setProperty('use_cardservice_ui', 'false');
     Logger.log('CardService disabled');
   }
   ```
   Run this function.

2. **Test HTML sidebar:**
   - Smart College > Open Control Panel
   - Should show HTML sidebar

3. **If HTML works:** Issue is CardService-specific
4. **If HTML broken too:** Issue is more general

---

### Fix Attempt 4: Use Simplified CardService

Replace the showSidebar() function with this simpler version:

```javascript
function showSidebar() {
  try {
    // Check feature flag
    const props = PropertiesService.getScriptProperties();
    const useCardService = props.getProperty('use_cardservice_ui') === 'true';

    if (useCardService) {
      Logger.log('Attempting to show CardService UI');

      // Simple test card
      const card = CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle('Client Manager'))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText('CardService UI is working!')))
        .build();

      return card;
    } else {
      Logger.log('Using HTML UI');
      const html = loadSidebarTemplate('control');
      SpreadsheetApp.getUi().showSidebar(html);
    }
  } catch (error) {
    Logger.log('ERROR in showSidebar: ' + error.message);
    // Fallback to HTML
    const html = loadSidebarTemplate('control');
    SpreadsheetApp.getUi().showSidebar(html);
  }
}
```

This simplified version will:
- Log what it's doing
- Try CardService with minimal card
- Fallback to HTML if error

---

## The Real Issue: CardService Context

### IMPORTANT: CardService Limitations

**CardService DOES NOT work in:**
- ❌ Container-bound scripts (scripts attached to Sheets)
- ❌ Standalone scripts
- ❌ Web apps
- ❌ Sidebar/Dialog in Sheets (uses HtmlService)

**CardService ONLY works in:**
- ✅ Workspace Add-ons (when published)
- ✅ Gmail add-ons
- ✅ Google Chat apps

### The Problem With Your Setup

Your script appears to be a **container-bound script** (attached to the Google Sheet). In this context:

- `SpreadsheetApp.getUi().showSidebar()` expects **HtmlService**, not CardService
- CardService cards are returned differently (not shown via showSidebar)

### The Solution

**Option 1: STAY WITH HTML UI (Recommended)**

CardService won't work in your current setup. The HTML UI is:
- ✅ Fully functional
- ✅ Customizable
- ✅ Works in container-bound scripts
- ✅ Easier to maintain for your use case

Just disable CardService and use the optimized HTML.

**Option 2: Convert to Workspace Add-on (Complex)**

To use CardService, you'd need to:
1. Create a new Workspace Add-on project
2. Set up proper manifest (appsscript.json)
3. Implement `onHomepage()` trigger (returns Card)
4. Publish to Google Workspace Marketplace
5. Install as an add-on (not container-bound)

This is a major architectural change and probably not worth it.

---

## Recommended Next Steps

### For Container-Bound Script (Your Current Setup):

1. **Disable CardService:**
   - Smart College > Settings > Feature Flags
   - Click "Disable"
   - Or run: `disableCardServiceUI();`

2. **Optimize HTML UI instead:**
   - Minify CSS
   - Lazy load components
   - Cache more aggressively
   - Use modern CSS (containment, etc.)

3. **Accept trade-off:**
   - HTML UI works great for container-bound
   - CardService requires add-on architecture
   - Not worth the migration effort

---

## Testing Commands

Run these in Apps Script console to diagnose:

**Test 1: Check if files loaded**
```javascript
function testFilesLoaded() {
  try {
    Logger.log('Testing getFeatureFlag: ' + typeof getFeatureFlag);
    Logger.log('Testing showCardServiceSidebar: ' + typeof showCardServiceSidebar);
    Logger.log('Testing FEATURE_FLAGS: ' + typeof FEATURE_FLAGS);
    Logger.log('All functions found!');
  } catch (error) {
    Logger.log('ERROR: ' + error.message);
  }
}
```

**Test 2: Check CardService availability**
```javascript
function testCardServiceAvailable() {
  try {
    Logger.log('CardService type: ' + typeof CardService);
    const card = CardService.newCardBuilder().build();
    Logger.log('CardService WORKS');
  } catch (error) {
    Logger.log('CardService NOT AVAILABLE: ' + error.message);
  }
}
```

**Test 3: Check feature flag**
```javascript
function testFeatureFlag() {
  const props = PropertiesService.getScriptProperties();
  const value = props.getProperty('use_cardservice_ui');
  Logger.log('Feature flag value: ' + value);
}
```

---

## Quick Fix: Fallback Version

Replace your showSidebar() with this error-resistant version:

```javascript
function showSidebar() {
  try {
    // Try to load feature flags safely
    let useCardService = false;
    try {
      useCardService = getFeatureFlag(FEATURE_FLAGS.USE_CARDSERVICE, false);
    } catch (e) {
      Logger.log('Feature flags not available, using HTML');
    }

    // If CardService enabled, try it
    if (useCardService) {
      try {
        const card = showCardServiceSidebar();
        Logger.log('CardService sidebar loaded');
        return card;
      } catch (cardError) {
        Logger.log('CardService failed: ' + cardError.message);
        Logger.log('Falling back to HTML');
        // Fall through to HTML
      }
    }

    // Default: HTML UI
    const html = loadSidebarTemplate('control');
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log('HTML sidebar loaded');

  } catch (error) {
    // Last resort: simple HTML
    Logger.log('Critical error: ' + error.message);
    const html = HtmlService.createHtmlOutput('<h2>Error loading sidebar</h2><p>' + error.message + '</p>');
    SpreadsheetApp.getUi().showSidebar(html);
  }
}
```

This version:
- ✅ Handles missing feature-flags.gs
- ✅ Handles CardService errors
- ✅ Falls back to HTML gracefully
- ✅ Logs what's happening
- ✅ Always shows something

---

## Most Likely Issue

Based on your setup, the **most likely issue** is:

**CardService doesn't work in container-bound scripts (scripts attached to Sheets)**

Your script is bound to a Google Sheet, so:
- `SpreadsheetApp.getUi().showSidebar()` only accepts HtmlService
- CardService.build() returns a Card object
- But there's no way to show a Card in a sidebar in this context

**Solution:** Disable CardService and stick with HTML UI.

---

## Next Steps

1. **Check Apps Script Execution Log** for actual error
2. **Run diagnostic tests** (above) to see what works
3. **Disable CardService** if it's the container-bound issue
4. **Let me know the exact error message** from the log

What error do you see in the Apps Script execution log?
