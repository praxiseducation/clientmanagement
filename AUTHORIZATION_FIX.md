# Fix Authorization Error - Step by Step

## Problem
After updating OAuth scopes in the manifest, you're still getting:
```
The script does not have permission to perform that action.
Required permissions: https://www.googleapis.com/auth/drive.file
```

## Solution: Force Reauthorization

### Method 1: Run Authorization Function (Recommended)

1. **Open your Apps Script project**

2. **Add this test function** to Code.gs (temporary - you can delete after):

```javascript
/**
 * Test function to force authorization
 * Run this once to grant all permissions
 */
function forceAuthorization() {
  // This function uses all the scopes to trigger auth
  try {
    // Calendar scope
    const calendarId = 'primary';
    Calendar.CalendarList.list();

    // Spreadsheet scope
    SpreadsheetApp.getActiveSpreadsheet();

    // Drive scope
    DriveApp.getRootFolder();

    // Script UI scope (implicit)
    Logger.log('All scopes authorized successfully!');

    return 'Authorization complete';
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return 'Error: ' + error.message;
  }
}
```

3. **Select the function** from dropdown at top: `forceAuthorization`

4. **Click Run** (▶️ button)

5. **Grant permissions** when prompted:
   - Click "Review Permissions"
   - Choose your Google account
   - Click "Advanced" (if you see unsafe warning)
   - Click "Go to Client Manager - Minimal (unsafe)"
   - Click "Allow"
   - You should see all 6 scopes listed

6. **Verify** - Check execution log shows "All scopes authorized successfully!"

7. **Delete the test function** (optional - clean up)

8. **Create new test deployment**:
   - Deploy → Test deployments
   - Click "Install"
   - Open any Sheet
   - Test the add-on

---

### Method 2: Fresh Test Deployment

1. **In Apps Script editor:**
   - Deploy → Test deployments
   - If installed, click "Uninstall"

2. **Clear previous authorizations:**
   - Go to: https://myaccount.google.com/permissions
   - Find "Client Manager - Minimal" (or your project name)
   - Click it → Remove Access
   - Confirm removal

3. **Create new test deployment:**
   - Deploy → Test deployments
   - Click "Install"
   - **Important:** Grant ALL permissions when prompted
   - Click "Done"

4. **Open Google Sheets:**
   - Open any sheet
   - Click puzzle icon (🧩)
   - Click "Client Manager - Minimal"
   - Should work now!

---

### Method 3: Head Deployment (Most Reliable)

If test deployment keeps having issues, use a head deployment:

1. **Deploy → New deployment**

2. **Select type:** Add-on

3. **Description:** "Client Manager Minimal v1.0"

4. **Execute as:** Me

5. **Who has access:** Only myself

6. **Click Deploy**

7. **Copy Deployment ID**

8. **Install via ID:**
   - Open any Google Sheet
   - Extensions → Add-ons → Get add-ons
   - Search for deployment ID (paste it)
   - Install
   - Grant all permissions

---

## Verification Checklist

After reauthorization, verify these permissions are granted:

- [ ] ✅ **Calendar (read-only)** - View calendar events
- [ ] ✅ **Spreadsheets** - Create and edit spreadsheets
- [ ] ✅ **Drive** - Access Drive folders
- [ ] ✅ **Drive File** - Manage files created by add-on
- [ ] ✅ **Script Container UI** - Show sidebar and cards
- [ ] ✅ **External Requests** - Make web requests (if needed)

---

## Common Issues

**"Unsafe app" warning**
→ Normal for test deployments. Click "Advanced" → "Go to [app] (unsafe)"

**Permissions not showing all 6 scopes**
→ Manifest not updated in project. Copy minimal-appsscript.json again.

**Still getting permission error after reauth**
→ Use Method 3 (head deployment) instead of test deployment

**Can't grant permissions**
→ Check Google Admin hasn't blocked unverified apps (if Workspace account)

---

## After Authorization Works

Once authorized successfully:

1. **Test creating a client sheet** - Should work without errors
2. **Test setting storage folder** - Should access Drive
3. **Test calendar sync** - Should fetch appointments
4. **Delete the test function** if you added `forceAuthorization()`

---

## Need Help?

If still having issues:
1. Check executions log for detailed error
2. Verify manifest matches minimal-appsscript.json exactly
3. Try incognito/private browsing mode
4. Contact Google Workspace admin (if using Workspace account)
