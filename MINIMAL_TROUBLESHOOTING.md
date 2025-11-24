# Minimal Client Manager - Troubleshooting Guide

## Error: "urlFetchWhitelist is required"

**Error Message:**
```
An explicit urlFetchWhitelist is required for all Google Workspace add-ons using UrlFetchApp.
```

### Cause
This error occurs when:
- The manifest has `script.external_request` scope BUT
- No `urlFetchWhitelist` is defined
- Or a cached version of the manifest still has the old scopes

### Solution: Clean Reinstall

**Step 1: Verify Manifest**

Your `minimal-appsscript.json` should have ONLY these 3 scopes:
```json
"oauthScopes": [
  "https://www.googleapis.com/auth/calendar.readonly",
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.container.ui"
]
```

**NO** `script.external_request` scope!

**Step 2: Complete Uninstall**

1. **In Apps Script project:**
   - Deploy → Test deployments
   - Click **Uninstall**
   - Confirm

2. **Clear OAuth authorization:**
   - Go to: https://myaccount.google.com/permissions
   - Find "Client Manager - Minimal" (or test deployment name)
   - Click **Remove Access**

3. **Close all Google Sheets tabs**

**Step 3: Update Manifest**

1. **In Apps Script editor:**
   - Click `appsscript.json`
   - **Verify** it matches exactly:

```json
{
  "timeZone": "America/Chicago",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Calendar",
        "version": "v3",
        "serviceId": "calendar"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.container.ui"
  ],
  "addOns": {
    "common": {
      "name": "Client Manager - Minimal",
      "logoUrl": "https://www.gstatic.com/images/branding/product/1x/calendar_2020q4_48dp.png",
      "layoutProperties": {
        "primaryColor": "#1a73e8"
      },
      "universalActions": [
        {
          "label": "Refresh",
          "runFunction": "refreshDashboard"
        }
      ]
    },
    "sheets": {
      "homepageTrigger": {
        "runFunction": "onHomepage_Minimal",
        "enabled": true
      }
    }
  }
}
```

2. **Save** (Ctrl+S)

**Step 4: Fresh Test Deployment**

1. **Deploy → Test deployments**
2. **Click Install**
3. **Grant permissions** (should show only 3 scopes)
4. **Open Google Sheets**
5. **Click add-on icon**

---

## Other Common Errors

### Error: "Calendar API not enabled"

**Solution:**
1. Apps Script editor → Services (+)
2. Add **Google Calendar API** v3
3. Save and redeploy

### Error: "Cannot read property 'clientName'"

**Solution:**
- No calendar events found or calendar not set
- Go to Settings → Change Calendar
- Ensure calendar has events formatted: `Client Name: Description`

### Error: "Clients sheet not found"

**Solution:**
- Click **New Client** to create Clients sheet automatically
- Or manually create sheet named "Clients" with headers:
  `Name | Email | Parent Email | Phone | Grade | Notes | Status | Date Added`

### Error: "Permission denied" when creating sheet

**Solution:**
- OAuth scope issue - make sure manifest has `spreadsheets.currentonly`
- Uninstall and reinstall to refresh permissions

---

## Verification Checklist

After successful installation, verify:

- [ ] Add-on icon appears in right sidebar (calendar icon)
- [ ] Clicking icon opens Client Manager card
- [ ] Settings shows calendar selection
- [ ] Can select calendar and save
- [ ] Today's sessions appear (if calendar has events)
- [ ] Search shows client suggestions
- [ ] Can click "New Client" and see form
- [ ] Can save new client
- [ ] Can click "Open Sheet" on client
- [ ] Sheet tab is created and activated
- [ ] "All Sheets" shows list of client sheets

---

## Debug Mode

To check what's happening:

**1. Check Execution Logs:**
- Apps Script editor → Executions tab
- Look for errors in recent runs
- Check which functions failed

**2. Test Functions Manually:**

Run these in Apps Script editor to test:

```javascript
// Test calendar connection
function testCalendar() {
  const calendars = Calendar.CalendarList.list();
  Logger.log('Found calendars: ' + calendars.items.length);
}

// Test client retrieval
function testClients() {
  const clients = getAllClients_Minimal();
  Logger.log('Found clients: ' + clients.length);
}

// Test appointments
function testAppointments() {
  const appointments = fetchCalendarAppointments_Minimal();
  Logger.log('Found appointments: ' + appointments.length);
}
```

**3. Check Properties:**

```javascript
function checkProperties() {
  const props = PropertiesService.getUserProperties();
  Logger.log('Calendar ID: ' + props.getProperty('MINIMAL_CALENDAR_ID'));
}
```

---

## Still Having Issues?

1. **Check manifest matches exactly** - copy/paste from this guide
2. **Verify no UrlFetchApp** in code - search for "UrlFetch"
3. **Clear browser cache** and try in incognito mode
4. **Wait 5 minutes** after uninstall before reinstalling
5. **Try different browser** (Chrome vs Firefox)

---

## Contact Info

If errors persist:
- Check Apps Script execution logs
- Verify all OAuth scopes are granted
- Ensure Calendar API is enabled in Services
- Test with a fresh Google Sheet

---

**Last Updated:** 2025-01-09
