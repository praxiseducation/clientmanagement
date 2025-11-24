# Converting to Workspace Add-on - Deployment Guide

## 🎯 Goal

Convert your container-bound script to a **standalone Workspace Add-on** that:
- ✅ Works with CardService (native UI)
- ✅ Can be installed by internal users
- ✅ Works across multiple spreadsheets
- ✅ Can be distributed within your organization

---

## 📋 Prerequisites

- Google Workspace account (not personal Gmail)
- Admin rights for internal deployment OR
- Ability to use "Test deployment" for your team

---

## 🚀 Deployment Options

### Option 1: Internal Deployment (Recommended for Teams)
- Deploy within your Google Workspace organization
- Users install from your organization's add-on gallery
- Requires Google Workspace admin approval
- **Best for:** 5+ users, ongoing use

### Option 2: Test Deployment (Quick Start)
- Share directly with specific users
- No admin approval needed
- Users install via deployment ID
- **Best for:** Testing, small teams (<10 users)

### Option 3: Public Marketplace (Not Recommended for Internal)
- Publish to Google Workspace Marketplace
- Available to all Google Workspace users
- Requires OAuth verification ($$$)
- **Best for:** Public distribution

**We'll cover Options 1 & 2 (internal use)**

---

## 📂 Step 1: Create Standalone Apps Script Project

### 1.1 Create New Standalone Project

1. **Go to:** https://script.google.com
2. **Click:** "New project"
3. **Rename:** "Client Manager Add-on" (click title at top)

**Important:** Do NOT open Apps Script from a spreadsheet. This must be a standalone project.

### 1.2 Upload Manifest File

1. **In Apps Script editor, click:**
   - View > Show manifest file (or Project Settings > Show "appsscript.json")
2. **You should now see** `appsscript.json` in the file list
3. **Replace entire contents** with:

```json
{
  "timeZone": "America/New_York",
  "dependencies": {
    "enabledAdvancedServices": []
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/userinfo.email"
  ],
  "addOns": {
    "common": {
      "name": "Client Manager",
      "logoUrl": "https://www.gstatic.com/images/branding/product/1x/apps_script_48dp.png",
      "layoutProperties": {
        "primaryColor": "#003366",
        "secondaryColor": "#1a73e8"
      },
      "homepageTrigger": {
        "runFunction": "onHomepage",
        "enabled": true
      },
      "universalActions": [
        {
          "label": "Refresh",
          "runFunction": "refreshCard"
        }
      ]
    },
    "sheets": {
      "homepageTrigger": {
        "runFunction": "onSheetsHomepage",
        "enabled": true
      },
      "onFileScopeGrantedTrigger": {
        "runFunction": "onFileScopeGranted"
      }
    }
  }
}
```

4. **Save** (Ctrl+S)

---

## 📦 Step 2: Upload All Script Files

### 2.1 Core Files (Upload in this order)

**File 1: Code.gs** (rename default Code.gs)
- Already exists
- You can rename or delete it
- We'll add a simple entry point

**File 2: addon-lifecycle.gs**
1. Click **+** next to Files → Script
2. Name: `addon-lifecycle`
3. Copy contents from [addon-lifecycle.gs](addon-lifecycle.gs)
4. Save

**File 3: cardservice-ui.gs**
1. Click **+** next to Files → Script
2. Name: `cardservice-ui`
3. Copy contents from [cardservice-ui.gs](cardservice-ui.gs)
4. Save

**File 4: cardservice-helpers.gs**
1. Click **+** next to Files → Script
2. Name: `cardservice-helpers`
3. Copy contents from [cardservice-helpers.gs](cardservice-helpers.gs)
4. Save

**File 5: feature-flags.gs**
1. Click **+** next to Files → Script
2. Name: `feature-flags`
3. Copy contents from [feature-flags.gs](feature-flags.gs)
4. Save

### 2.2 Optional Files (if you want HTML fallback)

If you want to keep Quick Notes HTML dialog:

**File 6: Create HTML file**
1. Click **+** next to Files → HTML
2. Name: `quick-notes`
3. Copy contents from `frontend/components/quick-notes.html`
4. Save

**Note:** You can skip HTML files and use pure CardService

---

## 🔧 Step 3: Update Code.gs (Entry Point)

Replace contents of `Code.gs` with:

```javascript
/**
 * Client Manager Workspace Add-on
 * Standalone version with CardService UI
 */

/**
 * Main entry point - called when add-on is opened
 */
function onHomepage(e) {
  return onSheetsHomepage(e);
}

/**
 * Called when opened in Sheets
 */
function onSheetsHomepage(e) {
  return buildHomepageCard(e);
}

/**
 * Called when add-on is installed
 */
function onInstall(e) {
  onHomepage(e);
}

/**
 * Called when file access is granted
 */
function onFileScopeGranted(e) {
  return buildMainClientCard(e);
}

/**
 * Universal refresh action
 */
function refreshCard(e) {
  return buildMainClientCard(e);
}

/**
 * Test function to verify deployment
 */
function testAddon() {
  Logger.log('Add-on is loaded correctly');
  return 'Success';
}
```

Save all files.

---

## 🧪 Step 4: Test the Add-on

### 4.1 Test Deployment

1. **In Apps Script editor, click:**
   - Deploy > Test deployments

2. **Click:** "Install"

3. **Select installation config:**
   - Sheets homepage trigger: Enabled
   - Click "Done"

4. **A test version is now installed** for you

### 4.2 Test in Google Sheets

1. **Open any Google Sheet** (or create new one)

2. **Look for the add-on:**
   - Click puzzle piece icon (🧩) in top right
   - You should see "Client Manager" (might say "Test - Client Manager")

3. **Click the add-on**
   - Should open in sidebar
   - Should show CardService UI!

4. **Test functionality:**
   - Add a client
   - Search for client
   - View client list
   - Everything should work

### 4.3 Troubleshooting Test Deployment

**Issue: Add-on doesn't appear**
- Wait 1-2 minutes and refresh
- Check Project Settings > OAuth scopes are populated
- Try "Refresh" in Test deployments dialog

**Issue: Permission errors**
- Add-on will request permissions first time
- Click "Continue" and grant permissions
- This is normal for new add-ons

---

## 📤 Step 5: Deploy for Internal Users

### Option A: Test Deployment Sharing (Quick, 10 minutes)

**Best for:** Small teams, testing

1. **In Apps Script, click:**
   - Deploy > Test deployments

2. **Copy the Deployment ID**
   - Looks like: `AKfycby...` (long string)

3. **Share with users:**
   - Send them the Deployment ID
   - They install via: Extensions > Add-ons > Get add-ons > Search by ID

**Limitations:**
- Max 10-20 test users
- Manual installation
- Not persistent (expires)

---

### Option B: Internal Deployment (Recommended, 30 minutes)

**Best for:** Organization-wide use, production

#### 5.1 Create Head Deployment

1. **In Apps Script, click:**
   - Deploy > New deployment

2. **Select type:** "Add-on"

3. **Configure:**
   - Description: "Client Manager - Internal"
   - Version: "1.0"
   - Visibility: "Only for users in [your organization]"

4. **Click:** "Deploy"

5. **Copy Deployment ID** (save this)

#### 5.2 Submit for Google Workspace Admin Approval

**If your organization requires admin approval:**

1. **Send to your Workspace admin:**
   ```
   Subject: Add-on Approval Request - Client Manager

   Hi [Admin],

   I've created an internal add-on for our team to manage client data
   in Google Sheets.

   Deployment ID: [paste ID]
   Script Project: [paste URL to Apps Script project]

   This add-on:
   - Manages client information
   - Tracks session notes
   - Sends automated recap emails
   - Uses only internal data (no external APIs)

   OAuth Scopes Required:
   - Spreadsheets (read/write client data)
   - Gmail (send recap emails)
   - User email (identify tutor)

   Please approve for internal installation.

   Thanks!
   ```

2. **Admin approves via:**
   - Google Admin Console
   - Apps > Google Workspace Marketplace apps
   - Allowlist the deployment ID

#### 5.3 Publish to Internal Marketplace

**After admin approval:**

1. **Go to:** Google Workspace Marketplace SDK
   - https://console.cloud.google.com/apis/library/appsmarket-component.googleapis.com

2. **Create Store Listing:**
   - Fill out add-on information
   - Upload screenshots
   - Set visibility: "Private (domain only)"
   - Submit

3. **Users can now install from:**
   - Extensions > Add-ons > Get add-ons
   - Search "Client Manager"
   - Install

---

## 👥 Step 6: User Installation Guide

### For End Users

**Installation Steps:**

1. **Open Google Sheets** (any sheet)

2. **Click Extensions > Add-ons > Get add-ons**

3. **Search:** "Client Manager"
   - Or paste Deployment ID if using test deployment

4. **Click "Install"**

5. **Grant permissions:**
   - Click "Continue"
   - Review permissions
   - Click "Allow"

6. **First use:**
   - Open any Google Sheet
   - Click puzzle icon (🧩) in top right
   - Click "Client Manager"
   - Follow welcome setup

**That's it!** The add-on is now installed.

---

## 🔄 Step 7: Updating the Add-on

### Push Updates to Users

1. **Make changes in Apps Script**

2. **Save all files**

3. **Create new deployment version:**
   - Deploy > Manage deployments
   - Click pencil icon (edit)
   - Version: "New version"
   - Description: "Version 1.1 - Bug fixes"
   - Click "Deploy"

4. **Users get update automatically:**
   - Next time they open a sheet
   - Or after cache expires (~6 hours)

**No reinstallation needed!**

---

## 📊 Step 8: Data Structure for Each User

### Important: Data is Per-Spreadsheet

**Container-bound script:**
- Data tied to ONE spreadsheet
- Everyone shares same sheet

**Workspace Add-on:**
- Each user can use their own spreadsheet
- OR multiple users share one spreadsheet
- OR team has multiple project sheets

### Setup Options

**Option 1: Team Shared Sheet**
- Create one master Client Manager sheet
- Share with team (Editor access)
- Everyone uses same sheet
- All see same clients

**Option 2: Individual Sheets**
- Each tutor creates their own sheet
- Use add-on to set it up
- Private client data per tutor

**Option 3: Hybrid**
- Shared master client list
- Individual session notes sheets
- Link via client IDs

---

## 🔐 Permissions & Security

### OAuth Scopes Explained

Your add-on requests these permissions:

1. **Spreadsheets** (`auth/spreadsheets`)
   - Read/write client data
   - Create sheets if needed

2. **Gmail** (`auth/gmail.send`)
   - Send recap emails
   - Only sends, doesn't read inbox

3. **User Email** (`auth/userinfo.email`)
   - Identify tutor for personalization
   - Track who created records

4. **Script Container UI** (`auth/script.container.ui`)
   - Show sidebar/dialogs
   - Required for add-ons

### Security Best Practices

- ✅ Data stays in user's Google Drive
- ✅ No external servers/databases
- ✅ No data collection or analytics
- ✅ All processing happens in Google's infrastructure
- ✅ Users control access via Sheet permissions

---

## 📱 Features in Workspace Add-on

### What Works Different

**Container-bound:**
- Sidebar opens automatically
- Menu bar integration
- Single spreadsheet

**Workspace Add-on:**
- Click icon to open
- Works across spreadsheets
- Homepage when no sheet open
- CardService native UI

### New Features Available

✅ **Homepage card** - Shows when no sheet open
✅ **Create new sheet** - Add-on can set up structure
✅ **Multi-spreadsheet** - Works with any sheet
✅ **CardService UI** - Native Google components
✅ **Universal actions** - Refresh anywhere
✅ **Better mobile** - Native cards work on mobile

---

## 🐛 Troubleshooting

### Issue: Add-on doesn't appear after install

**Solution:**
1. Refresh the Google Sheet page
2. Check if puzzle icon (🧩) is visible
3. Wait 1-2 minutes for installation to complete
4. Try opening a different sheet

### Issue: Permission errors

**Solution:**
1. Uninstall add-on
2. Re-install
3. Carefully grant all permissions
4. Check "Allow" on all scopes

### Issue: "Add-on not configured"

**Solution:**
1. First time opening, complete welcome setup
2. Enter your name and email
3. Setup is saved per-user, not per-sheet

### Issue: Can't see clients from old sheet

**Solution:**
- Add-on is now multi-sheet
- Open the specific sheet with your client data
- Or migrate data to new sheet
- Each sheet is independent

### Issue: Users can't find add-on

**Solution:**

**For test deployment:**
- Send them exact Deployment ID
- They install via: Extensions > Add-ons > Enter ID

**For internal deployment:**
- Verify admin approved
- Check visibility settings (domain only)
- Search might take 24 hours to index

---

## 📈 Monitoring Usage

### Check Who's Using It

**As developer:**
1. Go to: https://console.cloud.google.com
2. Select project: Client Manager Add-on
3. View: APIs & Services > Credentials > OAuth consent screen
4. See: User grant statistics

**As admin:**
1. Google Admin Console
2. Apps > Google Workspace Marketplace apps
3. Find: Client Manager
4. View: Usage statistics

---

## 🎯 Next Steps After Deployment

### For Your Team

1. **Create documentation:**
   - How to install
   - How to use
   - Where data is stored
   - Support contact

2. **Set up shared sheet (optional):**
   - Create master Client Manager sheet
   - Share with team
   - Set as default

3. **Train users:**
   - Walk through first use
   - Show key features
   - Answer questions

### For Developers

1. **Monitor logs:**
   - Apps Script > Executions
   - Check for errors
   - Fix issues quickly

2. **Collect feedback:**
   - Create feedback form
   - Track feature requests
   - Plan updates

3. **Iterate:**
   - Push updates regularly
   - Add requested features
   - Improve based on usage

---

## 🔄 Migration from Container-Bound

### Moving Existing Data

**Your users have existing container-bound sheets:**

**Option 1: Keep using them**
- Install add-on
- Open existing sheet
- Add-on works with existing data
- No migration needed

**Option 2: Migrate to new structure**
1. Create new sheet via add-on
2. Export old data (CSV)
3. Import to new sheet
4. Verify everything works
5. Archive old sheet

---

## 📋 Checklist

### Pre-Deployment
- [ ] Created standalone Apps Script project
- [ ] Uploaded appsscript.json manifest
- [ ] Uploaded all .gs files
- [ ] Tested with test deployment
- [ ] Verified all features work
- [ ] Prepared documentation for users

### Deployment
- [ ] Created deployment (test or production)
- [ ] Copied Deployment ID
- [ ] Submitted for admin approval (if needed)
- [ ] Listed in internal marketplace (optional)
- [ ] Shared installation instructions with team

### Post-Deployment
- [ ] Monitored for errors
- [ ] Responded to user questions
- [ ] Collected feedback
- [ ] Planned updates

---

## 🎉 You're Done!

Your Client Manager is now a **standalone Workspace Add-on** that:

✅ Uses native CardService UI (fast, modern)
✅ Works across multiple spreadsheets
✅ Can be installed by internal users
✅ Updates automatically
✅ Works on mobile
✅ Follows Google's add-on best practices

**Files created:**
1. [appsscript.json](appsscript.json) - Manifest
2. [addon-lifecycle.gs](addon-lifecycle.gs) - Lifecycle hooks
3. Plus existing CardService files

**Next:** Follow Step 1 to create your standalone project!

---

## 📞 Support

**Google Documentation:**
- Workspace Add-ons: https://developers.google.com/workspace/add-ons
- CardService: https://developers.google.com/apps-script/reference/card-service
- Deployment: https://developers.google.com/workspace/add-ons/how-tos/publish-add-on

**Internal Support:**
- Email: [your-email]
- Slack: [your-channel]
- Docs: [your-wiki]

---

**Version:** 1.0
**Date:** January 9, 2025
**Status:** Ready for deployment
