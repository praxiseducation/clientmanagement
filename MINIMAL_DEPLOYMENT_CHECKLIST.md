# Minimal Client Manager - Quick Deployment Checklist

## 🎯 5-Minute Setup

### Step 1: Create Standalone Project (1 min)
- [ ] Go to https://script.google.com
- [ ] Click **New project**
- [ ] Rename to "Client Manager - Minimal"

### Step 2: Upload Code (2 min)
- [ ] Delete default Code.gs content
- [ ] Copy entire contents from `minimal-client-manager.gs`
- [ ] Paste into Code.gs
- [ ] Save (Ctrl+S)

### Step 3: Upload Manifest (1 min)
- [ ] Click **⚙️ Project Settings**
- [ ] Check: ✅ **Show "appsscript.json" manifest file**
- [ ] Return to **Editor** tab
- [ ] Click `appsscript.json` in file list
- [ ] Replace with contents from `minimal-appsscript.json`
- [ ] Save

### Step 4: Enable Calendar API (30 sec)
- [ ] In Apps Script editor, click **Services** (+) in left sidebar
- [ ] Select **Google Calendar API**
- [ ] Version: v3
- [ ] Click **Add**

### Step 5: Test Deploy (30 sec)
- [ ] Click **Deploy** → **Test deployments**
- [ ] Click **Install**
- [ ] Click **Done**

### Step 6: Test in Sheets (30 sec)
- [ ] Open any Google Sheet (or create new one)
- [ ] Click puzzle icon (🧩) in top right
- [ ] Click "Client Manager - Minimal"
- [ ] Should open with calendar setup screen!

---

## ✅ Verification Tests

After installation, verify these work:

- [ ] Calendar selection appears on first open
- [ ] Can select a calendar and save
- [ ] Today's appointments appear (if you have events)
- [ ] Search bar shows and accepts input
- [ ] Can click "Open Sheet" next to an appointment
- [ ] Sheet is created with correct structure
- [ ] Can view "All Sheets" list
- [ ] Can access Settings
- [ ] Can refresh appointments from Settings

---

## 📋 What You Get

**Core Features:**
✅ Predictive client search (autocomplete)
✅ Today's appointments from Google Calendar
✅ One-click sheet access (open or create)
✅ Alphabetical list of all client sheets
✅ Minimal metadata storage (emails only)
✅ Auto-email extraction from calendar events

**What's NOT Included (By Design):**
❌ No complex forms
❌ No rich HTML UI
❌ No session note forms (use sheet directly)
❌ No email sending UI (use mailto links)
❌ No analytics
❌ No file uploads

**Performance:**
- Single file: 901 lines
- Load time: <1 second
- Mobile-friendly: Native CardService UI
- Cache: 6-hour TTL for calendar data

---

## 🔧 Calendar Event Format

For automatic client detection, format your Google Calendar events as:

```
Title: Client Name: Session Description
Description: [optional notes]
email1@example.com
email2@example.com
```

**Example:**
```
Title: John Smith: SAT Math Prep
Description: Working on algebra review
john.smith@email.com
parent@email.com
```

**Auto-extracted:**
- Client name: "John Smith"
- Emails: john.smith@email.com, parent@email.com
- Time: From calendar event

---

## 📂 Optional: Set Storage Folder

To organize client sheets in a specific Google Drive folder:

1. Create folder in Google Drive
2. Open folder, copy ID from URL: `/folders/[THIS_PART]`
3. In add-on: **Settings** → **Set Storage Folder**
4. Paste folder ID
5. Future sheets created in that folder

---

## 🐛 Quick Troubleshooting

**Add-on doesn't appear after install**
→ Refresh sheet, wait 1-2 minutes

**Permission errors**
→ Re-install, grant all permissions

**No appointments showing**
→ Check calendar has events today in format "Name: Description"
→ Settings → Refresh Appointments

**Can't find old clients**
→ This version creates individual spreadsheets per client
→ Check Google Drive for "Client: [Name]" files

---

## 🚀 Deployment for Team

Once tested, deploy internally:

**Option 1: Test Deployment (Quick)**
1. Deploy → Test deployments → Copy ID
2. Share ID with team
3. They install via Extensions → Add-ons → Enter ID

**Option 2: Internal Deployment (Production)**
1. Deploy → New deployment → Type: Add-on
2. Visibility: "Only for users in [your org]"
3. Submit for admin approval
4. Users install from add-on gallery

See [MINIMAL_README.md](MINIMAL_README.md) for complete guide.

---

## ⏱️ Total Time: ~5 minutes

**Result:**
✅ Standalone Workspace Add-on
✅ CardService native UI
✅ Works across any spreadsheet
✅ Ready to distribute internally

---

**Next Step:** Follow Step 1 above to create your standalone project!
