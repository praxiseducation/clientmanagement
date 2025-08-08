# Installation Guide

## Prerequisites

Before installing the Client Management System Beta 2, ensure you have:

- Google Workspace account (personal or business)
- Google Sheets access
- Basic familiarity with Google Apps Script (helpful but not required)

## Installation Methods

### Method 1: Direct Script Installation (Recommended)

1. **Open Google Sheets**
   - Create a new spreadsheet or open an existing one
   - Name it appropriately (e.g., "Client Management System")

2. **Access Apps Script Editor**
   - Go to `Extensions` → `Apps Script`
   - This opens the script editor in a new tab

3. **Install Main Application**
   - Delete any existing code in `Code.gs`
   - Copy the entire contents of `src/clientmanager.gs`
   - Paste into the script editor
   - Rename the file from `Code.gs` to `clientmanager.gs`

4. **Add Enhanced Dialogs Module**
   - Click the `+` next to Files
   - Select `Script`
   - Name it `enhanced_dialogs`
   - Copy contents of `src/enhanced_dialogs.js`
   - Paste into the new file

5. **Add Advanced Features Module**
   - Click the `+` next to Files again
   - Select `Script`
   - Name it `advanced_features`
   - Copy contents of `src/advanced_features.js`
   - Paste into the new file

6. **Save and Deploy**
   - Click the save icon or press `Ctrl+S` / `Cmd+S`
   - Close the Apps Script editor
   - Refresh your Google Sheet

7. **Initial Setup**
   - You should see a "Smart College" menu appear
   - Click `Smart College` → `Complete Setup`
   - Follow the setup wizard

### Method 2: Template Copy (If Available)

1. **Access Template**
   - Open the shared template link (if provided)
   - Click `File` → `Make a copy`
   - Name your copy appropriately

2. **Complete Setup**
   - The script will already be installed
   - Run `Smart College` → `Complete Setup`
   - Configure your personal settings

### Method 3: Add-on Installation (Future)

*Note: Add-on publication pending*

1. **Google Workspace Marketplace**
   - Search for "Smart College Client Manager"
   - Click Install
   - Grant necessary permissions
   - Complete initial setup

## Post-Installation Setup

### 1. Create Required Sheets

The system requires these sheets (created automatically on first client):

- **Master**: Central client registry
- **NewClient**: Template for new client sheets

### 2. Configure Permissions

When first running, you'll need to authorize the script:

1. Click `Smart College` menu
2. Select any option
3. Click "Review Permissions" in the authorization prompt
4. Sign in with your Google account
5. Click "Advanced" → "Go to Client Management (unsafe)"
6. Click "Allow"

*Note: "Unsafe" appears for unpublished scripts - this is normal*

### 3. Initial Configuration

Run the setup wizard to configure:

- **Tutor Information**
  - Your name
  - Your email address
  
- **Company Settings** (Optional)
  - Company name
  - Company website
  - Brand colors

- **Preferences**
  - Quick pill options
  - Auto-save settings
  - Theme preferences

### 4. Import Existing Clients (Optional)

If you have existing client data:

1. **From Another Sheet**
   - Use `Settings` → `Add Multiple Clients`
   - Follow the bulk import wizard

2. **From Individual Sheets**
   - Navigate to each client sheet
   - Use `Settings` → `Sync Active Sheet to Master`

## Troubleshooting

### Common Issues

**Menu Not Appearing**
- Refresh the spreadsheet (F5)
- Check Extensions → Apps Script for errors
- Ensure all files are saved

**Permission Errors**
- Re-run the authorization process
- Check that you have edit access to the spreadsheet
- Verify script permissions in Google Account settings

**Functions Not Working**
- Check the browser console for errors (F12)
- Ensure JavaScript is enabled
- Try using Chrome or Edge (recommended browsers)
- Clear browser cache and cookies

**Data Not Syncing**
- Check internet connection
- Verify auto-sync is enabled in preferences
- Manually trigger sync with `Refresh Cache`
- Check for conflicts in the error log

### Error Codes

- **E001**: Missing configuration - Run initial setup
- **E002**: Permission denied - Re-authorize the script
- **E003**: Data corruption - Run data integrity check
- **E004**: Network error - Check internet connection
- **E005**: Quota exceeded - Wait and retry later

## Updating

### From Beta 1.x to Beta 2.0

1. **Backup Current Data**
   - Export Master sheet as CSV
   - Save a copy of your spreadsheet

2. **Update Script Files**
   - Replace all three script files with Beta 2 versions
   - Do not delete user properties or script properties

3. **Run Migration**
   - The system will automatically detect and migrate old data
   - Monitor the execution log for any issues

4. **Verify Data**
   - Check that all clients appear correctly
   - Verify quick notes are preserved
   - Test sending a recap email

### Checking Version

To check your current version:
1. Open Apps Script editor
2. Search for "Version:" in clientmanager.gs
3. Check the README.md for version number

## Best Practices

### Performance Optimization

- **Regular Maintenance**
  - Run `Refresh Cache` weekly
  - Clear old session data monthly
  - Archive inactive clients quarterly

- **Data Management**
  - Keep under 500 active clients per sheet
  - Use pagination for large lists
  - Enable auto-sync for real-time updates

### Security

- **Access Control**
  - Only share with trusted users
  - Use view-only access for parents
  - Regularly review sharing settings

- **Data Protection**
  - Regular backups (weekly recommended)
  - Export critical data periodically
  - Use version history for recovery

## Support

For installation support:
1. Check the [Troubleshooting](#troubleshooting) section
2. Review the [FAQ](FAQ.md)
3. Contact support with:
   - Error messages/codes
   - Browser and OS version
   - Steps to reproduce issue
   - Screenshots if applicable

---

*Installation typically takes 10-15 minutes. For bulk installations or enterprise deployment, contact the development team.*