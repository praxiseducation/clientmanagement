# Quick Notes UI Implementation Guide for Apps Script

## Overview
The redesigned Quick Notes UI matches the main sidebar's design system with consistent styling, improved organization, and better user experience.

## Key Design Changes

### 1. Visual Consistency
- Matches sidebar's Google Sans font family
- Uses same color scheme (#1a73e8 for primary, #5f6368 for text)
- Consistent button styling with hover effects and transitions
- Unified spacing and border radius (6px)

### 2. Improved Layout
- **Collapsible sections** for better organization:
  - Session Highlights (Wins, Struggles)
  - Skills Development (Mastered, Practiced, Introduced)
  - Communication & Planning (Parent notes, Next session)
- **Current client display** at top matching sidebar style
- **Action buttons** grouped logically

### 3. Enhanced Features
- **Auto-save indicator** shows when notes are being saved
- **Quick buttons** for common phrases (pills style)
- **Inline messages** for success/error feedback
- **Loading states** for async operations

## Implementation Steps

### Step 1: Update the HTML Template
Replace your current Quick Notes HTML with `QUICK_NOTES_REDESIGNED.html`

### Step 2: Update Server-side Functions
Ensure these functions exist in your Apps Script:

```javascript
// Get current client information
function getCurrentClientInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Check if it's a client sheet
  const clientListSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Client List');
  
  if (!clientListSheet) {
    return { isClient: false, name: '' };
  }
  
  const clientData = clientListSheet.getDataRange().getValues();
  for (let i = 1; i < clientData.length; i++) {
    if (clientData[i][0] === sheetName) {
      return { isClient: true, name: sheetName };
    }
  }
  
  return { isClient: false, name: '' };
}

// Save Quick Notes
function saveQuickNotes(notes) {
  const clientInfo = getCurrentClientInfo();
  if (!clientInfo.isClient) {
    throw new Error('No client selected');
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const props = PropertiesService.getDocumentProperties();
  const key = `quickNotes_${sheet.getName()}`;
  
  props.setProperty(key, JSON.stringify(notes));
  return { success: true };
}

// Get Quick Notes
function getQuickNotes() {
  const clientInfo = getCurrentClientInfo();
  if (!clientInfo.isClient) {
    return { data: '{}' };
  }
  
  const sheet = SpreadsheetApp.getActiveSheet();
  const props = PropertiesService.getDocumentProperties();
  const key = `quickNotes_${sheet.getName()}`;
  
  const data = props.getProperty(key) || '{}';
  return { data: data };
}

// Get Quick Notes Button Settings
function getQuickNotesButtons() {
  const props = PropertiesService.getUserProperties();
  const settings = props.getProperty('quickNotesButtons');
  
  // Default buttons if not configured
  const defaults = {
    wins: ['Great focus today', 'Excellent problem solving', 'Showed persistence'],
    skillsMastered: ['Reading comprehension', 'Basic math facts', 'Writing sentences'],
    skillsPracticed: ['Multiplication', 'Essay writing', 'Critical thinking'],
    skillsIntroduced: ['Algebra concepts', 'Research skills', 'Time management'],
    struggles: ['Maintaining focus', 'Following directions', 'Completing tasks'],
    parent: ['Please review homework', 'Great session today', 'Needs extra practice'],
    next: ['Continue current topic', 'Review homework', 'Start new chapter']
  };
  
  return settings ? JSON.parse(settings) : defaults;
}

// Show Quick Notes Settings Dialog
function showQuickNotesSettings() {
  const html = HtmlService.createHtmlOutputFromFile('QuickNotesSettings')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Quick Notes Settings');
}

// Show Recap Dialog
function showRecapDialog() {
  const html = HtmlService.createHtmlOutputFromFile('RecapDialog')
    .setWidth(600)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send Session Recap');
}
```

### Step 3: Create Settings Dialog (Optional)
Create `QuickNotesSettings.html` for customizing quick buttons:

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; }
    .section { margin-bottom: 20px; }
    .section-title { font-weight: 500; margin-bottom: 10px; color: #5f6368; }
    textarea { width: 100%; height: 80px; padding: 8px; border: 1px solid #dadce0; border-radius: 4px; }
    button { padding: 10px 20px; margin: 5px; border-radius: 4px; cursor: pointer; }
    .primary { background: #1a73e8; color: white; border: none; }
    .secondary { background: white; color: #1a73e8; border: 1px solid #1a73e8; }
  </style>
</head>
<body>
  <h2>Quick Notes Button Settings</h2>
  <p>Enter phrases for quick buttons (one per line)</p>
  
  <div class="section">
    <div class="section-title">Wins</div>
    <textarea id="wins" placeholder="Enter win phrases, one per line"></textarea>
  </div>
  
  <div class="section">
    <div class="section-title">Skills Mastered</div>
    <textarea id="skillsMastered" placeholder="Enter mastered skills, one per line"></textarea>
  </div>
  
  <!-- Add more sections as needed -->
  
  <button class="primary" onclick="saveSettings()">Save Settings</button>
  <button class="secondary" onclick="google.script.host.close()">Cancel</button>
  
  <script>
    function saveSettings() {
      const settings = {
        wins: document.getElementById('wins').value.split('\n').filter(s => s.trim()),
        skillsMastered: document.getElementById('skillsMastered').value.split('\n').filter(s => s.trim()),
        // Add more fields
      };
      
      google.script.run
        .withSuccessHandler(() => google.script.host.close())
        .saveQuickNotesSettings(settings);
    }
    
    // Load existing settings
    google.script.run
      .withSuccessHandler(settings => {
        if (settings.wins) document.getElementById('wins').value = settings.wins.join('\n');
        if (settings.skillsMastered) document.getElementById('skillsMastered').value = settings.skillsMastered.join('\n');
      })
      .getQuickNotesButtons();
  </script>
</body>
</html>
```

### Step 4: Integration with Main Sidebar

Update your main sidebar's Quick Notes button to show the new interface:

```javascript
function openQuickNotes() {
  const html = HtmlService.createHtmlOutputFromFile('QUICK_NOTES_REDESIGNED')
    .setTitle('Quick Notes')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}
```

## Benefits of the New Design

1. **Consistent UX**: Matches the main sidebar perfectly
2. **Better Organization**: Collapsible sections reduce clutter
3. **Improved Workflow**: Auto-save and quick buttons speed up note-taking
4. **Professional Look**: Clean, modern Google Material Design
5. **Responsive Feedback**: Clear loading states and messages
6. **Accessibility**: Proper labels and keyboard navigation

## Migration Notes

- The new design maintains compatibility with existing data structure
- Quick buttons settings are stored in UserProperties
- Notes continue to be stored in DocumentProperties
- No data migration required - works with existing saved notes

## Customization Options

You can easily customize:
- Section names and organization
- Quick button phrases
- Color scheme (update CSS variables)
- Auto-save delay (default 2 seconds)
- Textarea heights and layouts