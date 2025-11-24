# Client Management System - Sheet Dependencies & Template Specification

## Overview
This document outlines all sheet cell dependencies in the Client Management System, allowing teams to customize their spreadsheet appearance while maintaining system functionality.

## Core Dependencies

### 1. Client Sheet Requirements

#### **REQUIRED CELLS** (Only for initial display)
- **Cell A1**: Client/Student Name
  - **Purpose**: Initial client identifier (written once during sheet creation)
  - **Used by**: Display purposes only - system uses cached data
  - **Format**: Plain text string
  - **Example**: "John Smith"
  - **⚠️ Note**: After sheet creation, this cell is NOT READ by the system

#### **LEGACY CELLS** (No longer used - safe to modify/remove)
- **Cell C2**: Client Services/Types *(Legacy - not read)*
- **Cell D2**: Dashboard Link *(Legacy - not read)*
- **Cell E2**: Parent Email *(Legacy - not read)*
- **Cell F2**: Email Addressees *(Legacy - not read)*

**All client data is now stored in Google Apps Script Properties Service and accessed by sheet name only.**

### 2. System Sheets (DO NOT MODIFY)

#### **Template Sheet**
- **Required Name**: "NewClient"
- **Purpose**: Template for creating new client sheets
- **Note**: Must exist for client creation to work

#### **SessionRecaps Sheet** (Auto-created if missing)
- **Purpose**: Logs all sent session recaps
- **Columns**:
  1. Date/Time sent
  2. Student Name
  3. Client Type
  4. Parent Email
  5. Focus Skill
  6. Today's Win
  7. Homework
  8. Next Session
  9. Status

#### **Master Sheet** (Optional, legacy support)
- **Purpose**: Central client registry (being phased out)
- **Note**: System works without it, using UnifiedClientDataStore instead

### 3. Excluded Sheet Names
The following sheet names are reserved and will not be treated as client sheets:
- NewClient
- NewACTClient
- NewAcademicSupportClient
- Template
- Settings
- Dashboard
- Summary
- SessionRecaps
- Master
- Any sheet starting with underscore (_)

## Data Storage Architecture

### Primary Storage: UnifiedClientDataStore
- **Location**: Google Apps Script Properties Service
- **Format**: JSON structure
- **Contains**: All client metadata, Quick Notes, settings
- **Note**: Sheet cells are for display only; actual data lives in Properties

### What This Means for Customization
✅ **You CAN**:
- Add any cells beyond the required ones
- Format cells with colors, borders, fonts
- Add formulas in unused cells
- Add charts, images, or other elements
- Create custom layouts after row 2
- Add additional columns beyond F

❌ **You CANNOT**:
- Move or delete cell A1
- Change the sheet naming convention (uses concatenated client names)
- Use reserved sheet names
- Move cells C2-F2 if you want to use those features

## Template Setup Guide

### Minimum Viable Template
1. Create a sheet named "NewClient"
2. Optionally style cell A1 for client name display
3. Optionally add labels/formatting for cells C2-F2
4. Add any custom design elements you want

### Example Template Structure
```
A1: [Client Name - Required]
A2: "Services:"
C2: [Services - Optional]
A3: "Dashboard:"
D2: [Dashboard Link - Optional]
A4: "Parent Email:"
E2: [Parent Email - Optional]
A5: "Other Recipients:"
F2: [Email Addressees - Optional]

Row 7+: Your custom layout/data
```

## Quick Notes Storage
- **Location**: Properties Service (not in sheets)
- **Accessed by**: Current active sheet name
- **No cell dependencies**: Notes are stored separately from sheet data

## Session Recap Email System
- **Reads from**: UnifiedClientDataStore (cached data) using sheet name as identifier
- **Writes to**: SessionRecaps sheet (auto-created)
- **No cell dependencies**: All data comes from Properties Service cache
- **No modifications to client sheets**: Completely independent of sheet content

## Testing Your Template

### Validation Checklist
- [ ] Sheet named "NewClient" exists
- [ ] Cell A1 is available for client name
- [ ] Cells C2-F2 are available (if using those features)
- [ ] No formulas reference the required cells
- [ ] Sheet name doesn't conflict with reserved names

### Test Process
1. Create your custom "NewClient" template
2. Add a test client through the sidebar
3. Verify client name appears in A1
4. Test Quick Notes functionality
5. Test Session Recap sending (if using)

## Migration Path for Existing Spreadsheets

### If you have existing client sheets:
1. Ensure each has client name in A1
2. Rename template sheet to "NewClient"
3. Move any data in C2-F2 if needed
4. System will auto-detect and work with existing sheets

### If starting fresh:
1. Create "NewClient" template with your design
2. Use the sidebar to add clients
3. System handles everything else

## Best Practices

### For Maximum Compatibility
- Keep A1 for client name only
- Use C2-F2 for system fields if needed
- Start custom content from row 7
- Use columns G+ for custom data

### For Custom Workflows
- Add your own tracking columns after column F
- Create custom formulas that reference but don't modify A1
- Add conditional formatting based on your needs
- Include your organization's branding

## Support & Troubleshooting

### Common Issues
1. **"Template sheet 'NewClient' not found"**
   - Solution: Create a sheet named exactly "NewClient"

2. **Client name not showing in sidebar**
   - Check: Is the client name in cell A1?

3. **Emails not sending to parents**
   - Check: Is parent email in cell E2?

4. **Sheet treated as non-client**
   - Check: Is the sheet name in the excluded list?

## Version Compatibility
- **Version 3.0.0+**: Full support for customizable templates
- **Version 2.x**: Legacy support, may have additional dependencies
- **Version 1.x**: Not recommended, upgrade required

## Summary for Teams

### Absolute Minimum Requirements
1. Sheet named "NewClient" exists (template for copying)
2. Don't use reserved sheet names
3. That's it! No cell dependencies.

### Everything Else is 100% Customizable!
- **Complete design freedom**: All cells beyond A1 are yours
- **No data dependencies**: System doesn't read any cells after creation
- **Custom layouts**: Create any structure you want
- **Team branding**: Add logos, colors, formats
- **Custom formulas**: Use any cells for your calculations
- **Data tracking**: Add your own columns and tracking

## NEW: Zero Cell Dependencies! 🎉

**Version 3.0.0 Update**: The system now operates completely independently of sheet cell contents:

- **Client identification**: By sheet name only
- **Email data**: From Properties Service cache
- **Session data**: From Properties Service cache  
- **Quick Notes**: From Properties Service cache

Your team can now customize sheets with **complete freedom** - the add-on will work regardless of your sheet design!