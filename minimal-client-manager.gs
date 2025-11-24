/**
 * Minimal Client Manager - Bare Bones CardService Version
 *
 * Core Features:
 * - Predictive client search
 * - Today's appointments from Google Calendar
 * - Click to open/create client sheet
 * - List all client sheets
 * - Minimal metadata storage (emails, files)
 *
 * Functionality > Form
 *
 * DEBUG_ID: MINIMAL_CLIENT_MANAGER_20250109
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG_MINIMAL = {
  CALENDAR_ID_KEY: 'MINIMAL_CALENDAR_ID',
  CACHE_KEY: 'MINIMAL_APPOINTMENTS_CACHE',
  CACHE_TIMESTAMP_KEY: 'MINIMAL_CACHE_TIMESTAMP',
  METADATA_KEY: 'MINIMAL_STUDENT_METADATA',
  SHEET_PREFIX: 'Client: ',  // Prefix for client sheets
  FOLDER_ID_KEY: 'MINIMAL_CLIENT_FOLDER_ID'  // Google Drive folder for client sheets
};

// ============================================================================
// ADD-ON LIFECYCLE
// ============================================================================

/**
 * Homepage - main entry point
 */
function onHomepage_Minimal(e) {
  const calendarId = PropertiesService.getUserProperties().getProperty(CONFIG_MINIMAL.CALENDAR_ID_KEY);

  if (!calendarId) {
    return buildCalendarSetupCard();
  }

  return buildTodayDashboard();
}

/**
 * Universal action: refresh
 */
function refreshDashboard(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(buildTodayDashboard()))
    .build();
}

// ============================================================================
// MAIN DASHBOARD - TODAY'S CLIENTS
// ============================================================================

/**
 * Builds the main dashboard showing today's appointments
 */
function buildTodayDashboard() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Client Manager')
      .setSubtitle('Today\'s Sessions'));

  // Search section
  const searchSection = CardService.newCardSection();

  const searchInput = CardService.newTextInput()
    .setFieldName('client_search')
    .setTitle('Search Client')
    .setHint('Start typing name...')
    .setSuggestionsAction(CardService.newAction()
      .setFunctionName('getClientSuggestions_Minimal'))
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onClientSelected'));

  searchSection.addWidget(searchInput);
  card.addSection(searchSection);

  // Today's appointments section
  const todayClients = getTodaysClients_Minimal();
  const todaySection = CardService.newCardSection()
    .setHeader(`Today - ${todayClients.length} session${todayClients.length === 1 ? '' : 's'}`);

  if (todayClients.length === 0) {
    todaySection.addWidget(CardService.newTextParagraph()
      .setText('<i>No sessions scheduled for today</i>'));
  } else {
    todayClients.forEach((client, idx) => {
      // Extract time from formatted string
      const timeMatch = client.formattedDateTime.match(/at (.+)/);
      const timeOnly = timeMatch ? timeMatch[1] : '';

      const clientCard = CardService.newDecoratedText()
        .setText(`<b>${client.clientName}</b>`)
        .setBottomLabel(timeOnly)
        .setButton(CardService.newTextButton()
          .setText('Open Sheet')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('openClientSheet')
            .setParameters({ clientName: client.clientName })));

      todaySection.addWidget(clientCard);

      if (idx < todayClients.length - 1) {
        todaySection.addWidget(CardService.newDivider());
      }
    });
  }

  card.addSection(todaySection);

  // Action buttons
  const actionsSection = CardService.newCardSection();

  const buttonSet1 = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('➕ New Client')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showNewClientCard')))
    .addButton(CardService.newTextButton()
      .setText('📋 All Sheets')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showAllSheetsCard')));

  const buttonSet2 = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('⚙️ Settings')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showSettingsCard')));

  actionsSection.addWidget(buttonSet1);
  actionsSection.addWidget(buttonSet2);
  card.addSection(actionsSection);

  return card.build();
}

// ============================================================================
// CALENDAR INTEGRATION
// ============================================================================

/**
 * Gets today's clients from calendar
 */
function getTodaysClients_Minimal() {
  try {
    const appointments = getCachedAppointments_Minimal();

    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);

    return appointments
      .filter(apt => {
        const aptDate = new Date(apt.dateTime);
        return aptDate >= todayStart && aptDate < todayEnd;
      })
      .sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

  } catch (error) {
    console.error('Error getting today\'s clients:', error);
    return [];
  }
}

/**
 * Fetches appointments from Google Calendar
 */
function fetchCalendarAppointments_Minimal() {
  try {
    const calendarId = PropertiesService.getUserProperties().getProperty(CONFIG_MINIMAL.CALENDAR_ID_KEY);
    if (!calendarId) return [];

    const now = new Date();
    const startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 30);
    const endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 90);

    const events = Calendar.Events.list(calendarId, {
      timeMin: startDate.toISOString(),
      timeMax: endDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });

    if (!events.items || events.items.length === 0) return [];

    const appointments = [];

    events.items.forEach(event => {
      const summary = event.summary || '';
      const colonIndex = summary.indexOf(':');

      if (colonIndex > 0) {
        const clientName = summary.substring(0, colonIndex).trim();
        const startDateTime = event.start.dateTime || event.start.date;

        if (startDateTime) {
          const description = event.description || '';
          const emailMatches = description.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z]{2,})/gi);
          const email = emailMatches ? emailMatches.join(',') : null;

          appointments.push({
            clientName: clientName,
            dateTime: startDateTime,
            formattedDateTime: formatDateTime_Minimal(startDateTime),
            eventId: event.id,
            email: email
          });

          // Auto-store email in metadata
          if (email) {
            const metadata = getClientMetadata_Minimal(clientName);
            const emailList = email.split(',').map(e => e.trim());
            let updated = false;

            emailList.forEach(addr => {
              if (!metadata.emails.includes(addr)) {
                metadata.emails.push(addr);
                updated = true;
              }
            });

            if (updated) {
              saveClientMetadata_Minimal(clientName, metadata);
            }
          }
        }
      }
    });

    return appointments;

  } catch (error) {
    console.error('Error fetching calendar:', error);
    return [];
  }
}

/**
 * Caches appointments
 */
function cacheAppointments_Minimal(appointments) {
  const props = PropertiesService.getUserProperties();
  props.setProperty(CONFIG_MINIMAL.CACHE_KEY, JSON.stringify(appointments));
  props.setProperty(CONFIG_MINIMAL.CACHE_TIMESTAMP_KEY, new Date().toISOString());
}

/**
 * Gets cached appointments (refreshes if stale)
 */
function getCachedAppointments_Minimal() {
  const props = PropertiesService.getUserProperties();
  const cached = props.getProperty(CONFIG_MINIMAL.CACHE_KEY);
  const timestamp = props.getProperty(CONFIG_MINIMAL.CACHE_TIMESTAMP_KEY);

  // Check if cache is stale (> 6 hours)
  if (timestamp) {
    const cacheDate = new Date(timestamp);
    const now = new Date();
    const hoursDiff = (now - cacheDate) / (1000 * 60 * 60);

    if (hoursDiff < 6 && cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        // Fall through to refresh
      }
    }
  }

  // Refresh cache
  const appointments = fetchCalendarAppointments_Minimal();
  cacheAppointments_Minimal(appointments);
  return appointments;
}

/**
 * Formats date/time
 */
function formatDateTime_Minimal(isoDateTime) {
  const date = new Date(isoDateTime);
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  const dayOfWeek = days[date.getDay()];
  const month = months[date.getMonth()];
  const day = date.getDate();

  let hours = date.getHours();
  const minutes = date.getMinutes();
  const ampm = hours >= 12 ? 'PM' : 'AM';

  hours = hours % 12 || 12;
  const minutesStr = minutes < 10 ? '0' + minutes : minutes;

  return `${dayOfWeek}, ${month} ${day} at ${hours}:${minutesStr} ${ampm}`;
}

// ============================================================================
// CLIENT SEARCH & SUGGESTIONS
// ============================================================================

/**
 * Gets all clients from Clients sheet
 */
function getAllClients_Minimal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientSheet = ss.getSheetByName('Clients') || ss.getSheetByName('Client Database');

    if (!clientSheet) {
      return [];
    }

    const data = clientSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // No clients (only header row)
    }

    const headers = data[0];
    const clients = [];

    // Find column indices
    const nameCol = headers.indexOf('Name');
    const emailCol = headers.indexOf('Email');
    const parentEmailCol = headers.indexOf('Parent Email');
    const phoneCol = headers.indexOf('Phone');
    const gradeCol = headers.indexOf('Grade');
    const notesCol = headers.indexOf('Notes');
    const statusCol = headers.indexOf('Status');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows
      if (!row[nameCol] || row[nameCol].toString().trim() === '') {
        continue;
      }

      clients.push({
        name: row[nameCol],
        email: emailCol >= 0 ? row[emailCol] : '',
        parentEmail: parentEmailCol >= 0 ? row[parentEmailCol] : '',
        phone: phoneCol >= 0 ? row[phoneCol] : '',
        grade: gradeCol >= 0 ? row[gradeCol] : '',
        notes: notesCol >= 0 ? row[notesCol] : '',
        status: statusCol >= 0 ? row[statusCol] : 'Active',
        rowIndex: i + 1
      });
    }

    return clients;
  } catch (error) {
    Logger.log('Error getting clients: ' + error.message);
    return [];
  }
}

/**
 * Gets a client by name
 */
function getClientByName_Minimal(clientName) {
  const clients = getAllClients_Minimal();
  return clients.find(c => c.name === clientName) || null;
}

/**
 * Searches for clients by name only
 */
function searchClients_Minimal(searchTerm) {
  if (!searchTerm || searchTerm.trim().length < 1) {
    return [];
  }

  const clients = getAllClients_Minimal();
  const term = searchTerm.toLowerCase().trim();

  return clients.filter(client => {
    return client.name && client.name.toLowerCase().includes(term);
  });
}

/**
 * Provides autocomplete suggestions
 */
function getClientSuggestions_Minimal(e) {
  const query = (e && e.formInput && e.formInput.client_search) ?
    e.formInput.client_search.toLowerCase().trim() : '';

  if (!query) {
    return CardService.newSuggestionsResponseBuilder()
      .setSuggestions(CardService.newSuggestions())
      .build();
  }

  // Search clients from Clients sheet
  const matchingClients = searchClients_Minimal(query);
  const suggestions = CardService.newSuggestions();

  matchingClients.slice(0, 10).forEach(client => {
    suggestions.addSuggestion(client.name);
  });

  return CardService.newSuggestionsResponseBuilder()
    .setSuggestions(suggestions)
    .build();
}

/**
 * Called when client is selected from search
 */
function onClientSelected(e) {
  const clientName = (e && e.formInput && e.formInput.client_search) ?
    e.formInput.client_search.trim() : '';

  if (!clientName || clientName.length < 2) {
    return null;
  }

  // Show client detail card
  return showClientDetail({ parameters: { clientName: clientName }});
}

// ============================================================================
// CLIENT DETAIL CARD
// ============================================================================

/**
 * Shows client detail with upcoming sessions (chronological order)
 */
function showClientDetail(e) {
  const clientName = e.parameters.clientName;

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(clientName));

  // Upcoming sessions in chronological order
  const appointments = getCachedAppointments_Minimal()
    .filter(apt => apt.clientName === clientName)
    .sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

  if (appointments.length > 0) {
    const apptSection = CardService.newCardSection()
      .setHeader('Upcoming Sessions');

    appointments.slice(0, 10).forEach((apt, idx) => {
      // Show name and time
      const timeMatch = apt.formattedDateTime.match(/at (.+)/);
      const timeOnly = timeMatch ? timeMatch[1] : apt.formattedDateTime;

      apptSection.addWidget(CardService.newDecoratedText()
        .setText(apt.formattedDateTime)
        .setIcon(CardService.Icon.CLOCK));

      if (idx < Math.min(appointments.length, 10) - 1) {
        apptSection.addWidget(CardService.newDivider());
      }
    });

    card.addSection(apptSection);
  } else {
    const noApptSection = CardService.newCardSection();
    noApptSection.addWidget(CardService.newTextParagraph()
      .setText('<i>No upcoming sessions scheduled</i>'));
    card.addSection(noApptSection);
  }

  // Actions
  const actionsSection = CardService.newCardSection();

  actionsSection.addWidget(CardService.newTextButton()
    .setText('📄 Open/Create Sheet')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('openClientSheet')
      .setParameters({ clientName: clientName })));

  actionsSection.addWidget(CardService.newTextButton()
    .setText('← Back')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('refreshDashboard')));

  card.addSection(actionsSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

// ============================================================================
// NEW CLIENT DIALOG
// ============================================================================

/**
 * Shows new client card (CardService form)
 */
function showNewClientCard(e) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('New Client'));

  const section = CardService.newCardSection();

  // Name (required)
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_name')
    .setTitle('Client Name *')
    .setHint('Required'));

  // Email
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_email')
    .setTitle('Email')
    .setHint('Optional'));

  // Parent Email
  section.addWidget(CardService.newTextInput()
    .setFieldName('parent_email')
    .setTitle('Parent Email')
    .setHint('Optional'));

  // Phone
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_phone')
    .setTitle('Phone')
    .setHint('Optional'));

  // Grade
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_grade')
    .setTitle('Grade')
    .setHint('Optional'));

  // Notes
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_notes')
    .setTitle('Notes')
    .setHint('Optional')
    .setMultiline(true));

  card.addSection(section);

  // Actions
  const actionsSection = CardService.newCardSection();

  actionsSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('Save')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
        .setFunctionName('saveNewClient')))
    .addButton(CardService.newTextButton()
      .setText('Cancel')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('refreshDashboard'))));

  card.addSection(actionsSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Saves new client to Clients sheet
 */
function saveNewClient(e) {
  const formInput = e.formInput;

  // Validate required fields
  if (!formInput.client_name || formInput.client_name.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client name is required'))
      .build();
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let clientSheet = ss.getSheetByName('Clients') || ss.getSheetByName('Client Database');

    // Create Clients sheet if it doesn't exist
    if (!clientSheet) {
      clientSheet = ss.insertSheet('Clients');
      const headers = ['Name', 'Email', 'Parent Email', 'Phone', 'Grade', 'Notes', 'Status', 'Date Added'];
      clientSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      clientSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      clientSheet.setFrozenRows(1);
    }

    // Add new client
    const newRow = [
      formInput.client_name.trim(),
      formInput.client_email || '',
      formInput.parent_email || '',
      formInput.client_phone || '',
      formInput.client_grade || '',
      formInput.client_notes || '',
      'Active',
      new Date()
    ];

    clientSheet.appendRow(newRow);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client added: ' + formInput.client_name))
      .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildTodayDashboard()))
      .build();

  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// SHEET MANAGEMENT
// ============================================================================

/**
 * Opens or creates a client sheet (activates it in current spreadsheet)
 */
function openClientSheet(e) {
  const clientName = e.parameters.clientName;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check if sheet exists
    let sheet = getClientSheet(clientName);

    if (!sheet) {
      // Create new sheet
      sheet = createClientSheet(clientName);
    }

    // Activate the sheet
    const sheetObj = ss.getSheetByName(sheet.name);
    if (sheetObj) {
      ss.setActiveSheet(sheetObj);
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Opened ' + clientName + ' sheet'))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();

  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

/**
 * Gets client sheet if it exists in current spreadsheet
 */
function getClientSheet(clientName) {
  const sheetName = CONFIG_MINIMAL.SHEET_PREFIX + clientName;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    return {
      name: sheetName,
      sheetId: sheet.getSheetId()
    };
  }

  return null;
}

/**
 * Creates a new client sheet within the current spreadsheet
 */
function createClientSheet(clientName) {
  const sheetName = CONFIG_MINIMAL.SHEET_PREFIX + clientName;

  // Get current spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create new sheet within current spreadsheet
  const sheet = ss.insertSheet(sheetName);

  // Set up structure
  const headers = [
    'Date',
    'Homework Assigned',
    'Topics Covered',
    'Progress Notes',
    'Strengths',
    'Challenges',
    'Next Session Goals',
    'Parent Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // Auto-resize
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Store metadata
  const metadata = getClientMetadata_Minimal(clientName);
  metadata.sheetId = sheet.getSheetId();
  metadata.sheetName = sheetName;
  saveClientMetadata_Minimal(clientName, metadata);

  return {
    name: sheetName,
    sheetId: sheet.getSheetId()
  };
}

/**
 * Gets all client sheets in current spreadsheet
 */
function getAllClientSheets() {
  const sheets = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  // Find all sheets with client prefix
  allSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(CONFIG_MINIMAL.SHEET_PREFIX)) {
      const clientName = name.replace(CONFIG_MINIMAL.SHEET_PREFIX, '');
      sheets.push({
        name: name,
        clientName: clientName,
        sheetId: sheet.getSheetId()
      });
    }
  });

  // Sort alphabetically
  sheets.sort((a, b) => a.clientName.localeCompare(b.clientName));

  return sheets;
}

// ============================================================================
// ALL SHEETS CARD
// ============================================================================

/**
 * Shows all client sheets in alphabetical order
 */
function showAllSheetsCard(e) {
  const sheets = getAllClientSheets();

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('All Client Sheets')
      .setSubtitle(`${sheets.length} client${sheets.length === 1 ? '' : 's'}`));

  const section = CardService.newCardSection();

  if (sheets.length === 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<i>No client sheets yet. Search for a client to create one.</i>'));
  } else {
    sheets.forEach((sheet, idx) => {
      section.addWidget(CardService.newDecoratedText()
        .setText(`<b>${sheet.clientName}</b>`)
        .setButton(CardService.newTextButton()
          .setText('Open')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('openClientSheetFromList')
            .setParameters({ clientName: sheet.clientName }))));

      if (idx < sheets.length - 1) {
        section.addWidget(CardService.newDivider());
      }
    });
  }

  card.addSection(section);

  // Back button
  const actionsSection = CardService.newCardSection();
  actionsSection.addWidget(CardService.newTextButton()
    .setText('← Back to Dashboard')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('refreshDashboard')));

  card.addSection(actionsSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Opens sheet from all sheets list (activates it)
 */
function openClientSheetFromList(e) {
  const clientName = e.parameters.clientName;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG_MINIMAL.SHEET_PREFIX + clientName;
    const sheet = ss.getSheetByName(sheetName);

    if (sheet) {
      ss.setActiveSheet(sheet);
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Opened ' + clientName))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();

  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error opening sheet'))
      .build();
  }
}

// ============================================================================
// METADATA STORAGE
// ============================================================================

/**
 * Gets all client metadata
 */
function getAllClientMetadata_Minimal() {
  const props = PropertiesService.getUserProperties();
  const data = props.getProperty(CONFIG_MINIMAL.METADATA_KEY);
  return data ? JSON.parse(data) : {};
}

/**
 * Gets metadata for specific client
 */
function getClientMetadata_Minimal(clientName) {
  const all = getAllClientMetadata_Minimal();
  return all[clientName] || {
    emails: [],
    sheetId: null,
    sheetName: null
  };
}

/**
 * Saves metadata for client
 */
function saveClientMetadata_Minimal(clientName, metadata) {
  const all = getAllClientMetadata_Minimal();
  all[clientName] = metadata;
  PropertiesService.getUserProperties().setProperty(
    CONFIG_MINIMAL.METADATA_KEY,
    JSON.stringify(all)
  );
}

// ============================================================================
// SETTINGS
// ============================================================================

/**
 * Shows settings card
 */
function showSettingsCard(e) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Settings'));

  const section = CardService.newCardSection();

  // Change calendar button
  section.addWidget(CardService.newTextButton()
    .setText('Change Calendar')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showCalendarSetup')));

  // Refresh cache button
  section.addWidget(CardService.newTextButton()
    .setText('Refresh Appointments')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('manualRefreshCache')));

  card.addSection(section);

  // Back button
  const actionsSection = CardService.newCardSection();
  actionsSection.addWidget(CardService.newTextButton()
    .setText('← Back')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('refreshDashboard')));

  card.addSection(actionsSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Shows calendar setup
 */
function showCalendarSetup(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(buildCalendarSetupCard()))
    .build();
}

/**
 * Builds calendar setup card
 */
function buildCalendarSetupCard() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Calendar Setup'));

  const section = CardService.newCardSection();

  try {
    const currentCalendarId = PropertiesService.getUserProperties().getProperty(CONFIG_MINIMAL.CALENDAR_ID_KEY);
    const calendars = Calendar.CalendarList.list();

    if (calendars.items && calendars.items.length > 0) {
      const dropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName('calendar_id')
        .setTitle('Select Calendar');

      calendars.items.forEach(cal => {
        const label = cal.summary + (cal.primary ? ' (Primary)' : '');
        const selected = currentCalendarId ? (cal.id === currentCalendarId) : cal.primary;
        dropdown.addItem(label, cal.id, selected);
      });

      section.addWidget(dropdown);
    }

  } catch (error) {
    section.addWidget(CardService.newTextParagraph()
      .setText('Error loading calendars: ' + error.message));
  }

  card.addSection(section);

  // Save button
  const actionsSection = CardService.newCardSection();
  actionsSection.addWidget(CardService.newTextButton()
    .setText('Save')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveCalendarSelection_Minimal')));

  card.addSection(actionsSection);

  return card.build();
}

/**
 * Saves calendar selection
 */
function saveCalendarSelection_Minimal(e) {
  const calendarId = e.formInput.calendar_id;

  if (!calendarId) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Please select a calendar'))
      .build();
  }

  PropertiesService.getUserProperties().setProperty(CONFIG_MINIMAL.CALENDAR_ID_KEY, calendarId);

  // Clear cache to force refresh
  PropertiesService.getUserProperties().deleteProperty(CONFIG_MINIMAL.CACHE_KEY);

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Calendar saved'))
    .setNavigation(CardService.newNavigation().popToRoot().updateCard(buildTodayDashboard()))
    .build();
}

/**
 * Manual cache refresh
 */
function manualRefreshCache(e) {
  const appointments = fetchCalendarAppointments_Minimal();
  cacheAppointments_Minimal(appointments);

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`Refreshed ${appointments.length} appointments`))
    .setNavigation(CardService.newNavigation().popCard())
    .build();
}

