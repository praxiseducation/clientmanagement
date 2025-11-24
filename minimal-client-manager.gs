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

  const buttonSet = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('📋 All Sheets')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showAllSheetsCard')))
    .addButton(CardService.newTextButton()
      .setText('⚙️ Settings')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showSettingsCard')));

  actionsSection.addWidget(buttonSet);
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

  // Get unique client names from calendar and sheets
  const appointments = getCachedAppointments_Minimal();
  const clientNames = new Set();

  appointments.forEach(apt => {
    if (apt.clientName.toLowerCase().includes(query)) {
      clientNames.add(apt.clientName);
    }
  });

  // Also check existing sheets
  const sheets = getAllClientSheets();
  sheets.forEach(sheet => {
    const name = sheet.name.replace(CONFIG_MINIMAL.SHEET_PREFIX, '');
    if (name.toLowerCase().includes(query)) {
      clientNames.add(name);
    }
  });

  const suggestions = CardService.newSuggestions();
  Array.from(clientNames).slice(0, 10).forEach(name => {
    suggestions.addSuggestion(name);
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
 * Shows client detail with upcoming appointments
 */
function showClientDetail(e) {
  const clientName = e.parameters.clientName;
  const metadata = getClientMetadata_Minimal(clientName);

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(clientName));

  // Email section
  if (metadata.emails && metadata.emails.length > 0) {
    const emailSection = CardService.newCardSection();
    metadata.emails.forEach(email => {
      emailSection.addWidget(CardService.newDecoratedText()
        .setText(email)
        .setIcon(CardService.Icon.EMAIL)
        .setButton(CardService.newImageButton()
          .setIcon(CardService.Icon.EMAIL)
          .setOpenLink(CardService.newOpenLink()
            .setUrl('mailto:' + email)
            .setOpenAs(CardService.OpenAs.FULL_SIZE))));
    });
    card.addSection(emailSection);
  }

  // Upcoming appointments
  const appointments = getCachedAppointments_Minimal()
    .filter(apt => apt.clientName === clientName)
    .sort((a, b) => new Date(a.dateTime) - new Date(b.dateTime));

  if (appointments.length > 0) {
    const apptSection = CardService.newCardSection()
      .setHeader('Upcoming Sessions');

    appointments.slice(0, 5).forEach((apt, idx) => {
      apptSection.addWidget(CardService.newTextParagraph()
        .setText(apt.formattedDateTime));

      if (idx < Math.min(appointments.length, 5) - 1) {
        apptSection.addWidget(CardService.newDivider());
      }
    });

    card.addSection(apptSection);
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
// SHEET MANAGEMENT
// ============================================================================

/**
 * Opens or creates a client sheet
 */
function openClientSheet(e) {
  const clientName = e.parameters.clientName;

  try {
    // Check if sheet exists
    let sheet = getClientSheet(clientName);

    if (!sheet) {
      // Create new sheet
      sheet = createClientSheet(clientName);
    }

    // Use URL from sheet object (no need to open by ID)
    const url = sheet.url || ('https://docs.google.com/spreadsheets/d/' + sheet.spreadsheetId + '/edit');

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Opening ' + clientName + ' sheet...'))
      .setOpenLink(CardService.newOpenLink()
        .setUrl(url)
        .setOpenAs(CardService.OpenAs.FULL_SIZE))
      .build();

  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

/**
 * Gets client sheet if it exists
 */
function getClientSheet(clientName) {
  const sheetName = CONFIG_MINIMAL.SHEET_PREFIX + clientName;
  const sheets = getAllClientSheets();

  return sheets.find(s => s.name === sheetName);
}

/**
 * Creates a new client sheet
 */
function createClientSheet(clientName) {
  const sheetName = CONFIG_MINIMAL.SHEET_PREFIX + clientName;

  // Create new spreadsheet
  const ss = SpreadsheetApp.create(sheetName);
  const sheet = ss.getActiveSheet();
  sheet.setName('Session Notes');

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

  // Move to folder if configured
  const folderId = PropertiesService.getUserProperties().getProperty(CONFIG_MINIMAL.FOLDER_ID_KEY);
  if (folderId) {
    try {
      const file = DriveApp.getFileById(ss.getId());
      const folder = DriveApp.getFolderById(folderId);
      file.moveTo(folder);
    } catch (e) {
      // Ignore folder errors
    }
  }

  // Store metadata
  const metadata = getClientMetadata_Minimal(clientName);
  metadata.spreadsheetId = ss.getId();
  metadata.spreadsheetUrl = ss.getUrl();
  saveClientMetadata_Minimal(clientName, metadata);

  return {
    name: sheetName,
    spreadsheetId: ss.getId(),
    url: ss.getUrl()
  };
}

/**
 * Gets all client sheets
 */
function getAllClientSheets() {
  const sheets = [];
  const metadata = getAllClientMetadata_Minimal();

  // From metadata
  Object.keys(metadata).forEach(clientName => {
    if (metadata[clientName].spreadsheetId) {
      sheets.push({
        name: CONFIG_MINIMAL.SHEET_PREFIX + clientName,
        clientName: clientName,
        spreadsheetId: metadata[clientName].spreadsheetId,
        url: metadata[clientName].spreadsheetUrl
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
            .setParameters({ spreadsheetId: sheet.spreadsheetId, clientName: sheet.clientName }))));

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
 * Opens sheet from all sheets list
 */
function openClientSheetFromList(e) {
  const spreadsheetId = e.parameters.spreadsheetId;
  const clientName = e.parameters.clientName;

  try {
    // Get URL from metadata instead of opening file
    const metadata = getClientMetadata_Minimal(clientName);
    const url = metadata.spreadsheetUrl;

    if (!url) {
      // Fallback: construct URL from ID
      const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit';
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Opening ' + clientName + '...'))
      .setOpenLink(CardService.newOpenLink()
        .setUrl(url)
        .setOpenAs(CardService.OpenAs.FULL_SIZE))
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
    files: [],
    spreadsheetId: null,
    spreadsheetUrl: null
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

  // Set folder button
  section.addWidget(CardService.newTextButton()
    .setText('Set Storage Folder')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showFolderSetup')));

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

/**
 * Shows folder setup
 */
function showFolderSetup(e) {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Storage Folder'));

  const section = CardService.newCardSection();

  section.addWidget(CardService.newTextParagraph()
    .setText('Enter the ID of the Google Drive folder where client sheets should be stored.\n\nTo get folder ID: Open folder in Drive, copy the last part of the URL after /folders/'));

  const currentFolder = PropertiesService.getUserProperties().getProperty(CONFIG_MINIMAL.FOLDER_ID_KEY);

  section.addWidget(CardService.newTextInput()
    .setFieldName('folder_id')
    .setTitle('Folder ID')
    .setHint('abc123xyz...')
    .setValue(currentFolder || ''));

  card.addSection(section);

  // Save button
  const actionsSection = CardService.newCardSection();
  actionsSection.addWidget(CardService.newTextButton()
    .setText('Save')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveFolderSelection')));

  actionsSection.addWidget(CardService.newTextButton()
    .setText('Cancel')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showSettingsCard')));

  card.addSection(actionsSection);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().pushCard(card.build()))
    .build();
}

/**
 * Saves folder selection
 */
function saveFolderSelection(e) {
  const folderId = e.formInput.folder_id;

  if (folderId && folderId.trim()) {
    try {
      // Validate folder exists
      DriveApp.getFolderById(folderId.trim());
      PropertiesService.getUserProperties().setProperty(CONFIG_MINIMAL.FOLDER_ID_KEY, folderId.trim());

      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText('Folder saved'))
        .setNavigation(CardService.newNavigation().popCard())
        .build();

    } catch (error) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
          .setText('Invalid folder ID'))
        .build();
    }
  } else {
    // Clear folder
    PropertiesService.getUserProperties().deleteProperty(CONFIG_MINIMAL.FOLDER_ID_KEY);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Folder cleared'))
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  }
}
