/**
 * Workspace Add-on Lifecycle Functions
 * Handles add-on installation, homepage, and CardService integration
 *
 * DEBUG_ID: ADDON_LIFECYCLE_20250109
 */

// ============================================================================
// ADD-ON LIFECYCLE TRIGGERS
// ============================================================================

/**
 * Called when the add-on is installed
 * Sets up initial configuration
 */
function onInstall(e) {
  Logger.log('Add-on installed');
  onOpen(e);
}

/**
 * Called when the add-on is opened (Sheets homepage)
 * This is the main entry point for the add-on
 */
function onSheetsHomepage(e) {
  Logger.log('Sheets homepage triggered');
  return buildHomepageCard(e);
}

/**
 * Generic homepage trigger (fallback)
 */
function onHomepage(e) {
  Logger.log('Generic homepage triggered');
  return buildHomepageCard(e);
}

/**
 * Called when file scope is granted (user gives permission to access specific sheet)
 */
function onFileScopeGranted(e) {
  Logger.log('File scope granted');
  return buildMainClientCard(e);
}

// ============================================================================
// HOMEPAGE CARD
// ============================================================================

/**
 * Builds the homepage card shown when add-on is opened
 */
function buildHomepageCard(e) {
  Logger.log('Building homepage card');

  // Check if user is configured
  const userConfig = getUserConfig();

  if (!userConfig.isConfigured) {
    return buildWelcomeCard();
  }

  // Check if we have access to active spreadsheet
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (spreadsheet) {
      return buildMainClientCard(e);
    }
  } catch (error) {
    Logger.log('No active spreadsheet, showing file selector: ' + error.message);
  }

  // No spreadsheet open - show file selector card
  return buildFileSelectorCard();
}

/**
 * Welcome card for first-time users
 */
function buildWelcomeCard() {
  const section = CardService.newCardSection()
    .setHeader('👋 Welcome to Client Manager!');

  section.addWidget(CardService.newTextParagraph()
    .setText('Get started by configuring your tutor profile.'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('tutor_name')
    .setTitle('Your Name')
    .setHint('e.g., John Smith'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('tutor_email')
    .setTitle('Your Email')
    .setHint('e.g., john@smartcollege.com'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('company_name')
    .setTitle('Company Name')
    .setHint('e.g., Smart College')
    .setValue('Smart College'));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Client Manager')
      .setSubtitle('Initial Setup'))
    .addSection(section)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Complete Setup')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('saveWelcomeSetup'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)))
    .build();

  return card;
}

/**
 * Saves welcome setup configuration
 */
function saveWelcomeSetup(e) {
  const formInput = e.formInput;

  if (!formInput.tutor_name || !formInput.tutor_email) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Name and email are required'))
      .build();
  }

  const config = {
    tutorName: formInput.tutor_name,
    tutorEmail: formInput.tutor_email,
    companyName: formInput.company_name || 'Smart College',
    companyUrl: 'https://www.smartcollege.com'
  };

  saveInitialSetup(config);

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText('Setup complete! 🎉'))
    .setNavigation(CardService.newNavigation()
      .updateCard(buildFileSelectorCard()))
    .build();
}

/**
 * File selector card - shown when no spreadsheet is open
 */
function buildFileSelectorCard() {
  const section = CardService.newCardSection()
    .setHeader('📊 Open a Spreadsheet');

  section.addWidget(CardService.newTextParagraph()
    .setText('To use Client Manager, please open a Google Sheet where you want to store client data.'));

  section.addWidget(CardService.newTextParagraph()
    .setText('<b>Options:</b><br>' +
            '• Open an existing sheet with client data<br>' +
            '• Create a new sheet for your clients'));

  section.addWidget(CardService.newTextButton()
    .setText('📝 Create New Sheet')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('createNewClientSheet'))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Client Manager')
      .setSubtitle('No spreadsheet open'))
    .addSection(section)
    .build();

  return card;
}

/**
 * Creates a new client sheet
 */
function createNewClientSheet(e) {
  try {
    // Create new spreadsheet
    const ss = SpreadsheetApp.create('Client Manager - ' + new Date().toLocaleDateString());

    // Set up initial structure
    setupClientSheet(ss);

    const url = ss.getUrl();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Sheet created! Opening...'))
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
 * Sets up a new client sheet with proper structure
 */
function setupClientSheet(spreadsheet) {
  // Create Clients sheet
  let clientSheet = spreadsheet.getSheetByName('Sheet1');
  if (clientSheet) {
    clientSheet.setName('Clients');
  } else {
    clientSheet = spreadsheet.insertSheet('Clients');
  }

  // Set up headers
  const headers = [
    'Name',
    'Email',
    'Parent Email',
    'Phone',
    'Grade',
    'Notes',
    'Status',
    'Date Added'
  ];

  clientSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  clientSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  clientSheet.setFrozenRows(1);

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    clientSheet.autoResizeColumn(i);
  }

  // Create Session Notes sheet
  const notesSheet = spreadsheet.insertSheet('Session Notes');
  const notesHeaders = [
    'Date',
    'Client Name',
    'Homework',
    'Topics',
    'Progress',
    'Strengths',
    'Challenges',
    'Next Session',
    'Parent Notes'
  ];

  notesSheet.getRange(1, 1, 1, notesHeaders.length).setValues([notesHeaders]);
  notesSheet.getRange(1, 1, 1, notesHeaders.length).setFontWeight('bold');
  notesSheet.setFrozenRows(1);

  // Create Recap History sheet
  const historySheet = spreadsheet.insertSheet('Recap History');
  const historyHeaders = ['Date', 'Client Name', 'Recipient', 'Subject'];

  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
  historySheet.getRange(1, 1, 1, historyHeaders.length).setFontWeight('bold');
  historySheet.setFrozenRows(1);

  // Create named range for current client (in a separate Config sheet)
  const configSheet = spreadsheet.insertSheet('Config');
  configSheet.hideSheet();
  configSheet.getRange('A1').setValue('CurrentClient');
  configSheet.getRange('B1').setValue('');

  spreadsheet.setNamedRange('CurrentClient', configSheet.getRange('B1'));

  Logger.log('Client sheet structure created');
}

// ============================================================================
// MAIN CLIENT CARD (When spreadsheet is open)
// ============================================================================

/**
 * Builds the main client management card
 * This replaces showCardServiceSidebar() for Workspace Add-ons
 */
function buildMainClientCard(e) {
  Logger.log('Building main client card');

  try {
    // Get context
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const clientInfo = getCurrentClientInfo();
    const userConfig = getUserConfig();

    // Use the existing CardService UI implementation
    return showCardServiceSidebar();

  } catch (error) {
    Logger.log('Error building main card: ' + error.message);
    return buildErrorCard(error.message);
  }
}

/**
 * Error card for troubleshooting
 */
function buildErrorCard(errorMessage) {
  const section = CardService.newCardSection()
    .setHeader('❌ Error');

  section.addWidget(CardService.newTextParagraph()
    .setText('An error occurred: ' + errorMessage));

  section.addWidget(CardService.newTextButton()
    .setText('Retry')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('refreshCard')));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Client Manager'))
    .addSection(section)
    .build();
}

// ============================================================================
// UNIVERSAL ACTIONS
// ============================================================================

/**
 * Refreshes the current card
 */
function refreshCard(e) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(buildMainClientCard(e)))
    .build();
}

// ============================================================================
// CONTEXT FUNCTIONS (for add-on compatibility)
// ============================================================================

/**
 * Gets context information from event object
 * Add-ons receive context differently than container-bound scripts
 */
function getAddonContext(e) {
  const context = {
    hasSpreadsheet: false,
    spreadsheetId: null,
    sheetId: null,
    userEmail: Session.getActiveUser().getEmail()
  };

  try {
    // Try to get spreadsheet from event
    if (e && e.docs && e.docs.id) {
      context.hasSpreadsheet = true;
      context.spreadsheetId = e.docs.id;
    }

    // Try to get active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      context.hasSpreadsheet = true;
      context.spreadsheetId = ss.getId();

      const sheet = ss.getActiveSheet();
      if (sheet) {
        context.sheetId = sheet.getSheetId();
      }
    }
  } catch (error) {
    Logger.log('Error getting context: ' + error.message);
  }

  return context;
}

/**
 * Gets current client info (add-on version)
 * Handles cases where spreadsheet may not be accessible
 */
function getCurrentClientInfoSafe() {
  try {
    return getCurrentClientInfo();
  } catch (error) {
    Logger.log('Error getting current client: ' + error.message);
    return {
      isClient: false,
      name: '',
      email: ''
    };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Saves initial setup (overload for add-on)
 */
function saveInitialSetup(config) {
  const userProperties = PropertiesService.getUserProperties();

  userProperties.setProperties({
    tutorName: config.tutorName,
    tutorEmail: config.tutorEmail,
    companyName: config.companyName || 'Smart College',
    companyUrl: config.companyUrl || 'https://www.smartcollege.com',
    isConfigured: 'true'
  });

  Logger.log('Initial setup saved');
  return { success: true };
}

/**
 * Gets user configuration (shared with container-bound version)
 */
function getUserConfig() {
  const userProperties = PropertiesService.getUserProperties();
  return {
    tutorName: userProperties.getProperty('tutorName') || '',
    tutorEmail: userProperties.getProperty('tutorEmail') || '',
    companyName: userProperties.getProperty('companyName') || 'Smart College',
    companyUrl: userProperties.getProperty('companyUrl') || 'https://www.smartcollege.com',
    isConfigured: userProperties.getProperty('isConfigured') === 'true'
  };
}

/**
 * Gets user role (for add-on context)
 */
function getUserRole() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (ss) {
      const owner = ss.getOwner();
      if (owner && owner.getEmail() === userEmail) {
        return 'owner';
      }
    }

    const userConfig = getUserConfig();
    return userConfig.role || 'tutor';

  } catch (error) {
    return 'tutor';
  }
}

/**
 * Gets current client info (add-on safe version)
 */
function getCurrentClientInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentClient = ss.getRangeByName('CurrentClient');

    if (!currentClient) {
      return { isClient: false };
    }

    const clientName = currentClient.getValue();

    if (clientName && clientName.trim() !== '') {
      const client = getClientByName(clientName);
      return {
        isClient: true,
        name: clientName,
        email: client ? client.email : '',
        id: clientName
      };
    }

    return { isClient: false };

  } catch (error) {
    Logger.log('Error getting current client: ' + error.message);
    return { isClient: false };
  }
}
