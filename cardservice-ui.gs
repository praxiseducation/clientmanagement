/**
 * CardService UI Implementation - Native Google Workspace Components
 * This replaces the HTML/CSS sidebar with native CardService widgets
 * Part of Hybrid Approach: CardService for navigation, HTML for Quick Notes
 *
 * DEBUG_ID: CARDSERVICE_UI_20250109
 */

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

/**
 * Shows the main CardService sidebar
 * Called from onOpen() and menu items
 */
function showCardServiceSidebar() {
  const clientInfo = getCurrentClientInfo();
  const userRole = getUserRole();
  const userInfo = getUserInfo();

  const card = buildMainCard(clientInfo, userRole, userInfo);
  return card;
}

/**
 * Builds the main control panel card
 */
function buildMainCard(clientInfo, userRole, userInfo) {
  const builder = CardService.newCardBuilder();

  // Header with branding
  builder.setHeader(buildMainHeader(userRole, userInfo));

  // Current client section
  builder.addSection(buildCurrentClientSection(clientInfo));

  // Client management section
  builder.addSection(buildClientManagementSection(clientInfo));

  // Session management section
  builder.addSection(buildSessionManagementSection(clientInfo));

  // Admin section (only for admins)
  if (userRole === 'admin' || userRole === 'owner') {
    builder.addSection(buildAdminSection());
  }

  // Footer with version info
  builder.setFixedFooter(buildFooter());

  return builder.build();
}

// ============================================================================
// HEADER COMPONENTS
// ============================================================================

/**
 * Builds the main card header
 */
function buildMainHeader(userRole, userInfo) {
  const config = CONFIG || {company: 'Smart College'};
  let subtitle = config.company;

  if (userInfo && userInfo.displayName) {
    subtitle = `${userInfo.displayName} - ${userRole}`;
  } else if (userRole) {
    subtitle = `${config.company} - ${userRole}`;
  }

  return CardService.newCardHeader()
    .setTitle('Client Manager')
    .setSubtitle(subtitle)
    .setImageStyle(CardService.ImageStyle.CIRCLE);
}

/**
 * Builds the footer with version info
 */
function buildFooter() {
  const config = CONFIG || {version: '3.0.0', build: '2025-01-09'};
  return CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText('Refresh')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('refreshMainCard')));
}

// ============================================================================
// CURRENT CLIENT SECTION
// ============================================================================

/**
 * Displays the currently selected client
 */
function buildCurrentClientSection(clientInfo) {
  const section = CardService.newCardSection()
    .setHeader('📍 Current Client');

  if (clientInfo && clientInfo.isClient) {
    // Client is selected
    section.addWidget(CardService.newDecoratedText()
      .setText(clientInfo.name)
      .setBottomLabel(clientInfo.email || 'No email')
      .setIcon(CardService.Icon.PERSON)
      .setWrapText(true)
      .setButton(CardService.newTextButton()
        .setText('Clear')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('clearCurrentClient'))));
  } else {
    // No client selected
    section.addWidget(CardService.newDecoratedText()
      .setText('No client selected')
      .setIcon(CardService.Icon.PERSON)
      .setBottomLabel('Search or select a client below'));
  }

  return section;
}

/**
 * Clears the current client selection
 */
function clearCurrentClient() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getRangeByName('CurrentClient').setValue('');

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client cleared'))
      .setStateChanged(true)
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// CLIENT MANAGEMENT SECTION
// ============================================================================

/**
 * Builds the client management section with search and actions
 */
function buildClientManagementSection(clientInfo) {
  const section = CardService.newCardSection()
    .setHeader('👥 Client Management');

  // Search input with suggestions
  section.addWidget(CardService.newTextInput()
    .setFieldName('client_search')
    .setTitle('Search Clients')
    .setHint('Type client name or email...')
    .setOnChangeAction(CardService.newAction()
      .setFunctionName('onClientSearchChange')));

  // Add New Client button
  section.addWidget(CardService.newTextButton()
    .setText('➕ Add New Client')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showAddClientCard'))
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));

  // View All Clients button
  section.addWidget(CardService.newTextButton()
    .setText('📋 View All Clients')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showClientListCard')));

  // Update Client button (only if client is selected)
  if (clientInfo && clientInfo.isClient) {
    section.addWidget(CardService.newTextButton()
      .setText('⚙️ Update Client Info')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showUpdateClientCard')
        .setParameters({clientId: clientInfo.id || clientInfo.name})));
  }

  return section;
}

/**
 * Handles client search input changes
 */
function onClientSearchChange(e) {
  const searchTerm = e.formInput.client_search;

  if (!searchTerm || searchTerm.trim().length < 2) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Type at least 2 characters to search'))
      .build();
  }

  // Search for clients
  const results = searchClientsInSheet(searchTerm);

  if (results.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('No clients found'))
      .build();
  }

  // Show results in a new card
  return showSearchResultsCard(results);
}

/**
 * Displays search results in a new card
 */
function showSearchResultsCard(results) {
  const section = CardService.newCardSection()
    .setHeader(`${results.length} client${results.length === 1 ? '' : 's'} found`);

  // Add each result as a selectable item
  results.slice(0, 20).forEach(client => { // Limit to 20 results
    section.addWidget(CardService.newDecoratedText()
      .setText(client.name)
      .setBottomLabel(client.email || 'No email')
      .setIcon(CardService.Icon.PERSON)
      .setButton(CardService.newTextButton()
        .setText('Select')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('selectClientFromSearch')
          .setParameters({
            clientName: client.name,
            clientEmail: client.email || ''
          }))));
  });

  if (results.length > 20) {
    section.addWidget(CardService.newTextParagraph()
      .setText(`<i>Showing first 20 of ${results.length} results. Refine your search for more precision.</i>`));
  }

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Search Results'))
    .addSection(section)
    .build();

  return CardService.newNavigation().pushCard(card);
}

/**
 * Selects a client from search results
 */
function selectClientFromSearch(e) {
  const clientName = e.parameters.clientName;
  const clientEmail = e.parameters.clientEmail;

  try {
    // Set current client
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getRangeByName('CurrentClient').setValue(clientName);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Selected: ${clientName}`))
      .setNavigation(CardService.newNavigation().popCard())
      .setStateChanged(true)
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// CLIENT LIST CARD
// ============================================================================

/**
 * Shows the full client list with filters
 */
function showClientListCard(e) {
  const filter = (e && e.parameters && e.parameters.filter) || 'all';
  const clients = getAllClients(); // Your existing function

  let filteredClients = clients;
  if (filter === 'active') {
    filteredClients = clients.filter(c => c.status === 'Active' || !c.status);
  } else if (filter === 'inactive') {
    filteredClients = clients.filter(c => c.status === 'Inactive');
  }

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Client List')
      .setSubtitle(`${filteredClients.length} clients`));

  // Filter buttons
  const filterSection = CardService.newCardSection()
    .setHeader('Filter By');

  const filterButtonSet = CardService.newButtonSet()
    .addButton(CardService.newTextButton()
      .setText('All')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showClientListCard')
        .setParameters({filter: 'all'}))
      .setDisabled(filter === 'all'))
    .addButton(CardService.newTextButton()
      .setText('Active')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showClientListCard')
        .setParameters({filter: 'active'}))
      .setDisabled(filter === 'active'))
    .addButton(CardService.newTextButton()
      .setText('Inactive')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('showClientListCard')
        .setParameters({filter: 'inactive'}))
      .setDisabled(filter === 'inactive'));

  filterSection.addWidget(filterButtonSet);
  builder.addSection(filterSection);

  // Client list
  const clientSection = CardService.newCardSection();

  if (filteredClients.length === 0) {
    clientSection.addWidget(CardService.newTextParagraph()
      .setText('<i>No clients found</i>'));
  } else {
    filteredClients.slice(0, 50).forEach(client => { // Limit to 50 for performance
      clientSection.addWidget(CardService.newDecoratedText()
        .setText(client.name)
        .setBottomLabel(client.email || 'No email')
        .setIcon(CardService.Icon.PERSON)
        .setButton(CardService.newTextButton()
          .setText('Select')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('selectClientFromSearch')
            .setParameters({
              clientName: client.name,
              clientEmail: client.email || ''
            }))));
    });

    if (filteredClients.length > 50) {
      clientSection.addWidget(CardService.newTextParagraph()
        .setText(`<i>Showing first 50 of ${filteredClients.length} clients</i>`));
    }
  }

  builder.addSection(clientSection);

  const card = builder.build();
  return CardService.newNavigation().pushCard(card);
}

// ============================================================================
// ADD CLIENT CARD
// ============================================================================

/**
 * Shows the add new client form
 */
function showAddClientCard() {
  const section = CardService.newCardSection()
    .setHeader('Client Information');

  section.addWidget(CardService.newTextInput()
    .setFieldName('client_name')
    .setTitle('Client Name *')
    .setHint('Full name of client'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('client_email')
    .setTitle('Email')
    .setHint('client@example.com'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('parent_email')
    .setTitle('Parent Email')
    .setHint('parent@example.com'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('phone')
    .setTitle('Phone')
    .setHint('(555) 123-4567'));

  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('grade')
    .setTitle('Grade')
    .addItem('Pre-K', 'pre-k', false)
    .addItem('Kindergarten', 'k', false)
    .addItem('1st Grade', '1', false)
    .addItem('2nd Grade', '2', false)
    .addItem('3rd Grade', '3', false)
    .addItem('4th Grade', '4', false)
    .addItem('5th Grade', '5', false)
    .addItem('6th Grade', '6', false)
    .addItem('7th Grade', '7', false)
    .addItem('8th Grade', '8', false)
    .addItem('9th Grade', '9', false)
    .addItem('10th Grade', '10', false)
    .addItem('11th Grade', '11', false)
    .addItem('12th Grade', '12', false)
    .addItem('College', 'college', false)
    .addItem('Adult', 'adult', false));

  section.addWidget(CardService.newTextInput()
    .setFieldName('notes')
    .setTitle('Notes')
    .setMultiline(true)
    .setHint('Any additional information...'));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Add New Client'))
    .addSection(section)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Add Client')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('submitNewClient'))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .setSecondaryButton(CardService.newTextButton()
        .setText('Cancel')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('cancelAndGoBack'))))
    .build();

  return CardService.newNavigation().pushCard(card);
}

/**
 * Submits the new client form
 */
function submitNewClient(e) {
  const formInput = e.formInput;
  const clientName = formInput.client_name;

  if (!clientName || clientName.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client name is required'))
      .build();
  }

  try {
    // Add client to sheet
    const clientData = {
      name: clientName.trim(),
      email: formInput.client_email || '',
      parentEmail: formInput.parent_email || '',
      phone: formInput.phone || '',
      grade: formInput.grade || '',
      notes: formInput.notes || '',
      status: 'Active',
      dateAdded: new Date()
    };

    addClientToSheet(clientData); // Your existing function

    // Set as current client
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getRangeByName('CurrentClient').setValue(clientName);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Added: ${clientName}`))
      .setNavigation(CardService.newNavigation().popCard())
      .setStateChanged(true)
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// UPDATE CLIENT CARD
// ============================================================================

/**
 * Shows the update client form
 */
function showUpdateClientCard(e) {
  const clientId = e.parameters.clientId;
  const client = getClientByName(clientId); // Your existing function

  if (!client) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client not found'))
      .build();
  }

  const section = CardService.newCardSection()
    .setHeader('Update Client Information');

  section.addWidget(CardService.newTextInput()
    .setFieldName('client_name')
    .setTitle('Client Name *')
    .setValue(client.name || ''));

  section.addWidget(CardService.newTextInput()
    .setFieldName('client_email')
    .setTitle('Email')
    .setValue(client.email || ''));

  section.addWidget(CardService.newTextInput()
    .setFieldName('parent_email')
    .setTitle('Parent Email')
    .setValue(client.parentEmail || ''));

  section.addWidget(CardService.newTextInput()
    .setFieldName('phone')
    .setTitle('Phone')
    .setValue(client.phone || ''));

  section.addWidget(CardService.newTextInput()
    .setFieldName('notes')
    .setTitle('Notes')
    .setMultiline(true)
    .setValue(client.notes || ''));

  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('status')
    .setTitle('Status')
    .addItem('Active', 'Active', (client.status === 'Active' || !client.status))
    .addItem('Inactive', 'Inactive', client.status === 'Inactive'));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Update Client')
      .setSubtitle(client.name))
    .addSection(section)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Save Changes')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('submitUpdateClient')
          .setParameters({originalName: client.name}))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .setSecondaryButton(CardService.newTextButton()
        .setText('Cancel')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('cancelAndGoBack'))))
    .build();

  return CardService.newNavigation().pushCard(card);
}

/**
 * Submits the update client form
 */
function submitUpdateClient(e) {
  const formInput = e.formInput;
  const originalName = e.parameters.originalName;
  const newName = formInput.client_name;

  if (!newName || newName.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client name is required'))
      .build();
  }

  try {
    const clientData = {
      name: newName.trim(),
      email: formInput.client_email || '',
      parentEmail: formInput.parent_email || '',
      phone: formInput.phone || '',
      notes: formInput.notes || '',
      status: formInput.status || 'Active'
    };

    updateClientInSheet(originalName, clientData); // Your existing function

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Updated: ${newName}`))
      .setNavigation(CardService.newNavigation().popCard())
      .setStateChanged(true)
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// SESSION MANAGEMENT SECTION
// ============================================================================

/**
 * Builds the session management section
 */
function buildSessionManagementSection(clientInfo) {
  const section = CardService.newCardSection()
    .setHeader('📝 Session Management');

  // Quick Notes button (opens HTML dialog - hybrid approach)
  const quickNotesDisabled = !clientInfo || !clientInfo.isClient;

  section.addWidget(CardService.newTextButton()
    .setText('📋 Quick Notes')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('openQuickNotesDialog'))
    .setDisabled(quickNotesDisabled)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED));

  // Email Recap button
  section.addWidget(CardService.newTextButton()
    .setText('📧 Send Email Recap')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showEmailRecapCard'))
    .setDisabled(quickNotesDisabled));

  // View Recap History
  section.addWidget(CardService.newTextButton()
    .setText('📜 View Recap History')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showRecapHistoryCard')));

  if (quickNotesDisabled) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<i>Select a client to access session features</i>'));
  }

  return section;
}

/**
 * Opens Quick Notes in HTML dialog (hybrid approach)
 */
function openQuickNotesDialog() {
  try {
    const html = HtmlService.createTemplateFromFile('frontend/components/quick-notes')
      .evaluate()
      .setWidth(700)
      .setHeight(800);

    SpreadsheetApp.getUi().showModalDialog(html, 'Quick Notes - Session Tracker');

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Quick Notes opened'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error opening Quick Notes: ' + error.message))
      .build();
  }
}

// ============================================================================
// EMAIL RECAP CARD
// ============================================================================

/**
 * Shows the email recap card
 */
function showEmailRecapCard() {
  const clientInfo = getCurrentClientInfo();

  if (!clientInfo || !clientInfo.isClient) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Please select a client first'))
      .build();
  }

  const client = getClientByName(clientInfo.name);
  const notes = getLatestNotes(clientInfo.name);

  if (!notes || !notes.homework) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('No session notes found. Please fill out Quick Notes first.'))
      .build();
  }

  const emailPreview = generateEmailBody(notes);
  const previewText = emailPreview.substring(0, 300) + '...';

  const detailsSection = CardService.newCardSection()
    .setHeader('Email Details');

  detailsSection.addWidget(CardService.newTextInput()
    .setFieldName('recipient_email')
    .setTitle('To')
    .setValue(client.parentEmail || client.email || ''));

  detailsSection.addWidget(CardService.newTextInput()
    .setFieldName('subject')
    .setTitle('Subject')
    .setValue(`Session Recap: ${clientInfo.name} - ${new Date().toLocaleDateString()}`));

  const previewSection = CardService.newCardSection()
    .setHeader('Preview')
    .addWidget(CardService.newTextParagraph()
      .setText(previewText));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Send Email Recap')
      .setSubtitle(clientInfo.name))
    .addSection(detailsSection)
    .addSection(previewSection)
    .setFixedFooter(CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
        .setText('Send Email')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('sendEmailRecapNow')
          .setParameters({clientName: clientInfo.name}))
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED))
      .setSecondaryButton(CardService.newTextButton()
        .setText('Cancel')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('cancelAndGoBack'))))
    .build();

  return CardService.newNavigation().pushCard(card);
}

/**
 * Sends the email recap
 */
function sendEmailRecapNow(e) {
  const formInput = e.formInput;
  const clientName = e.parameters.clientName;
  const recipient = formInput.recipient_email;
  const subject = formInput.subject;

  if (!recipient || recipient.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Email address is required'))
      .build();
  }

  try {
    const notes = getLatestNotes(clientName);
    const emailBody = generateEmailBody(notes);

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: emailBody
    });

    // Log to recap history
    logRecapToHistory(clientName, recipient, subject);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Email sent successfully!'))
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error sending email: ' + error.message))
      .build();
  }
}

/**
 * Shows recap history
 */
function showRecapHistoryCard() {
  const history = getRecapHistory(); // Your existing function

  const section = CardService.newCardSection()
    .setHeader(`Recent Recaps (${history.length})`);

  if (history.length === 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<i>No recap history found</i>'));
  } else {
    history.slice(0, 20).forEach(recap => {
      section.addWidget(CardService.newDecoratedText()
        .setText(recap.clientName)
        .setBottomLabel(`${recap.date} - Sent to: ${recap.recipient}`)
        .setIcon(CardService.Icon.EMAIL));
    });

    if (history.length > 20) {
      section.addWidget(CardService.newTextParagraph()
        .setText(`<i>Showing most recent 20 of ${history.length} recaps</i>`));
    }
  }

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Recap History'))
    .addSection(section)
    .build();

  return CardService.newNavigation().pushCard(card);
}

// ============================================================================
// ADMIN SECTION
// ============================================================================

/**
 * Builds the admin section
 */
function buildAdminSection() {
  const section = CardService.newCardSection()
    .setHeader('🔧 Admin Settings');

  section.addWidget(CardService.newTextButton()
    .setText('Refresh Cache')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminRefreshCache')));

  section.addWidget(CardService.newTextButton()
    .setText('Export Client List')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminExportClients')));

  section.addWidget(CardService.newTextButton()
    .setText('View All Settings')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('showAdminSettingsCard')));

  return section;
}

/**
 * Shows full admin settings card
 */
function showAdminSettingsCard() {
  const section = CardService.newCardSection()
    .setHeader('System Management');

  section.addWidget(CardService.newTextButton()
    .setText('Refresh Cache')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminRefreshCache')));

  section.addWidget(CardService.newTextButton()
    .setText('Debug Sheet Structure')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminDebugSheet')));

  section.addWidget(CardService.newTextButton()
    .setText('Migrate Data Store')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminMigrateData')));

  section.addWidget(CardService.newTextButton()
    .setText('Export Client List (CSV)')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminExportClients')));

  section.addWidget(CardService.newTextButton()
    .setText('Validate Dashboard Links')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('adminValidateLinks')));

  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Admin Settings')
      .setSubtitle('System Management'))
    .addSection(section)
    .build();

  return CardService.newNavigation().pushCard(card);
}

// Admin action handlers
function adminRefreshCache() {
  try {
    refreshClientCache(); // Your existing function
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Cache refreshed successfully'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

function adminExportClients() {
  try {
    exportClientListAsCSV(); // Your existing function
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Client list exported'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

function adminDebugSheet() {
  try {
    debugSheetStructure(); // Your existing function
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Debug info logged to console'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

function adminMigrateData() {
  try {
    migrateToUnifiedDataStore(); // Your existing function
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Data migration complete'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

function adminValidateLinks() {
  try {
    validateDashboardLinks(); // Your existing function
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Links validated'))
      .build();
  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Error: ' + error.message))
      .build();
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Cancels and goes back to previous card
 */
function cancelAndGoBack() {
  return CardService.newNavigation().popCard();
}

/**
 * Refreshes the main card
 */
function refreshMainCard() {
  return CardService.newActionResponseBuilder()
    .setStateChanged(true)
    .setNotification(CardService.newNotification()
      .setText('Refreshed'))
    .build();
}

/**
 * Gets current client info
 */
function getCurrentClientInfo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentClient = ss.getRangeByName('CurrentClient').getValue();

    if (currentClient && currentClient.trim() !== '') {
      const client = getClientByName(currentClient);
      return {
        isClient: true,
        name: currentClient,
        email: client ? client.email : '',
        id: currentClient
      };
    }

    return {isClient: false};
  } catch (error) {
    console.error('Error getting current client:', error);
    return {isClient: false};
  }
}

/**
 * Gets user role (for enterprise features)
 */
function getUserRole() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const owner = ss.getOwner().getEmail();

    if (userEmail === owner) {
      return 'owner';
    }

    // Check if admin in user config
    const userConfig = getUserConfig();
    if (userConfig && userConfig.role) {
      return userConfig.role;
    }

    return 'tutor';
  } catch (error) {
    return 'tutor';
  }
}

/**
 * Gets user info
 */
function getUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const userConfig = getUserConfig();

    return {
      email: userEmail,
      displayName: userConfig.tutorName || userEmail.split('@')[0]
    };
  } catch (error) {
    return null;
  }
}
