/**
 * CardService Helper Functions
 * Wrapper functions to integrate CardService with existing backend
 *
 * DEBUG_ID: CARDSERVICE_HELPERS_20250109
 */

// ============================================================================
// BACKEND INTEGRATION HELPERS
// ============================================================================

/**
 * Searches for clients in the sheet
 * @param {string} searchTerm - Search query
 * @return {Array<Object>} Array of matching clients
 */
function searchClientsInSheet(searchTerm) {
  if (!searchTerm || searchTerm.trim().length < 1) {
    return [];
  }

  const clients = getAllClients();
  const term = searchTerm.toLowerCase().trim();

  return clients.filter(client => {
    return (
      (client.name && client.name.toLowerCase().includes(term)) ||
      (client.email && client.email.toLowerCase().includes(term)) ||
      (client.parentEmail && client.parentEmail.toLowerCase().includes(term))
    );
  });
}

/**
 * Gets all clients from the sheet
 * @return {Array<Object>} Array of all clients
 */
function getAllClients() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientSheet = ss.getSheetByName('Clients') || ss.getSheetByName('Client Database');

    if (!clientSheet) {
      console.error('No Clients sheet found');
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
        rowIndex: i + 1 // Store row index for updates
      });
    }

    return clients;
  } catch (error) {
    console.error('Error getting all clients:', error);
    return [];
  }
}

/**
 * Gets a client by name
 * @param {string} clientName - Name of client
 * @return {Object|null} Client object or null
 */
function getClientByName(clientName) {
  const clients = getAllClients();
  return clients.find(c => c.name === clientName) || null;
}

/**
 * Adds a new client to the sheet
 * @param {Object} clientData - Client information
 */
function addClientToSheet(clientData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let clientSheet = ss.getSheetByName('Clients') || ss.getSheetByName('Client Database');

    if (!clientSheet) {
      // Create Clients sheet if it doesn't exist
      clientSheet = ss.insertSheet('Clients');
      clientSheet.appendRow([
        'Name',
        'Email',
        'Parent Email',
        'Phone',
        'Grade',
        'Notes',
        'Status',
        'Date Added'
      ]);
      clientSheet.getRange('A1:H1').setFontWeight('bold');
    }

    // Add client row
    clientSheet.appendRow([
      clientData.name,
      clientData.email || '',
      clientData.parentEmail || '',
      clientData.phone || '',
      clientData.grade || '',
      clientData.notes || '',
      clientData.status || 'Active',
      clientData.dateAdded || new Date()
    ]);

    // Clear cache
    refreshClientCache();

    console.log('Client added:', clientData.name);
  } catch (error) {
    console.error('Error adding client:', error);
    throw new Error('Failed to add client: ' + error.message);
  }
}

/**
 * Updates a client in the sheet
 * @param {string} originalName - Original client name
 * @param {Object} clientData - Updated client data
 */
function updateClientInSheet(originalName, clientData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientSheet = ss.getSheetByName('Clients') || ss.getSheetByName('Client Database');

    if (!clientSheet) {
      throw new Error('Clients sheet not found');
    }

    const data = clientSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indices
    const nameCol = headers.indexOf('Name');
    const emailCol = headers.indexOf('Email');
    const parentEmailCol = headers.indexOf('Parent Email');
    const phoneCol = headers.indexOf('Phone');
    const notesCol = headers.indexOf('Notes');
    const statusCol = headers.indexOf('Status');

    // Find client row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === originalName) {
        rowIndex = i + 1; // +1 for 1-based indexing
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Client not found: ' + originalName);
    }

    // Update row
    if (nameCol >= 0) clientSheet.getRange(rowIndex, nameCol + 1).setValue(clientData.name);
    if (emailCol >= 0) clientSheet.getRange(rowIndex, emailCol + 1).setValue(clientData.email || '');
    if (parentEmailCol >= 0) clientSheet.getRange(rowIndex, parentEmailCol + 1).setValue(clientData.parentEmail || '');
    if (phoneCol >= 0) clientSheet.getRange(rowIndex, phoneCol + 1).setValue(clientData.phone || '');
    if (notesCol >= 0) clientSheet.getRange(rowIndex, notesCol + 1).setValue(clientData.notes || '');
    if (statusCol >= 0) clientSheet.getRange(rowIndex, statusCol + 1).setValue(clientData.status || 'Active');

    // Update current client if it was changed
    if (originalName !== clientData.name) {
      const currentClient = ss.getRangeByName('CurrentClient');
      if (currentClient && currentClient.getValue() === originalName) {
        currentClient.setValue(clientData.name);
      }
    }

    // Clear cache
    refreshClientCache();

    console.log('Client updated:', clientData.name);
  } catch (error) {
    console.error('Error updating client:', error);
    throw new Error('Failed to update client: ' + error.message);
  }
}

/**
 * Gets the latest session notes for a client
 * @param {string} clientName - Client name
 * @return {Object|null} Notes object or null
 */
function getLatestNotes(clientName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const notesSheet = ss.getSheetByName('Session Notes') || ss.getSheetByName('Quick Notes');

    if (!notesSheet) {
      return null;
    }

    const data = notesSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return null;
    }

    const headers = data[0];
    const clientCol = headers.indexOf('Client Name') || headers.indexOf('Client');

    // Find most recent notes for this client
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][clientCol] === clientName) {
        // Map row to notes object
        return {
          clientName: clientName,
          date: data[i][headers.indexOf('Date')] || new Date(),
          homework: data[i][headers.indexOf('Homework')] || '',
          topics: data[i][headers.indexOf('Topics')] || '',
          progress: data[i][headers.indexOf('Progress')] || '',
          strengths: data[i][headers.indexOf('Strengths')] || '',
          challenges: data[i][headers.indexOf('Challenges')] || '',
          nextSession: data[i][headers.indexOf('Next Session')] || '',
          parentNotes: data[i][headers.indexOf('Parent Notes')] || ''
        };
      }
    }

    return null;
  } catch (error) {
    console.error('Error getting latest notes:', error);
    return null;
  }
}

/**
 * Saves notes to the sheet
 * @param {Object} notesData - Notes data
 */
function saveNotesToSheet(notesData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let notesSheet = ss.getSheetByName('Session Notes');

    if (!notesSheet) {
      notesSheet = ss.insertSheet('Session Notes');
      notesSheet.appendRow([
        'Date',
        'Client Name',
        'Homework',
        'Topics',
        'Progress',
        'Strengths',
        'Challenges',
        'Next Session',
        'Parent Notes'
      ]);
      notesSheet.getRange('A1:I1').setFontWeight('bold');
    }

    notesSheet.appendRow([
      notesData.date || new Date(),
      notesData.clientId || notesData.clientName,
      notesData.homework || '',
      notesData.topics || '',
      notesData.progress || '',
      notesData.strengths || '',
      notesData.challenges || '',
      notesData.nextSession || '',
      notesData.parentNotes || ''
    ]);

    console.log('Notes saved for:', notesData.clientName);
  } catch (error) {
    console.error('Error saving notes:', error);
    throw new Error('Failed to save notes: ' + error.message);
  }
}

/**
 * Generates email body from notes
 * @param {Object} notes - Session notes
 * @return {string} HTML email body
 */
function generateEmailBody(notes) {
  const config = CONFIG || {company: 'Smart College', tutorName: 'Your Tutor'};
  const tutorName = getUserConfig().tutorName || config.tutorName;

  let html = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
    .header { background: #003366; color: white; padding: 20px; text-align: center; }
    .content { padding: 20px; }
    .section { margin-bottom: 20px; }
    .section-title { font-weight: bold; color: #003366; margin-bottom: 5px; }
    .section-content { margin-left: 10px; }
    .footer { background: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #5f6368; }
  </style>
</head>
<body>
  <div class="header">
    <h2>Session Recap</h2>
    <p>${new Date().toLocaleDateString()}</p>
  </div>
  <div class="content">
    <p>Hello!</p>
    <p>Here's a summary of today's tutoring session:</p>
`;

  if (notes.homework) {
    html += `
    <div class="section">
      <div class="section-title">📚 Homework Assigned</div>
      <div class="section-content">${notes.homework}</div>
    </div>`;
  }

  if (notes.topics) {
    html += `
    <div class="section">
      <div class="section-title">📖 Topics Covered</div>
      <div class="section-content">${notes.topics}</div>
    </div>`;
  }

  if (notes.progress) {
    html += `
    <div class="section">
      <div class="section-title">📈 Progress Notes</div>
      <div class="section-content">${notes.progress}</div>
    </div>`;
  }

  if (notes.strengths) {
    html += `
    <div class="section">
      <div class="section-title">⭐ Strengths</div>
      <div class="section-content">${notes.strengths}</div>
    </div>`;
  }

  if (notes.challenges) {
    html += `
    <div class="section">
      <div class="section-title">💪 Areas for Growth</div>
      <div class="section-content">${notes.challenges}</div>
    </div>`;
  }

  if (notes.nextSession) {
    html += `
    <div class="section">
      <div class="section-title">🎯 Next Session Goals</div>
      <div class="section-content">${notes.nextSession}</div>
    </div>`;
  }

  if (notes.parentNotes) {
    html += `
    <div class="section">
      <div class="section-title">💬 Additional Notes</div>
      <div class="section-content">${notes.parentNotes}</div>
    </div>`;
  }

  html += `
    <p>Please feel free to reach out if you have any questions!</p>
    <p>Best regards,<br>${tutorName}</p>
  </div>
  <div class="footer">
    <p>${config.company}</p>
    <p>This email was generated by Client Manager</p>
  </div>
</body>
</html>`;

  return html;
}

/**
 * Logs recap to history
 * @param {string} clientName - Client name
 * @param {string} recipient - Email recipient
 * @param {string} subject - Email subject
 */
function logRecapToHistory(clientName, recipient, subject) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName('Recap History');

    if (!historySheet) {
      historySheet = ss.insertSheet('Recap History');
      historySheet.appendRow(['Date', 'Client Name', 'Recipient', 'Subject']);
      historySheet.getRange('A1:D1').setFontWeight('bold');
    }

    historySheet.appendRow([
      new Date(),
      clientName,
      recipient,
      subject
    ]);
  } catch (error) {
    console.error('Error logging to recap history:', error);
  }
}

/**
 * Gets recap history
 * @return {Array<Object>} Array of recap records
 */
function getRecapHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('Recap History');

    if (!historySheet) {
      return [];
    }

    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }

    const history = [];
    for (let i = data.length - 1; i >= 1 && history.length < 50; i--) {
      history.push({
        date: data[i][0],
        clientName: data[i][1],
        recipient: data[i][2],
        subject: data[i][3]
      });
    }

    return history;
  } catch (error) {
    console.error('Error getting recap history:', error);
    return [];
  }
}

/**
 * Gets user configuration
 * @return {Object} User config
 */
function getUserConfig() {
  try {
    const props = PropertiesService.getUserProperties();
    return {
      tutorName: props.getProperty('tutorName') || '',
      tutorEmail: props.getProperty('tutorEmail') || Session.getActiveUser().getEmail(),
      isConfigured: props.getProperty('isConfigured') === 'true',
      role: props.getProperty('role') || 'tutor'
    };
  } catch (error) {
    console.error('Error getting user config:', error);
    return {
      tutorName: '',
      tutorEmail: Session.getActiveUser().getEmail(),
      isConfigured: false,
      role: 'tutor'
    };
  }
}

/**
 * Refreshes client cache
 */
function refreshClientCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('client_list');
    console.log('Client cache cleared');
  } catch (error) {
    console.error('Error refreshing cache:', error);
  }
}
