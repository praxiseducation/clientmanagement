// TASK #9: Replace the showRecapDialog() function in clientmanager.gs

// FIND THIS (around line 8936):
// function showRecapDialog(clientData) {
//   const html = HtmlService.createHtmlOutput(`
//     <!-- 400+ lines of embedded HTML -->
//   `).setWidth(600).setHeight(800);
//   SpreadsheetApp.getUi().showModalDialog(html, 'Session Recap');
// }

// REPLACE WITH THIS:
function showRecapDialog(clientData) {
  if (!clientData) {
    clientData = getCurrentClientData();
  }
  
  const html = loadRecapDialog(clientData)
    .setWidth(600)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Session Recap');
}

// Also replace showRecapDialogLegacy() (around line 8941):
function showRecapDialogLegacy(clientData) {
  // Redirect to new template system
  showRecapDialog(clientData);
}