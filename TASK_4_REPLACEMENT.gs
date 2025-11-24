// TASK #4: Replace the showSidebar() function in clientmanager.gs

// FIND THIS (around line 7717):
// function showSidebar() {
//   const html = HtmlService.createHtmlOutput(`
//     <!-- 500+ lines of embedded HTML -->
//   `).setWidth(350).setHeight(600);
//   SpreadsheetApp.getUi().showSidebar(html);
// }

// REPLACE WITH THIS:
function showSidebar() {
  const html = loadSidebarTemplate('control');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Also replace showEnterpriseSidebar() (around line 7603):
function showEnterpriseSidebar() {
  const html = loadSidebarTemplate('control');
  SpreadsheetApp.getUi().showSidebar(html);
}

// And showUniversalSidebar() (around line 796):
function showUniversalSidebar(initialView = 'control') {
  const html = loadSidebarTemplate(initialView);
  SpreadsheetApp.getUi().showSidebar(html);
}