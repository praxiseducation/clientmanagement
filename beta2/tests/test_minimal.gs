function testMinimalSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
          .menu-button { 
            display: block; 
            width: 100%; 
            padding: 12px; 
            margin: 8px 0; 
            border: 1px solid #ddd; 
            background: white; 
            cursor: pointer; 
          }
        </style>
      </head>
      <body>
        <h1>Minimal Test</h1>
        
        <button class="menu-button" onclick="runFunction('addNewClient')">
          Add New Client
        </button>
        
        <button class="menu-button" onclick="runFunction('findClient')">
          Client List
        </button>
        
        <script>
          console.log('Minimal test loaded');
          
          let currentView = 'control';
          let clientInfo = {isClient: false, name: null, sheetName: null};
          
          function runFunction(functionName) {
            console.log('Running function:', functionName);
            google.script.run
              .withSuccessHandler(function(result) {
                console.log('Success:', result);
              })
              .withFailureHandler(function(error) {
                console.error('Error:', error);
              })[functionName]();
          }
          
          function switchToQuickNotes() {
            console.log('Switching to quick notes');
          }
          
          function openUpdateClientDialog() {
            console.log('Opening update dialog');
          }
          
          console.log('All functions defined');
        </script>
      </body>
    </html>
  `).setTitle('Minimal Test').setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}