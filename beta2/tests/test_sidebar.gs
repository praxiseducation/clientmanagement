function testSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; }
        </style>
      </head>
      <body>
        <h1>Test Sidebar</h1>
        <p>If you can see this, the HTML loads correctly.</p>
        
        <script>
          console.log('Test script loaded successfully');
          
          function testFunction() {
            console.log('Test function called');
            alert('Test function works!');
          }
          
          // Test if functions are defined
          console.log('testFunction is defined:', typeof testFunction);
        </script>
        
        <button onclick="testFunction()">Test Button</button>
      </body>
    </html>
  `).setTitle('Test').setWidth(300);
  
  SpreadsheetApp.getUi().showSidebar(html);
}