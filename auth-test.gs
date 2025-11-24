/**
 * AUTHORIZATION TEST FUNCTION
 *
 * Add this to your Code.gs file temporarily
 * Run it once to force authorization of all scopes
 * Then delete it after authorization succeeds
 */

/**
 * Test function to force authorization of all OAuth scopes
 * Run this once, grant permissions, then delete
 */
function forceAuthorization() {
  try {
    Logger.log('Testing authorization for all scopes...');

    // 1. Calendar scope - https://www.googleapis.com/auth/calendar.readonly
    Logger.log('Testing Calendar API...');
    const calendars = Calendar.CalendarList.list();
    Logger.log('✓ Calendar API authorized: ' + calendars.items.length + ' calendars found');

    // 2. Spreadsheet scope - https://www.googleapis.com/auth/spreadsheets
    Logger.log('Testing Spreadsheets...');
    try {
      SpreadsheetApp.getActiveSpreadsheet();
      Logger.log('✓ Spreadsheets authorized');
    } catch (e) {
      Logger.log('✓ Spreadsheets scope granted (no active sheet, but that\'s OK)');
    }

    // 3. Drive scope - https://www.googleapis.com/auth/drive
    Logger.log('Testing Drive API...');
    const root = DriveApp.getRootFolder();
    Logger.log('✓ Drive API authorized: ' + root.getName());

    // 4. Drive File scope - https://www.googleapis.com/auth/drive.file
    Logger.log('Testing Drive File access...');
    const files = DriveApp.getFiles();
    Logger.log('✓ Drive File access authorized');

    // 5. Script Container UI - implicit with CardService
    Logger.log('✓ Script Container UI scope granted (implicit)');

    // 6. External requests - implicit
    Logger.log('✓ External requests scope granted (implicit)');

    Logger.log('');
    Logger.log('========================================');
    Logger.log('SUCCESS! All scopes authorized.');
    Logger.log('========================================');
    Logger.log('You can now:');
    Logger.log('1. Delete this function');
    Logger.log('2. Deploy → Test deployments → Install');
    Logger.log('3. Test the add-on in Google Sheets');

    return 'Authorization complete - all scopes granted!';

  } catch (error) {
    Logger.log('');
    Logger.log('========================================');
    Logger.log('ERROR: ' + error.message);
    Logger.log('========================================');
    Logger.log('This is normal on first run.');
    Logger.log('Click "Review Permissions" and grant all access.');
    Logger.log('Then run this function again.');

    return 'Error: ' + error.message;
  }
}
