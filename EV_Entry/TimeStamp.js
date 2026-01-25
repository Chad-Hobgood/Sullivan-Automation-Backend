/** * Built by Chad Hobgood, Engineering lead 2025, 
 * Last Update: 11/12/2025 (Resilience Update)
 * * FOR FUTURE USERS:
 * Welcome lead lab assistant, 
 * the most important function that you need to be aware of to keep this sheet running is authorizeScript
 * Go up to the drop down at the top and change it to that function, and click run and a pop up should show up
 * Make sure that you fill that out and it will be able to send emails as you
 * 
 * 
 * * Dev notes:
 * This version has been refactored for improved resilience against race conditions.
 * The main thing that this is deisgned to do is automate the task I found with the entry sheet that I found annoying
 * This does make the sheet grow downwards, which I dont like but I got outvoted on that design choice
 * and the immediate deletion of the source row after archiving, which eliminates the erase on entry error
 * 
 * For Future Development:
 * Have something to make it look for ppl on the bad apples list
 * 
 * * The basic idea is this,
 * 1. At Midnight, make a new date stamp
 * 2. Everytime that someone swipes in, collect a time stamp so we have that information 
 * 
 * * * Known limitations:
 * 1. Google scripts has a limit of 100 emails sent in a day.
 * 2. Properties read/write is 50,0000/day.
 * 3. Triggers total runtime: 90min/day. - We have yet 
 */



/**
 * This function is included solely to force the user to complete the necessary
 * authorization steps for the script to run all its functions (like onEdit
 * and dailyDateStamper, which require permissions to edit the spreadsheet).
 *
 * The function itself does nothing functional for the sheet beyond triggering
 * the Google Authorization flow when run manually from the script editor.
 * You will have to accept the request to run the script on the spreadsheet tab to complete the process
 */
function authorizeScript() {
  // Accessing the active spreadsheet forces the script to request the
  // 'Spreadsheet' scope during the authorization process.
  try {
    const ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
    Logger.log(`Successfully accessed spreadsheet: ${ssName}`);
    // Optional: Use a simple UI alert to confirm the function ran successfully
    // after authorization is complete, which is helpful feedback for the user.
    SpreadsheetApp.getUi().alert(
      'Authorization Check Complete', 
      'The script has successfully run the authorization check. If you saw a request for permissions, you should now be fully authorized.', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Authorization function failed, usually indicating missing permissions or no active spreadsheet context: ' + e.toString());
    // If the error is caught, the authorization prompt should have already appeared.
  }
}





/**
 * This function automatically runs when a user edits a cell in the spreadsheet.
 * It checks if the edit happened in Column B (index 2) and, if so, places
 * the current timestamp into the corresponding cell in Column F (index 6).
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object provided by Google.
 */
function onEdit(e) {
  // Always include a check for the event object and range to prevent errors
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  
  // --- NEW: Exit immediately if not the target sheet 'EV Design Studio' ---
  const targetSheetName = 'EV Design studio';
  if (sheet.getName() !== targetSheetName) {
    return;
  }
  // ------------------------------------------------------------------------

  const editedColumn = range.getColumn();
  const editedRow = range.getRow();

  // --- Logic for Timestamp (Column B input -> Column F timestamp) ---
  // 1. Check if the edited column is Column B (index 2).
  // 2. Ensure only a single cell was edited.
  // 3. Ensure the cell is not empty after the edit (e.value exists).
  if (editedColumn === 2 && range.getNumColumns() === 1 && e.value) {

    // Column F is the 6th column
    const timestampCell = sheet.getRange(editedRow, 6);

    // Set the current date and time
    timestampCell.setValue(new Date());

    // Set the display format for clarity
    timestampCell.setNumberFormat('MM/dd/yyyy HH:mm:ss');
  }
}


/**
 * This function is designed to be run once per day using a time-driven trigger.
 * It finds the last row with content in Column B of the specified sheet and places the
 * current date into Column A (index 1) of the next empty row, based on Column B's data.
 * I have added extensive logging to help debug why the function is not making edits.
 */
function dailyDateStamper() {
  Logger.log('--- dailyDateStamper STARTED ---');
  // The sheet name is confirmed to be 'EV Design studio' (lowercase 's')
  const sheetName = 'EV Design studio'; 
  // Column B is the 2nd column
  const dataColumnIndex = 2; // B

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    // This message will appear in the log if the sheet is missing or the name is wrong
    Logger.log(`Error: Sheet named "${sheetName}" not found. Please check the sheetName variable.`);
    return; // Exit if the sheet is missing
  }

  Logger.log(`Sheet '${sheetName}' found successfully.`);
  
  // --- MODIFICATION: Find the last row with data in the specified column (B) ---
  
  // Get all data in Column B
  // sheet.getMaxRows() gets the total number of rows in the sheet.
  const columnData = sheet.getRange(1, dataColumnIndex, sheet.getMaxRows()).getValues();
  
  // Find the index of the last non-empty element in the flattened array
  // .reverse() changes the array order, .findIndex() finds the first index matching the condition (not empty)
  // The result is the index from the *end* of the original array.
  let lastRowWithData = columnData.flat().reverse().findIndex(row => row != '');
  
  // Convert the index from the end to a row number (1-based index)
  if (lastRowWithData !== -1) {
    lastRowWithData = columnData.length - lastRowWithData;
  } else {
    // If no data is found, we assume the first row (row 1) is the target.
    // If the sheet truly has no data, sheet.getLastRow() would be 0, and we'd target 1.
    // If column B is empty but other columns have data, sheet.getLastRow() might be > 0.
    // For consistency with an empty column B, we'll set it to 0 so targetRow becomes 1.
    lastRowWithData = 0; 
  }

  // Log the row number found
  Logger.log(`Last row with data in Column B found: ${lastRowWithData}`);

  // The first empty row is lastRowWithData + 1
  const targetRow = lastRowWithData + 1;
  // Log the row where the date will be placed
  Logger.log(`Targeting the next empty row: ${targetRow}`);

  // Column A is the 1st column
  const targetColumn = 1; // A
  const targetCell = sheet.getRange(targetRow, targetColumn);

  // Set the current date
  targetCell.setValue(new Date());

  // Set the display format for the date
  targetCell.setNumberFormat('yyyy-MM-dd');
  
  // Confirmation log
  Logger.log(`Successfully added date to A${targetRow}.`);
  Logger.log('--- dailyDateStamper FINISHED ---');
}
