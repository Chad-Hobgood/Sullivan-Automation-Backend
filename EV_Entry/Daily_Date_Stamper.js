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
