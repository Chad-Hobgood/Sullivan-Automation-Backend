/**
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
 * This function is intended to be run by a time-based trigger (e.g., nightly, I chose 1am-2am).
 * It finds and deletes all empty rows in the main QUEUE SHEET. Not any of the others
 * This acts as a resilience measure for any rows that were cleared but not deleted
 * due to previous errors or script interruptions.
 * * NOTE: The QUEUE_SHEET_NAME constant must be set correctly at the top of the script.
 * */
function cleanupEmptyQueueRows() {
  //this makes sure that the sheet exists with the right name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`Queue sheet named "${QUEUE_SHEET_NAME}" not found. Aborting cleanup.`);
    return;
  }
  //looks at the status column
  const statusColumn = 11; // Column K (Status)
  const lastRow = sheet.getLastRow();
  
  //if the last row has more than one thing in it
  if (lastRow > 1) {
    // Get values from Column K, starting from row 2 down to the last row.
    const range = sheet.getRange(2, statusColumn, lastRow - 1, 1);
    const values = range.getValues();
    
    // Loop backwards to handle row deletions correctly.
    for (let i = values.length - 1; i >= 0; i--) {
      // Check if the cell in column K is empty (empty string).
      if (String(values[i][0]).trim() === '') {
        // Delete the row if it's empty. Row index is (i + 2).
        sheet.deleteRow(i + 2);
        Logger.log(`Deleted empty row ${i + 2} from ${QUEUE_SHEET_NAME}.`);
      }
    }
  }

  Logger.log(`Completed Operation`);
}
