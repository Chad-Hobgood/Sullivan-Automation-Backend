/**
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
 * Enhanced onEdit handles high-frequency changes by processing 
 * the specific range and verifying row contents under lock.
 * 
 * THIS IS WHAT CALLS THE ARCHIVE FUNCTION, it runs when there has been an edit for the whole sheet
 */
function onEdit(e) {
  //get some information from the sheet
  const range = e.range;
  const sheet = range.getSheet();
  
  // Only process if the edit is on the Queue sheet and targets the Status column (11/K)
  if (sheet.getName() !== QUEUE_SHEET_NAME || range.getColumn() !== 11 || range.getRow() <= 1) {
    return;
  }

  //get the lock since this changes the row number and can mess with other rows
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for other instances to finish.
    lock.waitLock(30000); 

    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const emailColumn = 3; 
    const timestampColumn_V = 22;

    // Force any pending spreadsheet changes to complete before reading data
    SpreadsheetApp.flush();

    // Fetch the values again INSIDE the lock to ensure we aren't looking at stale indices
    const currentValues = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

    // Iterate backwards through the edited range to maintain index integrity during deletions
    for (let i = numRows - 1; i >= 0; i--) {
      const rowData = currentValues[i];
      const statusValue = String(rowData[10]).toLowerCase().trim(); // Column K
      const physicalRow = startRow + i;

      //if the thing is in progress but has not been given a timestamp yet  
      //and the cases is wrong
      if (statusValue === 'in progress') {
        //make sure to set the value as 'In Progress'
        sheet.getRange(physicalRow, 11).setValue('In Progress');
        const existingTimestamp = rowData[timestampColumn_V - 1];
        //set the value 
        if (!existingTimestamp) {
          sheet.getRange(physicalRow, timestampColumn_V).setValue(new Date());
        }
      } 

      //if we see that the value has been set to Completed or Flagged
      else if (statusValue === 'completed' || statusValue === 'flagged') {
        const recipient = rowData[emailColumn - 1]; //find the row that that has been flagged or completed
        const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive');

        if (recipient && archiveSheet) {  //call the function to archive
          archiveRowAndSendEmail(sheet, physicalRow, statusValue, recipient, rowData);
          // Flush after every deletion to ensure the NEXT row in the loop is still at the correct index
          SpreadsheetApp.flush();
        }
      }
    }
  } catch (error) {
    Logger.log('Error in onEdit: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}
