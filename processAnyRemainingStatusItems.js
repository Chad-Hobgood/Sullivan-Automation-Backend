/**
 * Sweeper function with safety buffer.
 * This does not have the param since its time driven 
 * This is similar to the original resilience update, but is meant to be the backup function
 */
function processAnyRemainingStatusItems() {
  //get the active spreadsheet, if it is deletec return
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  if (!sheet) return;

  //ensure that we are getting all of the resources and not overriding anything
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  //try so we only have it attempt to run oncee
  try {
    const lastRow = sheet.getLastRow(); //get the last row
    if (lastRow <= 1) return; //if last row is 0 (sheet empty) return 

    //get the data size
    //get the current date/time for moving data 
    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const now = new Date().getTime();
    const safetyBufferMs = 60000; // 1 minute safety buffer

    //giant for loop
    //this makes it so that we go from back to front checking for Flagged or Completed
    //if ppl mess up the dropdown we also make sure to fix that
    for (let i = data.length - 1; i >= 1; i--) {
      const rowData = data[i];
      const status = String(rowData[10]).toLowerCase().trim();
      const submissionTime = new Date(rowData[0]).getTime(); // Column A
      
      // Safety Check: Only process if status is ready AND it's not "Processing..." 
      // AND it's been at least 1 minute since submission (or since the edit was likely made)
      if ((status === 'completed' || status === 'flagged') && (now - submissionTime > safetyBufferMs)) {
        const physicalRow = i + 1; 
        const recipient = rowData[2]; // Column C
        
        // Final check: Is it still Completed/Flagged?
        const currentStatus = sheet.getRange(physicalRow, 11).getValue().toString().toLowerCase().trim();
        if (currentStatus === 'completed' || currentStatus === 'flagged') {
           sheet.getRange(physicalRow, 11).setValue('Processing...'); 
           SpreadsheetApp.flush(); //forces all changes to be done now

           //this makes sure that it moves and send the email
           archiveRowAndSendEmail(sheet, physicalRow, status, recipient, rowData);
           SpreadsheetApp.flush(); //we flush it asgain as it decriments the row number
        }
      }
    }
  } catch (e) {
    //error codes
    Logger.log("Error in Sweeper: " + e.toString());
  } finally { //finally keyword from JS docs  
    lock.releaseLock(); //allow other things to run
  }
}
