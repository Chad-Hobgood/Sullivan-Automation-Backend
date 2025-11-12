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
 * The core change is the use of the Lock Service to prevent concurrent editing issues
 * and the immediate deletion of the source row after archiving, which eliminates the erase on entry error
 * 
 * For Future Development:
 * 1. This was made over the course of a semester, with some debugging help using Gemini AI
 * 2. The Cleanup function has been the most problematic thus far, just be careful to handle it appropriately
 * 
 * * The basic idea is this,
 * 1. User submits a form, this populates the things in a row that they entered.
 * 2. onFormSubmit sets the status to "In Queue" and adds VLOOKUP formulas.
 * 3. When the status changes to "In Progress" or "Flagged/Completed", a timestamp is added.
 * 4. When marked as "Completed" or "Flagged," the script acquires a lock,
 *    sends the email, archives the data, and immediately deletes the original row.
 * 
 * * * Known limitations:
 * 1. Google scripts has a limit of 100 emails sent in a day.
 * 2. Properties read/write is 50,0000/day.
 * 3. Triggers total runtime: 90min/day. - We have yet 
 */


/**
 * This is a dummy function to trigger the authorization flow.
 * Run this function once from the Apps Script editor to grant permissions.
 */
function authorizeScript() {
  // This line simply calls a MailApp function. 
  MailApp.getRemainingDailyQuota();
  Logger.log('Authorization function executed. Please check the permissions pop-up.');
}


/**
 * This function handles all edits to the spreadsheet.
 * It uses the Lock Service to ensure only one heavy operation runs at a time,
 * preventing race conditions and improving data integrity during archiving.
 *
 * @param {Object} e The event object containing information about the edit.
 */
function onEdit(e) {
  // Acquire a lock to prevent concurrent script execution on critical sections.
  const lock = LockService.getScriptLock(); //https://www.reddit.com/r/learnjavascript/comments/mxs1tf/how_to_use_a_lock_object_in_javascript/  this is the thread describes what is going on here
  // Try to acquire the lock, waiting up to 15 seconds (15000 milliseconds).
  if (!lock.tryLock(15000)) {
    Logger.log("Could not acquire lock. Script is already running, skipping edit.");
    return;
  }

  try {
    // Define the sheet and columns to watch and update.
    const sheet = e.range.getSheet();
    const editedColumn = 11; // Column K (Status)
    const emailColumn = 3; // Column C (Recipient Email)
    const timestampColumn_V = 22; // Column V (In Progress Timestamp)
    
    // Get the row number immediately.
    const editedRow = e.range.getRow();

    // Check if the edited cell is in the correct column (K) and not in the header row.
    if (e.range.getColumn() === editedColumn && editedRow > 1) {
      // Get the status value and convert it to lowercase for robust comparison.
      const status = String(e.value).toLowerCase().trim();
      
      // --- Functionality for "in progress" status ---
      if (status === 'in progress') {
        e.range.setValue('In Progress');
        
        // Set the timestamp in Column V.
        sheet.getRange(editedRow, timestampColumn_V).setValue(new Date());
      }

      // --- New functionality for "Completed" or "Flagged" status ---
      if (status === 'completed' || status === 'flagged') {
        // Get the user's email from the specified column.
        const recipient = sheet.getRange(editedRow, emailColumn).getValue(); //gets the 
        const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Archive');

        // Check if the recipient email is valid and the archive sheet exists before proceeding.
        if (recipient && archiveSheet) {
          // Send email and archive/delete the row.
          archiveRowAndSendEmail(sheet, editedRow, status, recipient);
        } else {
          // Log an error if the email is not found or the archive sheet is missing.
          Logger.log(`Archiving failed for row ${editedRow}. Email: ${recipient ? 'Found' : 'Missing'}. Archive Sheet: ${archiveSheet ? 'Found' : 'Missing'}.`);
        }
      }
    }
  } catch (error) {
    // Log any errors that occur during the execution.
    Logger.log('Error in onEdit execution: ' + error.toString());
  } finally {
    // IMPORTANT: Release the lock in the finally block to ensure it is always released.
    lock.releaseLock();
  }
}

/**
 * This function handles archiving the row and sending an email.
 * After successful archiving, the original row is deleted immediately, 
 * eliminating the race condition with a scheduled cleanup job.
 *
 * @param {Object} sheet The sheet object where the edit occurred.
 * @param {number} editedRow The row number that was edited.
 * @param {string} status The status of the edited row (e.g., 'completed' or 'flagged').
 * @param {string} recipient The email address of the recipient.
 */
function archiveRowAndSendEmail(sheet, editedRow, status, recipient) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheetName = 'Archive';
  const timestampColumn_W = 23; // Column W (Completion Timestamp)
  const flagReasonColumn_L = 12; // Column L (Flag Reason)
  
  try {
    // Get the 'Archive' sheet.
    const archiveSheet = ss.getSheetByName(archiveSheetName);
    
    // Get the data from the entire edited row.
    const sourceRowRange = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn());
    const rowData = sourceRowRange.getValues()[0];
    
    // Append the row data to the bottom of the "Archive" sheet.
    archiveSheet.appendRow(rowData);
    
    // Get the row number of the newly added row in the archive sheet.
    const newlyAddedRowNumber = archiveSheet.getLastRow();
    
    // Add the completion/flag timestamp to column W of the newly added row.
    archiveSheet.getRange(newlyAddedRowNumber, timestampColumn_W).setValue(new Date());

  //this was made mainly using Gemini to put in the columns that I wanted from the Archive sheet
    // ----------------------------------------------------------------------
    // --- START: Implementing Functions for Columns Z, AA, AB, AC, AD in ARCHIVE
    // ----------------------------------------------------------------------
    
    const rowNum = newlyAddedRowNumber;
    
    // Set cell values as formulas using A1 notation for the archive sheet
    // Z: Time to In Progress (V - A)
    archiveSheet.getRange(`Z${rowNum}`).setFormula(`=IF(K${rowNum}="Completed", V${rowNum}-A${rowNum}, "Not Completed")`); 
    
    // AA: Time to Completion/Flagged (W - A) - Re-purposing the old Flagged column for total time
    archiveSheet.getRange(`AA${rowNum}`).setFormula(`=W${rowNum}-A${rowNum}`); 
    
    // AB: Max time between Z and AA (This max formula seems unusual, but retaining original logic)
    archiveSheet.getRange(`AB${rowNum}`).setFormula(`=MAX(Z${rowNum},AA${rowNum})`); 

    // AC: Time spent IN PROGRESS (W - V)
    archiveSheet.getRange(`AC${rowNum}`).setFormula(`=IF(W${rowNum} <>"" , IF(K${rowNum}="Completed", W${rowNum}-V${rowNum}, "Not Completed") , "Not Completed" )`); 

    // AD: Total Time from Submission to Completion (W - A) - Redundant with AA but retained for column layout
    archiveSheet.getRange(`AD${rowNum}`).setFormula(`=IF(W${rowNum} <>"" , IF(K${rowNum}="Completed", W${rowNum}-A${rowNum}, "Not Completed") , "Not Completed" )`); 

    // ----------------------------------------------------------------------
    // --- END: Implementing Functions for Columns Z, AA, AB, AC, AD in ARCHIVE
    // ----------------------------------------------------------------------

    // Define email content based on status.
    let subject, body;
    if (status === 'completed') {
      subject = 'Your 3D Print Is Ready for Pickup, Please Collect Within 72 Hours';
      body = `Hello Sullivan Student, \n\n` + 'Your 3D print is now complete and ready for pickup! \n\n' + 'You can collect it from the EV Studio/Lounge. \n\n' + 'Important: \n\n' + 'Please note that completed prints must be picked up within 72 hours of this notification.\n If this window includes a weekend, you have 96 hours instead. \n After that time, unclaimed prints will be discarded immediately to make space for new projects. \n\n\n' + 'Best wishes, \n' + 'Your Lab Assistants :)';
    } else if (status === 'flagged') {
      const flaggedReasons = ["Too big", "Couldn't view part", "Too Small", "Wrong File type", "Failed after 3+ print attempts", "24h+ print time", "Submitted A Folder Instead of a Part"];

      const flagReason = sheet.getRange(editedRow, flagReasonColumn_L).getValue();
      
      subject = '3D Print Job Flagged, Action Required';
      
      if (flaggedReasons.includes(flagReason)) {
        body = 'Hello Sullivan Student, \n' + `A task you submitted has been marked as Flagged. Reason: ${flagReason}. \n\n` + 'Please contact a lab assistant for more information regarding the issue. \n You are welcome to visit the lab during our operational hours to discuss it in person as well.\n\n' + 'Thank you for your understanding, and we look forward to resolving this with you soon. \n\n' + 'Best regards,\n Your Lab Assistants :)';
      } else {
        body = 'Hello Sullivan Student, \n' +`A task you submitted has been marked as Flagged and requires your review. \n \n` + 'Please contact a lab assistant for more information regarding the issue. \n You are welcome to visit the lab during our operational hours to discuss it in person as well.\n\n' + 'Thank you for your understanding, and we look forward to resolving this with you soon. \n\n' + 'Best regards,\n Your Lab Assistants :)';
      }
    }
    
    // Send the email.
    MailApp.sendEmail(recipient, subject, body);

    // --- CRITICAL RESILIENCE STEP: DELETE THE ROW IMMEDIATELY ---
    // This is the change that prevents the race condition and immediate-clear issue.
    // The row is now safely moved to the archive, and then instantly removed from the queue.
    //DO NOT MESS WITH THIS PLEASE, IF THIS STARTS MESSING UP EVALUATE THE REST OF THE CODE TOO
    sheet.deleteRow(editedRow);
    
  } catch (error) {
    // Log the error but DO NOT delete the row if the email or archiving failed.
    Logger.log('Error in archiveRowAndSendEmail. Row was NOT deleted: ' + error.toString());
  }
}

/**
 * This function is triggered by a form submission. It adds the
 * VLOOKUP formulas to columns T and U of the new row and sets the status in column K.
 *
 * @param {Object} e The event object containing information about the form submission.
 */
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const newRow = e.range.getRow();
  
  // Define the formulas. The row reference will be automatically updated by Google Sheets.
  const formulaT = '=IFERROR(VLOOKUP(M' + newRow + ', \'Current Printer Information\'!A:C, 3, FALSE), "In Queue or NA")';
  const formulaU = '=IFERROR(VLOOKUP(M' + newRow + ', \'Current Printer Information\'!A:C, 2, FALSE), "In Queue or NA")';
  
  // Set the formulas in the correct columns.
  sheet.getRange(newRow, 20).setFormula(formulaT); // Column T
  sheet.getRange(newRow, 21).setFormula(formulaU); // Column U

  // Set the initial status in column K.
  sheet.getRange(newRow, 11).setValue("In Queue");
}


/**
 * Removes duplicate rows from the 'Archive' sheet based on columns A through V.
 * This is still useful as a scheduled job to ensure the archive stays clean.
 */
function removeDuplicateArchiveRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('Archive');

  // Check if the 'Archive' sheet exists.
  if (!archiveSheet) {
    Logger.log('Archive sheet not found. Aborting.');
    return;
  }

  const lastRow = archiveSheet.getLastRow();
  const lastCol = archiveSheet.getLastColumn();

  // If there's only a header row or the sheet is empty, there are no duplicates to remove.
  if (lastRow <= 1) {
    Logger.log('Archive sheet is empty or only contains a header. No action needed.');
    return;
  }

  // Get all data from the sheet, including the header.
  const allRows = archiveSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = allRows[0];
  const dataRows = allRows.slice(1);

  // Use a map to store unique rows based on their columns A-V values.
  const uniqueRows = new Map();
  
  dataRows.forEach(row => {
    // Create a key from the values in columns A-V (indexes 0 to 21).
    const key = JSON.stringify(row.slice(0, 22));

    if (!uniqueRows.has(key)) {
      // If the key is new, add the row to the unique set, retaining the last/most recent entry.
      uniqueRows.set(key, row);
    }
  });

  // Rebuild the data without duplicates.
  const uniqueData = [header];
  uniqueRows.forEach(row => uniqueData.push(row));


  // Clear the entire sheet and rewrite the unique data.
  archiveSheet.getRange(1, 1, lastRow, lastCol).clearContent();
  if (uniqueData.length > 1) {
    archiveSheet.getRange(1, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
  }

  Logger.log(`Duplicate rows removed from 'Archive' sheet based on columns A-V. Initial rows: ${dataRows.length}, Unique rows: ${uniqueData.length - 1}`);
}

const QUEUE_SHEET_NAME = "Form_Responses";
/**
 * This function is intended to be run by a time-based trigger (e.g., nightly, I chose 1am-2am).
 * It finds and deletes all empty rows in the main QUEUE SHEET. Not any of the others
 * This acts as a resilience measure for any rows that were cleared but not deleted
 * due to previous errors or script interruptions.
 * * NOTE: The QUEUE_SHEET_NAME constant must be set correctly at the top of the script.
 * 
 */
function cleanupEmptyQueueRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  
  if (!sheet) {
    Logger.log(`Queue sheet named "${QUEUE_SHEET_NAME}" not found. Aborting cleanup.`);
    return;
  }
  
  const statusColumn = 11; // Column K (Status)
  const lastRow = sheet.getLastRow();
  
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
}

