/**
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
 * Modified archive function with full email text and optimized formula setting.
 * @param Sheet, this is the active sheet that we are working on
 * @param edited row, the row that has been changed
 * @param status, is the row been flagged or completed
 * @param recipent, the person the email needs to be sent to
 * @param rowData, relevant row data, mainly for flagged rows
 */
function archiveRowAndSendEmail(sheet, editedRow, status, recipient, rowData) {
  //get the data for the spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('Archive');
  const toolsSheet = ss.getSheetByName('Automation_Tools'); // Updated capitalization
  //these are the columns that we will be looking for in specific 
  const timestampColumn_W = 23; 
  const flagReasonColumn_L = 12; 

  try {
    // 1. Fetch Dynamic Flag Reasons from Automation_Tools D5:D
    const flaggedReasons = toolsSheet.getRange("D5:D")
                                     .getValues()
                                     .flat()
                                     .filter(String);

    // 2. Append the data we already have in memory
    archiveSheet.appendRow(rowData);
    const newArchiveRow = archiveSheet.getLastRow();
    
    // 3. Set completion timestamp
    archiveSheet.getRange(newArchiveRow, timestampColumn_W).setValue(new Date());
    
    // 4. Set formulas in Archive (Z, AA, AB, AC, AD)
    const formulas = [
      `=IF(K${newArchiveRow}="Completed", V${newArchiveRow}-A${newArchiveRow}, "Not Completed")`,
      `=W${newArchiveRow}-A${newArchiveRow}`,
      `=MAX(Z${newArchiveRow},AA${newArchiveRow})`,
      `=IF(W${newArchiveRow} <>"" , IF(K${newArchiveRow}="Completed", W${newArchiveRow}-V${newArchiveRow}, "Not Completed") , "Not Completed" )`,
      `=IF(W${newArchiveRow} <>"" , IF(K${newArchiveRow}="Completed", W${newArchiveRow}-A${newArchiveRow}, "Not Completed") , "Not Completed" )`
    ];
    archiveSheet.getRange(newArchiveRow, 26, 1, 5).setFormulas([formulas]);

    // 5. Prepare and Send Email
    // most of the email is kind of vibes based for the formatting
    
    let subject, body;
    if (status === 'completed') { //completed print case
      subject = 'Your 3D Print Is Ready for Pickup, Please Collect Within 72 Hours';
      body = `Hello Sullivan Student, \n\n` + 
             `Your 3D print is now complete and ready for pickup! \n\n` + 
             `You can collect it from the EV Studio/Lounge. \n\n` + 
             `Important: \n\n` + 
             `Please note that completed prints must be picked up within 72 hours of this notification.\n` +
             `If this window includes a weekend, you have 96 hours instead. \n` +
             `After that time, unclaimed prints will be discarded immediately to make space for new projects. \n\n\n` + 
             `Best wishes, \n` + 
             `Your Lab Assistants :)`;
    } else { //flagged reason case
      const flagReason = rowData[flagReasonColumn_L - 1]; 
      subject = '3D Print Job Flagged, Action Required';
      
      // Logic check: Is the reason in our list?
      let reasonText;
      if (flaggedReasons.includes(flagReason)) {
        reasonText = `Reason: ${flagReason}`;
      } else {
        reasonText = `and requires your review`;
        // Log the discrepancy for the Engineering Lead to review later
        console.warn(`Unlisted Flag Reason detected: "${flagReason}". This was not found in Automation_Tools D5:D.`);
      }
      
      body = `Hello Sullivan Student, \n\n` +
             `A task you submitted has been marked as Flagged ${reasonText}. \n\n` + 
             `Please contact a lab assistant for more information regarding the issue. \n` +
             `You are welcome to visit the lab during our operational hours to discuss it in person as well.\n\n` + 
             `Thank you for your understanding, and we look forward to resolving this with you soon. \n\n` + 
             `Best regards,\n Your Lab Assistants :)`;
    }

    //calls the mail app to send the email 
    MailApp.sendEmail(recipient, subject, body);

    // 6. Delete the row from the source sheet
    sheet.deleteRow(editedRow);
    
  } catch (error) {
    Logger.log('Critical Error in Archive/Delete: ' + error.toString());
  }
}
