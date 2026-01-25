/** * Built by Chad Hobgood, Engineering lead 2025, 
 * Last Update: 1/25/2025 (Metrics Update)
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
