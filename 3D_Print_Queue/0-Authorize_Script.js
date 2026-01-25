/** * Built by Chad Hobgood, Engineering lead 2025, 
 * Last Update: 1/25/2026 (Sweeper and Analytics)
 * * FOR FUTURE USERS:
 * Welcome lead lab assistant, 
 * the most important function that you need to be aware of to keep this sheet running is authorizeScript
 * Go up to the drop down at the top and change it to that function, and click run and a pop up should show up
 * Make sure that you fill that out and it will be able to send emails as you
 * 
 * TL:DR on the Way that this works
 * 1.  You authroize the script to run as you
 * 2.  When someone fills out the google form, On_form_submit happens (Timestamp, add in excel formulas)
 * 3A. When a lab assistant sees it and marks it as in progress, On_edit_runs (adds some timestamp info to queue sheet)
 * 3B. When a lab assistant marks it as complete Archive_Row runs (moves it to archive, makes timestamp info for archive, sends email)
 *     and the row is cleared by the Archive function
 * 4.  If there is things left over that are complete but not moved, the Sweeper runs and moves them 
 * 5.  We do some data cleaning nightly with the Lab Assistant Metricsbeing updated and removing duplicate rows from the spreadsheet 
 * 
 * * * * Dev notes:
 * 1. Each function is in its own file, I got tired of scrolling
 *    (Most of the functions don't have to call the other functions that we've made here)
 * 2. If you need to make edits, all of the functions should be within the context window of an AI
 * 3. I have tried my best to make it fairly easy to see where I was going with most of this with comments
 *    that make sense to me
 * 
 *  * * * * Known limitations:
 * 1. Google scripts has a limit of 100 emails sent in a day.
 * 2. Properties read/write is 50,0000/day.
 * 3. Triggers total runtime: 90min/day. - We have yet to get above 5mins/day
 */


const QUEUE_SHEET_NAME = "Form_Responses"; //if you change the sheet name this gets updated

/**
 * This is a dummy function to trigger the authorization flow.
 * Run this function once from the Apps Script editor to grant permissions.
 */
function authorizeScript() {
  // This line simply calls a MailApp function. 
  MailApp.getRemainingDailyQuota();
  Logger.log('Authorization function executed. Please check the permissions pop-up.');
}
