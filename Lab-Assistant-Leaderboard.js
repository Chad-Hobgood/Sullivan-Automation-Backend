/** 
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
 * This function is the one that makes the records of the lab assistants 
 * This could be done with a query, but that and the script had errors 
 * Most of this is done to be a bit silly
*/
function updateAssistantMetrics() {
  //get the activation scriot
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Get Assistant Names from Automation_Tools
  // this makes it so that as we change the people that are Lab assistants it dynamically updates
  const toolsSheet = ss.getSheetByName("Automation_Tools");
  const assistantNames = toolsSheet.getRange("A5:A" + toolsSheet.getLastRow())
                                  .getValues()
                                  .flat()
                                  .filter(name => name !== "");

  // 2. Get Data from Archive
  const archiveSheet = ss.getSheetByName("Archive");
  const archiveData = archiveSheet.getRange("A2:P" + archiveSheet.getLastRow()).getValues();
  const totalArchiveRows = archiveData.length;

  // 3. Prepare the Dashboard_Data_Link sheet
  // this makes sure we have the sheet if it ever gets deleted
  let dashSheet = ss.getSheetByName("Dashboard_Data_Link");
  if (!dashSheet) {
    dashSheet = ss.insertSheet("Dashboard_Data_Link");
  }
  dashSheet.clear();

  // 4. Initialize results with headers
  let results = [["Assistant Name", "Completed", "Flagged", "Total Actions", "% of Total Volume"]];

  // 5. Calculate metrics for each assistant
  assistantNames.forEach(assistant => {
    let completedCount = 0;
    let flaggedCount = 0;

    archiveData.forEach(row => {
      const rowAssistant = row[14]; // Column O (15th column, index 14)
      const rowStatus = row[10];    // Column K (11th column, index 10)

      // Check if the name matches (case-insensitive)
      if (rowAssistant && rowAssistant.toString().toLowerCase() === assistant.toString().toLowerCase()) {
        if (rowStatus === "Completed") completedCount++;
        if (rowStatus === "Flagged") flaggedCount++;
      }
    });

    let totalActions = completedCount + flaggedCount;
    let percentOfTotal = totalArchiveRows > 0 ? (totalActions / totalArchiveRows) : 0;

    results.push([assistant, completedCount, flaggedCount, totalActions, percentOfTotal]);
  });

  // 6. Write results to the sheet
  dashSheet.getRange(1, 1, results.length, 5).setValues(results);
  
  // Format the percentage column
  dashSheet.getRange(2, 5, results.length - 1, 1).setNumberFormat("0.00%");
  
  // Auto-resize columns for readability
  dashSheet.autoResizeColumns(1, 5);
}
