/**Made by Chadwick Hobgood, 1/26/2025
 * 
 * Counts occurrences of printer names from "Current Printer Information"
 * and outputs the results to "Dashboard_Data_Link" Columns G and H.
 */
function updateDashboardCounts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define Sheets
  const archiveSheet = ss.getSheetByName("Archive");
  const printerListSheet = ss.getSheetByName("Current Printer Information");
  const dashboardSheet = ss.getSheetByName("Dashboard_Data_Link");
  
  // 1. Get Archive data (Column M)
  // We use getFilter() logic or flat array for Column M
  const archiveData = archiveSheet.getRange("M1:M" + archiveSheet.getLastRow()).getValues().flat();
  
  // 2. Get Printer names from A8 downwards
  const lastPrinterRow = printerListSheet.getLastRow();
  if (lastPrinterRow < 8) return; 
  const printerNames = printerListSheet.getRange("A8:A" + lastPrinterRow).getValues().flat();
  
  // 3. Process data into a 2D array for Columns G & H
  const outputData = printerNames.map(printer => {
    if (printer === "" || printer === null) {
      return ["", ""]; // Keep rows clean if printer name is empty
    }
    
    // Count matches in Archive
    const matchCount = archiveData.reduce((acc, val) => (val === printer ? acc + 1 : acc), 0);
    
    return [printer, matchCount]; // [Col G, Col H]
  });
  
  // 4. Clear previous data in Dashboard G:H to avoid ghosting old data
  const dashboardLastRow = dashboardSheet.getLastRow();
  if (dashboardLastRow >= 1) {
    dashboardSheet.getRange("G1:H" + dashboardLastRow).clearContent();
  }
  
  // 5. Write the new data starting at G1
  dashboardSheet.getRange(1, 7, outputData.length, 2).setValues(outputData);
  
  console.log("Dashboard updated: " + outputData.length + " rows processed.");
}
