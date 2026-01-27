/** Made by Chadwick Hobgood, 1/26/2026
 * * Counts occurrences of printer names from "Current Printer Information"
 * and outputs the results to "Dashboard_Data_Link" Columns G and H with a header.
 */
function Printer_wear_leveling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define Sheets
  const archiveSheet = ss.getSheetByName("Archive");
  const printerListSheet = ss.getSheetByName("Current Printer Information");
  const dashboardSheet = ss.getSheetByName("Dashboard_Data_Link");
  
  // 1. Get Archive data (Column M)
  const archiveData = archiveSheet.getRange("M1:M" + archiveSheet.getLastRow()).getValues().flat();
  
  // 2. Get Printer names from A8 downwards
  const lastPrinterRow = printerListSheet.getLastRow();
  if (lastPrinterRow < 8) return; 
  const printerNames = printerListSheet.getRange("A8:A" + lastPrinterRow).getValues().flat();
  
  // 3. Process data into a 2D array for Columns G & H
  const bodyData = printerNames.map(printer => {
    if (printer === "" || printer === null) {
      return ["", ""]; 
    }
    
    // Count matches in Archive
    const matchCount = archiveData.reduce((acc, val) => (val === printer ? acc + 1 : acc), 0);
    
    return [printer, matchCount]; // [Col G, Col H]
  });

  // 4. Create the final output including the header row
  const finalOutput = [
    ["Printer Name", "Total Request Count"], // Header Row
    ...bodyData
  ];
  
  // 5. Clear previous data in Dashboard G:H to avoid ghosting old data
  const dashboardLastRow = dashboardSheet.getLastRow();
  if (dashboardLastRow >= 1) {
    dashboardSheet.getRange("G:H").clearContent();
  }
  
  // 6. Write the new data starting at G1
  dashboardSheet.getRange(1, 7, finalOutput.length, 2).setValues(finalOutput);

  // 7. Format header for better visibility
  dashboardSheet.getRange("G1:H1").setFontWeight("bold");
  
  console.log("Dashboard updated: " + bodyData.length + " printers processed.");
}
