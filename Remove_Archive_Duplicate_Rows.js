/**
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
 * The orginal version of this was much much more complicated, I have learned to read the docs since then
 * Takes the Archive sheet and checks to see if there is duplicates
 * I have this set to run nightly (but with its speed it could be run on the hour/minute frequency)
 */
function removeDuplicateArchiveRows() {  
  //make sure that the spreadhseet exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = ss.getSheetByName('Archive');
  if (!archiveSheet) return;

  //get the last row, if the last row is the first dont do anything
  const lastRow = archiveSheet.getLastRow();
  if (lastRow <= 1) return;

  // Define the range to check (Columns A through V)
  // 1 = Col A, 22 = Col V
  const columnsToCheck = [];
  for (let i = 1; i <= 22; i++) {
    columnsToCheck.push(i);
  }

  // Use the built-in removeDuplicates method on the entire data range
  const range = archiveSheet.getDataRange();
  range.removeDuplicates(columnsToCheck);

  //leave a logger message
  Logger.log('Duplicates removed using native method.');
}
