/**
 * Made by Chadwick Hobgood, Engineering lead, 1/25/2026
 * 
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
