/**
 * This function automatically runs when a user edits a cell in the spreadsheet.
 * It checks if the edit happened in Column B (index 2) and, if so, places
 * the current timestamp into the corresponding cell in Column F (index 6).
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object provided by Google.
 */
function onEdit(e) {
  // Always include a check for the event object and range to prevent errors
  if (!e || !e.range) {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  
  // --- NEW: Exit immediately if not the target sheet 'EV Design Studio' ---
  const targetSheetName = 'EV Design studio';
  if (sheet.getName() !== targetSheetName) {
    return;
  }
  // ------------------------------------------------------------------------

  const editedColumn = range.getColumn();
  const editedRow = range.getRow();

  // --- Logic for Timestamp (Column B input -> Column F timestamp) ---
  // 1. Check if the edited column is Column B (index 2).
  // 2. Ensure only a single cell was edited.
  // 3. Ensure the cell is not empty after the edit (e.value exists).
  if (editedColumn === 2 && range.getNumColumns() === 1 && e.value) {

    // Column F is the 6th column
    const timestampCell = sheet.getRange(editedRow, 6);

    // Set the current date and time
    timestampCell.setValue(new Date());

    // Set the display format for clarity
    timestampCell.setNumberFormat('MM/dd/yyyy HH:mm:ss');
  }
}
