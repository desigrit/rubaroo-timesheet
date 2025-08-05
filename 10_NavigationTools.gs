/**
 * Finds today's date in Column A of the active sheet and scrolls to that row.
 * Assign this script directly to a drawing/button on the sheet.
 */
function scrollToToday() {
  const ui = SpreadsheetApp.getUi(); // Get UI service for alerts
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // This will work on the sheet the user is currently viewing

  const today = new Date();
  today.setHours(0, 0, 0, 0); // Normalize today's date to the start of the day

  // Get all dates from Column A, starting from row 1 to the last row with content
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) { // Sheet is completely empty
    ui.alert('No Data', 'The sheet is empty.', ui.ButtonSet.OK); // Use ui.alert for feedback
    return;
  }

  // Read all values from Column A
  const datesRange = sheet.getRange('A1:A' + lastRow);
  const dates = datesRange.getValues();

  let targetRow = -1; // Initialize with -1 to indicate not found

  // Loop through the dates to find today's date
  for (let i = 0; i < dates.length; i++) {
    const cellValue = dates[i][0]; // Get the value from the current row in Column A

    // Ensure the cell value is a valid Date object before comparing
    if (cellValue instanceof Date) {
      const cellDate = new Date(cellValue.getTime()); // Create a copy to normalize
      cellDate.setHours(0, 0, 0, 0); // Normalize the cell's date

      if (cellDate.getTime() === today.getTime()) {
        targetRow = i + 1; // Found! +1 because arrays are 0-indexed, but sheet rows are 1-indexed
        break; // Exit the loop once found
      }
    }
  }

  // If today's date was found, activate the corresponding cell to scroll to it
  if (targetRow !== -1) {
    sheet.getRange(targetRow, 1).activate(); // Activate cell in Column A of the target row
    Logger.log(`Scrolled to row ${targetRow} for today's date.`);
  } else {
    // If not found, alert the user
    ui.alert('Date Not Found', 'Could not find today\'s date in Column A of this sheet. Please ensure the date exists and is formatted correctly.', ui.ButtonSet.OK);
    Logger.log('Today\'s date not found in Column A.');
  }
}
