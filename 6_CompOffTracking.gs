/**
 * This file contains logic for tracking and managing Comp Offs (earning, utilization, expiry, return).
 */

/**
 * REVISED: Handles the logic for utilizing Comp Offs when an employee enters a Comp Off leave type
 * in Column G of their timesheet. It now prevents using Comp Offs earned after the leave date.
 * Now includes a 60-day grace period to allow consumption of 'Expired' Comp Offs
 * for retroactive timesheet updates.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} employeeSheet The employee's timesheet.
 * @param {number} editedRow The row where the leave type was edited.
 */
function handleCompOffUtilization_(ss, employeeSheet, editedRow) {
  Logger.log(`[handleCompOffUtilization_] Starting for ${employeeSheet.getName()}, row ${editedRow}.`);

  const leaveTypeCell = employeeSheet.getRange(editedRow, CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX);
  const leaveType = String(leaveTypeCell.getValue() || '').trim();
  
  clearCompOffError_(employeeSheet, editedRow);

  if (![CONFIG.LEAVE_TYPE_COMP_OFF, CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY].includes(leaveType)) {
    return;
  }

  // ... (The core logic for validation and consumption remains the same) ...
  const compOffAmountToConsume = (leaveType === CONFIG.LEAVE_TYPE_COMP_OFF) ? 1 : CONFIG.HALF_DAY_VALUE;
  const lastRow = employeeSheet.getLastRow();
  if (lastRow < CONFIG.EMPLOYEE_DATA_START_ROW) {
    displayCompOffError_(employeeSheet, editedRow, "No comp offs recorded for this employee.", false);
    leaveTypeCell.clearContent();
    return;
  }
  const compOffDataRange = employeeSheet.getRange(
    CONFIG.EMPLOYEE_DATA_START_ROW, CONFIG.EMPLOYEE_DATE_COL,
    lastRow - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
    CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL + 1
  );
  const compOffData = compOffDataRange.getValues();
  let totalAvailableForConsumption = 0;
  let sortedEarnedCompOffsForConsumption = []; 
  const currentTimesheetDate = employeeSheet.getRange(editedRow, CONFIG.EMPLOYEE_DATE_COL).getValue();
  if (!(currentTimesheetDate instanceof Date) || isNaN(currentTimesheetDate.getTime())) {
      displayCompOffError_(employeeSheet, editedRow, "Invalid date in Column A for this leave entry.", true);
      leaveTypeCell.clearContent();
      return;
  }
  currentTimesheetDate.setHours(0,0,0,0);
  for (let i = 0; i < compOffData.length; i++) {
    const row = compOffData[i];
    const earnedDate = row[CONFIG.EMPLOYEE_DATE_COL - CONFIG.EMPLOYEE_DATE_COL];
    const earnedAmount = parseFloat(row[CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL - CONFIG.EMPLOYEE_DATE_COL]);
    let compOffStatus = String(row[CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL] || '').trim().toLowerCase();
    if (!(earnedDate instanceof Date) || isNaN(earnedAmount) || earnedAmount <= 0) { continue; }
    const normalizedEarnedDate = new Date(earnedDate.getTime());
    normalizedEarnedDate.setHours(0,0,0,0);
    if (normalizedEarnedDate > currentTimesheetDate) { continue; }
    const expiryDate = new Date(earnedDate.getTime());
    expiryDate.setDate(expiryDate.getDate() + CONFIG.COMP_OFF_EXPIRY_DAYS);
    expiryDate.setHours(0,0,0,0);
    if (currentTimesheetDate > expiryDate) {
      if (compOffStatus !== 'expired' && compOffStatus !== 'consumed') {
        employeeSheet.getRange(CONFIG.EMPLOYEE_DATA_START_ROW + i, CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL).setValue('Expired');
      }
      continue;
    }
    if (compOffStatus !== 'consumed') {
      let actualRemainingAmount = earnedAmount;
      if (compOffStatus.endsWith('remaining')) { actualRemainingAmount = parseFloat(compOffStatus); }
      if (actualRemainingAmount > 0) {
        totalAvailableForConsumption += actualRemainingAmount;
        sortedEarnedCompOffsForConsumption.push({
          date: earnedDate, amount: actualRemainingAmount,
          originalRow: CONFIG.EMPLOYEE_DATA_START_ROW + i,
          statusCellRange: employeeSheet.getRange(CONFIG.EMPLOYEE_DATA_START_ROW + i, CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL)
        });
      }
    }
  }
  sortedEarnedCompOffsForConsumption.sort((a, b) => a.date.getTime() - b.date.getTime());
  if (totalAvailableForConsumption < compOffAmountToConsume) {
    const errorMsg = `Insufficient Comp Offs available or valid for this date. Needed ${compOffAmountToConsume}, but only ${totalAvailableForConsumption} are available within the 60-day limit.`;
    displayCompOffError_(employeeSheet, editedRow, errorMsg, false);
    leaveTypeCell.clearContent();
    return;
  }
  let remainingToConsume = compOffAmountToConsume;
  for (let i = 0; i < sortedEarnedCompOffsForConsumption.length && remainingToConsume > 0; i++) {
    const earnedCo = sortedEarnedCompOffsForConsumption[i];
    if (earnedCo.amount >= remainingToConsume) {
      earnedCo.amount -= remainingToConsume;
      remainingToConsume = 0;
      if (earnedCo.amount === 0) { earnedCo.statusCellRange.setValue('Consumed'); } 
      else { earnedCo.statusCellRange.setValue(`${earnedCo.amount} Remaining`); }
    } else {
      remainingToConsume -= earnedCo.amount;
      earnedCo.statusCellRange.setValue('Consumed');
    }
  }

  // --- NEW: Log the successful utilization ---
  try {
    const logSheet = ss.getSheetByName(CONFIG.COMP_OFF_LOG_SHEET);
    if (logSheet) {
      const timezone = ss.getSpreadsheetTimeZone();
      const format = "dd-MMM-yyyy";
      
      const formattedUpdatedOn = Utilities.formatDate(new Date(), timezone, format);
      const dateUtilizedObject = employeeSheet.getRange(editedRow, CONFIG.EMPLOYEE_DATE_COL).getValue();
      const formattedDateUtilized = Utilities.formatDate(dateUtilizedObject, timezone, format); // Format this date too
      
      const employeeName = employeeSheet.getName();
      const usedAmount = leaveTypeCell.getValue();
      
      logSheet.appendRow([formattedUpdatedOn, employeeName, formattedDateUtilized, usedAmount, '']);
      Logger.log(`Logged Comp Off utilization for ${employeeName}.`);
    } else {
      Logger.log(`Warning: Compensatory Leave Log sheet not found.`);
    }
  } catch (logError) {
    Logger.log(`Error logging comp off utilization: ${logError.message}`);
  }
  // --- END NEW LOGIC ---

  Logger.log(`[handleCompOffUtilization_] Successfully processed consumption for row ${editedRow}.`);
}

/**
 * Handles the return of Comp Offs when a 'Comp Off' leave entry is deleted from Column G.
 * The returned amount is credited back to the most recently consumed/remaining comp offs (LIFO).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} employeeSheet The employee's timesheet.
 * @param {number} editedRow The row where the leave type was deleted.
 * @param {string} oldLeaveType The type of Comp Off that was removed (e.g., "Comp Off", "Comp Off - Half day").
 */
/**
 * REVISED: Handles the return of Comp Offs and deletes the corresponding log entry.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} employeeSheet The employee's timesheet.
 * @param {number} editedRow The row where the leave type was deleted.
 * @param {string} oldLeaveType The type of Comp Off that was removed.
 */
function handleCompOffReturn_(ss, employeeSheet, editedRow, oldLeaveType) {
  Logger.log(`[handleCompOffReturn_] Starting for ${employeeSheet.getName()}, row ${editedRow}.`);

  const compOffAmountToReturn = (oldLeaveType === CONFIG.LEAVE_TYPE_COMP_OFF) ? 1 : CONFIG.HALF_DAY_VALUE;
  
  // --- NEW: Delete the corresponding entry from the log ---
   try {
    const logSheet = ss.getSheetByName(CONFIG.COMP_OFF_LOG_SHEET);
    if (logSheet && logSheet.getLastRow() > 1) {
      const employeeName = employeeSheet.getName();
      const dateUtilizedObject = employeeSheet.getRange(editedRow, CONFIG.EMPLOYEE_DATE_COL).getValue();
      // Format the date from the timesheet to match the string format in the log
      const formattedDateToFind = Utilities.formatDate(dateUtilizedObject, ss.getSpreadsheetTimeZone(), "dd-MMM-yyyy");

      // Get display values to read the formatted date strings from the log
      const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 3).getDisplayValues();
      
      for (let i = logData.length - 1; i >= 0; i--) {
        const logRow = logData[i];
        const logEmployeeName = logRow[1]; // Column B
        const logDateUtilizedString = logRow[2]; // Column C

        // Compare the names and the formatted date strings
        if (logEmployeeName === employeeName && logDateUtilizedString === formattedDateToFind) {
            const rowToDelete = i + 2;
            logSheet.deleteRow(rowToDelete);
            Logger.log(`Deleted Comp Off log entry for ${employeeName} on row ${rowToDelete}.`);
            break; 
        }
      }
    }
  } catch (logError) {
      Logger.log(`Error deleting comp off log entry: ${logError.message}`);
  }
  // --- END NEW LOGIC ---

  // ... (The core logic for returning/crediting the leave remains the same) ...
  const lastRow = employeeSheet.getLastRow();
  if (lastRow < CONFIG.EMPLOYEE_DATA_START_ROW) { return; }
  const compOffDataRange = employeeSheet.getRange(
    CONFIG.EMPLOYEE_DATA_START_ROW, CONFIG.EMPLOYEE_DATE_COL,
    lastRow - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
    CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL + 1
  );
  const compOffData = compOffDataRange.getValues();
  let affectedEarnedCompOffs = []; 
  for (let i = 0; i < compOffData.length; i++) {
    const row = compOffData[i];
    const earnedDate = row[CONFIG.EMPLOYEE_DATE_COL - CONFIG.EMPLOYEE_DATE_COL];
    const earnedAmount = parseFloat(row[CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL - CONFIG.EMPLOYEE_DATE_COL]);
    let compOffStatus = String(row[CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL] || '').trim().toLowerCase();
    if (!(earnedDate instanceof Date) || isNaN(earnedAmount) || earnedAmount <= 0) { continue; }
    if (compOffStatus === 'consumed' || compOffStatus.endsWith('remaining')) {
      let actualAmountInStatus = earnedAmount;
      if (compOffStatus.endsWith('remaining')) { actualAmountInStatus = parseFloat(compOffStatus); }
      const currentUnconsumed = (compOffStatus === 'consumed') ? 0 : actualAmountInStatus;
      if (currentUnconsumed < earnedAmount) {
        affectedEarnedCompOffs.push({
          date: earnedDate, initialAmount: earnedAmount,
          currentAmountInStatus: currentUnconsumed,
          originalRow: CONFIG.EMPLOYEE_DATA_START_ROW + i,
          statusCellRange: employeeSheet.getRange(CONFIG.EMPLOYEE_DATA_START_ROW + i, CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL)
        });
      }
    }
  }
  affectedEarnedCompOffs.sort((a, b) => b.date.getTime() - a.date.getTime());
  let amountToCredit = compOffAmountToReturn;
  for (let i = 0; i < affectedEarnedCompOffs.length && amountToCredit > 0; i++) {
    const earnedCo = affectedEarnedCompOffs[i];
    const capacity = earnedCo.initialAmount - earnedCo.currentAmountInStatus;
    if (capacity <= 0) { continue; }
    const creditAmount = Math.min(amountToCredit, capacity);
    earnedCo.currentAmountInStatus += creditAmount;
    amountToCredit -= creditAmount;
    if (earnedCo.currentAmountInStatus === earnedCo.initialAmount) {
      earnedCo.statusCellRange.setValue('');
    } else {
      earnedCo.statusCellRange.setValue(`${earnedCo.currentAmountInStatus} Remaining`);
    }
  }
  if (amountToCredit > 0) {
    Logger.log(`[handleCompOffReturn_] Warning: Could not fully credit back ${amountToCredit} comp off.`);
  }
  Logger.log(`[handleCompOffReturn_] Finished processing comp off return for row ${editedRow}.`);
}


/**
 * Displays an error message on the timesheet cell (Column G) where Comp Off was attempted.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The employee's timesheet.
 * @param {number} row The row number.
 * @param {string} message The error message to display.
 * @param {boolean} [highlightRed=true] Optional. If true, highlights the cell red. Defaults to true.
 * @private
 */
function displayCompOffError_(sheet, row, message, highlightRed = true) {
  const cell = sheet.getRange(row, CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX);
  if (highlightRed) {
    cell.setBackground('red'); // Highlight in red only if true
  } else {
    cell.setBackground(null); // Ensure no highlight if highlightRed is false
  }
  cell.setNote(message);    // Add a note with the error
  SpreadsheetApp.getUi().alert('Comp Off Error', message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(`[displayCompOffError_] Displayed error for row ${row}: ${message}. Highlighted: ${highlightRed}`);
}

/**
 * Clears any error display (background color and note) from a timesheet cell.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The employee's timesheet.
 * @param {number} row The row number.
 * @private
 */
function clearCompOffError_(sheet, row) {
  const cell = sheet.getRange(row, CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX);
  cell.setBackground(null); // Clear background color
  cell.setNote('');         // Clear the note
  Logger.log(`[clearCompOffError_] Cleared error display for row ${row}.`);
}

/**
 * Periodically checks for expired Comp Offs and updates their status to 'Expired'.
 * This function is intended to be run by a time-driven trigger.
 */
function updateExpiredCompOffsAndDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheetNames = getActiveEmployeeSheetNames_();
  Logger.log('[updateExpiredCompOffsAndDashboard] Starting daily check for expired comp offs.');

  employeeSheetNames.forEach(sheetName => {
    const employeeSheet = ss.getSheetByName(sheetName);
    if (!employeeSheet) {
      Logger.log(`[updateExpiredCompOffsAndDashboard] Skipping ${sheetName}: sheet not found.`);
      return;
    }

    const lastRow = employeeSheet.getLastRow();
    if (lastRow < CONFIG.EMPLOYEE_DATA_START_ROW) {
      Logger.log(`[updateExpiredCompOffsAndDashboard] No data rows in ${sheetName}. Skipping expiry check.`);
      return;
    }

    // Read only the columns relevant for expiry check: Date (A), Comp Off Earned (H), Comp Off Status (I)
    const dataRange = employeeSheet.getRange(
      CONFIG.EMPLOYEE_DATA_START_ROW,
      CONFIG.EMPLOYEE_DATE_COL,
      lastRow - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
      CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL + 1 // Span from A to I
    );
    const values = dataRange.getValues();
    let dashboardUpdateNeeded = false;

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize today's date to midnight

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const earnedDate = row[CONFIG.EMPLOYEE_DATE_COL - CONFIG.EMPLOYEE_DATE_COL]; // Column A
      const earnedAmount = parseFloat(row[CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL - CONFIG.EMPLOYEE_DATE_COL]); // Column H
      const currentStatus = String(row[CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL] || '').trim().toLowerCase(); // Column I

      if (earnedDate instanceof Date && !isNaN(earnedAmount) && earnedAmount > 0) {
        const expiryDate = new Date(earnedDate.getTime());
        expiryDate.setDate(expiryDate.getDate() + CONFIG.COMP_OFF_EXPIRY_DAYS);
        expiryDate.setHours(0,0,0,0); // Normalize expiry date to midnight

        // Only mark as expired if it's past its expiry date AND it's not already expired AND it hasn't been fully consumed.
        // A 'Consumed' comp off is effectively utilized and doesn't need to be marked 'Expired' in addition.
        if (today > expiryDate && currentStatus !== 'expired' && currentStatus !== 'consumed') {
          employeeSheet.getRange(CONFIG.EMPLOYEE_DATA_START_ROW + i, CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL).setValue('Expired');
          Logger.log(`[updateExpiredCompOffsAndDashboard] Marked Comp Off from row ${CONFIG.EMPLOYEE_DATA_START_ROW + i} in ${sheetName} as Expired.`);
          dashboardUpdateNeeded = true;
        }
      }
    }
    
    if (dashboardUpdateNeeded) {
      Logger.log(`[updateExpiredCompOffsAndDashboard] Comp Offs expired in ${sheetName}. Updating dashboard calculations.`);
      updateDashboardCalculations(ss, sheetName);
    }
  });
  Logger.log('[updateExpiredCompOffsAndDashboard] Finished daily check for expired comp offs.');
}

// Helper functions (assuming they are defined in other files as per your project structure)
// getActiveEmployeeSheetNames_() // In 2_CoreLogic_TimesheetCreation.gs
// updateDashboardCalculations(ss, sheetName) // In 2_CoreLogic_TimesheetCreation.gs
