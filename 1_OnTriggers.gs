/**
 * This file contains functions that act as event triggers for the Google Sheet.
 * - onSheetEdit: Responds to cell edits for sheet creation, holiday updates, comp-off calculation, and status changes.
 */

/**
 * A special function that runs when a user changes a value in the spreadsheet.
 * REVISED: Includes a confirmation dialog when changing status to Inactive.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onSheetEdit(e) {
  Logger.log('--- onSheetEdit TRIGGER FIRED ---');
  const ss = e.source;
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const editedSheetName = editedSheet.getName();
  const editedRow = editedRange.getRow();
  const editedColumn = editedRange.getColumn();
  const newValue = e.value; // Using e.value for direct string comparison needed for checkbox logic (if reinstated) and status check
  const oldValue = e.oldValue;

  Logger.log(`[onSheetEdit] Trigger fired for range: ${editedRange.getA1Notation()} in sheet: ${editedSheetName}. Row: ${editedRow}, Col: ${editedColumn}. NewValue: '${newValue}', OldValue: '${oldValue}'.`);

  // --- Confirmation for changing status to Inactive ---
  if (
    editedSheetName === CONFIG.EMPLOYEE_LIST_SHEET &&
    editedColumn === CONFIG.STATUS_COL &&
    editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW &&
    newValue && newValue.toString().trim().toLowerCase() === 'inactive' && // Check if NEW value is 'inactive'
    oldValue && oldValue.toString().trim().toLowerCase() !== 'inactive' // Ensure OLD value wasn't already 'inactive'
  ) {
    const ui = SpreadsheetApp.getUi();
    const confirmationMessage = "You're changing this employee's status to Inactive. After you confirm, new Casual and Sick leave allocations will stop, and their comp-off expiry will be frozen. If the employee has been inactive for a while, you may need to update the casual and sick leave allocations manually.";
    const response = ui.alert('Confirm Status Change', confirmationMessage, ui.ButtonSet.YES_NO);

    if (response === ui.Button.NO) {
      // If user cancels, revert the change
      editedRange.setValue(oldValue);
      SpreadsheetApp.getActiveSpreadsheet().toast('Status change cancelled.');
      Logger.log(`[onSheetEdit] Status change to Inactive for row ${editedRow} cancelled by user.`);
      return; // Stop further processing
    } else {
      // If user confirms, log it and allow the edit to proceed
      Logger.log(`[onSheetEdit] User confirmed status change to Inactive for row ${editedRow}.`);
      // No 'return' here, let other logic like dashboard updates run if needed (though unlikely necessary for status change alone)
    }
  }

  // Get list of active employee sheets for faster lookup (moved down to avoid running if status change is cancelled)
  const activeEmployeeSheets = getActiveEmployeeSheetNames_();

  // --- Logic for creating new employee sheets from 'Employees' list ---
  if (editedSheetName === CONFIG.EMPLOYEE_LIST_SHEET && editedColumn === CONFIG.DATE_JOINING_COL && editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW && newValue) {
    handleNewEmployeeCreation_(ss, editedSheet, editedRow);
    return; // Exit after handling new employee creation
  }

  // --- Logic for updating employee timesheets when Birthday/Anniversary is edited in Employees list ---
  if (editedSheetName === CONFIG.EMPLOYEE_LIST_SHEET &&
    (editedColumn === CONFIG.BIRTHDAY_COL || editedColumn === CONFIG.ANNIVERSARY_COL) &&
    editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW) {
    handleEmployeeSpecialDateEdit_(ss, editedSheet, editedRow);
    return;
  }

  // --- Logic for updates within an individual employee's timesheet ---
  if (activeEmployeeSheets.includes(editedSheetName) && editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW) {

    // --- Specific handling for Comp Off Utilization/Return (Column G: Leave Type) ---
    if (editedColumn === CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX) {
      // Need to get the actual value from the range for comparisons, as e.value might be empty string for blank cells
      const actualNewValue = editedRange.getValue();
      const trimmedNewValue = String(actualNewValue || '').trim();
      const trimmedOldValue = String(oldValue || '').trim();

      if ([CONFIG.LEAVE_TYPE_COMP_OFF, CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY].includes(trimmedNewValue)) {
        Logger.log(`[onSheetEdit] Comp Off leave type detected in ${editedSheetName}. Calling handleCompOffUtilization_.`);
        handleCompOffUtilization_(ss, editedSheet, editedRow); // Handle consumption
      } else if ([CONFIG.LEAVE_TYPE_COMP_OFF, CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY].includes(trimmedOldValue) && !trimmedNewValue) {
        Logger.log(`[onSheetEdit] Comp Off leave type deleted/reset in ${editedSheetName}. Calling handleCompOffReturn_.`);
        handleCompOffReturn_(ss, editedSheet, editedRow, trimmedOldValue); // Handle return
      }

      Logger.log(`[onSheetEdit] Calling updateDashboardCalculations for sheet ${editedSheetName} due to Column G edit.`);
      updateDashboardCalculations(ss, editedSheetName);
      return; // Handled this edit, exit.
    }

    // --- Specific handling for Comp Off Earned (Column H) ---
    if (editedColumn === CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL) {
      Logger.log(`[onSheetEdit] Comp Off Earned detected in ${editedSheetName}. Calling updateDashboardCalculations.`);
      updateDashboardCalculations(ss, editedSheetName);
      return; // Handled this edit, exit.
    }

    // --- Generic handler for other relevant edits in employee timesheets ---
    if (editedColumn === CONFIG.EMPLOYEE_HOURS_COLUMN_INDEX ||
      editedColumn === CONFIG.EMPLOYEE_CHECK_IN_COL ||
      editedColumn === CONFIG.EMPLOYEE_CHECK_OUT_COL) {
      Logger.log(`[onSheetEdit] General timesheet edit in ${editedSheetName}. Calling updateDashboardCalculations.`);
      updateDashboardCalculations(ss, editedSheetName);
      return;
    }
  }
  Logger.log(`[onSheetEdit] No specific action for this edit: ${editedSheetName}, Col: ${editedColumn}, Row: ${editedRow}.`);
}

// Helper to update all dashboards (assuming updateDashboardCalculations can be called for all)
// This is not called by any menu item in the current configuration, can be removed if desired.
function updateAllDashboards_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheetNames = getActiveEmployeeSheetNames_();
  Logger.log('Manually refreshing dashboard for all active employees...');
  employeeSheetNames.forEach(sheetName => {
    updateDashboardCalculations(ss, sheetName);
  });
  SpreadsheetApp.getUi().alert('Dashboard Refresh', 'All active employee dashboards have been refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
}
