/**
 * This file contains functions that act as event triggers for the Google Sheet.
 * - onSheetEdit: Responds to cell edits for sheet creation, holiday updates, and comp-off calculation.
 */

/**
 * A special function that runs when a user changes a value in the spreadsheet.
 * This will create a new employee sheet from the template, update the dashboard,
 * and also trigger dashboard/timesheet updates for employee timesheet edits.
 * IMPORTANT: This function is linked via an installable trigger, NOT a simple trigger.
 * Make sure the function name here (onSheetEdit) matches the one in your Apps Script Triggers.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object.
 */
function onSheetEdit(e) {
  const ss = e.source;
  const editedRange = e.range;
  const editedSheet = editedRange.getSheet();
  const editedSheetName = editedSheet.getName();
  const editedRow = editedRange.getRow();
  const editedColumn = editedRange.getColumn();
  const newValue = editedRange.getValue();
  const oldValue = e.oldValue; // Get the old value for comparison if needed

  // Ensure CONFIG is loaded (though it should be global if defined correctly)
  if (typeof CONFIG === 'undefined' || Object.keys(CONFIG).length === 0) {
    Logger.log('CONFIG not loaded. Attempting to load...');
    // In a real scenario, you'd ensure CONFIG is loaded or handle gracefully.
    // Assuming 0_Constants.gs is always part of the project and loaded first.
  }

  // Get list of active employee sheets for faster lookup.
  const activeEmployeeSheets = getActiveEmployeeSheetNames_();
  Logger.log(`[onSheetEdit] Trigger fired for range: ${editedRange.getA1Notation()} in sheet: ${editedSheetName}. Row: ${editedRow}, Col: ${editedColumn}. NewValue: '${newValue}', OldValue: '${oldValue}'.`);

  // --- Logic for creating new employee sheets from 'Employees' list ---
  if (editedSheetName === CONFIG.EMPLOYEE_LIST_SHEET && editedColumn === CONFIG.DATE_JOINING_COL && editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW && newValue) {
    Logger.log(`[onSheetEdit] Calling handleNewEmployeeCreation_ for row ${editedRow} in ${CONFIG.EMPLOYEE_LIST_SHEET}.`);
    handleNewEmployeeCreation_(ss, editedSheet, editedRow);
    return; // Exit after handling new employee creation
  }

  // --- Logic for updating employee timesheets when Birthday/Anniversary is edited in Employees list ---
  if (editedSheetName === CONFIG.EMPLOYEE_LIST_SHEET &&
    (editedColumn === CONFIG.BIRTHDAY_COL || editedColumn === CONFIG.ANNIVERSARY_COL) &&
    editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW) {
    Logger.log(`[onSheetEdit] Calling handleEmployeeSpecialDateEdit_ for row ${editedRow} in ${CONFIG.EMPLOYEE_LIST_SHEET}.`);
    handleEmployeeSpecialDateEdit_(ss, editedSheet, editedRow);
    return;
  }

  // --- Logic for updates within an individual employee's timesheet ---
  if (activeEmployeeSheets.includes(editedSheetName) && editedRow >= CONFIG.EMPLOYEE_DATA_START_ROW) {

    // --- Specific handling for Comp Off Utilization/Return (Column G: Leave Type) ---
    if (editedColumn === CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX) { // Column G is Leave Type
      const trimmedNewValue = String(newValue || '').trim();
      const trimmedOldValue = String(oldValue || '').trim();

      // Case 1: Comp Off is entered (new value is a Comp Off type)
      if ([CONFIG.LEAVE_TYPE_COMP_OFF, CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY].includes(trimmedNewValue)) {
        Logger.log(`[onSheetEdit] Comp Off leave type detected in ${editedSheetName}. Calling handleCompOffUtilization_.`);
        handleCompOffUtilization_(ss, editedSheet, editedRow); // Handle consumption
      }
      // Case 2: Comp Off is deleted/removed (old value was Comp Off type, new value is empty)
      else if ([CONFIG.LEAVE_TYPE_COMP_OFF, CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY].includes(trimmedOldValue) && !trimmedNewValue) {
        Logger.log(`[onSheetEdit] Comp Off leave type deleted/reset in ${editedSheetName}. Calling handleCompOffReturn_.`);
        handleCompOffReturn_(ss, editedSheet, editedRow, trimmedOldValue); // Handle return
      }
      // Note: If old was Comp Off, new is non-Comp Off, it's treated as a deletion of Comp Off
      // and a new entry of something else. The `handleCompOffReturn_` will run,
      // and then `updateDashboardCalculations` will catch the new value if it's another leave type.

      // Regardless of comp off or other leave, update dashboard calculations.
      // This will ensure all leave types affect dashboard, and comp off calculations are refreshed.
      Logger.log(`[onSheetEdit] Calling updateDashboardCalculations for sheet ${editedSheetName} due to Column G edit.`);
      updateDashboardCalculations(ss, editedSheetName);
      return; // Handled this edit, exit.
    }

    // --- Specific handling for Comp Off Earned (Column H) ---
    if (editedColumn === CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL) { // Column H is Comp Off Earned
      Logger.log(`[onSheetEdit] Comp Off Earned detected in ${editedSheetName}. Calling updateDashboardCalculations.`);
      // If a comp off is earned, just update dashboard calculations.
      // The `handleCompOffUtilization_` function relies on `updateDashboardCalculations`
      // to correctly tally available/consumed amounts from Column I.
      updateDashboardCalculations(ss, editedSheetName);
      return; // Handled this edit, exit.
    }

    // --- Generic handler for other relevant edits in employee timesheets ---
    // If you have other columns on individual timesheets that trigger dashboard updates
    // (e.g., Check-in/Check-out, regular hours), add their conditions here.
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
// This might need to iterate through all active employee sheets.
function updateAllDashboards_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheetNames = getActiveEmployeeSheetNames_();
  Logger.log('Manually refreshing dashboard for all active employees...');
  employeeSheetNames.forEach(sheetName => {
    updateDashboardCalculations(ss, sheetName);
  });
  SpreadsheetApp.getUi().alert('Dashboard Refresh', 'All active employee dashboards have been refreshed.', SpreadsheetApp.getUi().ButtonSet.OK);
}
