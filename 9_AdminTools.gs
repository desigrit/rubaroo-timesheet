/**
 * Creates a custom menu in the Google Sheet when the spreadsheet is opened.
 * This consolidated function combines all admin tools into a single menu.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Timesheet Tools')
    .addItem('Apply All Protections', 'applyProtectionsToAllSheets') // New tool
    .addSeparator()
    .addItem('Send Monthly Summaries', 'sendMonthlySummaryEmails')
    .addItem('Recalculate All Leave Entitlements', 'recalculateAllLeaveEntitlements')
    .addSeparator()
    .addItem('Update Holidays in All Sheets', 'updateTemplateHolidays')
    .addItem('Update Leave Types in All Sheets', 'updateExistingTimesheetsWithNewDataValidation')
    .addItem('Update All Comp Offs & Dashboard', 'updateExpiredCompOffsAndDashboard')
    .addToUi();
  Logger.log('Consolidated custom menu created.');
}

/**
 * Admin tool to update the "Leave" dropdown on all timesheets
 * using the Holidays List sheet as the source.
 */
function updateExistingTimesheetsWithNewDataValidation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const validationSourceSheet = ss.getSheetByName(CONFIG.DATA_VALIDATION_SHEET_NAME);
    if (!validationSourceSheet) {
      ui.alert('Error', `The source sheet "${CONFIG.DATA_VALIDATION_SHEET_NAME}" was not found.`, ui.ButtonSet.OK);
      return;
    }

    const leaveTypesValues = validationSourceSheet.getRange(CONFIG.DATA_VALIDATION_LEAVE_TYPES_RANGE)
      .getValues()
      .filter(cell => cell[0] !== "");

    if (leaveTypesValues.length === 0) {
      ui.alert('Error', 'No leave types found in the source range on the "Holidays List" sheet.');
      return;
    }

    // --- UPDATED: Uses the new constant ---
    const startRow = CONFIG.DATA_VALIDATION_START_ROW; 
    const startCol = validationSourceSheet.getRange(CONFIG.DATA_VALIDATION_LEAVE_TYPES_RANGE).getColumn();
    const sourceRange = validationSourceSheet.getRange(startRow, startCol, leaveTypesValues.length);

    const validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(false)
        .setHelpText('Select a leave type from the list.')
        .build();

    const employeeSheetNames = getActiveEmployeeSheetNames_();
    employeeSheetNames.push(CONFIG.TEMPLATE_SHEET);
    
    const confirmProceed = ui.alert('Confirmation', `This will update the "Leave" dropdown on ${employeeSheetNames.length} sheets. Proceed?`, ui.ButtonSet.YES_NO);

    if (confirmProceed === ui.Button.YES) {
      employeeSheetNames.forEach(sheetName => {
        const employeeSheet = ss.getSheetByName(sheetName);
        if (employeeSheet) {
          const targetRange = employeeSheet.getRange(
            CONFIG.EMPLOYEE_DATA_START_ROW,
            CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX,
            employeeSheet.getMaxRows() - CONFIG.EMPLOYEE_DATA_START_ROW + 1
          );
          targetRange.setDataValidation(validationRule);
        }
      });
      ui.alert('Success', `Dropdown menus updated for ${employeeSheetNames.length} sheets.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Cancelled', 'Update cancelled by user.', ui.ButtonSet.OK);
    }
  } catch (e) {
    Logger.log(`Admin tool error: ${e.message}`);
    ui.alert('Error', `An unexpected error occurred: ${e.message}`, ui.ButtonSet.OK);
  }
}


/**
 * Applies advanced protections to all active employee timesheets.
 */
function applyProtectionsToAllSheets() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm', 'This will apply new protections to all active employee timesheets. Continue?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeDataSheet = ss.getSheetByName(CONFIG.EMPLOYEE_LIST_SHEET);
  const employeeData = employeeDataSheet.getRange(2, 1, employeeDataSheet.getLastRow() - 1, CONFIG.EMAIL_COL).getValues();
  
  let protectedCount = 0;
  employeeData.forEach(row => {
    const name = row[CONFIG.EMPLOYEE_NAME_COL - 1];
    const email = row[CONFIG.EMAIL_COL - 1];
    const status = row[CONFIG.STATUS_COL - 1];
    
    if (status === 'Active' && name) {
      const employeeSheet = ss.getSheetByName(name);
      if (employeeSheet) {
        // Remove any existing protections before applying new ones
        const protections = employeeSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        protections.forEach(p => p.remove());
        const rangeProtections = employeeSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        rangeProtections.forEach(p => p.remove());

        applyAdvancedSheetProtection_(employeeSheet, email);
        protectedCount++;
      }
    }
  });

  ui.alert('Complete', `Advanced protections applied to ${protectedCount} timesheets.`, ui.ButtonSet.OK);
}
