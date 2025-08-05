/**
 * This file contains logic for calculating prorated leave entitlements
 * based on joining date and probation periods.
 */

/**
 * Recalculates the total allocated Casual and Sick leave for all active employees
 * and updates the Dashboard. This is intended to be run manually by an admin from the menu.
 */
function recalculateAllLeaveEntitlements() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Action',
    'This will recalculate and overwrite the allocated Casual and Sick leave on the Dashboard for all active employees. Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'The operation was cancelled.', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = ss.getSheetByName(CONFIG.EMPLOYEE_LIST_SHEET);
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);

  if (!employeeSheet || !dashboardSheet) {
    ui.alert('Error', 'Could not find the "Employees" or "Dashboard" sheet.', ui.ButtonSet.OK);
    return;
  }

  // Get all employee data (Name and Date of Joining)
  const employeeData = employeeSheet.getRange(
    CONFIG.EMPLOYEE_DATA_START_ROW, 1,
    employeeSheet.getLastRow() - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
    CONFIG.DATE_JOINING_COL
  ).getValues();

  // Create a map of Dashboard names to their row number for efficient updates
  const dashboardNames = dashboardSheet.getRange(
    CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW, CONFIG.DASHBOARD_EMPLOYEE_NAME_COL,
    dashboardSheet.getLastRow() - CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW + 1, 1
  ).getValues();

  const dashboardRowMap = new Map();
  dashboardNames.forEach((name, index) => {
    if (name[0]) {
      dashboardRowMap.set(name[0].toString().trim(), index + CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW);
    }
  });

  const currentDate = new Date();
  let updatedCount = 0;

  // Loop through each employee and update their dashboard entry
  employeeData.forEach(row => {
    const name = row[CONFIG.EMPLOYEE_NAME_COL - 1];
    const status = row[CONFIG.STATUS_COL - 1];
    const joiningDate = row[CONFIG.DATE_JOINING_COL - 1];

    if (status === 'Active' && name && joiningDate instanceof Date) {
      const entitlements = calculateProratedLeave_(joiningDate, currentDate);
      const dashboardRow = dashboardRowMap.get(name.trim());

      if (dashboardRow) {
        dashboardSheet.getRange(dashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_COL).setValue(entitlements.casual);
        dashboardSheet.getRange(dashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_COL).setValue(entitlements.sick);
        updatedCount++;
        Logger.log(`Updated ${name}: Casual=${entitlements.casual}, Sick=${entitlements.sick}`);
      }
    }
  });

  ui.alert('Success', `Leave entitlements have been recalculated for ${updatedCount} active employees.`, ui.ButtonSet.OK);
}


/**
 * REVISED: Calculates the allocated sick and casual leaves for a single employee,
 * accounting for annual expiry on March 31st.
 * @private
 * @param {Date} joiningDate The employee's date of joining.
 * @param {Date} currentDate The date to calculate entitlements up to.
 * @returns {{sick: number, casual: number}} An object with total sick and casual leaves.
 */
function calculateProratedLeave_(joiningDate, currentDate) {
  let sickLeave = 0;
  let casualLeave = 0;

  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();
  
  const financialYearStartDate = (currentMonth < 3)
    ? new Date(currentYear - 1, 3, 1)
    : new Date(currentYear, 3, 1);

  const calculationStartDate = new Date(Math.max(joiningDate.getTime(), financialYearStartDate.getTime()));

  const probationEndDate = new Date(joiningDate.getTime());
  probationEndDate.setMonth(probationEndDate.getMonth() + 3);

  let monthIterator = new Date(calculationStartDate.getTime());
  monthIterator.setDate(1);

  while (monthIterator <= currentDate) {
    sickLeave++;

    if (monthIterator >= probationEndDate) {
      casualLeave++;
    }

    monthIterator.setMonth(monthIterator.getMonth() + 1);
  }

  return { sick: sickLeave, casual: casualLeave };
}
