/**
 * This file contains core logic related to the creation and initial setup
 * of new employee timesheets and their corresponding dashboard entries,
 * as well as the main dashboard calculation logic for existing employees.
 */

/**
 * Handles the creation of a new employee timesheet from the template
 * and updates the dashboard.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} editedSheet The 'Employees' sheet.
 * @param {number} editedRow The row number in the 'Employees' sheet that was edited.
 * @private
 */
function handleNewEmployeeCreation_(ss, editedSheet, editedRow) {
  Logger.log(`[handleNewEmployeeCreation_] Starting for row ${editedRow}.`);

  const employeeDataRange = editedSheet.getRange(editedRow, 1, 1, CONFIG.DATE_JOINING_COL);
  const employeeData = employeeDataRange.getValues()[0];

  const employeeName = String(employeeData[CONFIG.EMPLOYEE_NAME_COL - 1]).trim();
  const employeeStatus = String(employeeData[CONFIG.STATUS_COL - 1] || '').trim().toLowerCase();
  const employeeEmail = String(employeeData[CONFIG.EMAIL_COL - 1] || '').trim();
  const supervisorEmailsString = String(employeeData[CONFIG.SUPERVISOR_EMAIL_COL - 1] || '').trim();
  const birthday = employeeData[CONFIG.BIRTHDAY_COL - 1];
  const anniversary = employeeData[CONFIG.ANNIVERSARY_COL - 1];

  // --- Validations (no changes here) ---
  if (!employeeName || employeeStatus !== 'active' || ss.getSheetByName(employeeName)) {
    // Error messages are handled inside these checks
    return;
  }
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert('Error: "Template" sheet not found.');
    return;
  }

  const newSheet = templateSheet.copyTo(ss).setName(employeeName);
  SpreadsheetApp.getUi().alert('Successfully created a new timesheet for ' + employeeName + '.');
  
  // --- NEW: Add the employee as an editor to the spreadsheet file ---
  if (isValidEmail_(employeeEmail)) {
    ss.addEditor(employeeEmail);
    Logger.log(`Added ${employeeEmail} as an editor to the spreadsheet.`);
  }
  // --- END NEW ---

  applyAdvancedSheetProtection_(newSheet, employeeEmail);
  
  const joiningDate = employeeData[CONFIG.DATE_JOINING_COL - 1];
  updateDashboardForNewEmployee_(ss, employeeName, joiningDate);
  
  updateEmployeeSheetSpecialDates_(newSheet, birthday, anniversary);
  updateDashboardCalculations(ss, employeeName);

  if (isValidEmail_(employeeEmail)) {
    sendNewTimesheetNotification_(employeeName, employeeEmail, supervisorEmailsString, ss.getUrl());
  }
}

/**
 * REVISED: Applies advanced protection to an employee's timesheet.
 * 1. Protects the whole sheet, adding supervisors AND the employee as editors.
 * 2. Adds a second, more restrictive protection on specific columns that only supervisors can edit.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The employee's timesheet.
 * @param {string} employeeEmail The employee's email address.
 * @private
 */
function applyAdvancedSheetProtection_(sheet, employeeEmail) {
  const supervisorEmails = CONFIG.GLOBAL_HR_SUPERVISOR_EMAILS;

  // 1. Protect the entire sheet.
  const sheetProtection = sheet.protect().setDescription(`Full sheet protection for ${sheet.getName()}`);

  // 2. Add supervisors AND the specific employee as editors for the whole sheet.
  // This gives the employee baseline access.
  sheetProtection.addEditors(supervisorEmails);
  if (employeeEmail && isValidEmail_(employeeEmail)) {
    sheetProtection.addEditor(employeeEmail);
  }

  // 3. Define the columns the employee SHOULD NOT edit.
  const protectedColumns = [
    CONFIG.EMPLOYEE_DATE_COL,
    CONFIG.EMPLOYEE_DAY_COL,
    CONFIG.EMPLOYEE_HOURS_COLUMN_INDEX,
    CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX,
    CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL,
    CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL
  ];

  // 4. Apply a more restrictive protection to these columns.
  protectedColumns.forEach(colIndex => {
    const range = sheet.getRange(CONFIG.EMPLOYEE_DATA_START_ROW, colIndex, sheet.getMaxRows());
    const rangeProtection = range.protect().setDescription(`Protected column for supervisors only`);
    
    // Add supervisors as editors...
    rangeProtection.addEditors(supervisorEmails);
    // ...and explicitly remove the employee from this specific protection.
    if (employeeEmail && isValidEmail_(employeeEmail)) {
      rangeProtection.removeEditor(employeeEmail);
    }
  });
  
  Logger.log(`Applied advanced protections for ${sheet.getName()}.`);
}


/**
 * Initializes Comp Off related columns on the Dashboard for a newly created employee.
 * Comments updated to reflect new column layout.
 * @param {string} employeeName The name of the new employee.
 * @param {number} newDashboardRow The row number for the new employee on the Dashboard.
 * @private
 */
function updateCompOffUtilizationForNewEmployee_(employeeName, newDashboardRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!dashboardSheet) {
    Logger.log('[updateCompOffUtilizationForNewEmployee_] Error: Dashboard sheet not found.');
    return;
  }

  if (employeeName.toLowerCase() === CONFIG.TEMPLATE_SHEET.toLowerCase()) {
    // Clear the cells for Comp Off Earned (I), Available (J), Utilized Display (K), Storage (L), Expired (M)
    dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_EARNED_COL).clearContent();
    dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_AVAILABLE_COL).clearContent();
    dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL).clearContent();
    dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_STORAGE_COL).clearContent();
    dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL).clearContent();
    return;
  }

  // Set the Total Comp Off Earned formula (Column I on Dashboard)
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_EARNED_COL).setFormula(`=SUMIF('${employeeName}'!H:H, ">0")`);

  // Initialize the hidden Comp Off Utilized storage (Column L on Dashboard) to 0.
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_STORAGE_COL).setValue(0);

  // Set the Comp Off Utilized display formula (Column K on Dashboard) to reference the storage column L
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL).setFormula(`=L${newDashboardRow}`);

  // Initialize Expired Comp Offs (Column M on Dashboard) to 0.
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL).setValue(0);

  // Set the Comp Off Available formula (Column J on Dashboard) using the I-K-M structure
  const colI_dashboard = String.fromCharCode(64 + CONFIG.DASHBOARD_COMP_OFF_EARNED_COL);
  const colK_dashboard_display = String.fromCharCode(64 + CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL);
  const colM_dashboard = String.fromCharCode(64 + CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_AVAILABLE_COL).setFormula(`=${colI_dashboard}${newDashboardRow}-(${colK_dashboard_display}${newDashboardRow}+${colM_dashboard}${newDashboardRow})`);
}

/**
 * REVISED: Updates the Dashboard sheet with a new employee's entry.
 * It now automatically adds the two-color SPARKLINE visualization formulas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} employeeName The name of the new employee.
 * @param {Date} joiningDate The date the employee joined.
 * @private
 */
function updateDashboardForNewEmployee_(ss, employeeName, joiningDate) {
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Error: "Dashboard" sheet not found.');
    return;
  }

  const newDashboardRow = dashboardSheet.getLastRow() + 1;
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_EMPLOYEE_NAME_COL).setValue(employeeName);

  // Set Prorated Leave
  const entitlements = calculateProratedLeave_(joiningDate, new Date());
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_COL).setValue(entitlements.casual);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_COL).setValue(entitlements.sick);

  // Set Utilized Formulas
  const casualUtilizedFormula = `=COUNTIF('${employeeName}'!G:G, "${CONFIG.LEAVE_TYPE_CASUAL}") + COUNTIF('${employeeName}'!G:G, "${CONFIG.LEAVE_TYPE_CASUAL_HALF_DAY}")*0.5`;
  const sickUtilizedFormula = `=COUNTIF('${employeeName}'!G:G, "${CONFIG.LEAVE_TYPE_SICK}") + COUNTIF('${employeeName}'!G:G, "${CONFIG.LEAVE_TYPE_SICK_HALF_DAY}")*0.5`;
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL).setFormula(casualUtilizedFormula);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL).setFormula(sickUtilizedFormula);

  // Set up Comp Off Columns
  updateCompOffUtilizationForNewEmployee_(employeeName, newDashboardRow);

  // --- UPDATED Sparkline Formulas ---
  const clUtilized = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL).getA1Notation();
  const clAllocated = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_COL).getA1Notation();
  const casualSparklineFormula = `=IFERROR(SPARKLINE({${clUtilized}/${clAllocated}, 1-(${clUtilized}/${clAllocated})}, {"charttype","bar"; "max",1; "color1","green"; "color2","lightgrey"}))`;

  const slUtilized = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL).getA1Notation();
  const slAllocated = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_COL).getA1Notation();
  const sickSparklineFormula = `=IFERROR(SPARKLINE({${slUtilized}/${slAllocated}, 1-(${slUtilized}/${slAllocated})}, {"charttype","bar"; "max",1; "color1","green"; "color2","lightgrey"}))`;

  const coUtilized = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL).getA1Notation();
  const coEarned = dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_EARNED_COL).getA1Notation();
  const compOffSparklineFormula = `=IFERROR(SPARKLINE({${coUtilized}/${coEarned}, 1-(${coUtilized}/${coEarned})}, {"charttype","bar"; "max",1; "color1","green"; "color2","red"}))`;

  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_CASUAL_VISUAL_COL).setFormula(casualSparklineFormula);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_SICK_VISUAL_COL).setFormula(sickSparklineFormula);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_COMP_OFF_VISUAL_COL).setFormula(compOffSparklineFormula);
  // --- END UPDATED ---

  // Set Total Leave Formula
  const colC = String.fromCharCode(64 + CONFIG.DASHBOARD_CASUAL_LEAVE_COL);
  const colF = String.fromCharCode(64 + CONFIG.DASHBOARD_SICK_LEAVE_COL);
  const colJ = String.fromCharCode(64 + CONFIG.DASHBOARD_COMP_OFF_AVAILABLE_COL);
  dashboardSheet.getRange(newDashboardRow, CONFIG.DASHBOARD_TOTAL_LEAVE_COL).setFormula(`=SUM(${colC}${newDashboardRow},${colF}${newDashboardRow},${colJ}${newDashboardRow})`);

  Logger.log(`Fully configured new dashboard row for ${employeeName}.`);
}

/**
 * Updates various leave calculations on the Dashboard for a specific employee.
 * This function is typically triggered by edits on the employee's timesheet,
 * or explicitly called after new employee creation.
 * It calculates and sets values for Total Leave, Casual Utilized, Sick Utilized,
 * and Comp Off Utilized, Earned, and Expired based on the employee's timesheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {string} employeeName The name of the employee (which is also their timesheet name).
 */

function updateDashboardCalculations(ss, employeeName) {
  Logger.log(`[updateDashboardCalculations] Starting dashboard update for employee: ${employeeName}`);

  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!dashboardSheet) {
    Logger.log('Error: "Dashboard" sheet not found in updateDashboardCalculations. Cannot proceed.');
    return;
  }

  const employeeSheet = ss.getSheetByName(employeeName);
  if (!employeeSheet) {
    Logger.log(`Error: Employee sheet '${employeeName}' not found for dashboard calculations.`);
    return;
  }

  // 1. Find the employee's row on the Dashboard
  const employeeNamesRange = dashboardSheet.getRange(CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW, CONFIG.DASHBOARD_EMPLOYEE_NAME_COL, dashboardSheet.getLastRow() - CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW + 1, 1);
  const employeeNames = employeeNamesRange.getValues();
  let employeeDashboardRow = -1;

  for (let i = 0; i < employeeNames.length; i++) {
    if (String(employeeNames[i][0]).trim() === employeeName) {
      employeeDashboardRow = i + CONFIG.DASHBOARD_EMPLOYEE_DATA_START_ROW;
      break;
    }
  }

  if (employeeDashboardRow === -1) {
    Logger.log(`Employee '${employeeName}' not found on Dashboard for calculations. Cannot update dashboard row.`);
    return;
  }

  // 2. Read relevant data from the employee's timesheet
  const lastRowEmployeeSheet = employeeSheet.getLastRow();

  if (lastRowEmployeeSheet < CONFIG.EMPLOYEE_DATA_START_ROW) {
    // If no data rows (only header), set all counts to 0 and exit.
    dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL).setValue(0);
    dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL).setValue(0);
    dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_STORAGE_COL).setValue(0);
    dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_COMP_OFF_EARNED_COL).setValue(0);
    dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL).setValue(0);
    Logger.log(`[updateDashboardCalculations] No data in ${employeeName}'s timesheet. Resetting all related dashboard fields to 0.`);
    return;
  }

  const allEmployeeDataRange = employeeSheet.getRange(
    CONFIG.EMPLOYEE_DATA_START_ROW,
    CONFIG.EMPLOYEE_DATE_COL, // Column A
    lastRowEmployeeSheet - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
    CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL // Column I
  );
  const allEmployeeData = allEmployeeDataRange.getValues();

  // 3. Calculate leave totals
  let casualLeaveUtilizedDays = 0;
  let sickLeaveUtilizedDays = 0;
  let compOffUtilizedDays = 0;
  let trueTotalCompOffEarned = 0; // Correctly named variable for total earned
  let totalCompOffExpired = 0;

  const casualLeaveFull = CONFIG.LEAVE_TYPE_CASUAL.toLowerCase();
  const casualLeaveHalf = CONFIG.LEAVE_TYPE_CASUAL_HALF_DAY.toLowerCase();
  const sickLeaveFull = CONFIG.LEAVE_TYPE_SICK.toLowerCase();
  const sickLeaveHalf = CONFIG.LEAVE_TYPE_SICK_HALF_DAY.toLowerCase();
  const compOffFull = CONFIG.LEAVE_TYPE_COMP_OFF.toLowerCase();
  const compOffHalf = CONFIG.LEAVE_TYPE_COMP_OFF_HALF_DAY.toLowerCase();

  for (let i = 0; i < allEmployeeData.length; i++) {
    const rowData = allEmployeeData[i];
    const leaveType = String(rowData[CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX - CONFIG.EMPLOYEE_DATE_COL] || '').trim().toLowerCase(); // Col G
    const earnedCompOff = parseFloat(rowData[CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL - CONFIG.EMPLOYEE_DATE_COL]); // Col H
    const compOffStatus = String(rowData[CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_DATE_COL] || '').trim().toLowerCase(); // Col I

    // --- NEW, CORRECTED LOGIC ---

    // A. Calculate true total earned by summing Column H, regardless of status.
    if (!isNaN(earnedCompOff) && earnedCompOff > 0) {
      trueTotalCompOffEarned += earnedCompOff;
    }

    // B. Calculate utilized days from Column G.
    if (leaveType === casualLeaveFull) {
      casualLeaveUtilizedDays += 1;
    } else if (leaveType === casualLeaveHalf) {
      casualLeaveUtilizedDays += CONFIG.HALF_DAY_VALUE;
    } else if (leaveType === sickLeaveFull) {
      sickLeaveUtilizedDays += 1;
    } else if (leaveType === sickLeaveHalf) {
      sickLeaveUtilizedDays += CONFIG.HALF_DAY_VALUE;
    } else if (leaveType === compOffFull) {
      compOffUtilizedDays += 1;
    } else if (leaveType === compOffHalf) {
      compOffUtilizedDays += CONFIG.HALF_DAY_VALUE;
    }

    // C. Calculate total expired from Column I.
    if (!isNaN(earnedCompOff) && earnedCompOff > 0 && compOffStatus === 'expired') {
      totalCompOffExpired += earnedCompOff;
    }
  }

  // 4. Update Dashboard cells with the calculated totals
  dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL).setValue(casualLeaveUtilizedDays); // Column D
  dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL).setValue(sickLeaveUtilizedDays); // Column G

  // Update Comp Off specific dashboard columns
  dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_COMP_OFF_UTILIZED_STORAGE_COL).setValue(compOffUtilizedDays); // Column L (storage)
  dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_COMP_OFF_EARNED_COL).setValue(trueTotalCompOffEarned); // Column I (total earned)
  dashboardSheet.getRange(employeeDashboardRow, CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL).setValue(totalCompOffExpired); // Column M

  Logger.log(`[updateDashboardCalculations] Dashboard for ${employeeName} updated...`);
}

/**
 * Retrieves the names of all active employee timesheet sheets in the spreadsheet.
 * Filters out system sheets like 'Dashboard', 'Employees', etc.
 * @returns {string[]} An array of employee sheet names.
 * @private
 */
function getActiveEmployeeSheetNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const employeeSheetNames = [];

  // Use CONFIG.EXCLUDE_FROM_DASHBOARD for consistency
  const systemSheetNames = CONFIG.EXCLUDE_FROM_DASHBOARD.map(name => name.trim().toLowerCase()); // Normalize to lowercase

  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    if (!systemSheetNames.includes(sheetName.trim().toLowerCase())) { // Compare normalized names
      employeeSheetNames.push(sheetName.trim());
    }
  }
  Logger.log(`[getActiveEmployeeSheetNames_] Found employee sheets: ${employeeSheetNames.join(', ')}`);
  return employeeSheetNames;
}

// Helper to check for valid email format (can be in a separate Utilities file if exists)
function isValidEmail_(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
