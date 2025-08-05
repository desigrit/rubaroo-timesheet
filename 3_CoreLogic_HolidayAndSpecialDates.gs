/**
 * This file contains core logic for managing holidays and employee-specific special dates
 * (birthdays, anniversaries) across the Template and individual employee timesheets.
 */

/**
 * Handles the update of an existing employee's timesheet when their
 * Birthday or Anniversary is edited in the Employees list.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} editedSheet The 'Employees' sheet.
 * @param {number} editedRow The row number in the 'Employees' sheet that was edited.
 * @private
 */
function handleEmployeeSpecialDateEdit_(ss, editedSheet, editedRow) {
  const employeeDataRange = editedSheet.getRange(editedRow, 1, 1, CONFIG.DATE_JOINING_COL); // Fetch up to date joining to get all data
  const employeeData = employeeDataRange.getValues()[0];

  const employeeName = String(employeeData[CONFIG.EMPLOYEE_NAME_COL - 1]).trim();
  const employeeStatus = String(employeeData[CONFIG.STATUS_COL - 1]).trim().toLowerCase();
  const birthday = employeeData[CONFIG.BIRTHDAY_COL - 1];
  const anniversary = employeeData[CONFIG.ANNIVERSARY_COL - 1];

  if (employeeStatus === 'active' && employeeName) {
    const employeeSheet = ss.getSheetByName(employeeName);
    if (employeeSheet) {
      updateEmployeeSheetSpecialDates_(employeeSheet, birthday, anniversary);
      SpreadsheetApp.getUi().alert(`Timesheet for ${employeeName} updated with latest Birthday/Anniversary settings.`);
      Logger.log(`Updated special dates for existing sheet: ${employeeName} due to Birthday/Anniversary edit.`);
    } else {
      Logger.log(`Skipped special date update for ${employeeName}: Sheet not found.`);
    }
  }
}

/**
 * Handles updates within an individual employee's timesheet (e.g., hours, leave type).
 * Triggers comp-off formula application and dashboard recalculation.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} editedSheet The employee's timesheet.
 * @param {number} editedRow The row number that was edited.
 * @param {number} editedColumn The column number that was edited.
 * @param {*} newValue The new value of the edited cell.
 * @param {*} oldValue The old value of the edited cell.
 * @private
 */
function handleEmployeeTimesheetEdit_(ss, editedSheet, editedRow, editedColumn, newValue, oldValue) {
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);
  if (!dashboardSheet) {
    Logger.log('Error: "Dashboard" sheet not found. Cannot update dashboard.');
    return;
  }

  const employeeName = editedSheet.getName(); // Sheet name is employee name

  // --- Logic for Comp Off Utilization: Call the dedicated function ---
  // This must be triggered ONLY when the Leave Type column (G) is edited.
  if (editedColumn === CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX && editedRow > 1) {
      updateCompOffUtilization_(ss, employeeName, newValue, oldValue);
      // We don't 'return' here because other logic (like updating Comp Off Earned in H)
      // might still need to run if column G is also part of those conditions.
  }
  // --- END NEW LOGIC ---

  // --- Logic for updating employee timesheet's Column H (Comp Off Earned) ---
  // This logic runs if the edited column is A, C, D, E, or G.
  if (editedColumn === CONFIG.EMPLOYEE_DATE_COL || // Column A
      editedColumn === CONFIG.EMPLOYEE_CHECK_IN_COL || // Column C
      editedColumn === CONFIG.EMPLOYEE_CHECK_OUT_COL || // Column D
      editedColumn === CONFIG.EMPLOYEE_HOURS_COLUMN_INDEX || // Column E
      editedColumn === CONFIG.EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX) { // Column G

    if (editedRow > 1) { // Don't apply to header row
        const dateCell = editedSheet.getRange(editedRow, CONFIG.EMPLOYEE_DATE_COL); // Column A for date
        const dateValue = dateCell.getValue();
        const compOffEarnedCell = editedSheet.getRange(editedRow, CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL); // Column H
        const currentSheetName = editedSheet.getName();

        if (dateValue instanceof Date) {
          const dayOfWeek = dateValue.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
          const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

          const cellBackgroundColor = dateCell.getBackground();
          const isSpecialDayHighlighted = (cellBackgroundColor.toLowerCase() === CONFIG.HOLIDAY_COLOR.toLowerCase());

          if (isWeekend || isSpecialDayHighlighted) {
            compOffEarnedCell.setFormula(`=getDailyCompOffEarned(A${editedRow}, E${editedRow}, "${currentSheetName}", ${editedRow})`);
            Logger.log(`Applied Comp Off Earned formula to sheet ${employeeName}, row ${editedRow} (Weekend/Special Day).`);
          } else {
            compOffEarnedCell.clearContent();
            Logger.log(`Cleared Comp Off Earned cell on sheet ${employeeName}, row ${editedRow} (Working Day).`);
          }
        } else {
          compOffEarnedCell.clearContent();
          Logger.log(`Cleared Comp Off Earned cell on sheet ${employeeName}, row ${editedRow} (Invalid Date).`);
        }
    }
  }
  // --- End Comp Off Earned update ---

  // Original call to update dashboard (for other leave types/totals)
  // This function will still update other dashboard totals like Casual/Sick Leave.
  // This should run regardless of which column was edited, as it calculates
  // values for other dashboard columns (Total, Casual, Sick)
  updateDashboardCalculations(ss, employeeName);

} // <--- THIS IS THE CORRECT CLOSING BRACE for the handleEmployeeTimesheetEdit_ function

/**
 * Reads holidays from the 'Holidays List' sheet and applies the holiday color
 * to matching dates in the 'Template' sheet (Column A).
 * This function is now generalized to also include employee-specific special dates.
 */
function updateTemplateHolidays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  const holidaySourceSheet = ss.getSheetByName(CONFIG.HOLIDAY_LIST_SOURCE_SHEET);
  const employeeListSheet = ss.getSheetByName(CONFIG.EMPLOYEE_LIST_SHEET);

  if (!templateSheet || !holidaySourceSheet || !employeeListSheet) {
    SpreadsheetApp.getUi().alert('Error: Required sheets (Template, Holidays List, or Employees) not found. Please ensure they exist and are named correctly.');
    Logger.log(`Error: Template (${CONFIG.TEMPLATE_SHEET}), Holidays List (${CONFIG.HOLIDAY_LIST_SOURCE_SHEET}) or Employees (${CONFIG.EMPLOYEE_LIST_SHEET}) sheet not found.`);
    return;
  }

  // 1. Clear any existing backgrounds and description content from the Template sheet, starting from Row 2
  const templateMaxRows = templateSheet.getMaxRows();
  const startRow = 2;

  if (templateMaxRows >= startRow) {
      const fullClearRange = templateSheet.getRange(startRow, 1, templateMaxRows - startRow + 1, CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT);
      fullClearRange.setBackground('#ffffff');

      const descriptionColumnRange = templateSheet.getRange(startRow, CONFIG.TEMPLATE_DESCRIPTION_COLUMN_INDEX, templateMaxRows - startRow + 1, 1);
      descriptionColumnRange.clearContent();
  }

  // 2. Read holidays and descriptions from the source sheet (Holidays List)
  const sourceLastRow = holidaySourceSheet.getLastRow();
  const publicHolidaysMap = new Map();

  if (sourceLastRow >= 2) {
    const holidayDataRaw = holidaySourceSheet.getRange(2, 1, sourceLastRow - 1, 2).getValues();
    for (let i = 0; i < holidayDataRaw.length; i++) {
      const dateValue = holidayDataRaw[i][0];
      const description = holidayDataRaw[i][1] || '';
      if (dateValue instanceof Date) {
        const normalizedDateString = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()).toDateString();
        publicHolidaysMap.set(normalizedDateString, description);
      }
    }
  }
  Logger.log(`Found ${publicHolidaysMap.size} valid public holiday dates with descriptions for Template.`);

  // 3. Apply public holidays to the Template sheet
  const templateLastDataRow = templateSheet.getLastRow();
  if (templateLastDataRow < startRow) {
      Logger.log('Template sheet has no dates to mark or too few rows from row 2 onwards.');
      SpreadsheetApp.getUi().alert('Template sheet has no dates to mark (less than 2 rows of data).');
    return;
  }

  const numRowsToProcess = templateLastDataRow - startRow + 1;
  const templateDatesRange = templateSheet.getRange(startRow, CONFIG.HOLIDAY_DATE_COLUMN_INDEX, numRowsToProcess, 1);
  const templateDates = templateDatesRange.getValues();

  const newBackgrounds = [];
  const newDescriptions = [];

  for (let r = 0; r < numRowsToProcess; r++) {
      newBackgrounds.push(new Array(CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT).fill('#ffffff'));
      newDescriptions.push(['']);
  }

  for (let r = 0; r < templateDates.length; r++) {
    const templateDate = templateDates[r][0];
    if (templateDate instanceof Date) {
      const normalizedTemplateDate = new Date(templateDate.getFullYear(), templateDate.getMonth(), templateDate.getDate()).toDateString();

      if (publicHolidaysMap.has(normalizedTemplateDate)) {
        for (let c = 0; c < CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT; c++) {
            newBackgrounds[r][c] = CONFIG.HOLIDAY_COLOR;
        }
        newDescriptions[r][0] = publicHolidaysMap.get(normalizedTemplateDate);
      }
    } else {
        for (let c = 0; c < CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT; c++) {
            newBackgrounds[r][c] = '#ffffff';
        }
        newDescriptions[r][0] = '';
    }
  }

  templateSheet.getRange(startRow, 1, numRowsToProcess, CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT).setBackgrounds(newBackgrounds);
  templateSheet.getRange(startRow, CONFIG.TEMPLATE_DESCRIPTION_COLUMN_INDEX, numRowsToProcess, 1).setValues(newDescriptions);

  // After updating the template, also update all individual employee sheets
  updateAllEmployeeSheetsSpecialDates_();

  SpreadsheetApp.getUi().alert('Template sheet and all employee timesheets holidays and descriptions updated successfully.');
  Logger.log('Template sheet and all employee timesheets holidays and descriptions updated successfully.');
}

/**
 * REVISED: Helper function to update special dates and comp off formulas.
 * Now adds a formula to grant 0.5 comp off for 23:59 hours on a weekday.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} employeeSheet The specific employee's sheet.
 * @param {Date} birthday The employee's birthday date.
 * @param {Date} anniversary The employee's anniversary date.
 * @private
 */
function updateEmployeeSheetSpecialDates_(employeeSheet, birthday, anniversary) {
  const employeeSheetName = employeeSheet.getName();
  Logger.log(`[updateEmployeeSheetSpecialDates_] START for sheet: ${employeeSheetName}`);

  // ... (The first part of the function that finds and colors special dates remains the same) ...
  const startRow = 2;
  const numRowsToProcess = employeeSheet.getLastRow() - startRow + 1;
  if (numRowsToProcess <= 0) {
    Logger.log(`[updateEmployeeSheetSpecialDates_] No data rows to process for ${employeeSheetName}.`);
    return;
  }
  const employeeSheetDatesRange = employeeSheet.getRange(startRow, CONFIG.HOLIDAY_DATE_COLUMN_INDEX, numRowsToProcess, 1);
  const employeeSheetDates = employeeSheetDatesRange.getValues();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidaySourceSheet = ss.getSheetByName(CONFIG.HOLIDAY_LIST_SOURCE_SHEET);
  const publicHolidaysMap = new Map();
  if (holidaySourceSheet.getLastRow() >= 2) {
    const holidayDataRaw = holidaySourceSheet.getRange(2, 1, holidaySourceSheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < holidayDataRaw.length; i++) {
      const dateValue = holidayDataRaw[i][0];
      const description = holidayDataRaw[i][1] || '';
      if (dateValue instanceof Date) {
        const normalizedDateString = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()).toDateString();
        publicHolidaysMap.set(normalizedDateString, description);
      }
    }
  }
  const newBackgrounds = [];
  const newDescriptions = [];
  for (let r = 0; r < numRowsToProcess; r++) {
    newBackgrounds.push(new Array(CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT).fill('#ffffff'));
    newDescriptions.push(['']);
  }
  for (let r = 0; r < employeeSheetDates.length; r++) {
    const sheetDate = employeeSheetDates[r][0];
    if (!(sheetDate instanceof Date)) { continue; }
    const normalizedSheetDateString = new Date(sheetDate.getFullYear(), sheetDate.getMonth(), sheetDate.getDate()).toDateString();
    let isSpecialDay = false;
    let description = '';
    if (publicHolidaysMap.has(normalizedSheetDateString)) {
      isSpecialDay = true;
      description = publicHolidaysMap.get(normalizedSheetDateString);
    }
    if (birthday instanceof Date && sheetDate.getMonth() === birthday.getMonth() && sheetDate.getDate() === birthday.getDate()) {
      isSpecialDay = true;
      description = (description ? description + ", " : "") + "Birthday";
    }
    if (anniversary instanceof Date && sheetDate.getMonth() === anniversary.getMonth() && sheetDate.getDate() === anniversary.getDate()) {
        isSpecialDay = true;
        description = (description ? description + ", " : "") + "Anniversary";
    }
    if (isSpecialDay) {
      for (let c = 0; c < CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT; c++) {
        newBackgrounds[r][c] = CONFIG.HOLIDAY_COLOR;
      }
      newDescriptions[r][0] = description;
    }
  }
  employeeSheet.getRange(startRow, 1, numRowsToProcess, CONFIG.TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT).setBackgrounds(newBackgrounds);
  employeeSheet.getRange(startRow, CONFIG.TEMPLATE_DESCRIPTION_COLUMN_INDEX, numRowsToProcess, 1).setValues(newDescriptions);

  // --- Logic for Comp Off Earned (Column H) formulas ---
  const compOffEarnedColumn = CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL;
  const compOffFormulas = [];

  for (let r = 0; r < numRowsToProcess; r++) {
    const rowBackgroundColor = newBackgrounds[r][0];
    const rowDate = employeeSheetDates[r][0];

    if (rowDate instanceof Date) {
      const dayOfWeek = rowDate.getDay();
      const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

      if (isWeekend || rowBackgroundColor === CONFIG.HOLIDAY_COLOR) {
        // Keep the existing logic for weekends and special days
        const currentSheetName = employeeSheet.getName();
        compOffFormulas.push([`=getDailyCompOffEarned(A${startRow + r}, E${startRow + r}, "${currentSheetName}", ${startRow + r})`]);
      } else {
        // --- NEW: For weekdays, add a formula to check for 23:59 ---
        compOffFormulas.push([`=IF(TEXT(E${startRow + r}, "HH:mm")="23:59", 0.5, "")`]);
      }
    } else {
      compOffFormulas.push(['']); // Keep blank if no date
    }
  }

  employeeSheet.getRange(startRow, compOffEarnedColumn, numRowsToProcess, 1).setFormulas(compOffFormulas);

  Logger.log(`Finished updating Comp Off Earned (Col H) formulas for ${employeeSheetName}.`);
}


/**
 * Helper function to iterate through all active employee sheets and update their special dates.
 * This is called from `updateTemplateHolidays` to ensure consistency.
 * @private
 */
function updateAllEmployeeSheetsSpecialDates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeListSheet = ss.getSheetByName(CONFIG.EMPLOYEE_LIST_SHEET);
  if (!employeeListSheet) {
    Logger.log('Employee List sheet not found for batch update.');
    return;
  }

  const lastEmployeeRow = employeeListSheet.getLastRow();
  if (lastEmployeeRow < 2) {
    Logger.log('No employees found in the list to update special dates.');
    return;
  }

  const allEmployeeData = employeeListSheet.getRange(2, 1, lastEmployeeRow - 1, CONFIG.DATE_JOINING_COL).getValues();

  for (let i = 0; i < allEmployeeData.length; i++) {
    const rowData = allEmployeeData[i];
    const employeeName = String(rowData[CONFIG.EMPLOYEE_NAME_COL - 1]).trim();
    const employeeStatus = String(rowData[CONFIG.STATUS_COL - 1]).trim().toLowerCase();
    const birthday = rowData[CONFIG.BIRTHDAY_COL - 1];
    const anniversary = rowData[CONFIG.ANNIVERSARY_COL - 1];

    if (employeeStatus === 'active' && employeeName) {
      const employeeSheet = ss.getSheetByName(employeeName);
      if (employeeSheet) {
        updateEmployeeSheetSpecialDates_(employeeSheet, birthday, anniversary);
      } else {
        Logger.log(`Skipped special date update for ${employeeName}: Sheet not found.`);
      }
    } else {
      Logger.log(`Skipped special date update for row ${i + 2}: Employee is not active or name is missing.`);
    }
  }
  Logger.log('Finished updating special dates for all active employee sheets.');
}
