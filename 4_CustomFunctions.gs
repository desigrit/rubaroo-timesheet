/**
 * This file contains custom functions that can be used directly in the Google Sheet.
 * These functions perform calculations that might be too complex for standard sheet formulas.
 */

/**
 * Custom function to fetch all active employee sheet names.
 * This should be used on the Dashboard for lists of employees.
 * @customfunction
 * @returns {string[][]} A 2D array of employee names.
 */
function getEmployeeNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheets = getActiveEmployeeSheetNames_(); // Assumed to be defined in 2_CoreLogic_TimesheetCreation.gs
  return employeeSheets.map(name => [name]);
}

/**
 * Custom function to retrieve public holiday dates and descriptions.
 * This is intended for use in the 'Holidays' sheet.
 * @customfunction
 * @returns {Array<Array<Date|string>>} A 2D array of public holidays (Date, Description).
 */
function getHolidayDates_() { // Using _ as it's primarily a helper for the script, though exposed via a custom function as well
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidaySheet = ss.getSheetByName(CONFIG.HOLIDAY_LIST_SOURCE_SHEET);

  if (!holidaySheet) {
    Logger.log(`Error: Holiday sheet "${CONFIG.HOLIDAY_LIST_SOURCE_SHEET}" not found.`);
    // Return an empty array or throw an error depending on desired behavior in sheet
    return [];
  }

  const lastRow = holidaySheet.getLastRow();
  if (lastRow < CONFIG.HOLIDAY_LIST_START_ROW) {
    // No data rows found
    return [];
  }

  // Get data from Holiday Date column (A) and Holiday Description column (B)
  const dataRange = holidaySheet.getRange(
    CONFIG.HOLIDAY_LIST_START_ROW,
    CONFIG.HOLIDAY_DATE_COL,
    lastRow - CONFIG.HOLIDAY_LIST_START_ROW + 1,
    CONFIG.HOLIDAY_DESCRIPTION_COL - CONFIG.HOLIDAY_DATE_COL + 1
  );
  const values = dataRange.getValues();

  // Filter for valid dates and return
  return values.filter(row => row[0] instanceof Date);
}


/**
 * Custom function to calculate Comp Off earned for a given day based on weekend/holiday and hours worked.
 * To be used in Column H of employee timesheets.
 * This function gets the background color of the date cell (Column A) in the same row
 * to determine if it's a special day (holiday/birthday/anniversary).
 *
 * @customfunction
 * @param {Date} date The date of the entry (from Column A).
 * @param {number|Date} hoursDuration The hours worked, can be a number (e.g. 0.5 for 30 mins) or a Date object (e.g. 18:00) from Column E.
 * @param {string} sheetName The name of the current sheet (e.g., "Employee - John Doe"). Use SHEET_NAME() in the spreadsheet.
 * @param {number} rowNum The row number of the current cell. Use ROW() in the spreadsheet.
 * @returns {number} 1 for full Comp Off, 0.5 for half Comp Off, 0 otherwise.
 */
function getDailyCompOffEarned(date, hoursDuration, sheetName, rowNum) {
  // Return 0 silently if date is invalid (common for empty rows)
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return 0;
  }

  let actualHours;
  if (typeof hoursDuration === 'number') {
    actualHours = hoursDuration * 24; // If duration is a number (e.g., 0.5 for 30 min), convert to hours.
  } else if (hoursDuration instanceof Date) {
    // If duration is a Date object (e.g., 18:00), extract hours and minutes
    const hours = hoursDuration.getHours();
    const minutes = hoursDuration.getMinutes();
    const seconds = hoursDuration.getSeconds();
    actualHours = hours + (minutes / 60) + (seconds / 3600);
  } else {
    // Return 0 if hoursDuration is neither a number nor a Date object
    return 0;
  }

  // No Comp Off for 0 or negative hours
  if (actualHours <= 0) {
    return 0;
  }

  // Normalize date to check for weekend
  const dayOfWeek = date.getDay(); // 0 for Sunday, 6 for Saturday
  const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

  let isSpecialDayHighlighted = false;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Get the background color of the date cell (Column A) in the same row
      const dateCellRange = sheet.getRange(rowNum, CONFIG.EMPLOYEE_DATE_COL);
      const cellBackgroundColor = dateCellRange.getBackground();

      // Check if it's highlighted as a holiday, birthday, or anniversary
      // We check for all relevant colors, assuming CONFIG.HOLIDAY_COLOR covers all special day highlights
      if (cellBackgroundColor.toLowerCase() === CONFIG.HOLIDAY_COLOR.toLowerCase() ||
          cellBackgroundColor.toLowerCase() === CONFIG.BIRTHDAY_COLOR.toLowerCase() ||
          cellBackgroundColor.toLowerCase() === CONFIG.ANNIVERSARY_COLOR.toLowerCase()) {
        isSpecialDayHighlighted = true;
      }
    }
  } catch (e) {
    // Log errors in getting background, but proceed as if not a special day
    Logger.log(`[getDailyCompOffEarned] Error checking background color for row ${rowNum} on sheet ${sheetName}: ${e.message}`);
    isSpecialDayHighlighted = false;
  }

  // A Comp Off is earned on a weekend or a highlighted special day
  if (isWeekend || isSpecialDayHighlighted) {
    if (actualHours >= CONFIG.FULL_DAY_COMP_OFF_HOURS) { // E.g., > 5 hours
      return 1; // Full Comp Off
    } else if (actualHours >= CONFIG.HALF_DAY_COMP_OFF_HOURS && actualHours < CONFIG.FULL_DAY_COMP_OFF_HOURS) { // E.g., 3-5 hours
      return 0.5; // Half Comp Off
    }
  }

  return 0; // No Comp Off earned
}


/**
 * Custom function to calculate the total available Comp Offs for an employee.
 * This sums earned Comp Offs and subtracts consumed/expired ones.
 * To be used in Column J of employee timesheets and potentially on the Dashboard.
 *
 * @customfunction
 * @param {string} sheetName The name of the employee's timesheet.
 * @returns {number} The net available Comp Off balance.
 */
function getCompOffsAvailable(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = ss.getSheetByName(sheetName);

  if (!employeeSheet) {
    // Logger.log(`[getCompOffsAvailable] Sheet "${sheetName}" not found. Returning 0.`); // Remove verbose logging
    return 0;
  }

  const lastRow = employeeSheet.getLastRow();
  if (lastRow < CONFIG.EMPLOYEE_DATA_START_ROW) {
    return 0; // No data rows
  }

  // Read data from Column H (Comp Off Earned) and Column I (Comp Off Status/Consumed)
  const compOffDataRange = employeeSheet.getRange(
    CONFIG.EMPLOYEE_DATA_START_ROW,
    CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL,
    lastRow - CONFIG.EMPLOYEE_DATA_START_ROW + 1,
    CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL + 1
  );
  const compOffData = compOffDataRange.getValues();

  let totalEarned = 0;
  let totalUtilized = 0;
  let totalExpired = 0; // Keeping track of expired for accuracy, though not displayed directly in available

  for (let i = 0; i < compOffData.length; i++) {
    const earnedAmount = parseFloat(compOffData[i][0]); // Column H (index 0 in 'row' array)
    const status = String(compOffData[i][1] || '').trim().toLowerCase(); // Column I (index 1 in 'row' array)

    if (isNaN(earnedAmount) || earnedAmount <= 0) {
      continue; // Skip invalid or zero earned entries
    }

    totalEarned += earnedAmount;

    if (status === 'consumed') {
      totalUtilized += earnedAmount; // If fully consumed, add full earned amount to utilized
    } else if (status.endsWith('remaining')) {
      // If 'X.X Remaining', then (earnedAmount - X.X) was utilized
      const remainingAmount = parseFloat(status);
      if (!isNaN(remainingAmount)) {
        totalUtilized += (earnedAmount - remainingAmount);
      }
    } else if (status === 'expired') {
        totalExpired += earnedAmount;
    }
  }

  // Available = Total Earned - Total Utilized - Total Expired
  // Note: If a comp off is "Consumed", it's already counted in totalUtilized.
  // If a comp off is "Expired", it is not available for use, so it's effectively "used up" in terms of availability.
  // The logic in handleCompOffUtilization_ and handleCompOffReturn_ ensures that 'Expired' and 'Consumed' are distinct terminal states for an earned CO.
  // Therefore, for 'available', we subtract both.
  const availableBalance = totalEarned - totalUtilized - totalExpired;

  // Return the calculated balance, ensure it's not negative if floating point issues
  return Math.max(0, availableBalance);
}
