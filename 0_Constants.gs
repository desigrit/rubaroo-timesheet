/**
 * Configuration constants to make the script easy to manage.
 * This file contains all global settings for column indices, sheet names, and colors.
 */
const CONFIG = {
  EMPLOYEE_LIST_SHEET: 'Employees',
  TEMPLATE_SHEET: 'Template',
  EXCLUDE_FROM_DASHBOARD: ['Dashboard', 'Instructions', 'Comp Off Log', 'Template', 'Employees', 'Holidays List'],
  DASHBOARD_SHEET: 'Dashboard',
  HOLIDAY_SHEET: 'Template', // Sheet where holidays are applied and detected by getHolidayDates_()
  HOLIDAY_LIST_SOURCE_SHEET: 'Holidays List', // Sheet where the master list of holidays resides
  HOLIDAY_DATE_COLUMN_INDEX: 1, // Column A is index 1 for holiday dates in Template sheet
  HOLIDAY_COLOR: '#ffd966', // New color code for holidays (light yellow/gold)
  TEMPLATE_TIMESHEET_COLUMNS_TO_HIGHLIGHT: 10, // Number of columns (starting from A) to highlight for holidays in Template.
  TEMPLATE_DESCRIPTION_COLUMN_INDEX: 6, // Column F is index 6, where holiday descriptions will be written
  EMPLOYEE_HOURS_COLUMN_INDEX: 5, // Column E is index 5
 
  // Column indices for the 'Employees' sheet (1-indexed)
  EMPLOYEE_NAME_COL: 1,      // Column A
  STATUS_COL: 2,             // Column B
  EMAIL_COL: 3,              // Column C
  DESIGNATION_COL: 4,        // Column D
  SUPERVISOR_NAME_COL: 5,    // Column E
  SUPERVISOR_EMAIL_COL: 6,   // Column F (Source for dynamic protection emails)
  MENTOR_NAME_COL: 7,        // Column G
  BIRTHDAY_COL: 8,           // Column H for Birthday
  ANNIVERSARY_COL: 9,        // Column I for Anniversary
  DATE_JOINING_COL: 10,

  EMPLOYEE_DATA_START_ROW: 2,
  // Employee Timesheet Columns (used for reading data)
  EMPLOYEE_DATE_COL: 1, // Column A (Date)
  EMPLOYEE_DAY_COL: 2, // Column B (Day)
  EMPLOYEE_CHECK_IN_COL: 3, // Column C (Check-in time)
  EMPLOYEE_CHECK_OUT_COL: 4, // Column D (Check-out time)
  EMPLOYEE_HOURS_COLUMN_INDEX: 5, // Column E (Calculated/Entered Hours)
  EMPLOYEE_ACTIVITY_COL: 6, // Column F (Activity Description)
  EMPLOYEE_LEAVE_TYPE_COLUMN_INDEX: 7, // Column G (Leave Type: Casual, Sick, Comp off, etc.) - This is the critical one!
  EMPLOYEE_COMP_OFF_EARNED_COL: 8, // Column H (Comp Off Earned)
  EMPLOYEE_COMP_OFF_STATUS_COL: 9, // Column I (Hidden Comp off management column)
  EMPLOYEE_NOTES_COL: 10,     // Column J (Feelings)


// Dashboard Sheet Columns
DASHBOARD_EMPLOYEE_DATA_START_ROW: 3,

DASHBOARD_EMPLOYEE_NAME_COL: 1, // Column A 
DASHBOARD_TOTAL_LEAVE_COL: 2, // Column B 

DASHBOARD_CASUAL_LEAVE_COL: 3, // Column C (Now "Earned")
DASHBOARD_CASUAL_LEAVE_AVAILABLE_COL: 4, // Column D (New "Available" column)
DASHBOARD_CASUAL_LEAVE_UTILIZED_COL: 5, // Column E
DASHBOARD_CASUAL_VISUAL_COL: 6,  // Column F

DASHBOARD_SICK_LEAVE_COL: 7, // Column G (Now "Earned")
DASHBOARD_SICK_LEAVE_AVAILABLE_COL: 8, // Column H (New "Available" column)
DASHBOARD_SICK_LEAVE_UTILIZED_COL: 9, // Column I
DASHBOARD_SICK_VISUAL_COL: 10,    // Column J

DASHBOARD_COMP_OFF_EARNED_COL: 11, // Column K
DASHBOARD_COMP_OFF_AVAILABLE_COL: 12, // Column L
DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL: 13, // Column M
DASHBOARD_COMP_OFF_UTILIZED_STORAGE_COL: 14, // Column N
DASHBOARD_EXPIRED_COMP_OFFS_COL: 15,         // Column O
DASHBOARD_COMP_OFF_VISUAL_COL: 16,           // Column P


  // Leave Type Keywords
  LEAVE_TYPE_CASUAL: "Casual Leave",
  LEAVE_TYPE_SICK: "Sick Leave",
  LEAVE_TYPE_COMP_OFF: "Comp Off", // Exact string that employees will select for Comp Off. Case-sensitive by default, so ensure it matches.
  LEAVE_TYPE_CASUAL_HALF_DAY: "Casual Leave - Half day",
  LEAVE_TYPE_SICK_HALF_DAY: "Sick Leave - Half day",
  LEAVE_TYPE_COMP_OFF_HALF_DAY: "Comp Off - Half day",

  HALF_DAY_VALUE: 0.5,

  // --- Comp Off Calculation Thresholds ---
  FULL_DAY_COMP_OFF_HOURS: 5, // Hours >= this value for 1 Comp Off (e.g., 5 hours)
  HALF_DAY_COMP_OFF_HOURS: 3, // Hours >= this value for 0.5 Comp Off (e.g., 3 hours)

  COMP_OFF_LOG_SHEET: 'Comp Off Log',

  // List of leaves
  DATA_VALIDATION_SHEET_NAME: 'Holidays List',
  DATA_VALIDATION_LEAVE_TYPES_RANGE: 'D2:D',
  DATA_VALIDATION_START_ROW: 2, // The row where the leave type list starts

  // Comp Off Expiry (New constant)
  COMP_OFF_EXPIRY_DAYS: 60, // Set to 60 days as per common practice, adjust as needed.

  //List of emails that the timesheet and other notification emails go to
  NOTIFICATION_CC_EMAILS: ['people@rubaroo.org' , 'monisha@rubaroo.org', 'pavani@rubaroo.org', 'imtiaz@rubaroo.org', 'uma@rubaroo.org'],

  // List of global HR/Admin email addresses who always have access to protected ranges.
  GLOBAL_HR_SUPERVISOR_EMAILS: ['monisha@rubaroo.org', 'pavani@rubaroo.org', 'imtiaz@rubaroo.org' , 'people@rubaroo.org' , 'raunakoberoi@gmail.com','uma@rubaroo.org']

};
