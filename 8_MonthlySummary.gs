/**
 * This file contains logic for generating and emailing monthly leave summaries to employees.
 */

/**
 * Main function to be triggered manually or by a time-based trigger.
 */
function sendMonthlySummaryEmails() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirm Action', 'This will send a leave summary email to all active employees. Are you sure you want to proceed?', ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled', 'Operation cancelled.', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = ss.getSheetByName(CONFIG.EMPLOYEE_LIST_SHEET);
  const dashboardSheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET);

  if (!employeeSheet || !dashboardSheet) {
    ui.alert('Error', 'Could not find the "Employees" or "Dashboard" sheet.', ui.ButtonSet.OK);
    return;
  }

  const employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, CONFIG.DATE_JOINING_COL).getValues();
  const dashboardData = dashboardSheet.getRange(3, 1, dashboardSheet.getLastRow() - 2, CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL).getValues();
  const dashboardMap = new Map();
  dashboardData.forEach(row => {
    const name = row[CONFIG.DASHBOARD_EMPLOYEE_NAME_COL - 1];
    if (name) {
      dashboardMap.set(name, row);
    }
  });

  let sentCount = 0;
  employeeData.forEach(employee => {
    const name = employee[CONFIG.EMPLOYEE_NAME_COL - 1];
    const email = employee[CONFIG.EMAIL_COL - 1];
    const status = employee[CONFIG.STATUS_COL - 1];

    if (status === 'Active' && name && isValidEmail_(email)) {
      const dashRow = dashboardMap.get(name);
      if (dashRow) {
        generateAndSendIndividualSummary_(ss, name, email, dashRow);
        sentCount++;
      }
    }
  });

  ui.alert('Process Complete', `Monthly summary emails have been sent to ${sentCount} active employees.`, ui.ButtonSet.OK);
}


/**
 * Gathers data, creates charts, and sends the summary email for a single employee.
 * @private
 */
function generateAndSendIndividualSummary_(ss, employeeName, employeeEmail, dashboardRow) {
  try {
    // 1. Gather All Data from Dashboard
    const casualEarned = dashboardRow[CONFIG.DASHBOARD_CASUAL_LEAVE_COL - 1];
    const casualUtilized = dashboardRow[CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL - 1];
    const casualAvailable = dashboardRow[CONFIG.DASHBOARD_CASUAL_LEAVE_AVAILABLE_COL - 1];
    const sickEarned = dashboardRow[CONFIG.DASHBOARD_SICK_LEAVE_COL - 1];
    const sickUtilized = dashboardRow[CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL - 1];
    const sickAvailable = dashboardRow[CONFIG.DASHBOARD_SICK_LEAVE_AVAILABLE_COL - 1];
    const compOffEarned = dashboardRow[CONFIG.DASHBOARD_COMP_OFF_EARNED_COL - 1];
    const compOffAvailable = dashboardRow[CONFIG.DASHBOARD_COMP_OFF_AVAILABLE_COL - 1];
    const compOffUtilized = dashboardRow[CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL - 1];
    const compOffExpired = dashboardRow[CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL - 1];

    // 2. Calculate Comp Offs expiring this month (logic is unchanged)
    let expiringThisMonth = 0;
    // ...

    // 3. Create Charts (or get error text if balance is negative)
    const casualChartObject = createPieChart_('Casual Leave', casualUtilized, casualAvailable, casualEarned);
    const sickChartObject = createPieChart_('Sick Leave', sickUtilized, sickAvailable, sickEarned);
    const compOffChartObject = createThreeSlicePieChart_('Comp Off Summary', compOffUtilized, compOffAvailable, compOffExpired, compOffEarned);

    // 4. Build HTML and inlineImages object conditionally
    const inlineImages = {};
    let casualChartHtml, sickChartHtml, compOffChartHtml;

    if (typeof casualChartObject === 'string') {
      casualChartHtml = `<p style="color:red; font-family:sans-serif;"><b>Casual Leave:</b> ${casualChartObject}</p>`;
    } else {
      casualChartHtml = `<img src="cid:casualChart">`;
      inlineImages.casualChart = casualChartObject;
    }

    if (typeof sickChartObject === 'string') {
      sickChartHtml = `<p style="color:red; font-family:sans-serif;"><b>Sick Leave:</b> ${sickChartObject}</p>`;
    } else {
      sickChartHtml = `<img src="cid:sickChart">`;
      inlineImages.sickChart = sickChartObject;
    }
    
    if (typeof compOffChartObject === 'string') {
      compOffChartHtml = `<p style="color:red; font-family:sans-serif;"><b>Comp Offs:</b> ${compOffChartObject}</p>`;
    } else {
      compOffChartHtml = `<img src="cid:compOffChart">`;
      inlineImages.compOffChart = compOffChartObject;
    }

    // 5. Build and Send Email
    const monthName = new Date().toLocaleString('default', { month: 'long' });
    const subject = `${employeeName}, your summary for ${monthName} is here!`;
    const htmlBody = `
      <html><body>
        <p>Hi ${employeeName},</p>
        <p>Another fun month is here! You can find your leave balances below -</p>
        <table border="1" cellpadding="5" style="border-collapse: collapse; font-family:sans-serif;">
          <tr style="background-color:#f2f2f2;"><th>Casual Leaves Available</th><th>Sick Leaves Available</th><th>Comp Offs Available</th><th>Comp Offs Expiring This Month</th></tr>
          <tr><td style="text-align:center;">${casualAvailable}</td><td style="text-align:center;">${sickAvailable}</td><td style="text-align:center;">${compOffAvailable}</td><td style="text-align:center;">${expiringThisMonth}</td></tr>
        </table>
        <p>Your current leave utilization:</p>
        ${casualChartHtml}
        ${sickChartHtml}
        ${compOffChartHtml}
        <p>We're here to support your rest and recovery - so work with your supervisor to take time off!</p>
        <p>For the team, By the team</p>
      </body></html>`;

    MailApp.sendEmail({
      to: employeeEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: "For the team, By the team",
      inlineImages: inlineImages
    });

  } catch (e) {
    Logger.log(`Failed to send summary for ${employeeName}. Error: ${e.message}`);
  }
}

/**
 * Helper function to create a standardized two-slice pie chart.
 * Returns an error string if balance is negative.
 * @private
 */
function createPieChart_(title, utilized, available, earned) {
  // NEW: Check for negative balance
  if (utilized > earned) {
    return "You have a negative leave balance, so your data could not be visualized.";
  }
  if ((utilized + available) === 0) {
    const data = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days").addRow(["No Leave Data", 1]).build();
    return Charts.newPieChart().setDataTable(data).setColors(['#cccccc']).setTitle(title).setDimensions(300, 200).build().getAs('image/png');
  }
  const data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days")
      .addRow([`Utilized (${utilized})`, utilized]).addRow([`Available (${available})`, available]).build();
  return Charts.newPieChart().setDataTable(data)
      .setColors(['#2fafb6', '#0E3436']).setTitle(title).setDimensions(300, 200).build().getAs('image/png');
}

/**
 * Helper function to create a standardized three-slice pie chart for Comp Offs.
 * Returns an error string if balance is negative.
 * @private
 */
function createThreeSlicePieChart_(title, utilized, available, expired, earned) {
  // NEW: Check for negative balance
  if (utilized > earned) {
    return "You have a negative leave balance, so your data could not be visualized.";
  }
  if ((utilized + available + expired) === 0) {
    const data = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days").addRow(["No Comp Off Data", 1]).build();
    return Charts.newPieChart().setDataTable(data).setColors(['#cccccc']).setTitle(title).setDimensions(300, 200).build().getAs('image/png');
  }
  const data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days")
      .addRow([`Utilized (${utilized})`, utilized]).addRow([`Available (${available})`, available]).addRow([`Expired (${expired})`, expired]).build();
  return Charts.newPieChart().setDataTable(data)
      .setColors(['#2fafb6', '#0E3436', '#C0E7E9']).setTitle(title).setDimensions(300, 200).build().getAs('image/png');
}
