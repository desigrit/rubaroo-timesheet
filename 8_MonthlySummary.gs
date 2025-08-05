/**
 * This file contains logic for generating and emailing monthly leave summaries to employees.
 */

/**
 * Main function to be triggered manually or by a time-based trigger.
 * It iterates through all active employees and sends them a personalized summary email.
 */
function sendMonthlySummaryEmails() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Action',
    'This will send a leave summary email to all active employees. Are you sure you want to proceed?',
    ui.ButtonSet.YES_NO
  );

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
  const dashboardData = dashboardSheet.getRange(3, 1, dashboardSheet.getLastRow() - 2, CONFIG.DASHBOARD_COMP_OFF_VISUAL_COL).getValues();
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
    // 1. Gather Data from Dashboard
    const casualAvailable = dashboardRow[CONFIG.DASHBOARD_CASUAL_LEAVE_COL - 1];
    const casualUtilized = dashboardRow[CONFIG.DASHBOARD_CASUAL_LEAVE_UTILIZED_COL - 1];
    const sickAvailable = dashboardRow[CONFIG.DASHBOARD_SICK_LEAVE_COL - 1];
    const sickUtilized = dashboardRow[CONFIG.DASHBOARD_SICK_LEAVE_UTILIZED_COL - 1];
    const compOffAvailable = dashboardRow[CONFIG.DASHBOARD_COMP_OFF_AVAILABLE_COL - 1];
    const compOffUtilized = dashboardRow[CONFIG.DASHBOARD_COMP_OFF_UTILIZED_DISPLAY_COL - 1];
    const compOffExpired = dashboardRow[CONFIG.DASHBOARD_EXPIRED_COMP_OFFS_COL - 1];

    // 2. Calculate Comp Offs expiring this month
    let expiringThisMonth = 0;
    const employeeSheet = ss.getSheetByName(employeeName);
    if (employeeSheet) {
      const compOffData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL).getValues();
      const today = new Date();
      const currentMonth = today.getMonth();
      const currentYear = today.getFullYear();

      compOffData.forEach(row => {
        const earnedDate = row[CONFIG.EMPLOYEE_DATE_COL - 1];
        const earnedAmount = row[CONFIG.EMPLOYEE_COMP_OFF_EARNED_COL - 1];
        const status = row[CONFIG.EMPLOYEE_COMP_OFF_STATUS_COL - 1];

        if (earnedDate instanceof Date && earnedAmount > 0 && status !== 'Consumed' && status !== 'Expired') {
          const expiryDate = new Date(earnedDate.getTime());
          expiryDate.setDate(expiryDate.getDate() + CONFIG.COMP_OFF_EXPIRY_DAYS);
          if (expiryDate.getMonth() === currentMonth && expiryDate.getFullYear() === currentYear) {
            expiringThisMonth += earnedAmount;
          }
        }
      });
    }

    // 3. Create Charts
    const casualChart = createPieChart_('Casual Leave', casualUtilized, casualAvailable);
    const sickChart = createPieChart_('Sick Leave', sickUtilized, sickAvailable);
    const compOffChart = createThreeSlicePieChart_('Comp Off Summary', compOffUtilized, compOffAvailable, compOffExpired);

    // 4. Build and Send Email
    // --- UPDATED SUBJECT LINE ---
    const monthName = new Date().toLocaleString('default', { month: 'long' });
    const subject = `${employeeName}, your monthly summary for ${monthName} is here!`;
    
    const htmlBody = `
      <html><body>
        <p>Hi ${employeeName},</p>
        <p>Another fun month is here! You can find your leave balances below -</p>
        <table border="1" cellpadding="5" style="border-collapse: collapse;">
          <tr style="background-color:#f2f2f2;">
            <th>Casual Leaves Available</th>
            <th>Sick Leaves Available</th>
            <th>Comp Offs Available</th>
            <th>Comp Offs Expiring This Month</th>
          </tr>
          <tr>
            <td style="text-align:center;">${casualAvailable}</td>
            <td style="text-align:center;">${sickAvailable}</td>
            <td style="text-align:center;">${compOffAvailable}</td>
            <td style="text-align:center;">${expiringThisMonth}</td>
          </tr>
        </table>
        <p>Your current leave utilization:</p>
        <img src="cid:casualChart">
        <img src="cid:sickChart">
        <img src="cid:compOffChart">
        <p>We're here to support your rest and recovery - so work with your supervisor to take time off!</p>
        <p>For the team, By the team</p>
      </body></html>`;

    MailApp.sendEmail({
      to: employeeEmail,
      subject: subject,
      htmlBody: htmlBody,
      name: "For the team, By the team",
      inlineImages: {
        casualChart: casualChart,
        sickChart: sickChart,
        compOffChart: compOffChart
      }
    });
    Logger.log(`Sent monthly summary to ${employeeName}.`);

  } catch (e) {
    Logger.log(`Failed to send summary for ${employeeName}. Error: ${e.message}`);
  }
}

/**
 * Helper function to create a standardized two-slice pie chart.
 * @private
 */
function createPieChart_(title, utilized, available) {
  if ((utilized + available) === 0) {
     const data = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days").addRow(["No Leave Data", 1]).build();
    return Charts.newPieChart().setDataTable(data).setColors(['#cccccc']).setTitle(title).setDimensions(400, 300).build().getAs('image/png');
  }
  const data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Status")
      .addColumn(Charts.ColumnType.NUMBER, "Days")
      .addRow([`Utilized (${utilized})`, utilized])
      .addRow([`Available (${available})`, available])
      .build();
  return Charts.newPieChart().setDataTable(data)
      .setColors(['#0E3436', '#2fafb6']) // UPDATED COLORS
      .setTitle(title).setDimensions(400, 300).build().getAs('image/png');
}

/**
 * NEW: Helper function to create a standardized three-slice pie chart for Comp Offs.
 * @private
 */
function createThreeSlicePieChart_(title, utilized, available, expired) {
  if ((utilized + available + expired) === 0) {
    const data = Charts.newDataTable().addColumn(Charts.ColumnType.STRING, "Status").addColumn(Charts.ColumnType.NUMBER, "Days").addRow(["No Comp Off Data", 1]).build();
    return Charts.newPieChart().setDataTable(data).setColors(['#cccccc']).setTitle(title).setDimensions(400, 300).build().getAs('image/png');
  }
  const data = Charts.newDataTable()
      .addColumn(Charts.ColumnType.STRING, "Status")
      .addColumn(Charts.ColumnType.NUMBER, "Days")
      .addRow([`Utilized (${utilized})`, utilized])
      .addRow([`Available (${available})`, available])
      .addRow([`Expired (${expired})`, expired])
      .build();
  return Charts.newPieChart().setDataTable(data)
      .setColors(['#0E3436', '#2FAFB6', '#C0E7E9']) // UPDATED COLORS
      .setTitle(title).setDimensions(400, 300).build().getAs('image/png');
}
