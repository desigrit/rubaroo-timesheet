/**
 * This file contains general utility functions that support the main logic
 * but don't belong to a specific core feature.
 */

/**
 * Helper function to send an email notification when a new timesheet is created.
 * @param {string} employeeName The name of the new employee.
 * @param {string} employeeEmail The email of the new employee.
 * @param {string} supervisorEmailsString A comma-separated string of supervisor emails.
 * @param {string} timesheetUrl The URL of the newly created timesheet.
 * @private
 */

function isValidEmail_(email) {
  if (!email || typeof email !== 'string') {
    return false;
  }
  // A common regex for email validation (basic, covers most common cases)
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function sendNewTimesheetNotification_(employeeName, employeeEmail, supervisorEmailsString, timesheetUrl) {
  Logger.log(`[Email Debug] --- Starting sendNewTimesheetNotification_ for ${employeeName} ---`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetName = ss.getName(); //RAUNAK
  
  let toRecipient = isValidEmail_(employeeEmail) ? employeeEmail : '';

  if (!toRecipient) {
    Logger.log('[Email Debug] No valid employee email found. Email will NOT be sent.');
    SpreadsheetApp.getUi().alert(`No valid email found for ${employeeName}. Timesheet notification not sent.`);
    return;
  }

  // Build the CC list
  let ccList = [];
  if (supervisorEmailsString) {
    const supervisorEmails = supervisorEmailsString.split(',').map(e => e.trim()).filter(e => isValidEmail_(e));
    ccList = ccList.concat(supervisorEmails);
  }
  if (CONFIG.NOTIFICATION_CC_EMAILS && CONFIG.NOTIFICATION_CC_EMAILS.length > 0) {
    const staticEmails = CONFIG.NOTIFICATION_CC_EMAILS.filter(e => isValidEmail_(e));
    ccList = ccList.concat(staticEmails);
  }
  const uniqueCcList = [...new Set(ccList)];
  const ccString = uniqueCcList.join(',');

  const subject = `Your New Timesheet is Ready: ${employeeName} - Rubaroo Time Sheet Tracker`;
  
  // --- UPDATED EMAIL BODY ---
  const body = `Hi ${employeeName},\n\n` +
             `A huge welcome to the Rubaroo team! We're so excited to have you join us. ♥\n\n` +
             `As you settle into your role, we want to introduce this Timesheet to you. It might sound like just another administrative tool, but for us, it's a way to reinforce the culture of balance that our team has worked so hard to build!\n\n` +
             `Think of it as a way for us to all keep a lookout on everything - making sure you get credit for all your work and, more importantly, that you're taking the much-needed breaks. It's the key to making sure you can take all those Comp Offs you'll be earning for your extra efforts - we already know!\n\n` +
             `You can find the timesheet ${timesheetUrl}. Don't hesitate to ask if you have any questions at all.\n\n` +
             `We're so glad you're here!\n\n` +
             `Warmly,\nFor the team, By the team`;

  const options = {
    htmlBody: `<p>Hi ${employeeName},</p>` +
              `<p>A huge welcome to the Rubaroo team! We're so excited to have you join us. ♥</p>` +
              `<p>As you settle into your role, we want to introduce this Timesheet to you. It might sound like just another administrative tool, but for us, it's a way to reinforce the culture of balance that our team has worked so hard to build!</p>` +
              `<p>Think of it as a way for us to all keep a lookout on everything - making sure you get credit for all your work and, more importantly, that you're taking the much-needed breaks. It's the key to making sure you can take all those Comp Offs you'll be earning for your extra efforts - we already know!</p>` +
              `<p>You can find the timesheet <a href="${timesheetUrl}">here</a>. Don't hesitate to ask if you have any questions at all.</p>` +
              `<p>We're so glad you're here!</p>` +
              `<p>Warmly,<br>For the team, By the team</p>`,
    name: "For the team, By the team"
  };
  // --- END UPDATED BODY ---

  if (ccString) {
    options.cc = ccString;
    Logger.log(`[Email Debug] Sending with CC to: ${ccString}`);
  }
  
  try {
    MailApp.sendEmail(toRecipient, subject, body, options);
    Logger.log(`[Email Debug] Email sent successfully to: ${toRecipient}.`);
  } catch (error) {
    Logger.log(`[Email Debug] CRITICAL ERROR: MailApp.sendEmail failed. Error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Failed to send timesheet creation email. Error: ${error.message}.`);
  }
}
