/**
 * CONFIGURATION
 */
const FORM_ID = '1zLHiGyleU8pEHT6pLFy5sE7yprNIJXHpmdCERy0oDQQ';
const SPREADSHEET_ID = '1370PuPE1cxzt8vJgUpcw69AU5KPk04WBU6oh5xWUBKk';
const SEND_SHEET_NAME = 'Send_Form';
const RESPONSES_SHEET_NAME = 'Form Responses 1';

/**
 * Menu for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Feedback System')
    .addItem('Send Initial Emails (Batch)', 'sendEmails')
    .addItem('Send Reminders (Batch)', 'sendReminders')
    .addItem('Check Responses & Update Status', 'checkResponses')
    .addToUi();
}

/**
 * Sends initial emails to partners in 'Send_Form' who haven't received it yet.
 */
function sendEmails() {
  processEmails('email', 'Tu opini\u00f3n es clave para el 2026 - Google Cloud Readiness', false);
}

/**
 * Sends reminders to those who were shared the form but haven't responded yet.
 */
function sendReminders() {
  processEmails('reminder', 'Recordatorio: Tu visi\u00f3n es clave para el 2026', true);
}

/**
 * Core logic to process and send emails or reminders.
 */
function processEmails(templateName, subject, isReminder) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SEND_SHEET_NAME);
  if (!sheet) {
    Logger.log("Sheet not found: " + SEND_SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Read A (Email), D (Shared Status), E (Responded Status)
  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const data = range.getValues();

  const form = FormApp.openById(FORM_ID);
  const formUrl = form.getPublishedUrl();
  const htmlTemplate = HtmlService.createTemplateFromFile(templateName);
  htmlTemplate.formUrl = formUrl;
  const htmlBody = htmlTemplate.evaluate().getContent();

  let count = 0;
  const statusUpdates = [];

  for (let i = 0; i < data.length; i++) {
    const email = data[i][0];
    const sharedStatus = data[i][3]; // Column D
    const respondedStatus = data[i][4]; // Column E

    let shouldSend = false;
    if (isReminder) {
      // Send reminder if shared but not responded
      shouldSend = (sharedStatus === 'Shared' && respondedStatus !== 'Responded');
    } else {
      // Send initial if not shared yet
      shouldSend = (!sharedStatus || sharedStatus === '');
    }

    if (shouldSend && email && email.toString().includes('@')) {
      try {
        GmailApp.sendEmail(email, subject, '', {
          htmlBody: htmlBody,
          name: 'Google Cloud Readiness Team'
        });
        statusUpdates.push(['Shared']);
        count++;
      } catch (e) {
        Logger.log(`Failed to send to ${email}: ${e.message}`);
        statusUpdates.push([sharedStatus || 'Error: ' + e.message]);
      }
    } else {
      statusUpdates.push([sharedStatus]);
    }
  }

  // Update Column D
  if (statusUpdates.length > 0) {
    sheet.getRange(2, 4, statusUpdates.length, 1).setValues(statusUpdates);
  }

  Logger.log(`${isReminder ? 'Reminders' : 'Initial emails'} sent: ${count}`);
}

/**
 * Checks 'Form Responses 1' and updates 'Send_Form' Column E (Responded)
 */
function checkResponses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sendSheet = ss.getSheetByName(SEND_SHEET_NAME);
  const responsesSheet = ss.getSheetByName(RESPONSES_SHEET_NAME);

  if (!sendSheet || !responsesSheet) {
    Logger.log("Missing one or more sheets.");
    return;
  }

  // 1. Get all responses emails (Column D usually, but let's verify)
  // Warning: The user said "comparing with the email addes of column D in the sheet 'Form Responses 1'"
  // Let's assume Column D in Responses sheet has the email.
  const respLastRow = responsesSheet.getLastRow();
  const respondingEmails = new Set();

  if (respLastRow >= 2) {
    // Read Column D (index 4 in 1-based, index 3 in 0-based)
    const respData = responsesSheet.getRange(2, 4, respLastRow - 1, 1).getValues();
    for (let i = 0; i < respData.length; i++) {
      const email = respData[i][0];
      if (email) respondingEmails.add(email.toString().trim().toLowerCase());
    }
  }

  // 2. Update Send_Form Column E (Status: Responded?)
  // Actually user asked: "in column E i wnat to track if the form was actually responded"
  const sendLastRow = sendSheet.getLastRow();
  if (sendLastRow < 2) return;

  const emailRange = sendSheet.getRange(2, 1, sendLastRow - 1, 1);
  const emails = emailRange.getValues();

  const statusRange = sendSheet.getRange(2, 5, sendLastRow - 1, 1); // Column E
  const currentStatuses = statusRange.getValues();

  const newStatuses = [];

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i][0].toString().trim().toLowerCase();
    const currentStatus = currentStatuses[i][0];

    if (respondingEmails.has(email)) {
      newStatuses.push(['Responded']);
    } else {
      newStatuses.push([currentStatus]); // Keep existing or empty
    }
  }

  // Batch update Column E
  if (newStatuses.length > 0) {
    statusRange.setValues(newStatuses);
  }

  Logger.log("Response check complete.");
}