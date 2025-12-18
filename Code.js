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
    .addItem('Send Emails (Batch)', 'sendEmails')
    .addItem('Check Responses & Update Status', 'checkResponses')
    .addToUi();
}

/**
 * Sends emails to partners in 'Send_Form' who haven't received it yet.
 * Checks Column D (Status) to avoid duplicates.
 */
function sendEmails() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SEND_SHEET_NAME);
  if (!sheet) {
    Logger.log("Sheet not found: " + SEND_SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // No data

  // Columns: A=Email, B=..., C=..., D=Status
  // We read A:D
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
  const data = dataRange.getValues();

  // Load HTML template
  const form = FormApp.openById(FORM_ID);
  const formUrl = form.getPublishedUrl();
  const htmlTemplate = HtmlService.createTemplateFromFile('email');
  htmlTemplate.formUrl = formUrl;
  const htmlBody = htmlTemplate.evaluate().getContent();

  let emailsSent = 0;

  // We will update status in batch
  const statusUpdates = [];

  for (let i = 0; i < data.length; i++) {
    const email = data[i][0];
    const currentStatus = data[i][3]; // Column D (index 3)

    if (currentStatus === 'Shared' || currentStatus === 'Responded') {
      statusUpdates.push([currentStatus]); // Keep existing status
      continue;
    }

    if (email && email.toString().includes('@')) {
      try {
        GmailApp.sendEmail(email, 'Tu opinión es clave para el 2026 - Google Cloud Readiness', '', {
          htmlBody: htmlBody,
          name: 'Google Cloud Readiness Team'
        });
        statusUpdates.push(['Shared']);
        emailsSent++;
      } catch (e) {
        Logger.log(`Failed to send to ${email}: ${e.message}`);
        statusUpdates.push(['Error: ' + e.message]);
      }
    } else {
      statusUpdates.push([currentStatus || 'Invalid Email']);
    }
  }

  // Update Column D
  if (statusUpdates.length > 0) {
    sheet.getRange(2, 4, statusUpdates.length, 1).setValues(statusUpdates);
  }

  Logger.log(`Sent ${emailsSent} emails.`);
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