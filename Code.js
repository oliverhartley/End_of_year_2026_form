/**
 * CONFIGURATION
 */
const FORM_ID = '1zLHiGyleU8pEHT6pLFy5sE7yprNIJXHpmdCERy0oDQQ';
const FORM_ID_LEADERSHIP = ''; // To be filled after running setup
const SPREADSHEET_ID = '1370PuPE1cxzt8vJgUpcw69AU5KPk04WBU6oh5xWUBKk';
const SEND_SHEET_NAME = 'Send_Form';
const SEND_SHEET_LEADERSHIP = 'Send_Googlers';
const RESPONSES_SHEET_NAME = 'Partner Responses';
const RESPONSES_SHEET_LEADERSHIP = 'Googlers Responses';

/**
 * Menu for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Feedback System')
    .addSubMenu(ui.createMenu('Partners')
      .addItem('Send Initial Emails', 'sendEmails')
      .addItem('Send Reminders', 'sendReminders')
      .addItem('Check Responses', 'checkResponses'))
    .addSubMenu(ui.createMenu('Leadership (Internal)')
      .addItem('Setup Internal System (Create Form)', 'setupLeadershipSystem')
      .addItem('Send Initial Emails', 'sendLeadershipEmails')
      .addItem('Send Reminders', 'sendLeadershipReminders')
      .addItem('Check Responses', 'checkLeadershipResponses'))
    .addSeparator()
    .addItem('Rename Existing Sheets to New Names', 'renameResponseSheets')
    .addToUi();
}

/**
 * Sends initial emails to partners.
 */
function sendEmails() {
  processGeneralEmails(SEND_SHEET_NAME, FORM_ID, 'email', 'Tu opini\u00f3n es clave para el 2026 - Google Cloud Readiness', false);
}

/**
 * Sends reminders to partners.
 */
function sendReminders() {
  processGeneralEmails(SEND_SHEET_NAME, FORM_ID, 'reminder', 'Recordatorio: Tu visi\u00f3n es clave para el 2026', true);
}

/**
 * Checks responses for partners.
 */
function checkResponses() {
  processGeneralResponses(SEND_SHEET_NAME, RESPONSES_SHEET_NAME, 4); // Column D is #4
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
 * Leadership Wrapper Functions
 */
function sendLeadershipEmails() {
  processGeneralEmails(SEND_SHEET_LEADERSHIP, FORM_ID_LEADERSHIP, 'email', 'Feedback para el 2026: Una visi\u00f3n de liderazgo', false);
}

function sendLeadershipReminders() {
  processGeneralEmails(SEND_SHEET_LEADERSHIP, FORM_ID_LEADERSHIP, 'reminder', 'Recordatorio: Feedback para el 2026', true);
}

function checkLeadershipResponses() {
  processGeneralResponses(SEND_SHEET_LEADERSHIP, RESPONSES_SHEET_LEADERSHIP, 2); // Typically Column B for email in linked forms
}

/**
 * Generic Processor for Emails
 */
function processGeneralEmails(sheetName, formId, templateName, subject, isReminder) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    Logger.log("Error: Sheet '" + sheetName + "' no existe. Ejecuta 'Setup Internal System' primero.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const data = range.getValues();

  let currentFormId = formId;
  if (!currentFormId) {
    // Fallback to ScriptProperties
    currentFormId = PropertiesService.getScriptProperties().getProperty('FORM_ID_LEADERSHIP');
  }

  if (!currentFormId) {
    Logger.log("Error: FORM_ID_LEADERSHIP no configurado en el script.");
    Logger.log("Sugerencia: Ejecuta 'Setup Internal System' primero o pega el ID en el script.");
    return;
  }

  const form = FormApp.openById(currentFormId);
  let currentFormUrl = form.getPublishedUrl();

  const htmlTemplate = HtmlService.createTemplateFromFile(templateName);
  htmlTemplate.formUrl = currentFormUrl;
  const htmlBody = htmlTemplate.evaluate().getContent();

  let count = 0;
  const statusUpdates = [];

  for (let i = 0; i < data.length; i++) {
    const email = data[i][0];
    const sharedStatus = data[i][3];
    const respondedStatus = data[i][4];

    let shouldSend = false;
    if (isReminder) {
      shouldSend = (sharedStatus === 'Shared' && respondedStatus !== 'Responded');
    } else {
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

  if (statusUpdates.length > 0) {
    sheet.getRange(2, 4, statusUpdates.length, 1).setValues(statusUpdates);
  }

  Logger.log(`${isReminder ? 'Reminders' : 'Initial emails'} sent to ${sheetName}: ${count}`);
}

/**
 * Generic Processor for Responses
 */
function processGeneralResponses(sendSheetName, respSheetName, emailColIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sendSheet = ss.getSheetByName(sendSheetName);
  const responsesSheet = ss.getSheetByName(respSheetName);

  if (!sendSheet || !responsesSheet) {
    Logger.log("Missing sheets: " + sendSheetName + " or " + respSheetName);
    return;
  }

  const respLastRow = responsesSheet.getLastRow();
  const respondingEmails = new Set();

  if (respLastRow >= 2) {
    const respData = responsesSheet.getRange(2, emailColIndex, respLastRow - 1, 1).getValues();
    for (let i = 0; i < respData.length; i++) {
      const email = respData[i][0];
      if (email) respondingEmails.add(email.toString().trim().toLowerCase());
    }
  }

  const sendLastRow = sendSheet.getLastRow();
  if (sendLastRow < 2) return;

  const emails = sendSheet.getRange(2, 1, sendLastRow - 1, 1).getValues();
  const statusRange = sendSheet.getRange(2, 5, sendLastRow - 1, 1);
  const currentStatuses = statusRange.getValues();
  const newStatuses = [];

  for (let i = 0; i < emails.length; i++) {
    const email = emails[i][0].toString().trim().toLowerCase();
    if (respondingEmails.has(email)) {
      newStatuses.push(['Responded']);
    } else {
      newStatuses.push([currentStatuses[i][0]]);
    }
  }

  if (newStatuses.length > 0) statusRange.setValues(newStatuses);
  Logger.log(`Response check complete for ${sendSheetName}`);
}


/**
 * SETUP LEADERSHIP SYSTEM
 * Creates a new form and the 'Send_Googlers' sheet.
 */
function setupLeadershipSystem() {
  const SPREADSHEET_ID = '1370PuPE1cxzt8vJgUpcw69AU5KPk04WBU6oh5xWUBKk';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Create 'Send_Googlers' sheet if it doesn't exist
  let googlersSheet = ss.getSheetByName('Send_Googlers');
  if (!googlersSheet) {
    googlersSheet = ss.insertSheet('Send_Googlers');
    // Align with the original sheet structure but for internals
    googlersSheet.getRange('A1:E1').setValues([['Email', 'Name', 'Comments', 'Status', 'Responded Status']]);
    googlersSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e8eaed');
    googlersSheet.setFrozenRows(1);
    Logger.log("Created 'Send_Googlers' sheet.");
  }

  // 2. Create the Leadership Form
  const newForm = FormApp.create('Google Cloud Leadership Feedback 2026')
    .setTitle('Excellence in 2026: Leadership Vision')
    .setDescription('Feedback from Google Managers and Leaders regarding readiness and strategy for 2026.')
    .setCollectEmail(true)
    .setRequireLogin(true) // For internals
    .setAllowResponseEdits(true);

  // Add the open creativity question
  newForm.addParagraphTextItem()
    .setTitle('\u00bfEn qu\u00e9 deber\u00eda enfocarse el equipo de Readiness en 2026? Comparte tus ideas, feedback y visi\u00f3n creativa.')
    .setRequired(true);

  // Link to spreadsheet
  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);

  const formUrl = newForm.getPublishedUrl();
  const formId = newForm.getId();

  Logger.log("New Form Created: " + formUrl);
  Logger.log("Form ID (save this): " + formId);

  // Persist the ID automatically
  PropertiesService.getScriptProperties().setProperty('FORM_ID_LEADERSHIP', formId);

  Logger.log("Leadership System Initialized!\n\nNew Form ID: " + formId + "\n\nPlease note the secondary 'Form Responses' tab that was just created automatically.");
}

/**
 * Utility to rename existing sheets if they follow the default naming pattern.
 */
function renameResponseSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const renameMap = {
    'Form Responses 1': 'Partner Responses',
    'Form Responses 2': 'Googlers Responses'
  };

  let renamedCount = 0;
  for (let oldName in renameMap) {
    let sheet = ss.getSheetByName(oldName);
    if (sheet) {
      sheet.setName(renameMap[oldName]);
      renamedCount++;
    }
  }

  if (renamedCount > 0) {
    Logger.log(`Se han renombrado ${renamedCount} hoja(s).`);
  } else {
    Logger.log("No se encontraron hojas con los nombres predeterminados (Form Responses 1/2). Es posible que ya las hayas renombrado.");
  }
}