/**
 * CONFIGURATION
 */
const FORM_ID = '1zLHiGyleU8pEHT6pLFy5sE7yprNIJXHpmdCERy0oDQQ';
const FORM_ID_LEADERSHIP = ''; // To be filled after running setup
const SPREADSHEET_ID = '1370PuPE1cxzt8vJgUpcw69AU5KPk04WBU6oh5xWUBKk';
const SEND_SHEET_NAME = 'Send_Form';
const SEND_SHEET_LEADERSHIP = 'Send_Googlers';
const RESPONSES_SHEET_NAME = 'Partner Responses';
const RESPONSES_SHEET_LEADERSHIP = 'Form Responses 3';

/**
 * Menu for easy access
 */
function onOpen() {
  // Temporary Fix: Ensure the correct Leadership Form ID is set as per user report
  PropertiesService.getScriptProperties().setProperty('FORM_ID_LEADERSHIP', '1sjhxyGLCZoZLSveNg0CQUbvf2n_QLS8dSnRBhsFIFvs');

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Feedback System')
    .addSubMenu(ui.createMenu('Partners')
      .addItem('Send Initial Emails', 'sendEmails')
      .addItem('Send Reminders', 'sendReminders')
      .addItem('Check Responses', 'checkResponses'))
    .addSubMenu(ui.createMenu('Leadership (Internal)')
      .addItem('Setup Internal System (Create Form)', 'setupLeadershipSystem')
      .addItem('Link Existing Leadership Form ID', 'linkExistingLeadershipForm')
      .addItem('Send Initial Emails', 'sendLeadershipEmails')
      .addItem('Send Reminders', 'sendLeadershipReminders')
      .addItem('Check Responses', 'checkLeadershipResponses'))
    .addSeparator()
    .addItem('Rename Existing Sheets to New Names', 'renameResponseSheets')
    .addItem('Check System Status (Audit)', 'checkSystemStatus')
    .addToUi();
}

/**
 * Sends initial emails to partners.
 */
function sendEmails() {
  processGeneralEmails(SEND_SHEET_NAME, FORM_ID, 'email', false, true);
}

/**
 * Sends reminders to partners.
 */
function sendReminders() {
  processGeneralEmails(SEND_SHEET_NAME, FORM_ID, 'reminder', true, true);
}

/**
 * Checks responses for partners.
 */
function checkResponses() {
  processGeneralResponses(SEND_SHEET_NAME, RESPONSES_SHEET_NAME, 4); // Column D (4) has the email in the screenshot
}

/**
 * Sends initial emails to Leadership (English).
 * Now includes the Infographic as a robust inline image.
 */
function sendLeadershipEmails() {
  const infographicId = '1h1VBpbmY2iXH7gSiK9qspu7nlbLQwiaJ';
  processGeneralEmails(SEND_SHEET_LEADERSHIP, null, 'email_en', false, false, infographicId);
}

/**
 * Sends reminders to Leadership (English).
 */
function sendLeadershipReminders() {
  processGeneralEmails(SEND_SHEET_LEADERSHIP, null, 'reminder_en', true, false);
}

/**
 * Checks responses for Leadership.
 */
function checkLeadershipResponses() {
  processGeneralResponses(SEND_SHEET_LEADERSHIP, RESPONSES_SHEET_LEADERSHIP, 2); // Leadership form uses Column B (2) for auto-collected email
}

/**
 * Generic Processor for Emails
 */
function processGeneralEmails(sheetName, formId, baseTemplate, isReminder, useLanguageRouting, inlineImageId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Error: Sheet '" + sheetName + "' no existe.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay destinatarios en la hoja: " + sheetName);
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const data = range.getValues();

  let currentFormId = formId;
  if (!currentFormId) {
    currentFormId = PropertiesService.getScriptProperties().getProperty('FORM_ID_LEADERSHIP');
  }

  if (!currentFormId) {
    Logger.log("Error: Form ID no configurado para Leadership.");
    return;
  }

  const form = FormApp.openById(currentFormId);
  const formUrl = form.getPublishedUrl();

  const subjects = {
    'email_es': 'Tu opini\u00f3n es clave para el 2026 - Google Cloud Readiness',
    'reminder_es': 'Recordatorio: Tu visi\u00f3n es clave para el 2026',
    'email_pt': 'Sua opini\u00e3o \u00e9 fundamental para 2026 - Google Cloud Readiness',
    'reminder_pt': 'Lembrete: Sua vis\u00e3o 2026',
    'email_en': 'Feedback for 2026 - Google Cloud Readiness',
    'reminder_en': 'Reminder: Feedback for 2026'
  };

  // Prepare inline image blob if ID exists
  let inlineImages = {};
  if (inlineImageId) {
    try {
      const blob = DriveApp.getFileById(inlineImageId).getBlob();
      inlineImages['infographic'] = blob;
    } catch (e) {
      Logger.log("Warning: Could not fetch inline image: " + e.message);
    }
  }

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
      let langSuffix = '_es';
      if (useLanguageRouting && email.toString().toLowerCase().endsWith('.br')) {
        langSuffix = '_pt';
      } else if (!useLanguageRouting) {
        langSuffix = ''; // Leadership already provides 'email_en'
      }

      const finalTemplate = useLanguageRouting ? (baseTemplate + langSuffix) : baseTemplate;
      const subject = subjects[finalTemplate] || 'Feedback 2026';

      try {
        const htmlTemplate = HtmlService.createTemplateFromFile(finalTemplate);
        htmlTemplate.formUrl = formUrl;
        const htmlBody = htmlTemplate.evaluate().getContent();

        GmailApp.sendEmail(email, subject, '', {
          htmlBody: htmlBody,
          name: 'Google Cloud Readiness Team',
          inlineImages: inlineImages
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
    const updateRange = sheet.getRange(2, 4, statusUpdates.length, 1);
    updateRange.setValues(statusUpdates);
    SpreadsheetApp.flush(); // Force changes to appear in the UI immediately
    Logger.log(`Updated status column (D) for ${statusUpdates.length} rows in ${sheetName}.`);
  }

  Logger.log(`--- Email Summary for ${sheetName} ---`);
  Logger.log(`Total rows processed: ${data.length}`);
  Logger.log(`Emails successfully sent: ${count}`);
  Logger.log(`--------------------------------------`);
}

/**
 * Extracts LDAP (the part before @) for robust internal matching.
 */
function getLdap(email) {
  if (!email || !email.toString().includes('@')) return email ? email.toString().toLowerCase().trim() : '';
  return email.toString().split('@')[0].toLowerCase().trim();
}

/**
 * Generic Processor for Responses
 */
function processGeneralResponses(sendSheetName, respSheetName, emailColIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sendSheet = ss.getSheetByName(sendSheetName);
  let responsesSheet = ss.getSheetByName(respSheetName);

  // Fallback check: if the standard name isn't found, try common defaults
  if (!responsesSheet) {
    const fallbacks = ['Form Responses 1', 'Form Responses 2', 'Form Responses 3', 'Form_Responses2'];
    for (const fb of fallbacks) {
      responsesSheet = ss.getSheetByName(fb);
      if (responsesSheet) {
        Logger.log(`Note: Sheet '${respSheetName}' not found. Using fallback: '${fb}'`);
        break;
      }
    }
  }

  if (!sendSheet || !responsesSheet) {
    Logger.log("Error: Missing primary or fallback sheets for: " + respSheetName);
    return;
  }

  const respLastRow = responsesSheet.getLastRow();
  const respondingEmails = new Set();
  const respondingLdaps = new Set();

  if (respLastRow >= 2) {
    const respData = responsesSheet.getRange(2, emailColIndex, respLastRow - 1, 1).getValues();
    for (let i = 0; i < respData.length; i++) {
      const rawEmail = respData[i][0];
      if (rawEmail) {
        const cleanEmail = rawEmail.toString().trim().toLowerCase();
        respondingEmails.add(cleanEmail);
        respondingLdaps.add(getLdap(cleanEmail));
      }
    }
  }

  const sendLastRow = sendSheet.getLastRow();
  if (sendLastRow < 2) return;

  const emails = sendSheet.getRange(2, 1, sendLastRow - 1, 1).getValues();
  const statusRange = sendSheet.getRange(2, 5, sendLastRow - 1, 1);
  const currentStatuses = statusRange.getValues();
  const newStatuses = [];
  let updatedCount = 0;

  for (let i = 0; i < emails.length; i++) {
    const rawEmail = emails[i][0];
    const email = rawEmail ? rawEmail.toString().trim().toLowerCase() : "";
    const ldap = getLdap(email);

    // Internal matching is now more robust: Check Full Email OR just the LDAP
    const hasResponded = respondingEmails.has(email) || (ldap && respondingLdaps.has(ldap));

    if (hasResponded) {
      if (currentStatuses[i][0] !== 'Responded') {
        newStatuses.push(['Responded']);
        updatedCount++;
      } else {
        newStatuses.push([currentStatuses[i][0]]);
      }
    } else {
      newStatuses.push([currentStatuses[i][0]]);
    }
  }

  if (newStatuses.length > 0) {
    statusRange.setValues(newStatuses);
  }

  Logger.log(`--- Diagnostics for ${sendSheetName} ---`);
  Logger.log(`Sheet Used: ${responsesSheet.getName()}`);
  Logger.log(`Responding Emails found: ${Array.from(respondingEmails).join(", ")}`);
  Logger.log(`Total rows processed: ${emails.length}`);
  Logger.log(`New responses identified: ${updatedCount}`);
  Logger.log(`------------------------------------------`);
}


/**
 * SETUP LEADERSHIP SYSTEM
 * Creates a new form and the 'Send_Googlers' sheet.
 * Now includes a check to prevent accidental duplicates.
 */
function setupLeadershipSystem() {
  const SPREADSHEET_ID = '1370PuPE1cxzt8vJgUpcw69AU5KPk04WBU6oh5xWUBKk';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Check if we already have a Leadership Form ID
  const existingId = PropertiesService.getScriptProperties().getProperty('FORM_ID_LEADERSHIP');
  if (existingId) {
    Logger.log("⚠️ Ya existe un Formulario de Liderazgo vinculado (ID: " + existingId + ").");
    Logger.log("Si deseas crear uno nuevo, borra primero el ID de las 'Script Properties' o usa 'Link Existing Leadership Form ID' para cambiarlo.");
    return;
  }

  // 2. Create 'Send_Googlers' sheet if it doesn't exist
  let googlersSheet = ss.getSheetByName('Send_Googlers');
  if (!googlersSheet) {
    googlersSheet = ss.insertSheet('Send_Googlers');
    googlersSheet.getRange('A1:E1').setValues([['Email', 'Name', 'Comments', 'Status', 'Responded Status']]);
    googlersSheet.getRange('A1:E1').setFontWeight('bold').setBackground('#e8eaed');
    googlersSheet.setFrozenRows(1);
    Logger.log("Created 'Send_Googlers' sheet.");
  }

  // 3. Create the Leadership Form (English)
  const newForm = FormApp.create('Google Cloud Leadership Feedback 2026')
    .setTitle('Excellence in 2026: Leadership Vision')
    .setDescription('Feedback from Google Managers and Leaders regarding readiness and strategy for 2026.')
    .setCollectEmail(true)
    .setRequireLogin(true)
    .setAllowResponseEdits(true);

  newForm.addParagraphTextItem()
    .setTitle('Where should the Readiness team focus in 2026? Share your ideas, feedback, and creative vision.')
    .setRequired(true);

  newForm.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);

  const formUrl = newForm.getPublishedUrl();
  const formId = newForm.getId();

  // Persist the ID automatically
  PropertiesService.getScriptProperties().setProperty('FORM_ID_LEADERSHIP', formId);

  Logger.log("✅ Success: New Leadership Form Created: " + formUrl);
  Logger.log("Leadership System Initialized! Note the new tab created in this spreadsheet.");
}

/**
 * Identify and Audit: Shows which forms and sheets are currently active.
 */
function checkSystemStatus() {
  const props = PropertiesService.getScriptProperties();
  const partnerId = FORM_ID;
  const leadId = props.getProperty('FORM_ID_LEADERSHIP') || "NOT SET";

  Logger.log("--- SYSTEM AUDIT ---");
  Logger.log("1. Partner Form ID: " + partnerId);
  Logger.log("2. Leadership Form ID: " + leadId);

  try {
    const partnerForm = FormApp.openById(partnerId);
    Logger.log("✅ Partner Form is LIVE: " + partnerForm.getPublishedUrl());
  } catch (e) { Logger.log("❌ Partner Form Error: " + e.message); }

  try {
    const leadForm = FormApp.openById(leadId);
    Logger.log("✅ Leadership Form is LIVE: " + leadForm.getPublishedUrl());
  } catch (e) { Logger.log("❌ Leadership Form Error: " + e.message); }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());
  Logger.log("3. Current Spreadsheet Tabs: " + sheets.join(", "));
  Logger.log("--------------------");
}
/**
 * Manually link an existing form ID to the leadership system.
 */
function linkExistingLeadershipForm() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Link Existing Leadership Form', 'Please paste the Google Form ID for Leadership:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const formId = response.getResponseText().trim();
    if (formId) {
      PropertiesService.getScriptProperties().setProperty('FORM_ID_LEADERSHIP', formId);
      Logger.log("Linked Leadership Form ID: " + formId);
      ui.alert("Form ID successfully linked!");
    }
  }
}


/**
 * Utility to rename existing sheets if they follow the default naming pattern.
 */
function renameResponseSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Looking for common default names Google creates
  const renameMap = {
    'Form Responses 1': 'Partner Responses',
    'Form Responses 2': 'Googlers Responses',
    'Form Responses 3': 'Googlers Responses',
    'Form_Responses2': 'Googlers Responses'
  };

  let renamedCount = 0;
  for (let oldName in renameMap) {
    let sheet = ss.getSheetByName(oldName);
    let targetName = renameMap[oldName];

    if (sheet) {
      // Check if the target name already exists to avoid errors
      let existingTarget = ss.getSheetByName(targetName);
      if (!existingTarget) {
        sheet.setName(targetName);
        renamedCount++;
        Logger.log(`Renamed '${oldName}' to '${targetName}'`);
      } else {
        Logger.log(`Target name '${targetName}' already exists. Skipping '${oldName}'.`);
      }
    }
  }

  if (renamedCount > 0) {
    Logger.log(`Successfully renamed ${renamedCount} sheet(s).`);
  } else {
    Logger.log("No default Form Response sheets were found to rename. Your sheets might already be named correctly ('Partner Responses' and 'Googlers Responses').");
  }
}