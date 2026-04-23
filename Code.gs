/**
 * Configuration for sheet persistence.
 * Set SHEET_ID to target a specific spreadsheet when deployed as a standalone web app.
 */
var SHEET_CONFIG = {
  SHEET_ID: '',
  SHEET_NAME: 'Submissions',
  HEADERS: ['Timestamp', 'Name', 'Email', 'Message', 'AttachmentId'],
};

/**
 * Serves the form UI.
 * @return {HtmlOutput}
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

/**
 * Accepts form data and optional file attachment.
 *
 * @param {{name: string, email: string, message: string, attachment: {name: string, mimeType: string, data: string}|null}} payload
 * @return {{ok: boolean, message: string, attachmentId: string|null, rowNumber: number}}
 */
function submitForm(payload) {
  try {
    var sanitizedPayload = validateAndNormalizePayload_(payload);
    var attachmentId = saveAttachmentIfPresent_(sanitizedPayload.attachment);
    var rowData = mapPayloadToRow_(sanitizedPayload, attachmentId);
    var rowNumber = appendSubmissionRow_(rowData);

    return {
      ok: true,
      message: attachmentId
        ? 'Submission received and file uploaded.'
        : 'Submission received.',
      attachmentId: attachmentId,
      rowNumber: rowNumber,
    };
  } catch (error) {
    Logger.log('submitForm error: ' + (error && error.stack ? error.stack : error));
    throw new Error(error && error.message ? error.message : 'Submission failed.');
  }
}

/**
 * @param {Object} payload
 * @return {{name: string, email: string, message: string, attachment: Object|null}}
 * @private
 */
function validateAndNormalizePayload_(payload) {
  if (!payload) {
    throw new Error('Submission payload is required.');
  }

  var name = String(payload.name || '').trim();
  var email = String(payload.email || '').trim();
  var message = String(payload.message || '').trim();

  if (!name) {
    throw new Error('Name is required.');
  }

  if (!email) {
    throw new Error('Email is required.');
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('A valid email is required.');
  }

  return {
    name: name,
    email: email,
    message: message,
    attachment: payload.attachment || null,
  };
}

/**
 * @param {{name: string, mimeType: string, data: string}|null} attachment
 * @return {string|null}
 * @private
 */
function saveAttachmentIfPresent_(attachment) {
  if (!attachment || !attachment.data) {
    return null;
  }

  var bytes = Utilities.base64Decode(attachment.data);
  var blob = Utilities.newBlob(
    bytes,
    attachment.mimeType || 'application/octet-stream',
    attachment.name || 'attachment'
  );
  var file = DriveApp.createFile(blob);
  return file.getId();
}

/**
 * @param {{name: string, email: string, message: string}} payload
 * @param {string|null} attachmentId
 * @return {Array<*>}
 * @private
 */
function mapPayloadToRow_(payload, attachmentId) {
  return [
    new Date(),
    payload.name,
    payload.email,
    payload.message,
    attachmentId || '',
  ];
}

/**
 * @param {Array<*>} rowData
 * @return {number}
 * @private
 */
function appendSubmissionRow_(rowData) {
  var sheet = getOrCreateSubmissionSheet_();
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

/**
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 * @private
 */
function getOrCreateSubmissionSheet_() {
  var spreadsheet = SHEET_CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(SHEET_CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('No spreadsheet available. Set SHEET_CONFIG.SHEET_ID.');
  }

  var sheet = spreadsheet.getSheetByName(SHEET_CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_CONFIG.SHEET_NAME);
  }

  ensureHeaders_(sheet);
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @private
 */
function ensureHeaders_(sheet) {
  var headers = SHEET_CONFIG.HEADERS;
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  var existingHeaders = headerRange.getValues()[0];

  var missingHeaders = headers.some(function (header, index) {
    return existingHeaders[index] !== header;
  });

  if (missingHeaders) {
    headerRange.setValues([headers]);
  }
}
