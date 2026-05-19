/**
 * Container Arrival Tracker — Google Apps Script
 * 
 * Setup Instructions:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Paste this entire file into the editor
 * 4. Click Deploy → New Deployment
 *    - Type: Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy and copy the Web App URL
 * 6. Paste the URL into the tracker's Google Sheets Config panel
 *
 * Built by Testa Wu — Eshcol Studio | Build. Model. Deliver.
 * https://buildmodeldeliver.com
 */

const SHEET_NAME = 'Records';

const HEADERS = [
  'ID',
  'Cargo Name',
  'Project Code',
  'Invoice No.',
  'Packing List No.',
  'Customs Date',
  'Arrival Date',
  'Yard Zone',
  'Container No.',
  'Invoice File',
  'Packing List File',
  'Has Photo',
  'Created At'
];

// ── READ ──────────────────────────────────────────────────────────────────────
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: values }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── WRITE ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // Initialize headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  // ── APPEND (new record) ──
  if (action === 'append') {
    sheet.appendRow([
      data.id,
      data.cargo,
      data.project,
      data.invoiceNo,
      data.plNo,
      data.customsDate,
      data.arrivalDate,
      data.zone,
      data.containerNo,
      data.invoiceFile,
      data.plFile,
      data.hasPhoto,
      data.createdAt
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', action: 'append' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── UPDATE (edit record) ──
  if (action === 'update') {
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 1, 1, 13).setValues([[
          data.id,
          data.cargo,
          data.project,
          data.invoiceNo,
          data.plNo,
          data.customsDate,
          data.arrivalDate,
          data.zone,
          data.containerNo,
          data.invoiceFile,
          data.plFile,
          data.hasPhoto,
          data.createdAt
        ]]);
        break;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', action: 'update' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── DELETE ──
  if (action === 'delete') {
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', action: 'delete' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── UNKNOWN ACTION ──
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action: ' + action }))
    .setMimeType(ContentService.MimeType.JSON);
}
