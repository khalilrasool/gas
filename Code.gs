const SPREADSHEET_ID = 'REPLACE_WITH_SHEET_ID';
const SHEET_NAME = 'Sheet1';

/**
 * Generates a PDF of member cards from the configured Google Sheet.
 * @returns {GoogleAppsScript.Base.Blob} PDF blob that was saved to Drive.
 */
function exportMemberCards() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" was not found in the spreadsheet.`);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    throw new Error('No member data found in the sheet.');
  }

  const rows = values.slice(1).filter((row) => row.some((cell) => cell !== ''));
  const members = rows.map((row) => ({
    name: row[2] || '',
    cpr: row[1] || '',
    address: row[4] || '',
    phone: row[3] || '',
    photoUrl: resolveImageDataUrl(row[5]),
    cardUrl: resolveImageDataUrl(row[6]),
  }));

  const template = HtmlService.createTemplateFromFile('card-template');
  template.members = members;
  const html = template.evaluate().setTitle('Member Cards').getContent();

  const pdfBlob = Utilities.newBlob(html, 'text/html', 'member-cards.html')
    .getAs('application/pdf')
    .setName('member-cards.pdf');

  DriveApp.createFile(pdfBlob);
  return pdfBlob;
}

/**
 * Allows including additional HTML files when using HtmlService templates.
 * @param {string} filename
 * @returns {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Converts an image URL (Drive link or external URL) to a data URL for embedding in the PDF.
 * @param {string} url
 * @returns {string}
 */
function resolveImageDataUrl(url) {
  if (!url) return '';

  try {
    const driveId = extractDriveId(url);
    if (driveId) {
      return blobToDataUrl(DriveApp.getFileById(driveId).getBlob());
    }

    if (/^[a-z][a-z0-9+.-]*:\/\//i.test(url)) {
      const response = UrlFetchApp.fetch(url);
      return blobToDataUrl(response.getBlob());
    }
  } catch (error) {
    Logger.log('Image fetch failed for %s: %s', url, error);
  }

  return '';
}

/**
 * Converts a blob to a base64 data URL string.
 * @param {GoogleAppsScript.Base.Blob} blob
 * @returns {string}
 */
function blobToDataUrl(blob) {
  const contentType = blob.getContentType() || 'image/png';
  const base64 = Utilities.base64Encode(blob.getBytes());
  return `data:${contentType};base64,${base64}`;
}

/**
 * Extracts the Drive file ID from a sharing link.
 * @param {string} url
 * @returns {string | null}
 */
function extractDriveId(url) {
  const patterns = [
    /[-\w]{25,}/, // generic file ID pattern
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = url.match(patterns[i]);
    if (match) return match[0];
  }

  return null;
}
