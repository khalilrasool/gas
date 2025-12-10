const SPREADSHEET_ID = 'REPLACE_WITH_SHEET_ID';
const SHEET_NAME = 'Sheet1';

/**
 * Generates a PDF of member cards from the configured Google Sheet.
 * @returns {GoogleAppsScript.Base.Blob} PDF blob that was saved to Drive.
 */
function exportMemberCards() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'REPLACE_WITH_SHEET_ID') {
    throw new Error('Set SPREADSHEET_ID to the target spreadsheet ID before running the script.');
  }

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet "${SHEET_NAME}" was not found in the spreadsheet.`);
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  if (values.length <= 1) {
    throw new Error('No member data found in the sheet.');
  }

  const rows = values.slice(1).filter((row) => row.some((cell) => cell !== ''));
  const imageLookup = resolveImageBatch(rows);
  const members = rows.map((row) => ({
    name: row[2] || '',
    cpr: row[1] || '',
    address: row[4] || '',
    phone: row[3] || '',
    photoUrl: imageLookup[row[5]] || '',
    cardUrl: imageLookup[row[6]] || '',
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
 * Resolves image URLs for all rows in batches to reduce execution time.
 * @param {any[][]} rows
 * @returns {Object<string, string>}
 */
function resolveImageBatch(rows) {
  const urls = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][5]) urls.push(rows[i][5]);
    if (rows[i][6]) urls.push(rows[i][6]);
  }

  const uniqueUrls = Array.from(new Set(urls));
  const lookup = {};
  const driveIds = [];
  const fetchRequests = [];

  uniqueUrls.forEach(function (url) {
    const driveId = extractDriveId(url);
    if (driveId) {
      driveIds.push({ url: url, id: driveId });
    } else if (/^[a-z][a-z0-9+.-]*:\/\//i.test(url)) {
      fetchRequests.push({ url: url, request: { url: url, muteHttpExceptions: true } });
    }
  });

  // Fetch Drive images sequentially.
  driveIds.forEach(function (entry) {
    try {
      lookup[entry.url] = blobToDataUrl(DriveApp.getFileById(entry.id).getBlob());
    } catch (error) {
      Logger.log('Drive image fetch failed for %s: %s', entry.url, error);
    }
  });

  // Fetch external images in parallel where possible.
  if (fetchRequests.length) {
    try {
      var responses = UrlFetchApp.fetchAll(fetchRequests.map(function (item) { return item.request; }));
      responses.forEach(function (response, index) {
        if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
          lookup[fetchRequests[index].url] = blobToDataUrl(response.getBlob());
        }
      });
    } catch (error) {
      Logger.log('External image fetch failed: %s', error);
    }
  }

  return lookup;
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
