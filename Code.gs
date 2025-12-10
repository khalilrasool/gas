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
    name: row[3] || '',
    cpr: row[2] || '',
    address: row[4] || '',
    phone: row[5] || '',
    photoUrl: row[6] || '',
    cardUrl: row[7] || '',
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
