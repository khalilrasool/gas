# gas

Google Apps Script snippet to export member cards from a Google Sheet into a PDF.

## Usage
1. Create a new Apps Script project attached to the Google Sheet that holds your member data.
2. Add the contents of `Code.gs` and `card-template.html` to the project (File → New → Script / HTML file).
3. Set the `SPREADSHEET_ID` and `SHEET_NAME` constants in `Code.gs` to match your sheet.
4. Run `exportMemberCards()` to generate a PDF. The script reads rows in the following order: timestamp, CPR, name, phone, address, photo URL (`img`), card image URL (`card_img`). The PDF file (`member-cards.pdf`) is saved to your Drive and also returned as a blob.
5. Images for the personal photo and ID card are fetched and embedded as data URLs so they render correctly in the PDF. Use sharing links from Drive or direct URLs in the `img` and `card_img` columns.
