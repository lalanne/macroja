// ============================================================
// Google Sheets Email Sender — Google Apps Script
// ============================================================
//
// SHEET LAYOUT (expected columns, row 1 = headers):
//
//   A: Email        — recipient email address
//   B: Name         — recipient name (used in greeting)
//   C: Subject      — email subject line
//   D: Body         — email body (plain text; supports {{name}} placeholder)
//   E: Status       — filled automatically after sending ("Sent" / "Error: …")
//
// You can customize columns in the CONFIG object below.
// ============================================================

const CONFIG = {
  sheetName: "Sheet1",       // name of the sheet tab to read from
  headerRow: 1,              // row number that contains column headers
  cols: {
    email:   1,  // column A
    name:    2,  // column B
    subject: 3,  // column C
    body:    4,  // column D
    status:  5,  // column E
  },
  senderName: "Your Name",  // display name on outgoing emails
  skipAlreadySent: true,     // skip rows where Status is already "Sent"
};

/**
 * Adds a custom menu to the Google Sheets UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Email Sender")
    .addItem("Send Emails", "sendEmails")
    .addItem("Send Test (first row only)", "sendTestEmail")
    .addToUi();
}

/**
 * Main function — iterates over every data row and sends a personalized email.
 */
function sendEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${CONFIG.sheetName}" not found.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.headerRow) {
    SpreadsheetApp.getUi().alert("No data rows found below the header.");
    return;
  }

  const dataRange = sheet.getRange(
    CONFIG.headerRow + 1, 1, lastRow - CONFIG.headerRow, sheet.getLastColumn()
  );
  const rows = dataRange.getValues();

  let sentCount = 0;
  let errorCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowNumber = CONFIG.headerRow + 1 + i;
    const email   = String(row[CONFIG.cols.email   - 1]).trim();
    const name    = String(row[CONFIG.cols.name    - 1]).trim();
    const subject = String(row[CONFIG.cols.subject - 1]).trim();
    const rawBody = String(row[CONFIG.cols.body    - 1]).trim();
    const status  = String(row[CONFIG.cols.status  - 1]).trim();

    if (!email) continue;
    if (CONFIG.skipAlreadySent && status === "Sent") continue;

    const body = personalizeBody(rawBody, name);

    try {
      GmailApp.sendEmail(email, subject, body, {
        name: CONFIG.senderName,
        htmlBody: convertToHtml(body),
      });
      sheet.getRange(rowNumber, CONFIG.cols.status).setValue("Sent");
      sentCount++;
    } catch (err) {
      sheet.getRange(rowNumber, CONFIG.cols.status).setValue("Error: " + err.message);
      errorCount++;
    }
  }

  SpreadsheetApp.getUi().alert(
    `Done!\n\nSent: ${sentCount}\nErrors: ${errorCount}`
  );
}

/**
 * Sends only the first data row — useful for testing your template.
 */
function sendTestEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${CONFIG.sheetName}" not found.`);
    return;
  }

  const rowNumber = CONFIG.headerRow + 1;
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const email   = String(row[CONFIG.cols.email   - 1]).trim();
  const name    = String(row[CONFIG.cols.name    - 1]).trim();
  const subject = String(row[CONFIG.cols.subject - 1]).trim();
  const rawBody = String(row[CONFIG.cols.body    - 1]).trim();

  if (!email) {
    SpreadsheetApp.getUi().alert("First data row has no email address.");
    return;
  }

  const body = personalizeBody(rawBody, name);

  try {
    GmailApp.sendEmail(email, subject, body, {
      name: CONFIG.senderName,
      htmlBody: convertToHtml(body),
    });
    sheet.getRange(rowNumber, CONFIG.cols.status).setValue("Sent");
    SpreadsheetApp.getUi().alert(`Test email sent to ${email}`);
  } catch (err) {
    sheet.getRange(rowNumber, CONFIG.cols.status).setValue("Error: " + err.message);
    SpreadsheetApp.getUi().alert("Failed: " + err.message);
  }
}

/**
 * Replaces {{name}} (case-insensitive) with the recipient's name.
 * Add more placeholders here as needed.
 */
function personalizeBody(body, name) {
  return body.replace(/\{\{name\}\}/gi, name);
}

/**
 * Wraps plain text in minimal HTML so line breaks render properly in email.
 */
function convertToHtml(plainText) {
  const escaped = plainText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
  return `<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">${escaped}</div>`;
}
