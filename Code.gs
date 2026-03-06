const CONFIG = {
  sheetName: "Sheet1",
  headerRow: 1,
  cols: {
    email:   1,
    name:    2,
    subject: 3,
    body:    4,
    status:  5,
  },
  senderName: "Your Name",
  skipAlreadySent: true,
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Email Sender")
    .addItem("Send Emails", "sendEmails")
    .addItem("Send Test (first row only)", "sendTestEmail")
    .addToUi();
}

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

function personalizeBody(body, name) {
  return body.replace(/\{\{name\}\}/gi, name);
}

function convertToHtml(plainText) {
  const escaped = plainText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>");
  return `<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">${escaped}</div>`;
}
