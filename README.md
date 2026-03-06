# Google Sheets Email Sender

Send personalized emails directly from a Google Sheet using Google Apps Script and Gmail.

## Setup

### 1. Prepare your Google Sheet

Create a Google Sheet with the following columns in **Row 1** (headers):

| A       | B      | C        | D     | E      |
|---------|--------|----------|-------|--------|
| Email   | Name   | Subject  | Body  | Status |

Then fill in your data starting from **Row 2**:

| A                  | B       | C                  | D                                                    | E |
|--------------------|---------|--------------------|----------------------------------------------------|---|
| alice@example.com  | Alice   | Hello Alice!       | Hi {{name}}, just wanted to say hello!              |   |
| bob@example.com    | Bob     | Meeting follow-up  | Dear {{name}}, thanks for attending the meeting.    |   |

- **Email** — recipient address
- **Name** — used to personalize the body via `{{name}}`
- **Subject** — each row can have a unique subject
- **Body** — plain text; use `{{name}}` where you want the name inserted
- **Status** — left blank; the script fills it with "Sent" or an error message

### 2. Add the script to your sheet

1. Open your Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Delete any code in the editor and paste the contents of `Code.gs`.
4. Click the **Save** icon (or Ctrl/Cmd + S).
5. Close the Apps Script editor tab.

### 3. Run it

1. Reload your Google Sheet — a new menu **Email Sender** will appear in the menu bar.
2. Click **Email Sender → Send Test (first row only)** to verify with a single email.
3. Google will ask you to authorize the script the first time:
   - Click **Review Permissions** → choose your Google account → **Advanced** → **Go to (project name)** → **Allow**.
4. Once the test works, click **Email Sender → Send Emails** to process all rows.

## Features

- **Personalization** — use `{{name}}` in the Body column; the script replaces it with each row's Name.
- **Skip already sent** — rows with "Sent" in the Status column are skipped, so you can safely re-run.
- **Error tracking** — if an email fails, the Status column shows the error message.
- **Test mode** — send only the first row to verify your template before sending to everyone.

## Configuration

Edit the `CONFIG` object at the top of `Code.gs` to change:

- `sheetName` — if your tab isn't called "Sheet1"
- `senderName` — the display name on outgoing emails
- `cols` — if your columns are in a different order
- `skipAlreadySent` — set to `false` to re-send to all rows

## Gmail daily limits

Google imposes sending limits:
- **Free Gmail**: ~100 emails/day
- **Google Workspace**: ~1,500 emails/day

The script sends one email per row. If you exceed the limit, remaining rows will show an error in the Status column — just run again the next day.
