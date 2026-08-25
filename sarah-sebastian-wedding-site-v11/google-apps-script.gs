/**
 * Wedding RSVP -> Google Sheets
 *
 * 1. Create a Google Sheet and add a sheet/tab named "RSVPs".
 * 2. Put these headers in row 1:
 *    Timestamp | Name | Email | Attending | Other Guests | Bunker Bay | Dietary | Note
 * 3. Paste your spreadsheet ID below.
 * 4. In script.google.com, create a project and paste this file.
 * 5. Deploy > New deployment > Web app.
 * 6. Execute as: Me.
 * 7. Who has access: Anyone.
 * 8. Copy the Web App URL into GOOGLE_APPS_SCRIPT_URL in index.html.
 */

const SPREADSHEET_ID = "PASTE_YOUR_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "RSVPs";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const payload = JSON.parse(e.postData.contents || "{}");

    // Honeypot: silently accept bots without writing a row.
    if (payload.website) {
      return json_({ ok: true });
    }

    if (!payload.name || !payload.email || !payload.attending) {
      return json_({ ok: false, error: "Missing required fields" });
    }

    const sheet = SpreadsheetApp
      .openById(SPREADSHEET_ID)
      .getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet tab "' + SHEET_NAME + '" not found.');
    }

    sheet.appendRow([
      new Date(),
      clean_(payload.name),
      clean_(payload.email),
      clean_(payload.attending),
      clean_(payload.guests),
      clean_(payload.bunkerBay),
      clean_(payload.dietary),
      clean_(payload.note)
    ]);

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function clean_(value) {
  const text = String(value || "").trim();

  // Prevent spreadsheet formula injection.
  if (/^[=+\-@]/.test(text)) return "'" + text;
  return text;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
