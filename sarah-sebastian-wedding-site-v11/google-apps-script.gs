/**
 * Wedding RSVP System — source supplied for the existing Google Apps Script deployment.
 *
 * Guests:
 * Invite ID | Primary Guest | Guest 2 | Guest 3 | Email | Magic Code | RSVP Link
 *
 * RSVPs (column order is significant):
 * Timestamp | Invite ID | Guest Name | Attending | Dietary Requirements | Email | Notes
 *
 * GET /exec?invite=MAGIC_CODE returns the invitation's guest names.
 * POST { invite, responses: [{ name, attending, dietary }], notes, website }
 * writes one row per guest. Email is optional and defaults to the Guests sheet.
 * Each attending guest's responses[].dietary is written to column E.
 *
 * This file mirrors the supplied backend. A Git push does not redeploy Apps Script.
 */
const SPREADSHEET_ID = "1XsNAFnWTsdNTvb89ndlO3-oWrvR6WTzRxpt--qnVa6Y";
const GUEST_SHEET_NAME = "Guests";
const RSVP_SHEET_NAME = "RSVPs";
const RSVP_BASE_URL = "https://www.sarahsebastianwedding.site";

function doGet(e) {
  try {
    const code = cleanCode_(e.parameter.invite);
    if (!code) {
      return json_({ ok: false, valid: false, error: "Missing invitation code" });
    }
    const invitation = findInvitation_(code);
    if (!invitation) return json_({ ok: true, valid: false });
    return json_({
      ok: true,
      valid: true,
      inviteId: invitation.inviteId,
      guests: invitation.guests,
      email: invitation.email
    });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, valid: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : "{}");

    // Honeypot: ignore automated submissions without writing any rows.
    if (payload.website) return json_({ ok: true });
    const code = cleanCode_(payload.invite);
    if (!code) return json_({ ok: false, error: "Missing invitation code" });
    const invitation = findInvitation_(code);
    if (!invitation) return json_({ ok: false, error: "Invalid invitation" });
    if (!Array.isArray(payload.responses) || payload.responses.length === 0) {
      return json_({ ok: false, error: "No RSVP responses supplied" });
    }

    const allowedGuests = invitation.guests;
    const responses = payload.responses.filter(response => {
      if (!response || !response.name) return false;
      return allowedGuests.includes(String(response.name).trim());
    });
    if (responses.length === 0) {
      return json_({ ok: false, error: "No valid guest responses supplied" });
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(RSVP_SHEET_NAME);
    if (!sheet) throw new Error('Sheet tab "' + RSVP_SHEET_NAME + '" not found.');

    // Guests may resubmit their invitation; replace its previous responses.
    removeExistingRSVPs_(sheet, invitation.inviteId);
    const timestamp = new Date();
    const email = clean_(payload.email || invitation.email || "");
    const notes = clean_(payload.notes);
    const rows = responses.map(response => {
      const attending = response.attending === true ||
        String(response.attending).toLowerCase() === "yes" ? "Yes" : "No";
      const dietary = attending === "Yes" ? clean_(response.dietary) : "";
      return [
        timestamp,
        clean_(invitation.inviteId),
        clean_(response.name),
        attending,
        dietary,
        email,
        notes
      ];
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return json_({ ok: true, inviteId: invitation.inviteId });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function findInvitation_(code) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(GUEST_SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + GUEST_SHEET_NAME + '" not found.');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(h => String(h).trim());
  const inviteIdCol = headers.indexOf("Invite ID");
  const primaryGuestCol = headers.indexOf("Primary Guest");
  const guest2Col = headers.indexOf("Guest 2");
  const guest3Col = headers.indexOf("Guest 3");
  const guest4Col = headers.indexOf("Guest 4");
  const emailCol = headers.indexOf("Email");
  const codeCol = headers.indexOf("Magic Code");
  if (inviteIdCol === -1 || primaryGuestCol === -1 || codeCol === -1) {
    throw new Error("Guests sheet headers are incorrect.");
  }
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (cleanCode_(row[codeCol]) !== code) continue;
    const guests = [
      row[primaryGuestCol],
      guest2Col >= 0 ? row[guest2Col] : "",
      guest3Col >= 0 ? row[guest3Col] : ""
    ].map(value => String(value || "").trim()).filter(Boolean);
    return {
      inviteId: String(row[inviteIdCol] || "").trim(),
      email: emailCol >= 0 ? String(row[emailCol] || "").trim() : "",
      guests: guests
    };
  }
  return null;
}

function removeExistingRSVPs_(sheet, inviteId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    const existingInviteId = String(data[i][1] || "").trim();
    if (existingInviteId === String(inviteId).trim()) sheet.deleteRow(i + 2);
  }
}

// Only fill codes for guests whose Magic Code cell is empty.
function generateMagicCodes() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(GUEST_SHEET_NAME);
  if (!sheet) throw new Error('Sheet tab "' + GUEST_SHEET_NAME + '" not found.');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(h => String(h).trim());
  const codeCol = headers.indexOf("Magic Code");
  const primaryGuestCol = headers.indexOf("Primary Guest");
  if (codeCol === -1) throw new Error('Missing "Magic Code" column.');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const primaryGuest = primaryGuestCol >= 0 ? sheet.getRange(row, primaryGuestCol + 1).getValue() : "";
    if (!primaryGuest) continue;
    const codeCell = sheet.getRange(row, codeCol + 1);
    if (!codeCell.getValue()) codeCell.setValue(Utilities.getUuid().replace(/-/g, ""));
  }
}

// Keep existing invitation IDs untouched.
function generateInviteIds() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(GUEST_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idCol = headers.indexOf("Invite ID");
  const primaryGuestCol = headers.indexOf("Primary Guest");
  if (idCol === -1) throw new Error('Missing "Invite ID" column.');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const primaryGuest = sheet.getRange(row, primaryGuestCol + 1).getValue();
    if (!primaryGuest) continue;
    const idCell = sheet.getRange(row, idCol + 1);
    if (!idCell.getValue()) idCell.setValue("INV" + String(row - 1).padStart(3, "0"));
  }
}

function generateRSVPLinks() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(GUEST_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const codeCol = headers.indexOf("Magic Code");
  const linkCol = headers.indexOf("RSVP Link");
  if (codeCol === -1 || linkCol === -1) {
    throw new Error('Guests sheet requires "Magic Code" and "RSVP Link" columns.');
  }
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const code = String(sheet.getRange(row, codeCol + 1).getValue() || "").trim();
    if (!code) continue;
    const link = RSVP_BASE_URL + "?invite=" + encodeURIComponent(code);
    sheet.getRange(row, linkCol + 1).setValue(link);
  }
}

function setupGuestLinks() {
  generateInviteIds();
  generateMagicCodes();
  generateRSVPLinks();
}

function cleanCode_(value) {
  return String(value || "").trim().toLowerCase();
}

function clean_(value) {
  const text = String(value || "").trim();
  if (/^[=+\-@]/.test(text)) return "'" + text;
  return text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
