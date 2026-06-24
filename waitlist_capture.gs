/**
 * Loomful waitlist capture — Google Apps Script.
 *
 * Receives POSTs from the loomful.app early-access form and appends each
 * signup as a row in the bound Google Sheet. No third-party service, no
 * account beyond the Google account that owns the sheet.
 *
 * SETUP (one time, ~5 clicks):
 *   1. Create a new Google Sheet (name it e.g. "Loomful waitlist").
 *   2. Extensions -> Apps Script. Delete the stub, paste this whole file.
 *   3. Deploy -> New deployment -> type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Deploy, authorise when prompted, copy the Web app URL (ends in /exec).
 *   4. Send that /exec URL back — it gets wired into the form's action.
 *
 * Re-deploying after edits: Deploy -> Manage deployments -> edit (pencil)
 * -> Version: New version -> Deploy. The /exec URL stays the same.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // header row on first write
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['timestamp', 'email', 'source']);
    }
    var params = (e && e.parameter) ? e.parameter : {};
    var email = (params.email || '').toString().trim();
    if (!email) {
      return _json({ ok: false, error: 'no email' });
    }
    sheet.appendRow([new Date(), email, params.source || 'loomful.app']);
    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

// A GET just confirms the endpoint is alive (handy for a quick browser check).
function doGet() {
  return _json({ ok: true, service: 'loomful-waitlist' });
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
