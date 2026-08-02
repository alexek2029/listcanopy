/**
 * List Canopy lead capture, Google Apps Script.
 *
 * Receives POSTs from the listcanopy.com form, appends each enquiry as a row
 * in the bound Google Sheet, and emails a notification to NOTIFY_EMAIL.
 * No third-party service, no account beyond the Google account that owns the
 * sheet. Lead data never leaves your own Google Workspace.
 *
 * SETUP (one time, about 5 clicks, must be done by the account owner):
 *   1. Create a new Google Sheet (name it e.g. "List Canopy leads").
 *   2. Extensions -> Apps Script. Delete the stub, paste this whole file.
 *   3. Deploy -> New deployment -> type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Deploy, authorise when prompted (Google will warn the app is
 *      unverified, that is expected for your own script: Advanced ->
 *      Go to project), then copy the Web app URL (ends in /exec).
 *   4. Paste that /exec URL into index.html as the form's action attribute,
 *      replacing REPLACE_WITH_APPS_SCRIPT_URL.
 *
 * Re-deploying after edits: Deploy -> Manage deployments -> edit (pencil)
 * -> Version: New version -> Deploy. The /exec URL stays the same.
 *
 * Test without the site: visit the /exec URL in a browser. A GET returns
 * {"ok":true,...}. To test the email path, use "Run -> testNotification"
 * inside the Apps Script editor.
 */

var NOTIFY_EMAIL = 'alex@listcanopy.com';
var SITE = 'listcanopy.com';

function doPost(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var email = (params.email || '').toString().trim();
    if (!email) {
      return _json({ ok: false, error: 'no email' });
    }

    var when = new Date();
    var source = params.source || SITE;

    _appendRow(when, email, source);

    // Notification is best-effort: a mail failure must not lose the lead,
    // which is already safely written to the sheet above.
    try {
      _notify(email, when, source);
    } catch (mailErr) {
      _appendRow(when, 'NOTIFY FAILED: ' + String(mailErr), source);
    }

    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _appendRow(when, email, source) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['timestamp', 'email', 'source']);
  }
  sheet.appendRow([when, email, source]);
}

function _notify(email, when, source) {
  var subject = 'New List Canopy enquiry: ' + email;
  var body = [
    'A new automation map request came in.',
    '',
    'Email:  ' + email,
    'When:   ' + when,
    'Source: ' + source,
    '',
    'Reply straight to this address to respond.',
    '',
    'Logged in the leads sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].join('\n');

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    body: body,
    replyTo: email,
    name: 'List Canopy site'
  });
}

// A GET just confirms the endpoint is alive (handy for a quick browser check).
function doGet() {
  return _json({ ok: true, service: 'listcanopy-leads' });
}

// Run this manually from the Apps Script editor to verify email delivery.
function testNotification() {
  _notify('test@example.com', new Date(), 'manual test');
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
