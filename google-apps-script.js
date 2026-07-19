// ═══════════════════════════════════════════════════════════════
// Google Apps Script — RSVP to Google Sheets
//
// SETUP INSTRUCTIONS:
//
// 1. Go to https://sheets.google.com and create a new spreadsheet
//
// 2. In Row 1 add these headers (one per column):
//      A: Timestamp
//      B: Name
//      C: Email
//      D: Guests
//      E: Attendance
//      F: Dietary
//      G: Message
//
// 3. Go to Extensions → Apps Script
//
// 4. Delete any code in the editor and paste this entire file
//
// 5. Click Deploy → New deployment
//      - Type: Web app
//      - Description: "Wedding RSVP"
//      - Execute as: Me
//      - Who has access: Anyone
//      - Click Deploy
//
// 6. Authorize the app when prompted (allow access to Sheets)
//
// 7. Copy the Web App URL and paste it into rsvp.js as the
//    value of GOOGLE_SCRIPT_URL
//
// That's it! RSVP submissions will now appear in your sheet.
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name       || '',
      data.email      || '',
      data.guests     || '',
      data.attendance || '',
      data.dietary    || '',
      data.message    || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Required for GET requests (optional — useful for testing)
function doGet() {
  return ContentService
    .createTextOutput('RSVP endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
