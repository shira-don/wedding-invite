// ═══════════════════════════════════════════════════════════════
// Google Apps Script — Wedding RSVP (guest-list matched)
//
// This is a STANDALONE script (do NOT paste it over the entourage
// script). It opens the same spreadsheet by ID and reads/writes a
// "Guests" tab, so the working entourage endpoint is left untouched.
//
// ── SETUP ──────────────────────────────────────────────────────
// 1. Open the spreadsheet that already powers the entourage list.
//    Copy its ID from the URL:
//      https://docs.google.com/spreadsheets/d/<THIS_IS_THE_ID>/edit
//    Paste it into SHEET_ID below.
//
// 2. Add a tab named exactly "Guests" with these Row 1 headers:
//      A: Party        (invitation display name, e.g. "The Dela Cruz Family")
//      B: MatchNames   (comma-separated names a guest may type, e.g.
//                       "Juan Dela Cruz, Maria Dela Cruz")
//      C: Seats        (allotted headcount, a number)
//      D: Side         (optional: Bride / Groom)
//      E: Table        (leave blank — for the future "find your seat" page)
//      F: SeatNote     (leave blank — future)
//      G: Status       (filled by this script: Attending / Declined)
//      H: Count        (filled by this script)
//      I: Mobile       (filled by this script)
//      J: Message      (filled by this script)
//      K: RespondedAt  (filled by this script)
//    Fill A–D for every invited party. Leave E–K empty.
//
// 3. (Auto) A second tab "RSVP_Log" is created on first submit for history.
//
// 4. Go to https://script.google.com → New project → paste this whole
//    file. Set SHEET_ID. Save.
//
// 5. Deploy → New deployment → Type: Web app
//      - Execute as: Me
//      - Who has access: Anyone
//    Deploy, authorize when prompted, and COPY the Web App URL.
//
// 6. Paste that URL into rsvp.js as RSVP_SCRIPT_URL.
// ═══════════════════════════════════════════════════════════════

var SHEET_ID   = 'YOUR_SPREADSHEET_ID';   // ← paste the spreadsheet ID here
var GUESTS_TAB = 'Guests';
var LOG_TAB    = 'RSVP_Log';

// Column indexes (1-based) in the Guests tab.
var COL = { PARTY:1, MATCH:2, SEATS:3, SIDE:4, TABLE:5, SEATNOTE:6,
            STATUS:7, COUNT:8, MOBILE:9, MESSAGE:10, RESPONDED_AT:11 };

function doGet(e) {
  var action = (e && e.parameter && e.parameter.rsvp) || '';
  try {
    if (action === 'lookup') return json(lookup_(e.parameter.name || ''));
    if (action === 'submit') return json(submit_(e.parameter));
    return json({ status: 'ok', message: 'Wedding RSVP endpoint is running.' });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
}

// ── Lookup: return only the matched parties (no full-list leakage) ──
function lookup_(rawName) {
  var key = tokenKey_(rawName);
  if (!key) return { status: 'success', matches: [] };

  var rows = guestsSheet_().getDataRange().getValues();
  var matches = [];
  for (var i = 1; i < rows.length; i++) {          // skip header
    var r = rows[i];
    if (!r[COL.PARTY - 1]) continue;
    var candidates = [r[COL.PARTY - 1]].concat(
      String(r[COL.MATCH - 1] || '').split(','));
    var hit = candidates.some(function (c) { return tokenKey_(c) === key; });
    if (hit) {
      var status = String(r[COL.STATUS - 1] || '');
      matches.push({
        id: i + 1,                                  // 1-based sheet row
        party: String(r[COL.PARTY - 1]),
        seats: Number(r[COL.SEATS - 1]) || 1,
        responded: status !== '',
        status: status,
        count: Number(r[COL.COUNT - 1]) || 0,
        mobile: String(r[COL.MOBILE - 1] || ''),
        message: String(r[COL.MESSAGE - 1] || '')
      });
    }
  }
  return { status: 'success', matches: matches };
}

// ── Submit: re-validate, cap to Seats, upsert the row, log it ──
function submit_(p) {
  var sheet = guestsSheet_();
  var id = Number(p.id);
  if (!id || id < 2) return { status: 'error', message: 'Invalid invitation reference.' };

  var lastRow = sheet.getLastRow();
  if (id > lastRow) return { status: 'error', message: 'Invitation not found.' };

  var row = sheet.getRange(id, 1, 1, COL.RESPONDED_AT).getValues()[0];
  var party = String(row[COL.PARTY - 1] || '');
  if (!party || tokenKey_(party) !== tokenKey_(p.party || '')) {
    return { status: 'error', message: 'We could not verify your invitation. Please try again.' };
  }

  var seats = Number(row[COL.SEATS - 1]) || 1;
  var attending = String(p.attending || 'yes') === 'no' ? 'Declined' : 'Attending';
  var count = attending === 'Declined' ? 0 : Math.max(1, Math.min(seats, Number(p.count) || 1));
  var mobile = String(p.mobile || '').trim();
  var message = String(p.message || '').trim();
  var now = new Date();

  if (attending === 'Attending' && Number(p.count) > seats) {
    return { status: 'error', message: 'That is more than your allotted seats (' + seats + ').' };
  }
  if (!mobile) return { status: 'error', message: 'A mobile number is required.' };

  sheet.getRange(id, COL.STATUS,       1, 1).setValue(attending);
  sheet.getRange(id, COL.COUNT,        1, 1).setValue(count);
  sheet.getRange(id, COL.MOBILE,       1, 1).setValue(mobile);
  sheet.getRange(id, COL.MESSAGE,      1, 1).setValue(message);
  sheet.getRange(id, COL.RESPONDED_AT, 1, 1).setValue(now);

  logSheet_().appendRow([now, party, attending, count, mobile, message]);

  return { status: 'success', party: party, attending: attending, count: count };
}

// ── Helpers ──
function guestsSheet_() {
  var s = SpreadsheetApp.openById(SHEET_ID).getSheetByName(GUESTS_TAB);
  if (!s) throw new Error('Tab "' + GUESTS_TAB + '" not found.');
  return s;
}

function logSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s = ss.getSheetByName(LOG_TAB);
  if (!s) {
    s = ss.insertSheet(LOG_TAB);
    s.appendRow(['Timestamp', 'Party', 'Status', 'Count', 'Mobile', 'Message']);
  }
  return s;
}

// Order-insensitive, title/case/diacritic/punctuation-tolerant name key.
function tokenKey_(name) {
  var s = String(name || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')   // strip accents
    .toLowerCase()
    .replace(/[.,]/g, ' ')
    .replace(/\b(mr|mrs|ms|miss|sir|madam|dr|doctor|engr|atty|rev|fr|hon|prof|sr|jr|iii|iv|ii)\b/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return '';
  return s.split(' ').sort().join(' ');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
