// ═══════════════════════════════════════════════════════════════
// Google Apps Script — Entourage + Wedding RSVP (one deployment)
//
// This REPLACES your existing entourage script. It keeps the entourage
// endpoint exactly as-is and adds the RSVP endpoints, so the whole thing
// runs on the SAME Web App URL you already have.
//
// The guest list is one row per person (tab "GUEST LIST", data from row 4).
// Invitations are the blocks of rows between BLANK rows — searching any
// member pulls up the whole block so one person can RSVP for everyone.
// RSVP writes to RSVP RESPONSE (col G) and NOTES (col H) only.
//
// Columns:  A NO · B NAME · C RELATIONSHIP · D TABLE NO. ·
//           E ROLE · F SENT · G RSVP RESPONSE · H NOTES
//
// ── HOW TO UPDATE (keeps the same URL) ─────────────────────────
// 1. Open your existing entourage Apps Script project.
// 2. Replace ALL of its code with this file. Save.
// 3. Deploy → Manage deployments → (your active Web app) → ✎ Edit →
//    Version: "New version" → Deploy.  ← same URL is kept.
//    (Do NOT create a *new* deployment — that would change the URL.)
// The RSVP form in rsvp.js already points at this same URL.
// ═══════════════════════════════════════════════════════════════

const SHEET_ID   = '1t5cbsdMCdoAq2igyebTX4-TFycz4nDVNPwXPwyvYdZE';
const GUESTS_TAB = 'GUEST LIST';
const LOG_TAB    = 'RSVP_Log';
const DATA_START = 4;   // first data row (rows 1-3 are the title/header)

// 1-based column indexes in the GUEST LIST tab.
const COL = { NO:1, NAME:2, RELATIONSHIP:3, TABLE:4, ROLE:5, SENT:6, RESPONSE:7, NOTES:8 };

function doGet(e) {
  const p = (e && e.parameter) || {};
  const which = (p.list || '').toLowerCase();
  if (which === 'entourage') return json(getEntourage());
  if (which === 'seats') return json(getSeats());

  const action = (p.rsvp || '').toLowerCase();
  try {
    if (action === 'lookup') return json(lookup_(p.name || ''));
    if (action === 'submit') return json(submit_(p));
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

// ── Entourage (unchanged behavior) ──
function getEntourage() {
  const sheet = guestsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { updated: new Date().toISOString(), people: [] };
  const values = sheet.getRange(DATA_START, 1, lastRow - (DATA_START - 1), 8).getValues();
  const people = values
    .map(r => ({
      name: cleanName_(r[COL.NAME - 1]),
      relationship: String(r[COL.RELATIONSHIP - 1] || '').trim(),
      table: String(r[COL.TABLE - 1] || '').trim(),
      role: String(r[COL.ROLE - 1] || '').trim(),
    }))
    .filter(p => p.name && p.role)
    .map(p => ({
      ...p,
      side: /groom/i.test(p.relationship) ? 'groom'
          : /bride/i.test(p.relationship) ? 'bride' : 'mutual',
    }));
  return { updated: new Date().toISOString(), people };
}

// ── RSVP lookup: find the invitation block(s) a typed name belongs to ──
// Match is by name-word SUBSET: every word typed must appear in a member's
// name (order/case/accents/titles ignored). So a first name alone — or any
// subset of the full name — pulls up that member's whole invitation block.
function lookup_(rawName) {
  const inTokens = nameTokens_(rawName);
  if (!inTokens.length) return { status: 'success', groups: [] };

  const groups = buildGroups_(guestsSheet_().getDataRange().getValues());
  const hits = groups.filter(g => g.some(m => subsetOf_(inTokens, m.tokens)));

  return {
    status: 'success',
    groups: hits.map(g => ({
      headId: g[0].id,
      seats: g.length,
      members: g.map(m => ({ id: m.id, name: m.name, role: m.role, response: m.response })),
    })),
  };
}

// ── RSVP submit: write yes/no per member (col G); mobile+message to head Notes (col H) ──
function submit_(p) {
  const sheet = guestsSheet_();
  const members = String(p.m || '').split(';').filter(String).map(s => {
    const a = s.split(':');
    return { id: Number(a[0]), resp: (a[1] === 'yes' ? 'yes' : 'no') };
  });
  if (!members.length) return { status: 'error', message: 'No guests were selected.' };

  const mobile = String(p.mobile || '').trim();
  if (!mobile) return { status: 'error', message: 'A mobile number is required.' };

  const lastRow = sheet.getLastRow();
  let yes = 0, no = 0;
  members.forEach(m => {
    if (m.id >= DATA_START && m.id <= lastRow) {
      const name = String(sheet.getRange(m.id, COL.NAME, 1, 1).getValue() || '').trim();
      if (!name) return;                      // never write to a blank separator row
      sheet.getRange(m.id, COL.RESPONSE, 1, 1).setValue(m.resp);
      m.resp === 'yes' ? yes++ : no++;
    }
  });

  const head = Number(p.head) || members[0].id;
  const headName = String(sheet.getRange(head, COL.NAME, 1, 1).getValue() || '').trim();
  const note = 'RSVP · ' + mobile + (p.message ? (' · ' + String(p.message).trim()) : '');
  sheet.getRange(head, COL.NOTES, 1, 1).setValue(note);

  logSheet_().appendRow([new Date(), headName, mobile, yes, no, String(p.message || '').trim()]);

  return { status: 'success', attending: yes, declined: no };
}

// ── Seating: every guest who has not declined, with their TABLE NO. (col D) ──
function getSeats() {
  const sheet = guestsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return { updated: new Date().toISOString(), guests: [] };
  const values = sheet.getRange(DATA_START, 1, lastRow - (DATA_START - 1), 8).getValues();
  const guests = values
    .map(r => ({
      name: cleanName_(r[COL.NAME - 1]),
      relationship: String(r[COL.RELATIONSHIP - 1] || '').trim(),
      table: String(r[COL.TABLE - 1] || '').trim(),
      vip: String(r[COL.NAME - 1] || '').trim().slice(-1) === '^',
      response: String(r[COL.RESPONSE - 1] || '').trim().toLowerCase(),
    }))
    .filter(g => g.name && g.response !== 'no');
  return { updated: new Date().toISOString(), guests: guests };
}

// ── One-off: write the generated seat plan into TABLE NO. (col D). ──
// Run this ONCE from the Apps Script editor (Run > applySeatPlan), then edit
// column D by hand from then on — the website reads whatever is in that column.
// Only fills rows it can match by name; never clears a cell it cannot match.
function applySeatPlan() {
  const PLAN = {
    "1 babysitter dodge": "6",
    "2 babysitter dodge": "6",
    "abonador mike": "4",
    "abris ayrton john": "5",
    "actorani cahuya ii": "4",
    "adela lara": "VIP 1",
    "adelina magbanua": "12",
    "aduna engr vincent": "4",
    "agabon janessa": "11",
    "agabon rogelio": "11",
    "aguihap eve jasmine": "1",
    "aguihap ivy jarmaine": "1",
    "aguihap jenny rose": "1",
    "aguihap john": "1",
    "aguihap justine vnz": "1",
    "aj b dhustine lobrido": "8",
    "alimoot engr marian mi": "5",
    "aliyah calunsag": "10",
    "alyssa blaza": "8",
    "amante analie": "11",
    "andrea dimamay karriel": "3",
    "angelie morales": "11",
    "ann jenny razonales": "12",
    "ann mary sulmaca": "VIP 2",
    "annabelle moscoso": "6",
    "annabelle rojo": "VIP 2",
    "anne buno judy": "5",
    "anthony legaspi mark": "7",
    "anthony romeo tubianosa": "VIP 1",
    "arimas engr leo": "9",
    "arimas gabriel leo": "9",
    "arimas khaeleigh": "9",
    "arimas rutchelie": "9",
    "arlu dominic perlas": "6",
    "austria dimamay goefrey jr": "3",
    "austria dimamay lenie": "3",
    "auxillador avitus dofelis": "6",
    "baluya lorenz": "5",
    "bangcaya dimamay nenita": "1",
    "beth pangan": "12",
    "bill jun sulmaca": "VIP 2",
    "bolneo steline": "13",
    "boyet rojo": "VIP 2",
    "brent engr jeremy rojo": "VIP 2",
    "brian verdadero": "3",
    "brize nathaniel verdadero": "3",
    "calunsag colt": "10",
    "calunsag pinky": "10",
    "capua christopher": "13",
    "carlito lobrido": "9",
    "caspillo donato": "VIP 2",
    "caspillo felicitas": "VIP 2",
    "castillo chavez edmund fr rev very": "VIP 2",
    "ceniza june patrick": "4",
    "cenon vergara": "VIP 2",
    "christiana gimenez marie": "6",
    "christine sulmaca": "10",
    "claire young": "5",
    "condrado iii pangan": "12",
    "cruz dela xian": "11",
    "cumla james tyrone": "6",
    "dainn evangelio izellah": "13",
    "dale rhylvin tinador": "6",
    "dalumpines esther labrague": "VIP 1",
    "dalumpines gerardo": "VIP 1",
    "daniella lobrido": "7",
    "de gerdalym guzman": "2",
    "de grano mae via": "11",
    "de grano verline": "11",
    "de guzman gwen michaela": "2",
    "de guzman michael": "2",
    "de guzman perlita": "11",
    "dee fabilla lysa": "4",
    "desiree legaspi": "7",
    "deviza edwin": "4",
    "dexter john vergara": "8",
    "dianara lobrido": "7",
    "dimamay edna erta": "1",
    "dimamay goefrey": "3",
    "dimamay grace melody": "VIP 1",
    "dimamay jimiswil": "1",
    "dimamay kaiden": "2",
    "dimamay kailah": "3",
    "dimamay karl": "2",
    "dimamay karmila": "2",
    "dimamay keishaun": "2",
    "dimamay kristian": "2",
    "dimamay kristina": "2",
    "dimamay l leo": "VIP 1",
    "dimamay leenie lou": "VIP 1",
    "dimamay lenn leonard": "VIP 1",
    "dimamay leonardo": "VIP 1",
    "dimamay meriel": "3",
    "dimamay nena pacheco": "1",
    "dimamay o": "VIP 1",
    "dimamay oshea": "2",
    "dimamay rosie": "VIP 1",
    "dofelis honey joyce": "6",
    "dolorica manuel": "6",
    "dylan lobrido": "9",
    "edna gindap labrague": "VIP 1",
    "edna morales": "11",
    "edwin lobrido": "7",
    "elma vergara": "VIP 2",
    "elsie labrague moscoso": "VIP 1",
    "elvie razonales": "12",
    "engr fher leongson noah": "4",
    "engr jr manuel sulmaca": "VIP 2",
    "erma sulmaca": "VIP 2",
    "ernesto lobrido": "8",
    "ernie john lobrido": "8",
    "erta stevien": "1",
    "esparagoza shenry": "5",
    "estelito petate": "9",
    "esteron kim": "11",
    "fr jonathan rev sevegan": "VIP 2",
    "gabrielle vergara": "8",
    "gayares rosmarie": "13",
    "gerlie sulmaca": "12",
    "gerry laurent tarroja": "4",
    "gindap wilnah": "3",
    "grace ianne lim": "VIP 1",
    "graceky sulmaca": "10",
    "gualberto javier": "VIP 1",
    "gwendolyn mary tubianosa": "VIP 1",
    "henry sulmaca": "10",
    "hermie sulmaca": "VIP 2",
    "inigo legaspi": "7",
    "isabelle kiara sulmaca": "12",
    "jane publico zan": "4",
    "jean ray sulmaca": "10",
    "jenifer lobrido": "9",
    "jezza villanueva": "5",
    "joana mariel vergara": "10",
    "john morales paul": "12",
    "jomir morales": "12",
    "jonas vergara": "10",
    "joshua kyll sulmaca": "VIP 2",
    "judy verdadero": "3",
    "kross lobrido": "7",
    "lababit meriam yumen": "VIP 1",
    "lababit rhodel": "VIP 1",
    "labanon mirabel": "VIP 1",
    "labrague nena sister": "VIP 1",
    "lobrido rechilda": "7",
    "lobrido samara": "9",
    "magbanua mercedita": "VIP 2",
    "megan veloso": "4",
    "michaela vergara": "8",
    "mirobela morales": "12",
    "natasha vergara": "8",
    "nathaniel vergara": "8",
    "pagarita ray saimon": "7",
    "pagarita salvacion": "7",
    "paul young": "5",
    "petate veronica": "9",
    "renjan solis": "5",
    "rizalina villarta": "13",
    "sulmaca vanjay": "10",
    "tabea vergara": "8"
};
  const sheet = guestsSheet_();
  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(DATA_START, COL.NAME, lastRow - (DATA_START - 1), 1).getValues();
  const tables = sheet.getRange(DATA_START, COL.TABLE, lastRow - (DATA_START - 1), 1).getValues();
  let filled = 0, missed = [];
  for (let i = 0; i < names.length; i++) {
    const raw = String(names[i][0] || '').trim();
    if (!raw) continue;
    const key = nameTokens_(raw).slice().sort().join(' ');
    if (PLAN[key]) { tables[i][0] = PLAN[key]; filled++; }
    else missed.push(raw);
  }
  sheet.getRange(DATA_START, COL.TABLE, tables.length, 1).setValues(tables);
  Logger.log('Seated ' + filled + ' guests. Not in plan (declined or new): ' + missed.join(', '));
}

// ── Group the sheet into invitations ──
// Family (Groom/Bride side) invitations are the blocks of consecutive rows
// between blank rows. "Mutual" guests (relationship not Groom/Bride, e.g. the
// couple's shared friends) are NOT grouped — each is its own single invitation.
function buildGroups_(rows) {
  const groups = [];
  let fam = null;
  const closeFam = () => { if (fam) { groups.push(fam); fam = null; } };
  for (let i = DATA_START - 1; i < rows.length; i++) {   // 0-based; DATA_START-1 = first data row
    const raw = String(rows[i][COL.NAME - 1] || '').trim();
    if (!raw) { closeFam(); continue; }                 // blank row ends a family block
    const m = {
      id: i + 1,                                         // 1-based sheet row
      name: cleanName_(raw),
      tokens: nameTokens_(raw),
      role: String(rows[i][COL.ROLE - 1] || '').trim(),
      response: String(rows[i][COL.RESPONSE - 1] || '').trim(),
    };
    if (sideOf_(rows[i][COL.RELATIONSHIP - 1]) === 'mutual') {
      closeFam();                                        // mutual guest = individual invitation
      groups.push([m]);
    } else {
      if (!fam) fam = [];
      fam.push(m);
    }
  }
  closeFam();
  return groups;
}

// Which side a relationship belongs to (matches the entourage logic).
function sideOf_(relationship) {
  const r = String(relationship || '');
  return /groom/i.test(r) ? 'groom' : /bride/i.test(r) ? 'bride' : 'mutual';
}

// ── Helpers ──
function guestsSheet_() {
  const s = SpreadsheetApp.openById(SHEET_ID).getSheetByName(GUESTS_TAB);
  if (!s) throw new Error('Tab "' + GUESTS_TAB + '" not found.');
  return s;
}

function logSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(LOG_TAB);
  if (!s) {
    s = ss.insertSheet(LOG_TAB);
    s.appendRow(['Timestamp', 'Invitation (head)', 'Mobile', 'Attending', 'Declined', 'Message']);
  }
  return s;
}

// Strip the sheet's trailing name markers (^, *) for display / matching.
function cleanName_(name) {
  return String(name || '').trim().replace(/\s*[\^*]+\s*$/, '');
}

// Normalize a name to its comparable word tokens (title/case/accent/punct-insensitive).
function nameTokens_(name) {
  const s = String(name || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')   // strip accents
    .toLowerCase()
    .replace(/[.,]/g, ' ')
    .replace(/\b(mr|mrs|ms|miss|sir|madam|dr|doctor|engr|atty|rev|fr|hon|prof|sr|jr|iii|iv|ii)\b/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return s ? s.split(' ') : [];
}

// True if every token in `a` also appears in `b`.
function subsetOf_(a, b) {
  return a.length > 0 && a.every(t => b.indexOf(t) !== -1);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
