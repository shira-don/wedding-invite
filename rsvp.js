// ─────────────────────────────────────────────────────────────
// RSVP → Google Sheets (guest-list matched, two-step)
//
// Paste the deployed **standalone** Apps Script Web App URL below.
// See google-apps-script.js for setup instructions.
// ─────────────────────────────────────────────────────────────

const RSVP_SCRIPT_URL = 'YOUR_RSVP_APPS_SCRIPT_WEB_APP_URL';

// Shown in the strict "we couldn't find your name" message.
// TODO: replace with the real coordinator name / mobile number.
const COORDINATOR_CONTACT = 'our wedding coordinator';

// PH mobile: 09xxxxxxxxx or +639xxxxxxxxx (spaces/dashes tolerated).
const PH_MOBILE_RE = /^(09\d{9}|\+639\d{9})$/;

let currentParty = null;   // { id, party, seats, responded, status, count, mobile, message }

document.addEventListener('DOMContentLoaded', () => {
  const findBtn   = document.getElementById('rsvp-find');
  const nameInput = document.getElementById('rsvp-name');
  const attendSel = document.getElementById('rsvp-attendance');
  const submitBtn = document.getElementById('rsvp-submit');
  const backBtn   = document.getElementById('rsvp-back');
  if (!findBtn) return;   // RSVP not on this page

  findBtn.addEventListener('click', lookupInvitation);
  nameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); lookupInvitation(); }
  });
  attendSel.addEventListener('change', updateCountVisibility);
  submitBtn.addEventListener('click', submitRsvp);
  if (backBtn) backBtn.addEventListener('click', backToSearch);
});

function configured() {
  return RSVP_SCRIPT_URL && RSVP_SCRIPT_URL !== 'YOUR_RSVP_APPS_SCRIPT_WEB_APP_URL';
}

function lookupInvitation() {
  const statusEl = document.getElementById('rsvp-status');
  const chooser  = document.getElementById('rsvp-chooser');
  const findBtn  = document.getElementById('rsvp-find');
  const name = document.getElementById('rsvp-name').value.trim();

  chooser.hidden = true;
  chooser.innerHTML = '';
  if (!name) { showStatus(statusEl, 'Please enter your full name.', 'error'); return; }
  if (!configured()) {
    showStatus(statusEl, 'RSVP is not set up yet. Please contact ' + COORDINATOR_CONTACT + '.', 'error');
    return;
  }

  const original = findBtn.textContent;
  findBtn.disabled = true; findBtn.textContent = 'Searching…';
  showStatus(statusEl, '', '');

  fetch(RSVP_SCRIPT_URL + '?rsvp=lookup&name=' + encodeURIComponent(name))
    .then((r) => r.json())
    .then((res) => {
      const matches = (res && res.matches) || [];
      if (matches.length === 0) {
        showStatus(statusEl,
          'We couldn’t find that name on our guest list. Please check the spelling, ' +
          'or reach out to ' + COORDINATOR_CONTACT + '.', 'error');
      } else if (matches.length === 1) {
        selectParty(matches[0]);
      } else {
        renderChooser(matches);
      }
    })
    .catch(() => showStatus(statusEl,
      'Something went wrong looking up your invitation. Please try again.', 'error'))
    .finally(() => { findBtn.disabled = false; findBtn.textContent = original; });
}

function renderChooser(matches) {
  const chooser = document.getElementById('rsvp-chooser');
  chooser.innerHTML =
    '<p class="rsvp-chooser-label">We found a few matches — which one is you?</p>' +
    matches.map((m, i) =>
      '<label class="rsvp-chooser-opt"><input type="radio" name="rsvp-party" value="' + i + '"> ' +
      escapeHtml(m.party) + '</label>').join('');
  chooser.hidden = false;
  chooser.querySelectorAll('input[name="rsvp-party"]').forEach((el, i) => {
    el.addEventListener('change', () => selectParty(matches[i]));
  });
}

function selectParty(m) {
  currentParty = m;

  document.getElementById('rsvp-step1').hidden = true;
  document.getElementById('rsvp-step2').hidden = false;
  showStatus(document.getElementById('rsvp-status'), '', '');

  document.getElementById('rsvp-welcome').innerHTML =
    'Welcome, <strong>' + escapeHtml(m.party) + '</strong>! ' +
    'You have up to <strong>' + m.seats + '</strong> seat' + (m.seats > 1 ? 's' : '') + '.';

  // Build the count options 1..seats
  const countSel = document.getElementById('rsvp-count');
  countSel.innerHTML = '';
  for (let n = 1; n <= m.seats; n++) {
    countSel.insertAdjacentHTML('beforeend', '<option value="' + n + '">' + n + '</option>');
  }

  // Prefill if they already responded (updates are allowed)
  const attendSel = document.getElementById('rsvp-attendance');
  const noteEl = document.getElementById('rsvp-update-note');
  if (m.responded) {
    attendSel.value = m.status === 'Declined' ? 'no' : 'yes';
    countSel.value = String(Math.min(m.count || 1, m.seats) || 1);
    document.getElementById('rsvp-mobile').value = m.mobile || '';
    document.getElementById('rsvp-message').value = m.message || '';
    noteEl.hidden = false;
  } else {
    attendSel.value = 'yes';
    countSel.value = String(m.seats);
    document.getElementById('rsvp-mobile').value = '';
    document.getElementById('rsvp-message').value = '';
    noteEl.hidden = true;
  }
  updateCountVisibility();
}

function updateCountVisibility() {
  const attending = document.getElementById('rsvp-attendance').value !== 'no';
  const seats = currentParty ? currentParty.seats : 1;
  // Hide the count picker when declining, or when only a single seat is allotted.
  document.getElementById('rsvp-count-group').hidden = !attending || seats <= 1;
}

function submitRsvp() {
  const statusEl = document.getElementById('rsvp-status');
  const submitBtn = document.getElementById('rsvp-submit');
  if (!currentParty) { showStatus(statusEl, 'Please find your invitation first.', 'error'); return; }

  const attending = document.getElementById('rsvp-attendance').value;   // 'yes' | 'no'
  const count = attending === 'no' ? 0 : Number(document.getElementById('rsvp-count').value || 1);
  const mobileRaw = document.getElementById('rsvp-mobile').value.trim();
  const mobile = mobileRaw.replace(/[\s-]/g, '');
  const message = document.getElementById('rsvp-message').value.trim();

  if (!PH_MOBILE_RE.test(mobile)) {
    showStatus(statusEl, 'Please enter a valid PH mobile number (e.g. 0917 123 4567).', 'error');
    document.getElementById('rsvp-mobile').focus();
    return;
  }
  if (attending === 'yes' && count > currentParty.seats) {
    showStatus(statusEl, 'That is more than your allotted seats (' + currentParty.seats + ').', 'error');
    return;
  }

  const original = submitBtn.textContent;
  submitBtn.disabled = true; submitBtn.textContent = 'Sending…';
  showStatus(statusEl, '', '');

  const params = new URLSearchParams({
    rsvp: 'submit',
    id: currentParty.id,
    party: currentParty.party,
    attending: attending,
    count: String(count),
    mobile: mobile,
    message: message
  });

  fetch(RSVP_SCRIPT_URL + '?' + params.toString())
    .then((r) => r.json())
    .then((res) => {
      if (res && res.status === 'success') {
        const msg = attending === 'no'
          ? 'Thank you for letting us know — you will be missed! 💜'
          : 'Thank you! Your RSVP is confirmed for ' + count + ' — we can’t wait to celebrate with you! 💜';
        showStatus(statusEl, msg, 'success');
        submitBtn.textContent = 'Sent!';
        submitBtn.style.background = '#7a9b78';
      } else {
        showStatus(statusEl, (res && res.message) || 'Something went wrong. Please try again.', 'error');
        submitBtn.disabled = false; submitBtn.textContent = original;
      }
    })
    .catch(() => {
      showStatus(statusEl, 'Something went wrong sending your RSVP. Please try again.', 'error');
      submitBtn.disabled = false; submitBtn.textContent = original;
    });
}

function backToSearch() {
  currentParty = null;
  document.getElementById('rsvp-step2').hidden = true;
  document.getElementById('rsvp-step1').hidden = false;
  const submitBtn = document.getElementById('rsvp-submit');
  submitBtn.disabled = false; submitBtn.textContent = 'Send RSVP'; submitBtn.style.background = '';
  showStatus(document.getElementById('rsvp-status'), '', '');
}

function showStatus(el, message, type) {
  if (!el) return;
  el.textContent = message;
  el.className = 'rsvp-status' + (type ? ' ' + type : '');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
