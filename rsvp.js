// ─────────────────────────────────────────────────────────────
// RSVP → Google Sheets (guest-list matched, per-family)
//
// The guest list is one row per person; invitations are the blocks of
// rows between blank rows. Searching any name pulls up the whole block,
// and one person confirms who's attending. Writes yes/no to col G and
// the mobile/message to the head row's Notes (col H).
//
// Paste the deployed standalone Apps Script Web App URL below.
// ─────────────────────────────────────────────────────────────

// Same Web App as the entourage endpoint — the script now handles both.
const RSVP_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyQlKkVbL9LLWWtE4l1jPspYy4kftYR46H5QsnKaEpg3imBotvS-CNAAghHkEdUqnnL/exec';

// Shown in the strict "we couldn't find your name" message.
// TODO: replace with the real coordinator name / mobile number.
const COORDINATOR_CONTACT = 'our wedding coordinator';

// PH mobile: 09xxxxxxxxx or +639xxxxxxxxx (spaces/dashes tolerated).
const PH_MOBILE_RE = /^(09\d{9}|\+639\d{9})$/;

// RSVP closes end of this day (PH time). Past it, the form is replaced with a notice.
const RSVP_DEADLINE = new Date('2026-11-08T23:59:59+08:00');
const RSVP_DEADLINE_LABEL = 'November 8, 2026';

let currentGroup = null;   // { headId, members: [{ id, name, role, response }] }

document.addEventListener('DOMContentLoaded', () => {
  const findBtn   = document.getElementById('rsvp-find');
  const nameInput = document.getElementById('rsvp-name');
  const submitBtn = document.getElementById('rsvp-submit');
  const backBtn   = document.getElementById('rsvp-back');
  if (!findBtn) return;   // RSVP not on this page

  if (Date.now() > RSVP_DEADLINE.getTime()) { showClosed(); return; }   // deadline passed

  findBtn.addEventListener('click', lookupInvitation);
  nameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); lookupInvitation(); }
  });
  submitBtn.addEventListener('click', submitRsvp);
  if (backBtn) backBtn.addEventListener('click', backToSearch);
});

function configured() {
  return RSVP_SCRIPT_URL && RSVP_SCRIPT_URL !== 'YOUR_RSVP_APPS_SCRIPT_WEB_APP_URL';
}

// Replace the form with a "closed" notice once the deadline has passed.
function showClosed() {
  const closed = document.getElementById('rsvp-closed');
  const step1 = document.getElementById('rsvp-step1');
  const step2 = document.getElementById('rsvp-step2');
  const note = document.getElementById('rsvp-deadline');
  if (step1) step1.hidden = true;
  if (step2) step2.hidden = true;
  if (note) note.textContent = 'RSVPs closed on ' + RSVP_DEADLINE_LABEL + '.';
  if (closed) {
    const body = closed.querySelector('.rsvp-closed-body');
    if (body) {
      body.textContent = 'The deadline to respond has passed. If you still need to reach us ' +
        'about your attendance, please contact ' + COORDINATOR_CONTACT + '.';
    }
    closed.hidden = false;
  }
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
      const groups = (res && res.groups) || [];
      if (groups.length === 0) {
        showStatus(statusEl,
          'We couldn’t find that name on our guest list. Please check the spelling, ' +
          'or reach out to ' + COORDINATOR_CONTACT + '.', 'error');
      } else if (groups.length === 1) {
        selectGroup(groups[0]);
      } else {
        renderChooser(groups);
      }
    })
    .catch(() => showStatus(statusEl,
      'Something went wrong looking up your invitation. Please try again.', 'error'))
    .finally(() => { findBtn.disabled = false; findBtn.textContent = original; });
}

function renderChooser(groups) {
  const chooser = document.getElementById('rsvp-chooser');
  const label = (g) => {
    const first = g.members[0] ? g.members[0].name : 'Invitation';
    const more = g.members.length - 1;
    return more > 0 ? (first + ' + ' + more + ' more') : first;
  };
  chooser.innerHTML =
    '<p class="rsvp-chooser-label">We found a few invitations — which one is yours?</p>' +
    groups.map((g, i) =>
      '<label class="rsvp-chooser-opt"><input type="radio" name="rsvp-group" value="' + i + '"> ' +
      escapeHtml(label(g)) + '</label>').join('');
  chooser.hidden = false;
  chooser.querySelectorAll('input[name="rsvp-group"]').forEach((el, i) => {
    el.addEventListener('change', () => selectGroup(groups[i]));
  });
}

function selectGroup(g) {
  currentGroup = g;

  document.getElementById('rsvp-step1').hidden = true;
  document.getElementById('rsvp-step2').hidden = false;
  showStatus(document.getElementById('rsvp-status'), '', '');

  const n = g.members.length;
  const single = n === 1;
  const welcome = document.getElementById('rsvp-welcome');
  const label = document.getElementById('rsvp-attend-label');
  const list = document.getElementById('rsvp-members');

  const alreadyResponded = g.members.some((m) => m.response === 'yes' || m.response === 'no');
  document.getElementById('rsvp-update-note').hidden = !alreadyResponded;

  if (single) {
    const m = g.members[0];
    welcome.innerHTML = 'Welcome, <strong>' + escapeHtml(m.name) + '</strong>! ' +
      'Please let us know if you can make it.';
    label.textContent = 'Will You Attend?';
    const declined = m.response === 'no';
    list.innerHTML =
      '<select class="rsvp-single-attend" data-id="' + m.id + '">' +
        '<option value="yes"' + (declined ? '' : ' selected') + '>Joyfully Accept</option>' +
        '<option value="no"' + (declined ? ' selected' : '') + '>Respectfully Decline</option>' +
      '</select>';
  } else {
    welcome.innerHTML = 'Your invitation includes <strong>' + n + '</strong> guests. ' +
      'Please tick who will be joining us.';
    label.textContent = "Who's Attending?";
    list.innerHTML = g.members.map((m) => {
      const checked = m.response === 'no' ? '' : 'checked';   // default attending
      const role = m.role ? '<span class="rsvp-mem-role">' + escapeHtml(m.role) + '</span>' : '';
      return '<label class="rsvp-member">' +
        '<input type="checkbox" class="rsvp-mem-check" data-id="' + m.id + '" ' + checked + '>' +
        '<span class="rsvp-mem-name">' + escapeHtml(m.name) + role + '</span>' +
        '<span class="rsvp-mem-state"></span>' +
      '</label>';
    }).join('');
  }

  document.getElementById('rsvp-mobile').value = '';
  document.getElementById('rsvp-message').value = '';
}

function submitRsvp() {
  const statusEl = document.getElementById('rsvp-status');
  const submitBtn = document.getElementById('rsvp-submit');
  if (!currentGroup) { showStatus(statusEl, 'Please find your invitation first.', 'error'); return; }

  const checks = Array.from(document.querySelectorAll('.rsvp-mem-check'));
  let m, attending;
  if (checks.length) {                       // multi-person: per-member checkboxes
    m = checks.map((c) => c.dataset.id + ':' + (c.checked ? 'yes' : 'no')).join(';');
    attending = checks.filter((c) => c.checked).length;
  } else {                                    // single-person: accept/decline select
    const sel = document.querySelector('.rsvp-single-attend');
    if (!sel) { showStatus(statusEl, 'Please find your invitation first.', 'error'); return; }
    m = sel.dataset.id + ':' + sel.value;
    attending = sel.value === 'yes' ? 1 : 0;
  }

  const mobileRaw = document.getElementById('rsvp-mobile').value.trim();
  const mobile = mobileRaw.replace(/[\s-]/g, '');
  const message = document.getElementById('rsvp-message').value.trim();

  if (!PH_MOBILE_RE.test(mobile)) {
    showStatus(statusEl, 'Please enter a valid PH mobile number (e.g. 0917 123 4567).', 'error');
    document.getElementById('rsvp-mobile').focus();
    return;
  }

  const original = submitBtn.textContent;
  submitBtn.disabled = true; submitBtn.textContent = 'Sending…';
  showStatus(statusEl, '', '');

  const params = new URLSearchParams({
    rsvp: 'submit',
    head: currentGroup.headId,
    m: m,
    mobile: mobile,
    message: message
  });

  fetch(RSVP_SCRIPT_URL + '?' + params.toString())
    .then((r) => r.json())
    .then((res) => {
      if (res && res.status === 'success') {
        const msg = attending === 0
          ? 'Thank you for letting us know — you will be missed! 💜'
          : 'Thank you! We’ve got ' + attending + ' attending — we can’t wait to celebrate with you! 💜';
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
  currentGroup = null;
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
