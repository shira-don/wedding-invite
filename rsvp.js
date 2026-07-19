// ─────────────────────────────────────────────────────────────
// RSVP → Google Sheets
//
// Replace the URL below with your deployed Google Apps Script
// Web App URL. See google-apps-script.js for setup instructions.
// ─────────────────────────────────────────────────────────────

const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL';

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function handleRsvp(e) {
  e.preventDefault();

  const form = e.target;
  const btn = form.querySelector('.btn');
  const statusEl = document.getElementById('rsvp-status');

  // Collect form data
  const data = {
    name: form.querySelector('#name').value.trim(),
    email: form.querySelector('#email').value.trim(),
    guests: form.querySelector('#guests').value,
    attendance: form.querySelector('#attendance').value,
    dietary: form.querySelector('#dietary').value.trim(),
    message: form.querySelector('#message').value.trim(),
    timestamp: new Date().toISOString(),
  };

  // Validate
  if (!data.name) {
    showStatus(statusEl, 'Please tell us your name.', 'error');
    form.querySelector('#name').focus();
    return;
  }
  if (!EMAIL_RE.test(data.email)) {
    showStatus(statusEl, 'Please enter a valid email address.', 'error');
    form.querySelector('#email').focus();
    return;
  }

  // Guard: endpoint not configured yet
  if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL === 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL') {
    showStatus(statusEl, 'RSVP is not set up yet. Please contact the couple directly.', 'error');
    console.warn('[RSVP] GOOGLE_SCRIPT_URL is not configured. See google-apps-script.js for setup.');
    return;
  }

  // Disable button & show loading
  const originalLabel = btn.textContent;
  btn.disabled = true;
  btn.textContent = 'Sending…';
  statusEl.className = 'rsvp-status';
  statusEl.textContent = '';

  fetch(GOOGLE_SCRIPT_URL, {
    method: 'POST',
    mode: 'no-cors', // Apps Script doesn't send CORS headers
    // text/plain is a "simple" request → no CORS preflight, which
    // Apps Script can't answer. doPost still JSON.parse()s the body.
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(data),
  })
    .then(() => {
      // no-cors means the response is opaque and can't be read, so a
      // resolved promise only confirms the request was dispatched.
      const msg = data.attendance === 'no'
        ? 'Thank you for letting us know — you will be missed! 💜'
        : 'Thank you! Your RSVP has been received. We can\'t wait to celebrate with you! 💜';
      showStatus(statusEl, msg, 'success');
      btn.textContent = 'Sent!';
      btn.style.background = '#7a9b78';
      form.reset();
    })
    .catch(() => {
      showStatus(statusEl, 'Something went wrong sending your RSVP. Please try again, or message the couple directly.', 'error');
      btn.textContent = originalLabel;
      btn.disabled = false;
    });
}

function showStatus(el, message, type) {
  el.textContent = message;
  el.className = 'rsvp-status ' + type;
}
