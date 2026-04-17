/**
 * taskpane.js — ECC Lead Quick-Add task pane controller
 *
 * Checkpoint 6: skeleton only — initialises Office, reads mailbox context,
 * runs prefill lookups, then hands off to renderForm() (Checkpoint 7).
 *
 * State machine:
 *   loading → app          (happy path)
 *   loading → auth-error   (SSO/provisioning failure)
 *   loading → error        (schema/network failure)
 *   app     → success      (after form submit)
 *   success → app          ("Create another" resets form)
 */

/* global Office, Auth, API */

// ── DOM refs ─────────────────────────────────────────────────────────────────

const Views = {
  loading:    document.getElementById('loading'),
  authError:  document.getElementById('auth-error'),
  error:      document.getElementById('error'),
  app:        document.getElementById('app'),
  success:    document.getElementById('success'),
};

const El = {
  authErrorMsg:     document.getElementById('auth-error-message'),
  authRetryBtn:     document.getElementById('auth-retry-btn'),
  authFallbackBtn:  document.getElementById('auth-fallback-btn'),
  errorTitle:       document.getElementById('error-title'),
  errorMsg:         document.getElementById('error-message'),
  errorRetryBtn:    document.getElementById('error-retry-btn'),
  prefillBanner:    document.getElementById('prefill-banner'),
  prefillCompany:   document.getElementById('prefill-company-name'),
  prefillContact:   document.getElementById('prefill-contact-name'),
  prefillUseBtn:    document.getElementById('prefill-use-btn'),
  prefillNewBtn:    document.getElementById('prefill-new-btn'),
  formSections:     document.getElementById('form-sections'),
  submitBtn:        document.getElementById('submit-btn'),
  successCompany:   document.getElementById('success-company-name'),
  viewLeadBtn:      document.getElementById('view-lead-btn'),
  createAnotherBtn: document.getElementById('create-another-btn'),
};

// ── View helpers ─────────────────────────────────────────────────────────────

function showView(name) {
  Object.entries(Views).forEach(([key, el]) => {
    el.classList.toggle('hidden', key !== name);
  });
}

// ── Mailbox context ───────────────────────────────────────────────────────────

function getMailboxContext() {
  const item = Office.context.mailbox.item;
  if (!item) return {};

  return {
    senderEmail:     item.from?.emailAddress ?? null,
    senderName:      item.from?.displayName  ?? null,
    senderDomain:    item.from?.emailAddress?.split('@')[1] ?? null,
    emailSubject:    item.subject ?? null,
    emailReceivedAt: item.dateTimeCreated?.toISOString() ?? null,
    mailboxAddress:  Office.context.mailbox.userProfile?.emailAddress ?? null,
  };
}

// ── Main init ─────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  El.authRetryBtn.addEventListener('click', () => {
    Auth.clearToken();
    boot();
  });

  El.authFallbackBtn.addEventListener('click', async () => {
    try {
      Auth.clearToken();
      await Auth.getToken();
      boot();
    } catch (err) {
      El.authErrorMsg.textContent = err.message;
    }
  });

  El.errorRetryBtn.addEventListener('click', boot);

  El.createAnotherBtn.addEventListener('click', () => {
    showView('app');
    El.formSections.innerHTML = '';
    El.submitBtn.disabled = true;
    boot();
  });

  boot();
});

async function boot() {
  showView('loading');

  // 1. Acquire token
  let token;
  try {
    token = await Auth.getToken();
  } catch (err) {
    El.authErrorMsg.textContent = err.message ?? 'Could not sign you in.';
    showView('authError');
    return;
  }

  // 2. Read mailbox context
  const ctx = getMailboxContext();

  // 3. Load form schema + run prefill lookups in parallel
  let schema, companyMatch, contactMatch;
  try {
    [{ schema }, companyMatch, contactMatch] = await Promise.all([
      API.getFormSchema('lead_quick_add', ctx),
      ctx.senderDomain
        ? API.matchCompany({ domain: ctx.senderDomain })
        : Promise.resolve({ matches: [] }),
      ctx.senderEmail
        ? API.matchContact(ctx.senderEmail)
        : Promise.resolve({ match: null }),
    ]);
  } catch (err) {
    if (err instanceof API.ApiError && err.isNotProvisioned) {
      El.authErrorMsg.textContent = err.message;
      showView('authError');
      return;
    }
    El.errorTitle.textContent = 'Could not load form';
    El.errorMsg.textContent   = err.message ?? 'Check your internet connection and try again.';
    showView('error');
    return;
  }

  // 4. Show prefill banner if we matched anything
  const matchedCompany = companyMatch.matches?.[0] ?? null;
  const matchedContact = contactMatch.match ?? null;

  if (matchedCompany || matchedContact) {
    El.prefillCompany.textContent = matchedCompany?.name ?? '';
    El.prefillContact.textContent = matchedContact
      ? `· ${matchedContact.name}`
      : '';
    El.prefillBanner.classList.remove('hidden');
  }

  // 5. Render form — Checkpoint 7 fills this in
  renderForm(schema, {
    ctx,
    matchedCompany,
    matchedContact,
    useMatch: true,
  });

  showView('app');
}

// ── Form rendering (stub — fully implemented in Checkpoint 7) ──────────────

function renderForm(schema, { ctx, matchedCompany, matchedContact, useMatch }) {
  // Checkpoint 7: build DOM from schema.fields grouped by section
  // For now, show a placeholder so the pane is not blank during testing
  El.formSections.innerHTML = `
    <div class="form-placeholder">
      <p>Form schema loaded (${schema?.fields?.length ?? 0} fields).</p>
      <p>Sender: ${ctx.senderEmail ?? '—'}</p>
      ${matchedCompany ? `<p>Matched company: <strong>${matchedCompany.name}</strong></p>` : ''}
      ${matchedContact ? `<p>Matched contact: <strong>${matchedContact.name}</strong></p>` : ''}
      <p class="form-placeholder__note">Full form renders in Checkpoint 7.</p>
    </div>
  `;
  // Enable submit so Checkpoint 7 wiring can be tested
  El.submitBtn.disabled = false;
}

// ── Form submit ───────────────────────────────────────────────────────────────

document.getElementById('lead-form').addEventListener('submit', async e => {
  e.preventDefault();

  El.submitBtn.disabled = true;
  El.submitBtn.textContent = 'Adding…';

  try {
    // Checkpoint 7 replaces this with real payload collection
    const result = await API.createLead({
      form_version: 1,
      company_data: {},
      contact_data: {},
      lead_data: {},
    });

    El.successCompany.textContent = `Lead saved. View: ${result.hub_url}`;
    El.viewLeadBtn.onclick = () => window.open(result.hub_url, '_blank');
    showView('success');
  } catch (err) {
    if (err instanceof API.ApiError) {
      if (err.isSchemaOutdated) {
        El.errorTitle.textContent = 'Form is outdated';
        El.errorMsg.textContent   = 'Reload the add-in to get the latest form.';
      } else if (err.isDuplicate) {
        El.errorTitle.textContent = 'Already submitted';
        El.errorMsg.textContent   = 'This email was already added as a lead recently.';
      } else {
        El.errorTitle.textContent = 'Failed to save lead';
        El.errorMsg.textContent   = err.message;
      }
    } else {
      El.errorTitle.textContent = 'Unexpected error';
      El.errorMsg.textContent   = String(err);
    }
    showView('error');
  } finally {
    El.submitBtn.disabled    = false;
    El.submitBtn.textContent = 'Add Lead';
  }
});
