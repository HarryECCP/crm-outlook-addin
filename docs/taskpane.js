/**
 * taskpane.js — ECC Lead Quick-Add task pane controller
 *
 * State machine:
 *   loading → app          happy path
 *   loading → auth-error   SSO / provisioning failure
 *   loading → error        schema / network failure
 *   app → success          after submit
 *   success → app          "Create another"
 */

/* global Office, Auth, API */

// ── Pane state ────────────────────────────────────────────────────────────────

let _schema          = null;
let _users           = [];
let _selectedCoId    = null;   // UUID of existing company (if matched/selected)
let _selectedCtId    = null;   // UUID of existing contact (if matched/selected)
let _hubUrl          = null;
let _emailMessageId  = null;
let _ctx             = {};

// ── DOM refs ─────────────────────────────────────────────────────────────────

const Views = {
  loading:   document.getElementById('loading'),
  authError: document.getElementById('auth-error'),
  error:     document.getElementById('error'),
  app:       document.getElementById('app'),
  success:   document.getElementById('success'),
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

function showView(name) {
  Object.entries(Views).forEach(([key, el]) =>
    el.classList.toggle('hidden', key !== name)
  );
}

// ── Mailbox context ───────────────────────────────────────────────────────────

function getMailboxContext() {
  const item = Office.context?.mailbox?.item;
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

function getEmailMessageId() {
  return Office.context?.mailbox?.item?.itemId ?? null;
}

// ── Boot ──────────────────────────────────────────────────────────────────────

Office.onReady(async () => {
  El.authRetryBtn.addEventListener('click',     () => { Auth.clearToken(); boot(); });
  El.authFallbackBtn.addEventListener('click',  async () => {
    try { Auth.clearToken(); await Auth.getToken(); boot(); }
    catch (err) { El.authErrorMsg.textContent = err.message; }
  });
  El.errorRetryBtn.addEventListener('click', boot);
  El.createAnotherBtn.addEventListener('click', () => {
    _selectedCoId = null;
    _selectedCtId = null;
    boot();
  });

  document.getElementById('lead-form').addEventListener('submit', handleSubmit);

  boot();
});

async function boot() {
  showView('loading');
  _schema = null;
  _users  = [];

  try {
    await Auth.getToken();
  } catch (err) {
    El.authErrorMsg.textContent = err.message ?? 'Could not sign you in.';
    showView('authError');
    return;
  }

  _ctx           = getMailboxContext();
  _emailMessageId = getEmailMessageId();

  let schema, users, companyMatches, contactMatch;
  try {
    [{ schema }, { users }, companyMatches, contactMatch] = await Promise.all([
      API.getFormSchema('lead_quick_add', _ctx),
      API.getUsers(),
      _ctx.senderDomain
        ? API.matchCompany({ domain: _ctx.senderDomain })
        : Promise.resolve({ matches: [] }),
      _ctx.senderEmail
        ? API.matchContact(_ctx.senderEmail)
        : Promise.resolve({ match: null }),
    ]);
  } catch (err) {
    if (err instanceof API.ApiError && err.isNotProvisioned) {
      El.authErrorMsg.textContent = err.message;
      showView('authError');
      return;
    }
    El.errorTitle.textContent = 'Could not load form';
    El.errorMsg.textContent   = err.message ?? 'Check your connection and try again.';
    showView('error');
    return;
  }

  _schema = schema;
  _users  = users ?? [];

  const matchedCompany = (companyMatches.matches ?? [])[0] ?? null;
  const matchedContact = contactMatch.match ?? null;

  // Auto-select existing records (can be overridden via prefill banner)
  _selectedCoId = matchedCompany?.id ?? null;
  _selectedCtId = matchedContact?.id  ?? null;

  // Prefill banner
  if (matchedCompany || matchedContact) {
    El.prefillCompany.textContent = matchedCompany?.name ?? '';
    El.prefillContact.textContent = matchedContact ? `· ${matchedContact.name}` : '';
    El.prefillBanner.classList.remove('hidden');

    El.prefillUseBtn.onclick = () => {
      _selectedCoId = matchedCompany?.id ?? null;
      _selectedCtId = matchedContact?.id  ?? null;
      El.prefillBanner.classList.add('hidden');
      renderForm(schema, _users, { ctx: _ctx, matchedCompany, matchedContact, useMatch: true });
    };
    El.prefillNewBtn.onclick = () => {
      _selectedCoId = null;
      _selectedCtId = null;
      El.prefillBanner.classList.add('hidden');
      renderForm(schema, _users, { ctx: _ctx, matchedCompany: null, matchedContact: null, useMatch: false });
    };
  } else {
    El.prefillBanner.classList.add('hidden');
  }

  renderForm(schema, _users, { ctx: _ctx, matchedCompany, matchedContact, useMatch: true });
  showView('app');
}

// ── Form rendering ────────────────────────────────────────────────────────────

function renderForm(schema, users, { ctx, matchedCompany, matchedContact, useMatch }) {
  El.formSections.innerHTML = '';

  const grouped = groupBySection(schema.fields);

  for (const [sectionKey, fields] of Object.entries(grouped)) {
    const sectionMeta = schema.sections_config?.[sectionKey] ?? {};
    const label       = sectionMeta.label ?? titleCase(sectionKey);
    const collapsed   = sectionMeta.collapsed_by_default ?? false;

    const sectionEl = buildSection(sectionKey, label, collapsed);
    const body      = sectionEl.querySelector('.form-section__body');

    for (const field of fields) {
      if (!field.is_visible) {
        // Hidden field — store value in a real hidden input so collectPayload() finds it
        const hiddenEl = buildHiddenInput(field);
        body.appendChild(hiddenEl);
        continue;
      }

      const prefillValue = getPrefillValue(field, { ctx, matchedCompany, matchedContact, useMatch });
      const fieldEl      = buildFieldRow(field, users, prefillValue);
      body.appendChild(fieldEl);
    }

    El.formSections.appendChild(sectionEl);
  }

  // Wire up live dedup + assignment after DOM is built
  wireEmailDedup();
  wireWebsiteDedup();
  wireStateAssignment();

  updateSubmitState();
}

// ── Section accordion ─────────────────────────────────────────────────────────

function buildSection(key, label, collapsed) {
  const div = document.createElement('div');
  div.className = 'form-section';
  div.dataset.section = key;

  const header = document.createElement('div');
  header.className = 'form-section__header';
  header.innerHTML = `
    <span class="form-section__title">${escHtml(label)}</span>
    <span class="form-section__toggle">${collapsed ? '▸' : '▾'}</span>
  `;

  const body = document.createElement('div');
  body.className = `form-section__body${collapsed ? ' collapsed' : ''}`;

  header.addEventListener('click', () => {
    const isCollapsed = body.classList.toggle('collapsed');
    header.querySelector('.form-section__toggle').textContent = isCollapsed ? '▸' : '▾';
  });

  div.appendChild(header);
  div.appendChild(body);
  return div;
}

// ── Field row ─────────────────────────────────────────────────────────────────

function buildFieldRow(field, users, prefillValue) {
  const wrap = document.createElement('div');
  wrap.className = 'form-field';
  wrap.dataset.table  = field.table_name;
  wrap.dataset.column = field.column_name;

  const labelEl = document.createElement('label');
  labelEl.className = `form-field__label${field.is_required ? ' form-field__label--required' : ''}`;
  labelEl.textContent = field.label ?? titleCase(field.column_name);
  labelEl.htmlFor     = fieldId(field);

  const widget = buildWidget(field, users, prefillValue);
  widget.id = fieldId(field);

  const errorEl = document.createElement('span');
  errorEl.className = 'form-field__error';
  errorEl.id = `${fieldId(field)}-error`;

  wrap.appendChild(labelEl);
  wrap.appendChild(widget);

  if (field.help_text) {
    const help = document.createElement('span');
    help.className   = 'form-field__help';
    help.textContent = field.help_text;
    wrap.appendChild(help);
  }

  wrap.appendChild(errorEl);

  // Validation on change
  widget.addEventListener('change', () => {
    validateField(field, widget, errorEl);
    updateSubmitState();
  });
  widget.addEventListener('input', () => updateSubmitState());

  return wrap;
}

function buildHiddenInput(field) {
  const input = document.createElement('input');
  input.type           = 'hidden';
  input.dataset.table  = field.table_name;
  input.dataset.column = field.column_name;
  input.value          = field.default_value ?? '';
  return input;
}

// ── Widget factory ────────────────────────────────────────────────────────────

function buildWidget(field, users, prefillValue) {
  const hint = field.widget_hint ?? inferWidget(field);

  switch (hint) {
    case 'textarea':      return buildTextarea(field, prefillValue);
    case 'select':        return buildSelect(field, prefillValue);
    case 'toggle':        return buildToggle(field, prefillValue);
    case 'users_lookup':  return buildUsersLookup(field, users, prefillValue);
    case 'lookup_or_text':return buildLookupOrText(field, prefillValue);
    case 'email':         return buildTextInput(field, prefillValue, 'email');
    case 'phone':         return buildTextInput(field, prefillValue, 'tel');
    default:              return buildTextInput(field, prefillValue, 'text');
  }
}

function buildTextInput(field, value, type = 'text') {
  const input        = document.createElement('input');
  input.type         = type;
  input.className    = 'form-field__input';
  input.dataset.table  = field.table_name;
  input.dataset.column = field.column_name;
  input.placeholder  = field.placeholder ?? '';
  input.value        = value ?? '';
  if (field.is_required) input.required = true;
  return input;
}

function buildTextarea(field, value) {
  const ta         = document.createElement('textarea');
  ta.className     = 'form-field__textarea';
  ta.dataset.table  = field.table_name;
  ta.dataset.column = field.column_name;
  ta.placeholder   = field.placeholder ?? '';
  ta.value         = value ?? '';
  if (field.is_required) ta.required = true;
  return ta;
}

function buildSelect(field, value) {
  const sel         = document.createElement('select');
  sel.className     = 'form-field__select';
  sel.dataset.table  = field.table_name;
  sel.dataset.column = field.column_name;
  if (field.is_required) sel.required = true;

  const blank = document.createElement('option');
  blank.value   = '';
  blank.textContent = field.placeholder ?? '— select —';
  sel.appendChild(blank);

  const opts = field.options ?? [];
  for (const opt of opts) {
    const o    = document.createElement('option');
    o.value    = opt.value;
    o.textContent = opt.label ?? opt.value;
    if (value !== null && value !== undefined && String(value) === String(opt.value)) {
      o.selected = true;
    }
    sel.appendChild(o);
  }

  return sel;
}

function buildToggle(field, value) {
  const wrap = document.createElement('div');
  wrap.className = 'form-field__toggle-wrap';
  wrap.dataset.table  = field.table_name;
  wrap.dataset.column = field.column_name;

  const chk         = document.createElement('input');
  chk.type          = 'checkbox';
  chk.className     = 'form-field__toggle';
  chk.dataset.table  = field.table_name;
  chk.dataset.column = field.column_name;
  chk.checked       = value === true || value === 'true';

  const lbl = document.createElement('span');
  lbl.textContent = chk.checked ? 'On' : 'Off';
  chk.addEventListener('change', () => { lbl.textContent = chk.checked ? 'On' : 'Off'; });

  wrap.appendChild(chk);
  wrap.appendChild(lbl);
  return wrap;
}

function buildUsersLookup(field, users, value) {
  const sel         = document.createElement('select');
  sel.className     = 'form-field__select';
  sel.dataset.table  = field.table_name;
  sel.dataset.column = field.column_name;
  if (field.is_required) sel.required = true;

  const blank = document.createElement('option');
  blank.value       = '';
  blank.textContent = '— select CP —';
  sel.appendChild(blank);

  for (const u of users) {
    const o    = document.createElement('option');
    o.value    = u.display_name;   // leads.assigned_cp stores display_name text
    o.textContent = u.display_name;
    if (value && u.display_name === value) o.selected = true;
    sel.appendChild(o);
  }

  return sel;
}

/**
 * lookup_or_text — text input with live typeahead against the companies table.
 * Selecting a result sets a companion hidden input (data-column="name__id")
 * which collectPayload() uses to populate company_id.
 */
function buildLookupOrText(field, value) {
  const wrap = document.createElement('div');
  wrap.style.position = 'relative';
  wrap.dataset.table  = field.table_name;
  wrap.dataset.column = field.column_name;

  const input         = document.createElement('input');
  input.type          = 'text';
  input.className     = 'form-field__input';
  input.dataset.table  = field.table_name;
  input.dataset.column = field.column_name;
  input.placeholder   = field.placeholder ?? 'Search or type a name…';
  input.value         = value ?? '';
  if (field.is_required) input.required = true;

  // Hidden companion that stores selected record UUID
  const idInput         = document.createElement('input');
  idInput.type          = 'hidden';
  idInput.dataset.table  = field.table_name;
  idInput.dataset.column = `${field.column_name}__id`;

  // Dropdown container
  const dropdown = document.createElement('div');
  dropdown.className = 'typeahead-dropdown hidden';

  let debounceTimer = null;

  input.addEventListener('input', () => {
    // Clear the stored ID when user edits manually
    idInput.value = '';
    clearTimeout(debounceTimer);
    const q = input.value.trim();
    if (q.length < 2) { dropdown.classList.add('hidden'); return; }

    debounceTimer = setTimeout(async () => {
      try {
        const { results } = await API.lookup(field.table_name, q);
        renderTypeaheadDropdown(dropdown, results, item => {
          input.value   = item.label;
          idInput.value = item.id;
          dropdown.classList.add('hidden');
          // Update _selectedCoId so the submit payload uses the existing record
          if (field.table_name === 'companies') _selectedCoId = item.id;
          updateSubmitState();
        });
      } catch {
        dropdown.classList.add('hidden');
      }
    }, 280);
  });

  // Close dropdown on click outside
  document.addEventListener('click', e => {
    if (!wrap.contains(e.target)) dropdown.classList.add('hidden');
  });

  wrap.appendChild(input);
  wrap.appendChild(idInput);
  wrap.appendChild(dropdown);
  return wrap;
}

function renderTypeaheadDropdown(dropdown, results, onSelect) {
  dropdown.innerHTML = '';
  if (!results.length) { dropdown.classList.add('hidden'); return; }

  for (const item of results) {
    const row = document.createElement('div');
    row.className = 'typeahead-item';
    row.innerHTML = `<span class="typeahead-label">${escHtml(item.label)}</span>${item.sublabel ? `<span class="typeahead-sub">${escHtml(item.sublabel)}</span>` : ''}`;
    row.addEventListener('mousedown', e => { e.preventDefault(); onSelect(item); });
    dropdown.appendChild(row);
  }

  dropdown.classList.remove('hidden');
}

// ── Live dedup and assignment wiring ─────────────────────────────────────────

let _dedupeTimer = null;

function wireEmailDedup() {
  const emailInput = getFieldInput('contacts', 'email');
  if (!emailInput) return;

  emailInput.addEventListener('blur', async () => {
    const email = emailInput.value.trim().toLowerCase();
    if (!email || email === _ctx.senderEmail?.toLowerCase()) return;

    try {
      const { match } = await API.matchContact(email);
      if (match && match.id !== _selectedCtId) {
        showInlineMatchBanner('contact', match, () => {
          _selectedCtId = match.id;
          prefillContactFields(match);
        });
      }
    } catch {}
  });
}

function wireWebsiteDedup() {
  const websiteInput = getFieldInput('companies', 'website');
  if (!websiteInput) return;

  websiteInput.addEventListener('blur', async () => {
    const website = websiteInput.value.trim();
    if (!website) return;
    const domain = website.replace(/^https?:\/\//, '').split('/')[0];

    try {
      const { matches } = await API.matchCompany({ domain });
      const match = matches?.[0];
      if (match && match.id !== _selectedCoId) {
        showInlineMatchBanner('company', match, () => {
          _selectedCoId = match.id;
          prefillCompanyFields(match);
        });
      }
    } catch {}
  });
}

function wireStateAssignment() {
  const stateSelect = getFieldInput('companies', 'state');
  const assignedCpInput = getFieldInput('leads', 'assigned_cp');
  if (!stateSelect || !assignedCpInput) return;

  stateSelect.addEventListener('change', async () => {
    const state = stateSelect.value;
    if (!state) return;
    try {
      const { assignment } = await API.resolveAssignment(state);
      if (assignment?.display_name) {
        // Update the users_lookup select
        Array.from(assignedCpInput.options).forEach(opt => {
          opt.selected = opt.value === assignment.display_name;
        });
      }
    } catch {}
  });

  // Run immediately if state already prefilled
  if (stateSelect.value) {
    stateSelect.dispatchEvent(new Event('change'));
  }
}

// ── Inline match banner ───────────────────────────────────────────────────────

function showInlineMatchBanner(type, match, onAccept) {
  const existing = document.getElementById('inline-match-banner');
  if (existing) existing.remove();

  const banner = document.createElement('div');
  banner.id = 'inline-match-banner';
  banner.className = 'inline-match-banner';
  banner.innerHTML = `
    <span>Found existing ${type}: <strong>${escHtml(match.name)}</strong></span>
    <button class="btn btn--small btn--primary" id="inline-use-btn">Use it</button>
    <button class="btn btn--small btn--ghost" id="inline-dismiss-btn">✕</button>
  `;

  banner.querySelector('#inline-use-btn').onclick = () => {
    onAccept();
    banner.remove();
    updateSubmitState();
  };

  banner.querySelector('#inline-dismiss-btn').onclick = () => banner.remove();

  El.formSections.prepend(banner);
}

// ── Prefill helpers ───────────────────────────────────────────────────────────

function getPrefillValue(field, { ctx, matchedCompany, matchedContact, useMatch }) {
  // Server has already resolved {{sender.*}} etc. into field.default_value
  if (useMatch) {
    if (field.table_name === 'companies' && matchedCompany) {
      return matchedCompany[field.column_name] ?? field.default_value ?? null;
    }
    if (field.table_name === 'contacts' && matchedContact) {
      return matchedContact[field.column_name] ?? field.default_value ?? null;
    }
  }
  return field.default_value ?? null;
}

function prefillCompanyFields(company) {
  for (const [col, val] of Object.entries(company)) {
    const input = getFieldInput('companies', col);
    if (input && val !== null && val !== undefined) setInputValue(input, val);
  }
}

function prefillContactFields(contact) {
  for (const [col, val] of Object.entries(contact)) {
    const input = getFieldInput('contacts', col);
    if (input && val !== null && val !== undefined) setInputValue(input, val);
  }
}

function setInputValue(el, val) {
  if (el.type === 'checkbox') { el.checked = Boolean(val); return; }
  el.value = val ?? '';
}

// ── Validation ────────────────────────────────────────────────────────────────

function validateField(field, widget, errorEl) {
  const val = getWidgetValue(widget);
  if (field.is_required && (val === '' || val === null || val === undefined)) {
    const errorId = `${fieldId(field)}-error`;
    errorEl.textContent = `${field.label ?? titleCase(field.column_name)} is required`;
    errorEl.classList.add('visible');
    widget.classList?.add('error');
    return false;
  }
  errorEl.classList.remove('visible');
  widget.classList?.remove('error');
  return true;
}

function updateSubmitState() {
  if (!_schema) { El.submitBtn.disabled = true; return; }

  const visibleRequired = _schema.fields.filter(f => f.is_visible && f.is_required);
  const allFilled = visibleRequired.every(f => {
    const input = getFieldInput(f.table_name, f.column_name);
    if (!input) return false;
    const val = getWidgetValue(input);
    return val !== '' && val !== null && val !== undefined;
  });

  El.submitBtn.disabled = !allFilled;
}

// ── Payload collection ────────────────────────────────────────────────────────

function collectPayload() {
  const company_data = {};
  const contact_data = {};
  const lead_data    = {};

  // Collect all inputs (including hidden)
  const allInputs = El.formSections.querySelectorAll('[data-table][data-column]');

  for (const input of allInputs) {
    const table  = input.dataset.table;
    const column = input.dataset.column;

    // Skip the companion __id fields — handled separately below
    if (column.endsWith('__id')) continue;

    const val = getWidgetValue(input);
    if (val === '' || val === null || val === undefined) continue;

    if (table === 'companies') company_data[column] = val;
    else if (table === 'contacts') contact_data[column] = val;
    else if (table === 'leads')    lead_data[column]    = val;
  }

  return {
    form_version:              _schema.version,
    company_id:                _selectedCoId   ?? null,
    company_data,
    contact_id:                _selectedCtId   ?? null,
    contact_data,
    lead_data,
    activity_content:          buildActivityContent(),
    activity_email_message_id: _emailMessageId ?? null,
  };
}

function buildActivityContent() {
  const parts = [];
  if (_ctx.emailSubject)  parts.push(`Subject: ${_ctx.emailSubject}`);
  if (_ctx.senderEmail)   parts.push(`From: ${_ctx.senderName ? `${_ctx.senderName} <${_ctx.senderEmail}>` : _ctx.senderEmail}`);
  if (_ctx.emailReceivedAt) parts.push(`Received: ${new Date(_ctx.emailReceivedAt).toLocaleString('en-AU')}`);
  return parts.join('\n') || 'Lead created via Outlook add-in';
}

// ── Submit ────────────────────────────────────────────────────────────────────

async function handleSubmit(e) {
  e.preventDefault();
  El.submitBtn.disabled    = true;
  El.submitBtn.textContent = 'Adding…';

  try {
    const payload = collectPayload();
    const result  = await API.createLead(payload);

    _hubUrl = result.hub_url;

    const companyName = payload.company_data?.name ?? 'lead';
    El.successCompany.textContent = result.created_new_company
      ? `New company created: ${companyName}`
      : `Added to existing: ${companyName}`;

    El.viewLeadBtn.onclick = () => {
      Office.context.ui.openBrowserWindow(_hubUrl);
    };

    showView('success');
  } catch (err) {
    if (err instanceof API.ApiError) {
      if (err.isSchemaOutdated) {
        El.errorTitle.textContent = 'Form is outdated';
        El.errorMsg.textContent   = 'Close and reopen the add-in to get the latest form.';
      } else if (err.isDuplicate) {
        El.errorTitle.textContent = 'Already submitted';
        El.errorMsg.textContent   = 'This email was already added as a lead recently.';
      } else {
        El.errorTitle.textContent = 'Could not save lead';
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
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function groupBySection(fields) {
  const order = ['company', 'contact', 'lead', 'qualification'];
  const map   = {};

  for (const f of fields) {
    const sec = f.section ?? 'other';
    if (!map[sec]) map[sec] = [];
    map[sec].push(f);
  }

  // Sort within each section by display_order
  for (const sec of Object.keys(map)) {
    map[sec].sort((a, b) => (a.display_order ?? 99) - (b.display_order ?? 99));
  }

  // Return in canonical section order + any extras at end
  const sorted = {};
  for (const sec of order) {
    if (map[sec]) sorted[sec] = map[sec];
  }
  for (const sec of Object.keys(map)) {
    if (!sorted[sec]) sorted[sec] = map[sec];
  }

  return sorted;
}

function getFieldInput(table, column) {
  return El.formSections.querySelector(
    `[data-table="${CSS.escape(table)}"][data-column="${CSS.escape(column)}"]`
  );
}

function getWidgetValue(el) {
  if (!el) return null;
  if (el.type === 'checkbox') return el.checked;
  // toggle-wrap div: find the checkbox inside
  if (el.classList?.contains('form-field__toggle-wrap')) {
    const chk = el.querySelector('input[type="checkbox"]');
    return chk ? chk.checked : null;
  }
  // lookup_or_text wrap div: find the text input
  if (el.style?.position === 'relative') {
    const txt = el.querySelector('input[type="text"]');
    return txt ? txt.value.trim() || null : null;
  }
  return el.value?.trim() || null;
}

function fieldId(field) {
  return `f_${field.table_name}_${field.column_name}`;
}

function inferWidget(field) {
  if (field.column_name.match(/email/i)) return 'email';
  if (field.column_name.match(/phone|mobile/i)) return 'phone';
  if (field.column_name.match(/notes|description/i)) return 'textarea';
  return 'text';
}

function titleCase(str) {
  return (str ?? '').replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
