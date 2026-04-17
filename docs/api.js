/**
 * api.js — HTTP client for ecc-crmplus Function App endpoints
 *
 * Every method acquires a fresh-or-cached Bearer token via Auth.getToken().
 * On 401 the token cache is cleared and the call is retried once.
 */

const API = (() => {
  const BASE = 'https://ecc-crmplus.azurewebsites.net/api';

  // ── Core fetch wrapper ──────────────────────────────────────────────────────

  async function _fetch(path, options = {}, retried = false) {
    let token;
    try {
      token = await Auth.getToken();
    } catch (err) {
      throw new ApiError('Authentication failed', 401, err.message);
    }

    const res = await fetch(`${BASE}${path}`, {
      ...options,
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`,
        ...(options.headers ?? {}),
      },
    });

    if (res.status === 401 && !retried) {
      Auth.clearToken();
      return _fetch(path, options, true);
    }

    if (!res.ok) {
      let body = {};
      try { body = await res.json(); } catch {}
      throw new ApiError(
        body.message ?? `Request failed (${res.status})`,
        res.status,
        body.error ?? 'unknown_error'
      );
    }

    return res.json();
  }

  // ── Endpoints ────────────────────────────────────────────────────────────────

  /**
   * GET /crm/forms/{formKey}
   * Returns the active schema for the given form key.
   */
  async function getFormSchema(formKey, context = {}) {
    const params = new URLSearchParams();
    if (context.senderEmail)     params.set('senderEmail',     context.senderEmail);
    if (context.senderName)      params.set('senderName',      context.senderName);
    if (context.emailSubject)    params.set('emailSubject',    context.emailSubject);
    if (context.emailReceivedAt) params.set('emailReceivedAt', context.emailReceivedAt);
    if (context.mailboxAddress)  params.set('mailboxAddress',  context.mailboxAddress);
    const qs = params.toString();
    return _fetch(`/crm/forms/${encodeURIComponent(formKey)}${qs ? `?${qs}` : ''}`);
  }

  /**
   * GET /crm/match/company?domain=&abn=&name=
   * Returns up to 10 matching companies.
   */
  async function matchCompany({ domain, abn, name } = {}) {
    const params = new URLSearchParams();
    if (domain) params.set('domain', domain);
    if (abn)    params.set('abn',    abn);
    if (name)   params.set('name',   name);
    return _fetch(`/crm/match/company?${params}`);
  }

  /**
   * GET /crm/match/contact?email=
   * Returns the matching contact or {match: null}.
   */
  async function matchContact(email) {
    return _fetch(`/crm/match/contact?email=${encodeURIComponent(email)}`);
  }

  /**
   * GET /crm/lookup/{table}?q=
   * Typeahead lookup (companies or contacts), min 2 chars, top 20.
   */
  async function lookup(table, q) {
    return _fetch(`/crm/lookup/${encodeURIComponent(table)}?q=${encodeURIComponent(q)}`);
  }

  /**
   * GET /crm/users
   * Returns all hub_users for the "Assign to" picker.
   */
  async function getUsers() {
    return _fetch('/crm/users');
  }

  /**
   * POST /crm/assignment/resolve
   * Returns the CP assigned to the given state (falls back to default rule).
   */
  async function resolveAssignment(state) {
    return _fetch('/crm/assignment/resolve', {
      method: 'POST',
      body: JSON.stringify({ state }),
    });
  }

  /**
   * POST /crm/leads
   * Creates company + contact + lead + activity in one atomic transaction.
   * Returns { lead_id, company_id, contact_id, hub_url, created_new_company, created_new_contact }.
   */
  async function createLead(payload) {
    return _fetch('/crm/leads', {
      method: 'POST',
      body: JSON.stringify(payload),
    });
  }

  // ── Error class ──────────────────────────────────────────────────────────────

  class ApiError extends Error {
    constructor(message, status = 500, code = 'unknown_error') {
      super(message);
      this.name   = 'ApiError';
      this.status = status;
      this.code   = code;
    }

    get isSchemaOutdated() { return this.code === 'schema_mismatch'; }
    get isDuplicate()      { return this.code === 'duplicate_lead';  }
    get isNotProvisioned() { return this.code === 'user_not_provisioned'; }
  }

  return {
    getFormSchema,
    matchCompany,
    matchContact,
    lookup,
    getUsers,
    resolveAssignment,
    createLead,
    ApiError,
  };
})();
