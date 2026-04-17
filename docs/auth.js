/**
 * auth.js — Office SSO token acquisition with MSAL fallback
 *
 * Usage:
 *   const token = await Auth.getToken()   // throws AuthError on failure
 *   Auth.clearToken()                     // force re-acquire on next call
 */

const Auth = (() => {
  const FALLBACK_URL = 'https://harryeccp.github.io/crm-outlook-addin/fallback.html';
  const DIALOG_WIDTH  = 40;   // % of screen
  const DIALOG_HEIGHT = 60;

  let _cachedToken = null;

  /**
   * Acquire an access token for the CRM+ API.
   * 1. Tries Office SSO (Office.auth.getAccessToken).
   * 2. On 13xxx SSO errors, opens fallback.html in a dialog for interactive MSAL auth.
   * 3. Token is cached in memory until clearToken() is called.
   */
  async function getToken() {
    if (_cachedToken) return _cachedToken;

    try {
      const token = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false,
      });
      _cachedToken = token;
      return token;
    } catch (err) {
      const code = err.code ?? 0;

      // 13000-series: SSO unavailable (shared mailbox, consent needed, etc.)
      // Fall back to interactive MSAL dialog.
      if (code >= 13000 && code < 14000) {
        return _fallbackDialog();
      }

      // 13003: user not signed in to Office — same fallback
      if (code === 13003 || code === 13005) {
        return _fallbackDialog();
      }

      throw new AuthError(`SSO failed (${code}): ${err.message ?? err}`, code);
    }
  }

  function clearToken() {
    _cachedToken = null;
  }

  /**
   * Opens fallback.html in an Office dialog for interactive MSAL auth.
   * Resolves with the access token posted back via Office.context.ui.messageParent.
   */
  function _fallbackDialog() {
    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        FALLBACK_URL,
        { height: DIALOG_HEIGHT, width: DIALOG_WIDTH, promptBeforeOpen: false },
        result => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            return reject(new AuthError(`Dialog failed: ${result.error.message}`, result.error.code));
          }

          const dialog = result.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, event => {
            dialog.close();
            try {
              const msg = JSON.parse(event.message);
              if (msg.type === 'auth_success' && msg.token) {
                _cachedToken = msg.token;
                resolve(msg.token);
              } else {
                reject(new AuthError('Fallback auth did not return a token', 0));
              }
            } catch {
              reject(new AuthError('Invalid message from auth dialog', 0));
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, event => {
            dialog.close();
            reject(new AuthError(`Auth dialog closed (event ${event.error})`, event.error));
          });
        }
      );
    });
  }

  class AuthError extends Error {
    constructor(message, code = 0) {
      super(message);
      this.name  = 'AuthError';
      this.code  = code;
    }
  }

  return { getToken, clearToken, AuthError };
})();
