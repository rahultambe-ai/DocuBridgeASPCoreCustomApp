/**
 * Sign-in is handled by the portal (SSL Core Web App): the user types a user name + password on /Account/Login
 * and the server sets a secure session cookie. This React app lives on the same origin, so the browser
 * sends that cookie with every /api call automatically. There is NO Microsoft sign-in here and no token
 * handling in the browser: SharePoint and Graph are called by the server with the application's own identity.
 */

const LOGIN_URL = `/Account/Login?returnUrl=${encodeURIComponent('/Dms')}`;

/** Kept so App.tsx reads the same as before; nothing to do because the cookie already exists. */
export async function bootstrapAuth(): Promise<void> {
  /* intentionally empty */
}

/**
 * fetch() for this app's own /api endpoints.
 *  - X-Requested-With is required by the server for every non-GET request (CSRF defence).
 *  - If the session has expired (401) the user is sent back to the portal login page.
 */
export async function apiFetch(input: string, init: RequestInit = {}): Promise<Response> {
  const headers = new Headers(init.headers);
  headers.set('X-Requested-With', 'XMLHttpRequest');

  const res = await fetch(input, { ...init, headers, credentials: 'same-origin' });

  if (res.status === 401) {
    window.location.assign(LOGIN_URL);
    return new Promise<Response>(() => { /* page is navigating away */ });
  }
  return res;
}
