import { SPBrowser } from '@pnp/sp';

/**
 * Drop-in replacement for PnPjs's `SPFx(context)` behavior.
 *
 *   const sp = spfi().using(SPFx(context));   // original code keeps working unchanged
 *
 * Instead of calling SharePoint directly with SPFx's token provider, every PnPjs request is sent to
 * <this app>/api/sharepoint/_api/... . The browser's session cookie authenticates the request and the
 * ASP.NET Core server forwards it to SharePoint with the application's own identity.
 *
 * IMPORTANT: `init.headers` MUST stay a plain object here, never a `Headers` instance.
 *  - PnPjs copies headers after the auth step with `{ ...init.headers }`, and spreading a `Headers`
 *    object gives `{}`, which silently drops the headers added here.
 *  - PnPjs' RequestDigest behavior checks `hasOwnProperty("Authorization")` on that same object to know
 *    it must NOT fetch a form digest. The server never forwards this header (it adds its own token), so the
 *    value is only a marker and carries no secret.
 */
export function SPFx(context: { proxyBaseUrl: string }) {
  return (instance: any) => {
    instance.using(SPBrowser({ baseUrl: context.proxyBaseUrl }));

    instance.on.auth.replace(async (url: URL, init: RequestInit) => {
      const existing: Record<string, string> =
        typeof Headers !== 'undefined' && init.headers instanceof Headers
          ? Object.fromEntries((init.headers as Headers).entries())
          : { ...((init.headers as Record<string, string>) || {}) };

      init.headers = {
        ...existing,
        Authorization: 'Bearer portal-session',
        'X-Requested-With': 'XMLHttpRequest'
      };
      return [url, init];
    });

    return instance;
  };
}
