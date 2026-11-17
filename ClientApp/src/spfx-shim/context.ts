import { apiFetch } from '../auth/auth';
import { createGraphClient } from './graph';

/** Shape returned by GET /api/config */
export interface AppConfig {
  siteUrl: string;
  tenantId: string;
  siteId: string;
  webId: string;
  webTitle: string;
  webServerRelativeUrl: string;
  user: { displayName: string; email: string; loginName: string };
}

export const SHAREPOINT_PROXY = `${window.location.origin}/api/sharepoint`;

/**
 * Link that streams a SharePoint file THROUGH this app (GET /api/files/open).
 * Linking to SharePoint's own URL would ask the user for a Microsoft sign-in, which this app avoids.
 */
export function fileOpenUrl(serverRelativeUrl: string): string {
  return `${window.location.origin}/api/files/open?path=${encodeURIComponent(serverRelativeUrl || '')}`;
}

/** Profile picture of a person, served by this app (Graph, application permission) for <img src>. */
export function photoUrlFor(email: string): string {
  return `${window.location.origin}/api/photo?email=${encodeURIComponent(email || '')}`;
}

/**
 * Turns whatever URL the original SPFx code built (absolute site URL, or server-relative)
 * into the path the ASP.NET Core proxy expects, e.g. "_api/web/lists/...".
 */
export function toSharePointRelative(url: string, cfg: AppConfig): string {
  const site = cfg.siteUrl.replace(/\/+$/, '');
  const lower = url.toLowerCase();

  if (lower.startsWith(site.toLowerCase() + '/')) return url.slice(site.length + 1);

  const rel = cfg.webServerRelativeUrl.replace(/\/+$/, '');
  if (rel && lower.startsWith((rel + '/_api/').toLowerCase())) return url.slice(rel.length + 1);

  if (url.startsWith('/_api/')) return url.slice(1);
  if (url.startsWith('_api/')) return url;

  throw new Error(`Blocked: "${url}" is outside the configured SharePoint site.`);
}

/** Replacement for SPFx's SPHttpClient (only get/post were used by the app). */
export class SPHttpClient {
  public static readonly configurations = { v1: { flags: 0 } };
}
export type SPHttpClientResponse = Response;

function createSpHttpClient(cfg: AppConfig) {
  const send = (method: 'GET' | 'POST', url: string, _config?: unknown, options?: any): Promise<Response> => {
    const target = `${SHAREPOINT_PROXY}/${toSharePointRelative(url, cfg)}`;
    return apiFetch(target, {
      method,
      headers: options?.headers,
      body: options?.body
    });
  };
  return {
    get: (url: string, config?: unknown, options?: any) => send('GET', url, config, options),
    post: (url: string, config?: unknown, options?: any) => send('POST', url, config, options)
  };
}

/**
 * A stand-in for the SPFx WebPartContext. The 13 components that used `context`
 * only touch: pageContext.{web,site,user,aadInfo}, spHttpClient.{get,post},
 * msGraphClientFactory.getClient() and (via PnPjs) the SPFx() behavior.
 */
export function createSpContext(cfg: AppConfig) {
  return {
    proxyBaseUrl: SHAREPOINT_PROXY,
    appConfig: cfg,
    isServedFromLocalhost: window.location.hostname === 'localhost',
    sdks: { microsoftTeams: undefined as undefined },
    pageContext: {
      web: {
        absoluteUrl: cfg.siteUrl,
        serverRelativeUrl: cfg.webServerRelativeUrl,
        id: cfg.webId,
        title: cfg.webTitle
      },
      site: { id: cfg.siteId },
      aadInfo: { tenantId: cfg.tenantId },
      user: {
        displayName: cfg.user.displayName,
        email: cfg.user.email,
        loginName: cfg.user.loginName
      }
    },
    spHttpClient: createSpHttpClient(cfg),
    msGraphClientFactory: {
      getClient: async (_version?: string) => createGraphClient()
    }
  };
}

export type SpContext = ReturnType<typeof createSpContext>;
