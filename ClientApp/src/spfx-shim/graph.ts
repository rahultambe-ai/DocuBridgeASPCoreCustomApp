import { apiFetch } from '../auth/auth';

const GRAPH_PROXY = `${window.location.origin}/api/graph`;

type ResponseKind = 'json' | 'blob' | 'text';

/** Minimal clone of the @microsoft/microsoft-graph-client fluent API used by the app. */
class GraphRequest {
  private readonly query = new URLSearchParams();
  private readonly extraHeaders: Record<string, string> = {};
  private kind: ResponseKind = 'json';

  constructor(private readonly path: string) {}

  select(fields: string | string[]) { this.query.set('$select', Array.isArray(fields) ? fields.join(',') : fields); return this; }
  expand(fields: string | string[]) { this.query.set('$expand', Array.isArray(fields) ? fields.join(',') : fields); return this; }
  filter(expr: string)  { this.query.set('$filter', expr); return this; }
  orderby(expr: string) { this.query.set('$orderby', expr); return this; }
  top(n: number)        { this.query.set('$top', String(n)); return this; }
  skip(n: number)       { this.query.set('$skip', String(n)); return this; }
  count(v = true)       { this.query.set('$count', String(v)); return this; }
  search(expr: string)  { this.query.set('$search', expr); return this; }
  header(name: string, value: string) { this.extraHeaders[name] = value; return this; }
  responseType(kind: string) { this.kind = (kind as ResponseKind) || 'json'; return this; }

  get()                 { return this.send('GET'); }
  post(body?: unknown)  { return this.send('POST', body); }
  patch(body?: unknown) { return this.send('PATCH', body); }
  put(body?: unknown)   { return this.send('PUT', body); }
  delete()              { return this.send('DELETE'); }

  private async send(method: string, body?: unknown): Promise<any> {
    // Build the query by hand so '$' and ',' stay readable and the Graph proxy sees what Graph expects.
    const qs = Array.from(this.query.entries())
      .map(([k, v]) => `${k}=${encodeURIComponent(v).replace(/%2C/g, ',')}`)
      .join('&');
    const cleanPath = this.path.replace(/^\/+/, '');
    const url = `${GRAPH_PROXY}/${cleanPath}${qs ? (cleanPath.includes('?') ? '&' : '?') + qs : ''}`;

    const headers: Record<string, string> = { ...this.extraHeaders };
    let payload: BodyInit | undefined;
    if (body !== undefined && body !== null) {
      if (typeof body === 'string' || body instanceof Blob || body instanceof ArrayBuffer) {
        payload = body as BodyInit;
      } else {
        payload = JSON.stringify(body);
        headers['Content-Type'] = 'application/json';
      }
    }

    const res = await apiFetch(url, { method, headers, body: payload });

    if (!res.ok) {
      let code = 'graphError';
      let message = res.statusText;
      try {
        const err = await res.json();
        code = err?.error?.code || code;
        message = err?.error?.message || message;
      } catch { /* body was not JSON */ }
      // Same fields the real GraphError exposes (UserProfile.tsx reads statusCode)
      const error: any = new Error(message);
      error.statusCode = res.status;
      error.code = code;
      throw error;
    }

    if (this.kind === 'blob') return res.blob();
    if (this.kind === 'text') return res.text();
    if (res.status === 204) return null;
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }
}

export function createGraphClient() {
  return { api: (path: string) => new GraphRequest(path) };
}
