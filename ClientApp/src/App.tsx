import * as React from 'react';
import Dms from './components/Dms';
import { bootstrapAuth, apiFetch } from './auth/auth';
import { createSpContext, type AppConfig } from './spfx-shim';

/**
 * Replaces DmsWebPart.ts: load the "page context" from the API (the user is already signed in to the portal), then render the original UI.
 */
export default function App(): React.ReactElement {
  const [context, setContext] = React.useState<any>(null);
  const [error, setError] = React.useState<string>('');

  React.useEffect(() => {
    (async () => {
      try {
        await bootstrapAuth();
        const res = await apiFetch('/api/config');
        if (!res.ok) {
          let detail = '';
          try { detail = (await res.json())?.error?.message ?? ''; } catch { /* ignore */ }
          throw new Error(detail || `Could not load configuration (${res.status}).`);
        }
        const cfg: AppConfig = await res.json();
        setContext(createSpContext(cfg));
      } catch (e: any) {
        setError(e?.message ?? String(e));
      }
    })();
  }, []);

  if (error) {
    return <div style={{ padding: 24, color: '#a4262c' }}><b>Could not start the app.</b><br />{error}</div>;
  }
  if (!context) {
    return <div style={{ padding: 24 }}>Loading…</div>;
  }
  return <Dms context={context} />;
}
