import * as React from 'react';

// -----------------------------------------------------------------------------
// License Expiry Banner
// Shown by AnalyticsDashboard.tsx just below the top header while the trial is
// still running (daysRemaining > 0). It tells the user how many trial days are
// left. When the days reach 0, this banner disappears and LicenseExpiryPopup
// takes over and blocks the app.
//
// Colour changes with the days left:
//   16 days or more  -> GREEN  (all good)
//   8 to 15 days     -> YELLOW (ending soon)
//   7 days or less   -> RED    (renew now)
// -----------------------------------------------------------------------------

// >>> CHANGE THESE to your real contact details (same as LicenseExpiryPopup) <<<
const CONTACT_EMAIL = 'support@speteq.com';
const CONTACT_PHONE = '+91 95866-61642';

export interface ILicenseExpiryBannerProps {
  /** Trial days left from the Licence API. null = not loaded yet / API unreachable. */
  daysRemaining: number | null;
}

type Tier = 'safe' | 'warning' | 'danger';

interface ITheme {
  accent: string;       // left bar + divider
  iconBg: string;       // circle behind the icon
  iconColor: string;    // clock icon
  bgFrom: string;       // banner gradient start
  bgTo: string;         // banner gradient end
  border: string;       // banner outline
  contactBg: string;    // email/phone box
  chipBg: string;       // days chip
  chipColor: string;
}

const THEMES: { [key in Tier]: ITheme } = {
  safe: {
    accent: '#16a34a', iconBg: '#dcfce7', iconColor: '#15803d',
    bgFrom: '#f0fdf4', bgTo: '#f7fef9', border: '#bbf7d0', contactBg: 'rgba(220, 252, 231, 0.55)',
    chipBg: '#dcfce7', chipColor: '#166534'
  },
  warning: {
    accent: '#f59e0b', iconBg: '#fde68a', iconColor: '#b45309',
    bgFrom: '#fffbeb', bgTo: '#fef9ee', border: '#fde8b0', contactBg: 'rgba(254, 243, 199, 0.55)',
    chipBg: '#fef3c7', chipColor: '#b45309'
  },
  danger: {
    accent: '#dc2626', iconBg: '#fecaca', iconColor: '#b91c1c',
    bgFrom: '#fef2f2', bgTo: '#fff5f5', border: '#fecaca', contactBg: 'rgba(254, 226, 226, 0.55)',
    chipBg: '#fee2e2', chipColor: '#b91c1c'
  }
};

const getTier = (days: number): Tier => (days <= 7 ? 'danger' : days <= 15 ? 'warning' : 'safe');

const getMessage = (tier: Tier): string => {
  if (tier === 'danger') {
    return 'Your trial is about to end. Please contact your administrator or our support team now to renew your licence and avoid losing access to your documents and workspace.';
  }
  if (tier === 'warning') {
    return 'Your trial period is ending soon. To keep using the Document Management System without interruption, please contact your administrator or our support team to renew your licence.';
  }
  return 'Your Document Management System trial is active and all features are available. Contact our support team any time if you want to upgrade your licence.';
};

const LicenseExpiryBanner: React.FC<ILicenseExpiryBannerProps> = ({ daysRemaining }) => {
  if (daysRemaining === null || daysRemaining === undefined || isNaN(daysRemaining) || daysRemaining <= 0) return null;

  const tier = getTier(daysRemaining);
  const theme = THEMES[tier];
  const dayLabel = `${daysRemaining} day${daysRemaining === 1 ? '' : 's'}`;

  return (
    <div
      className="licb-root"
      role="status"
      style={{
        ...styles.root,
        background: `linear-gradient(90deg, ${theme.bgFrom} 0%, ${theme.bgTo} 100%)`,
        border: `1px solid ${theme.border}`,
        borderLeft: `5px solid ${theme.accent}`
      }}
    >
      <style>{`
        @keyframes licbSlideIn { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .licb-root { animation: licbSlideIn 0.25s ease; }
        .licb-btn { transition: transform 0.15s ease, filter 0.15s ease; }
        .licb-btn:hover { transform: translateY(-1px); filter: brightness(1.1); }
        .licb-btn:active { transform: translateY(0); }
        @media (max-width: 1100px) { .licb-contact { display: none !important; } }
        @media (max-width: 700px) { .licb-root { gap: 10px !important; padding: 10px 12px !important; } .licb-divider { display: none !important; } .licb-btn { width: 100%; justify-content: center; } }
      `}</style>

      {/* Icon */}
      <div style={{ ...styles.iconWrap, backgroundColor: theme.iconBg }}>
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={theme.iconColor} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
          <circle cx="12" cy="12" r="10" />
          <polyline points="12 6 12 12 16 14" />
        </svg>
      </div>

      <div className="licb-divider" style={{ ...styles.divider, backgroundColor: theme.accent }} />

      {/* Title + message */}
      <div style={styles.textWrap}>
        <div style={styles.titleRow}>
          <span style={styles.title}>Trial period: {dayLabel} remaining</span>
          <span style={{ ...styles.chip, backgroundColor: theme.chipBg, color: theme.chipColor }}>
            {tier === 'danger' ? 'Renew now' : tier === 'warning' ? 'Ending soon' : 'Active'}
          </span>
        </div>
        <div style={styles.message}>{getMessage(tier)}</div>
      </div>

      {/* Contact box */}
      <div className="licb-contact" style={{ ...styles.contactBox, backgroundColor: theme.contactBg }}>
        <div style={styles.contactRow}>
          <span style={styles.contactIcon}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="5" width="18" height="14" rx="2" />
              <polyline points="3 7 12 13 21 7" />
            </svg>
          </span>
          <a href={`mailto:${CONTACT_EMAIL}`} style={styles.contactLink}>{CONTACT_EMAIL}</a>
        </div>
        <div style={styles.contactRow}>
          <span style={styles.contactIcon}>
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.13.96.36 1.9.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.91.34 1.85.57 2.81.7A2 2 0 0 1 22 16.92z" />
            </svg>
          </span>
          <span style={styles.contactText}>{CONTACT_PHONE}</span>
        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  root: { display: 'flex', alignItems: 'center', gap: '16px', padding: '10px 16px', margin: '10px 16px 0 16px', borderRadius: '10px', flexShrink: 0, boxSizing: 'border-box', flexWrap: 'wrap', fontFamily: '"Segoe UI", Arial, sans-serif' },
  iconWrap: { width: '40px', height: '40px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  divider: { width: '2px', alignSelf: 'stretch', borderRadius: '2px', opacity: 0.7, flexShrink: 0 },
  textWrap: { flex: '1 1 320px', minWidth: 0 },
  titleRow: { display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', marginBottom: '2px' },
  title: { fontSize: '14px', fontWeight: 700, color: '#0f172a' },
  chip: { display: 'inline-block', padding: '1px 9px', borderRadius: '999px', fontSize: '10.5px', fontWeight: 700 },
  message: { fontSize: '12px', lineHeight: 1.5, color: '#334155' },
  contactBox: { display: 'flex', flexDirection: 'column', gap: '6px', padding: '8px 14px', borderRadius: '10px', flexShrink: 0 },
  contactRow: { display: 'flex', alignItems: 'center', gap: '10px' },
  contactIcon: { width: '22px', height: '22px', borderRadius: '50%', backgroundColor: '#ffffff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  contactLink: { fontSize: '12px', fontWeight: 600, color: '#004d3d', textDecoration: 'none' },
  contactText: { fontSize: '12px', fontWeight: 600, color: '#1e293b' },
  button: { display: 'inline-flex', alignItems: 'center', gap: '8px', backgroundColor: '#004d3d', color: '#ffffff', borderRadius: '8px', padding: '10px 18px', fontSize: '13px', fontWeight: 700, textDecoration: 'none', cursor: 'pointer', flexShrink: 0, boxSizing: 'border-box' }
};

export default LicenseExpiryBanner;
