import * as React from 'react';

// -----------------------------------------------------------------------------
// License Expiry popup
// Shown by AnalyticsDashboard.tsx when the Licence API says the trial is over
// (daysRemaining <= 0). It sits on top of the whole app (dimmed + blurred
// backdrop, centered card with 16px rounded corners - same look as the
// "Process Master Form" popup) and has NO close button, so the app stays
// blocked until the licence is renewed.
// -----------------------------------------------------------------------------

// >>> CHANGE THESE to your real contact details <<<
const CONTACT_EMAIL = 'support@speteq.com';
const CONTACT_PHONE = '+91 95866-61642';

export interface ILicenseExpiryPopupProps {
  isOpen: boolean;
}

const LicenseExpiryPopup: React.FC<ILicenseExpiryPopupProps> = ({ isOpen }) => {
  if (!isOpen) return null;

  return (
    <div className="lic-overlay" style={styles.overlay}>
      <style>{`
        @keyframes licFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes licScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        .lic-overlay { animation: licFadeIn 0.18s ease; }
        .lic-card { animation: licScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
        .lic-btn { transition: transform 0.15s ease, filter 0.15s ease; }
        .lic-btn:hover { transform: translateY(-1px); filter: brightness(1.08); }
        .lic-btn:active { transform: translateY(0); }
      `}</style>

      <div className="lic-card" style={styles.card} role="alertdialog" aria-modal="true" aria-labelledby="lic-title">
        <div style={styles.iconWrap}>
          <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#b45309" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="10" />
            <polyline points="12 6 12 12 16 14" />
          </svg>
        </div>

        <h2 id="lic-title" style={styles.title}>Licence Renewal Required</h2>
        <p style={styles.message}>
          Your application licence has expired, and access to the Document Management System is currently restricted. 
          Please contact your administrator or our support team to renew your licence and restore access.
        </p>

        <div style={styles.infoBox}>
          <div style={styles.infoRow}>
            <span style={styles.infoLabel}>Status</span>
            <span style={styles.expiredChip}>Expired</span>
          </div>
          <div style={styles.infoRow}>
            <span style={styles.infoLabel}>Email</span>
            <a href={`mailto:${CONTACT_EMAIL}`} style={styles.infoLink}>{CONTACT_EMAIL}</a>
          </div>
          <div style={{ ...styles.infoRow, borderBottom: 'none' }}>
            <span style={styles.infoLabel}>Phone</span>
            <span style={styles.infoValue}>{CONTACT_PHONE}</span>
          </div>
        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(6px)', WebkitBackdropFilter: 'blur(6px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 99999, padding: '24px 16px', boxSizing: 'border-box' },
  card: { backgroundColor: '#ffffff', borderRadius: '16px', width: '100%', maxWidth: '440px', padding: '32px 28px 28px 28px', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)', border: '1px solid #e2e8f0', boxSizing: 'border-box', textAlign: 'center', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
  iconWrap: { width: '60px', height: '60px', borderRadius: '50%', backgroundColor: '#fef3c7', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 16px auto' },
  title: { margin: '0 0 8px 0', fontSize: '20px', fontWeight: 700, color: '#0f172a' },
  message: { margin: '0 0 20px 0', fontSize: '13px', lineHeight: 1.6, color: '#475569' },
  infoBox: { border: '1px solid #e2e8f0', borderRadius: '10px', backgroundColor: '#f8fafc', padding: '4px 14px', marginBottom: '20px', textAlign: 'left' },
  infoRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 0', borderBottom: '1px solid #e2e8f0', gap: '12px' },
  infoLabel: { fontSize: '12px', fontWeight: 600, color: '#64748b' },
  infoValue: { fontSize: '12px', fontWeight: 600, color: '#1e293b' },
  infoLink: { fontSize: '12px', fontWeight: 600, color: '#004d3d', textDecoration: 'none' },
  expiredChip: { display: 'inline-block', padding: '2px 10px', borderRadius: '999px', fontSize: '11px', fontWeight: 700, backgroundColor: '#fee2e2', color: '#b91c1c' },
  contactBtn: { display: 'block', width: '100%', boxSizing: 'border-box', backgroundColor: '#004d3d', color: '#ffffff', border: 'none', borderRadius: '8px', padding: '11px 16px', fontSize: '13px', fontWeight: 700, textDecoration: 'none', cursor: 'pointer' }
};

export default LicenseExpiryPopup;
