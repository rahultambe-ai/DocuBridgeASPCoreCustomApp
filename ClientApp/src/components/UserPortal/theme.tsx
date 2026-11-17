import * as React from 'react';

/* ============================================================
   DMS GREEN THEME — single source of truth for the whole app.
   Every value below was taken from the colors / sizes already
   used in the existing screens, so switching a screen over to
   this theme keeps it looking the same.

   Usage in any component file:
     import { theme, getStatusStyle } from './theme';
     <button style={theme.components.buttonPrimary}>Save</button>
     <div style={{ color: theme.colors.primary }} />
   ============================================================ */

export const colors = {
  // ---- Brand green ----
  primary: '#005a43',        // main brand green (buttons, active items, links)
  primaryHover: '#004d3d',   // darker green for hover / pressed
  primaryText: '#0d6b52',    // green used for secondary green text
  primaryLight: '#e6f4ea',   // light green background (active nav item, Active/Approved pill)
  primarySoft: '#e6f4f1',    // very light green (tags, doc-number chips)
  primaryBorder: '#ccece6',  // light green border

  // ---- Neutrals ----
  white: '#ffffff',
  bgPage: '#f8fafc',         // page background
  bgMuted: '#f1f5f9',        // row dividers / muted surfaces
  border: '#e2e8f0',         // default border
  borderStrong: '#cbd5e1',   // input / dropdown border

  textStrong: '#0f172a',     // headings
  text: '#1e293b',           // main text
  textBody: '#334155',       // body text
  textNav: '#384c66',        // navigation text
  textSecondary: '#475569',
  textMuted: '#64748b',      // labels / sub-text
  textSubtle: '#94a3b8',     // placeholders / hints

  // ---- Status / semantic ----
  success: '#15803d',
  successBg: '#f0fdf4',
  warning: '#d97706',
  warningBg: '#fef3c7',
  danger: '#dc2626',
  dangerBg: '#fee2e2',
  info: '#0284c7',
  infoBg: '#e0f2fe',
  orange: '#c2410c',
  orangeBg: '#ffedd5',
  gray: '#4b5563',
  grayBg: '#f3f4f6',

  // ---- Extra shades used across screens ----
  gray500: '#6b7280',        // secondary grey text
  greenInk: '#0f2b22',       // very dark green text (workflow step labels)
  surfaceTint: '#fafcfc',    // faint tinted surface (upload drop zone)
  teal: '#0d9488',           // doc-number text / Approved chart color
  successStrong: '#059669',  // Active dot / chart color
  successSoft: '#dcfce7',    // Active pill background
  successText: '#137333',    // Active text on light green
  infoStrong: '#2563eb',     // In Review dot
  orangeStrong: '#ea580c',   // Not in DMS dot / chart color
  warningDark: '#b45309',    // pending text on light yellow
  warningText: '#92400e',    // warning banner text
  warningBorder: '#fde68a',  // warning banner border
  error: '#ef4444',          // error toast background
  dangerDark: '#b91c1c',     // error / optional-step text
  dangerSoft: '#fef2f2',     // error tinted background
  dangerBorder: '#fecaca'    // error tinted border
};

/* ---- Notifications (toast) — success is the brand green, error is red ---- */
export const notification = {
  success: { backgroundColor: colors.primary, color: colors.white },
  error: { backgroundColor: colors.error, color: colors.white },
  warning: { backgroundColor: colors.warningBg, color: colors.warningText, borderColor: colors.warningBorder },
  info: { backgroundColor: colors.infoBg, color: colors.info }
};

/* ---- Overlays / focus (rgba) ---- */
export const overlay = {
  backdrop: 'rgba(15, 23, 42, 0.55)',       // modal backdrop
  backdropLight: 'rgba(15, 23, 42, 0.45)',
  focusRing: 'rgba(0, 90, 67, 0.15)',       // green glow on focused inputs
  primaryTint: 'rgba(0, 90, 67, 0.10)'      // faint green wash
};

/* ---- Colors for chart segments and status dots ---- */
export const statusDotColors: { [status: string]: string } = {
  'Active': colors.successStrong,
  'Approved': colors.teal,
  'Pending Approval': colors.warning,
  'In Review': colors.info,
  'Draft': colors.textSubtle,
  'Archived': colors.textMuted,
  'Not in DMS': colors.orangeStrong
};

export const typography = {
  fontFamily: '"Segoe UI", Arial, sans-serif',
  size: { xs: '10px', sm: '11px', md: '12px', base: '13px', lg: '14px', xl: '16px', xxl: '20px' },
  weight: { regular: 400, medium: 500, semibold: 600, bold: 700 }
};

export const radius = {
  xs: '4px',
  sm: '6px',
  md: '8px',
  lg: '12px',
  pill: '999px'
};

export const shadows = {
  menu: '0 4px 12px rgba(0,0,0,0.1)',
  dropdown: '0 8px 20px rgba(0,0,0,0.12)',
  modal: '0 10px 25px rgba(0,0,0,0.15)'
};

export const spacing = {
  xs: '4px',
  sm: '8px',
  md: '12px',
  lg: '16px',
  xl: '24px'
};

/* ---- Document status colors (same mapping the dashboard uses today) ---- */
export const statusColors: { [status: string]: { backgroundColor: string; color: string } } = {
  'Active': { backgroundColor: colors.primaryLight, color: colors.primary },
  'Approved': { backgroundColor: colors.primaryLight, color: colors.primary },
  'In Review': { backgroundColor: colors.infoBg, color: colors.info },
  'Pending Approval': { backgroundColor: colors.warningBg, color: colors.warning },
  'Draft': { backgroundColor: colors.grayBg, color: colors.gray },
  'Archived': { backgroundColor: colors.border, color: colors.textSecondary },
  'Not in DMS': { backgroundColor: colors.orangeBg, color: colors.orange }
};

export const getStatusStyle = (status: string): { backgroundColor: string; color: string } =>
  statusColors[status] || { backgroundColor: colors.grayBg, color: colors.gray };

export type ToastType = 'success' | 'error' | 'warning' | 'info';

/* Style for a notification of the given type: style={getToastStyle(toastMessage.type)} */
export const getToastStyle = (type: ToastType): React.CSSProperties => ({
  ...components.toast,
  ...notification[type]
});

/* ---- Ready-made component styles (spread and override as needed:
        style={{ ...theme.components.buttonPrimary, width: '100%' }}) ---- */
const components: { [name: string]: React.CSSProperties } = {
  buttonPrimary: {
    padding: '8px 16px',
    border: 'none',
    backgroundColor: colors.primary,
    color: colors.white,
    borderRadius: radius.md,
    cursor: 'pointer',
    fontSize: typography.size.md,
    fontWeight: typography.weight.semibold,
    fontFamily: typography.fontFamily
  },
  buttonSecondary: {
    padding: '8px 16px',
    border: `1px solid ${colors.borderStrong}`,
    backgroundColor: colors.white,
    color: colors.textBody,
    borderRadius: radius.md,
    cursor: 'pointer',
    fontSize: typography.size.md,
    fontWeight: typography.weight.semibold,
    fontFamily: typography.fontFamily
  },
  card: {
    backgroundColor: colors.white,
    borderRadius: radius.lg,
    border: `1px solid ${colors.border}`,
    padding: spacing.lg,
    boxSizing: 'border-box'
  },
  input: {
    padding: '8px 12px',
    border: `1px solid ${colors.borderStrong}`,
    borderRadius: radius.md,
    fontSize: typography.size.md,
    fontFamily: typography.fontFamily,
    color: colors.text,
    backgroundColor: colors.white,
    boxSizing: 'border-box',
    outline: 'none'
  },
  statusPill: {
    fontSize: typography.size.xs,
    fontWeight: typography.weight.semibold,
    padding: '2px 8px',
    borderRadius: radius.lg,
    textTransform: 'uppercase',
    flexShrink: 0
  },
  tag: {
    backgroundColor: colors.primarySoft,
    color: colors.primary,
    fontSize: typography.size.sm,
    fontWeight: typography.weight.semibold,
    padding: '2px 8px',
    borderRadius: radius.xs
  },
  navItemActive: {
    backgroundColor: colors.primaryLight,
    color: colors.primary,
    fontWeight: typography.weight.bold
  },
  toast: {
    position: 'fixed',
    top: '20px',
    right: '20px',
    padding: '12px 20px',
    borderRadius: radius.md,
    boxShadow: shadows.modal,
    display: 'flex',
    alignItems: 'center',
    gap: spacing.lg,
    fontFamily: typography.fontFamily,
    fontSize: typography.size.base,
    fontWeight: typography.weight.semibold,
    zIndex: 99999
  },
  toastCloseBtn: {
    background: 'none',
    border: 'none',
    color: 'inherit',
    fontSize: '18px',
    fontWeight: typography.weight.bold,
    cursor: 'pointer'
  },
  modalOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundColor: overlay.backdrop,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 9999
  },
  modalContainer: {
    backgroundColor: colors.white,
    borderRadius: radius.lg,
    boxShadow: shadows.modal,
    fontFamily: typography.fontFamily,
    overflow: 'hidden'
  },
  alertWarning: {
    backgroundColor: colors.warningBg,
    color: colors.warningText,
    border: `1px solid ${colors.warningBorder}`,
    fontSize: typography.size.md
  },
  alertError: {
    backgroundColor: colors.dangerSoft,
    color: colors.dangerDark,
    border: `1px solid ${colors.dangerBorder}`,
    fontSize: typography.size.md
  },
  pageContainer: {
    fontFamily: typography.fontFamily,
    backgroundColor: colors.bgPage,
    color: colors.text,
    boxSizing: 'border-box'
  }
};

export const theme = {
  colors,
  typography,
  radius,
  shadows,
  spacing,
  notification,
  overlay,
  statusColors,
  statusDotColors,
  components
};

export default theme;
