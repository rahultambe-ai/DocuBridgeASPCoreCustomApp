import * as React from 'react';

// ---------------------------------------------------------------------------
// Shared design tokens for every Settings module.
// These values match the DMS dashboard theme (green #005c42, Segoe UI).
// Change a colour here and all Settings screens update together.
// ---------------------------------------------------------------------------

export const COLORS = {
  primary: '#005c42',
  primaryHover: '#00492f',
  primarySoft: '#e6f4ea',
  primaryTint: '#f0f9f4',

  textStrong: '#1e293b',
  textBody: '#334155',
  textNav: '#384c66',
  textMuted: '#64748b',
  textFaint: '#94a3b8',

  border: '#e2e8f0',
  borderStrong: '#cbd5e1',

  surface: '#ffffff',
  background: '#f8fafc',
  backgroundAlt: '#f1f5f9',

  danger: '#dc2626',
  dangerSoft: '#fee2e2',
  warning: '#d97706',
  warningSoft: '#fef3c7',
  info: '#0284c7',
  infoSoft: '#e0f2fe'
};

export const FONT_FAMILY = '"Segoe UI", Arial, sans-serif';

// Width at which the Settings layout switches from two columns to one.
export const NARROW_BREAKPOINT = 1000;

// Reusable style objects used across the Settings modules.
export const sharedStyles: { [key: string]: React.CSSProperties } = {
  card: {
    backgroundColor: COLORS.surface,
    border: `1px solid ${COLORS.border}`,
    borderRadius: '16px',
    boxSizing: 'border-box'
  },
  panelHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
    marginBottom: '20px'
  },
  panelTitle: {
    margin: 0,
    fontSize: '20px',
    fontWeight: 700,
    color: COLORS.textStrong,
    fontFamily: FONT_FAMILY
  },
  panelSubtitle: {
    margin: '2px 0 0 0',
    fontSize: '13px',
    color: COLORS.textMuted,
    fontFamily: FONT_FAMILY
  },
  sectionTitle: {
    margin: '0 0 12px 0',
    fontSize: '14px',
    fontWeight: 700,
    color: COLORS.textStrong,
    fontFamily: FONT_FAMILY
  },
  bodyText: {
    margin: 0,
    fontSize: '13.5px',
    lineHeight: 1.65,
    color: COLORS.textBody,
    fontFamily: FONT_FAMILY
  },
  labelText: {
    fontSize: '12px',
    fontWeight: 600,
    color: COLORS.textMuted,
    fontFamily: FONT_FAMILY
  },
  primaryButton: {
    backgroundColor: COLORS.primary,
    color: '#fff',
    border: 'none',
    borderRadius: '8px',
    padding: '9px 18px',
    fontSize: '13px',
    fontWeight: 600,
    cursor: 'pointer',
    fontFamily: FONT_FAMILY
  },
  secondaryButton: {
    backgroundColor: COLORS.surface,
    color: COLORS.textBody,
    border: `1px solid ${COLORS.borderStrong}`,
    borderRadius: '8px',
    padding: '9px 18px',
    fontSize: '13px',
    fontWeight: 600,
    cursor: 'pointer',
    fontFamily: FONT_FAMILY
  },
  pill: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '6px',
    padding: '3px 10px',
    borderRadius: '999px',
    fontSize: '11.5px',
    fontWeight: 600,
    fontFamily: FONT_FAMILY
  }
};
