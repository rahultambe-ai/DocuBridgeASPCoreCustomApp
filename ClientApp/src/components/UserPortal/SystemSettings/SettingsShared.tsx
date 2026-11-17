import * as React from 'react';
import { useState, useEffect } from 'react';
import { COLORS, FONT_FAMILY } from './settingsTheme';

// ---------------------------------------------------------------------------
// Shared building blocks used by every Settings module:
// icons, the Microsoft logo, the panel header and a responsive-width hook.
// ---------------------------------------------------------------------------

const ICON_PATHS: { [key: string]: React.ReactNode } = {
  settings: (
    <>
      <circle cx="12" cy="12" r="3" />
      <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33h.01a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51h.01a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82v.01a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
    </>
  ),
  user: (
    <>
      <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
      <circle cx="12" cy="7" r="4" />
    </>
  ),
  sliders: (
    <>
      <line x1="4" y1="7" x2="20" y2="7" />
      <line x1="4" y1="17" x2="20" y2="17" />
      <circle cx="9" cy="7" r="2.5" fill="#fff" />
      <circle cx="15" cy="17" r="2.5" fill="#fff" />
    </>
  ),
  contrast: (
    <>
      <circle cx="12" cy="12" r="9" />
      <path d="M12 3a9 9 0 0 1 0 18z" fill="currentColor" />
    </>
  ),
  headset: (
    <>
      <path d="M4 14v-2a8 8 0 0 1 16 0v2" />
      <path d="M4 14h2.5a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1z" />
      <path d="M20 14h-2.5a1 1 0 0 0-1 1v3a1 1 0 0 0 1 1H19a1 1 0 0 0 1-1z" />
      <path d="M18 19a4 3 0 0 1-4 2h-1.5" />
    </>
  ),
  shield: (
    <path d="M12 3l8 3v6c0 4.5-3.2 8.2-8 9-4.8-.8-8-4.5-8-9V6z" />
  ),
  infoCircle: (
    <>
      <circle cx="12" cy="12" r="9" />
      <line x1="12" y1="11" x2="12" y2="16.5" />
      <line x1="12" y1="7.5" x2="12.01" y2="7.5" />
    </>
  ),
  chevronRight: <polyline points="9 6 15 12 9 18" />,
  arrowRight: (
    <>
      <line x1="5" y1="12" x2="19" y2="12" />
      <polyline points="12 5 19 12 12 19" />
    </>
  ),
  phone: (
    <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z" />
  ),
  chevronDown: <polyline points="6 9 12 15 18 9" />,
  mail: (
    <>
      <rect x="3" y="5" width="18" height="14" rx="2" />
      <polyline points="3 7 12 13 21 7" />
    </>
  ),
  building: (
    <>
      <rect x="5" y="3" width="14" height="18" rx="1.5" />
      <line x1="9" y1="8" x2="9.01" y2="8" />
      <line x1="15" y1="8" x2="15.01" y2="8" />
      <line x1="9" y1="12" x2="9.01" y2="12" />
      <line x1="15" y1="12" x2="15.01" y2="12" />
      <path d="M10 21v-4h4v4" />
    </>
  ),
  users: (
    <>
      <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
      <circle cx="9" cy="7" r="4" />
      <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
      <path d="M16 3.13a4 4 0 0 1 0 7.75" />
    </>
  ),
  briefcase: (
    <>
      <rect x="3" y="7" width="18" height="13" rx="2" />
      <path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2" />
      <line x1="3" y1="13" x2="21" y2="13" />
    </>
  ),
  check: <polyline points="20 6 9 17 4 12" />,
  file: (
    <>
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
    </>
  ),
  processes: (
    <>
      <rect x="3.5" y="3.5" width="17" height="17" rx="2.5" />
      <path d="M7.5 12.5h4.5v4.5H7.5z" fill="currentColor" stroke="none" />
    </>
  ),
  bug: (
    <>
      <rect x="8" y="8" width="8" height="12" rx="4" />
      <path d="M9 8a3 3 0 0 1 6 0" />
      <line x1="3" y1="13" x2="8" y2="13" />
      <line x1="16" y1="13" x2="21" y2="13" />
      <line x1="4" y1="7" x2="8.5" y2="10" />
      <line x1="20" y1="7" x2="15.5" y2="10" />
    </>
  ),
  copy: (
    <>
      <rect x="9" y="9" width="12" height="12" rx="2" />
      <path d="M5 15H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v1" />
    </>
  ),
  tag: (
    <>
      <path d="M20.6 13.4l-7.2 7.2a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8z" />
      <line x1="7.5" y1="7.5" x2="7.51" y2="7.5" />
    </>
  ),
  calendar: (
    <>
      <rect x="3" y="5" width="18" height="16" rx="2" />
      <line x1="3" y1="10" x2="21" y2="10" />
      <line x1="8" y1="3" x2="8" y2="7" />
      <line x1="16" y1="3" x2="16" y2="7" />
    </>
  ),
  globe: (
    <>
      <circle cx="12" cy="12" r="9" />
      <line x1="3" y1="12" x2="21" y2="12" />
      <path d="M12 3a14 14 0 0 1 0 18a14 14 0 0 1 0-18z" />
    </>
  ),
  link: (
    <>
      <path d="M10 13a5 5 0 0 0 7 0l3-3a5 5 0 0 0-7-7l-1 1" />
      <path d="M14 11a5 5 0 0 0-7 0l-3 3a5 5 0 0 0 7 7l1-1" />
    </>
  )
};

export const SVGIcon = ({ name, color = 'currentColor', size = 20 }: { name: string; color?: string; size?: number }) => (
  <svg
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke={color}
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }}
    aria-hidden="true"
  >
    {ICON_PATHS[name] || null}
  </svg>
);

// Official four-colour Microsoft mark, drawn with plain SVG squares.
export const MicrosoftLogo = ({ size = 16 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }} aria-hidden="true">
    <rect x="1" y="1" width="10" height="10" fill="#f25022" />
    <rect x="13" y="1" width="10" height="10" fill="#7fba00" />
    <rect x="1" y="13" width="10" height="10" fill="#00a4ef" />
    <rect x="13" y="13" width="10" height="10" fill="#ffb900" />
  </svg>
);

// Rounded square that holds an icon (used in the menu, headers and rows).
export const IconBadge = ({ iconName, size = 44, active = false }: { iconName: string; size?: number; active?: boolean }) => (
  <div
    style={{
      width: size,
      height: size,
      minWidth: size,
      borderRadius: Math.round(size * 0.3),
      backgroundColor: active ? COLORS.primarySoft : COLORS.primaryTint,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      color: COLORS.primary
    }}
  >
    <SVGIcon name={iconName} color={COLORS.primary} size={Math.round(size * 0.5)} />
  </div>
);

// Small rounded notice bar (info / success / warning).
export const NoticeBar = ({ tone = 'info', title, text }: { tone?: 'info' | 'success' | 'warning'; title: string; text?: string }) => {
  const toneColors = {
    info: { bg: COLORS.primaryTint, fg: COLORS.primary },
    success: { bg: COLORS.primarySoft, fg: COLORS.primary },
    warning: { bg: COLORS.warningSoft, fg: COLORS.warning }
  }[tone];

  return (
    <div style={{ display: 'flex', alignItems: 'flex-start', gap: '12px', padding: '14px 16px', borderRadius: '12px', backgroundColor: toneColors.bg, fontFamily: FONT_FAMILY }}>
      <SVGIcon name="infoCircle" color={toneColors.fg} size={18} />
      <div>
        <div style={{ fontSize: '12.5px', fontWeight: 700, color: toneColors.fg }}>{title}</div>
        {text && <div style={{ fontSize: '12px', color: COLORS.textMuted, marginTop: '2px' }}>{text}</div>}
      </div>
    </div>
  );
};

// Returns true while the viewport is narrower than the given breakpoint.
export const useIsNarrow = (breakpoint: number): boolean => {
  const [isNarrow, setIsNarrow] = useState<boolean>(typeof window !== 'undefined' ? window.innerWidth < breakpoint : false);

  useEffect(() => {
    const handleResize = () => setIsNarrow(window.innerWidth < breakpoint);
    window.addEventListener('resize', handleResize);
    handleResize();
    return () => window.removeEventListener('resize', handleResize);
  }, [breakpoint]);

  return isNarrow;
};

// Reads values from the SharePoint page context in one safe place.
export const getContextInfo = (context: any) => ({
  siteTitle: (context?.pageContext?.web?.title as string) || '',
  siteUrl: (context?.pageContext?.web?.absoluteUrl as string) || '',
  tenantId: (context?.pageContext?.aadInfo?.tenantId?.toString() as string) || '',
  userEmail: ((context?.pageContext?.user?.email as string) || '').toLowerCase(),
  userName: (context?.pageContext?.user?.displayName as string) || ''
});
