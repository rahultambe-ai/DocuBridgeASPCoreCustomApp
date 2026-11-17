import * as React from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { IconBadge, getContextInfo } from './SettingsShared';
import { APP_NAME, APP_FULL_NAME, APP_TAGLINE, ABOUT_DESCRIPTION, ABOUT_HIGHLIGHTS, RELEASES } from './settingsConfig';

export interface IAboutProps {
  context: any;
}

const About: React.FC<IAboutProps> = ({ context }) => {
  const info = getContextInfo(context);
  const currentRelease = RELEASES.length > 0 ? RELEASES[0] : undefined;

  // Product facts. Site values are read live from the SharePoint page context.
  const productRows: { label: string; value: string }[] = [
    { label: 'Product', value: `${APP_NAME} - ${APP_FULL_NAME}` },
    { label: 'Version', value: currentRelease ? currentRelease.version : '-' },
    { label: 'Platform', value: 'SharePoint Framework (SPFx) on Microsoft 365' },
    { label: 'Site', value: info.siteTitle || '-' },
    { label: 'Site URL', value: info.siteUrl || '-' },
    { label: 'Tenant ID', value: info.tenantId || '-' }
  ];

  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      {/* Overview */}
      <div style={{ ...sharedStyles.card, padding: '22px', backgroundColor: COLORS.primaryTint, marginBottom: '20px' }}>
        <div style={{ fontSize: '18px', fontWeight: 700, color: COLORS.primary }}>{APP_FULL_NAME}</div>
        <div style={{ fontSize: '13px', color: COLORS.textMuted, marginTop: '4px' }}>{APP_TAGLINE}</div>
        <p style={{ ...sharedStyles.bodyText, marginTop: '14px' }}>{ABOUT_DESCRIPTION}</p>
      </div>

      {/* Purpose and key capabilities */}
      <h3 style={sharedStyles.sectionTitle}>What DMS is for</h3>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '14px', marginBottom: '24px' }}>
        {ABOUT_HIGHLIGHTS.map(item => (
          <div key={item.title} style={{ ...sharedStyles.card, padding: '16px', display: 'flex', gap: '14px', alignItems: 'flex-start' }}>
            <IconBadge iconName={item.iconName} size={40} />
            <div>
              <div style={{ fontSize: '14px', fontWeight: 700, color: COLORS.textStrong }}>{item.title}</div>
              <div style={{ fontSize: '12.5px', color: COLORS.textMuted, marginTop: '3px', lineHeight: 1.55 }}>{item.text}</div>
            </div>
          </div>
        ))}
      </div>

      {/* Product information */}
      <h3 style={sharedStyles.sectionTitle}>Product information</h3>
      <div style={{ ...sharedStyles.card, padding: '4px 20px' }}>
        {productRows.map((row, index) => (
          <div
            key={row.label}
            style={{
              display: 'grid',
              gridTemplateColumns: 'minmax(110px, 180px) minmax(0, 1fr)',
              gap: '12px',
              padding: '13px 0',
              borderBottom: index === productRows.length - 1 ? 'none' : `1px solid ${COLORS.backgroundAlt}`
            }}
          >
            <div style={{ fontSize: '13px', color: COLORS.textMuted }}>{row.label}</div>
            <div style={{ fontSize: '13.5px', color: COLORS.textStrong, wordBreak: 'break-all' }}>{row.value}</div>
          </div>
        ))}
      </div>
    </div>
  );
};

export default About;
