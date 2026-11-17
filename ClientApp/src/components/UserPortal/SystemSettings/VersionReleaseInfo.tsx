import * as React from 'react';
import { useState } from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { SVGIcon } from './SettingsShared';
import { APP_NAME, RELEASES, ReleaseType } from './settingsConfig';

// Formats an ISO date (YYYY-MM-DD) as a readable date such as "21 Oct 2026".
const formatReadableDate = (isoDate: string): string => {
  const parsed = new Date(`${isoDate}T00:00:00`);
  if (isNaN(parsed.getTime())) return isoDate;
  return parsed.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

// Badge colours for each kind of release.
const RELEASE_TYPE_COLORS: { [key in ReleaseType]: { bg: string; fg: string } } = {
  Major: { bg: COLORS.primarySoft, fg: COLORS.primary },
  Minor: { bg: COLORS.infoSoft, fg: COLORS.info },
  Patch: { bg: COLORS.warningSoft, fg: COLORS.warning }
};

const VersionReleaseInfo: React.FC = () => {
  const [showAllHistory, setShowAllHistory] = useState<boolean>(false);

  const currentRelease = RELEASES.length > 0 ? RELEASES[0] : undefined;
  const olderReleases = RELEASES.slice(1);
  const visibleOlderReleases = showAllHistory ? olderReleases : olderReleases.slice(0, 3);

  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      {!currentRelease ? (
        <p style={sharedStyles.bodyText}>No release information is available yet.</p>
      ) : (
        <>
          {/* Current version */}
          <div style={{ ...sharedStyles.card, padding: '22px', backgroundColor: COLORS.primaryTint, marginBottom: '22px' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px', flexWrap: 'wrap' }}>
              <div>
                <div style={{ fontSize: '12px', fontWeight: 600, color: COLORS.textMuted }}>Current version</div>
                <div style={{ fontSize: '30px', fontWeight: 700, color: COLORS.primary, marginTop: '2px' }}>
                  {APP_NAME} {currentRelease.version}
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <span style={{ ...sharedStyles.pill, backgroundColor: COLORS.surface, color: COLORS.textBody, border: `1px solid ${COLORS.border}` }}>
                  <SVGIcon name="calendar" color={COLORS.textMuted} size={14} />
                  Released {formatReadableDate(currentRelease.date)}
                </span>
                <span style={{ ...sharedStyles.pill, backgroundColor: RELEASE_TYPE_COLORS[currentRelease.type].bg, color: RELEASE_TYPE_COLORS[currentRelease.type].fg }}>
                  {currentRelease.type} release
                </span>
              </div>
            </div>
          </div>

          {/* Latest updates */}
          <h3 style={sharedStyles.sectionTitle}>Latest updates</h3>
          <div style={{ ...sharedStyles.card, padding: '16px 20px', marginBottom: '24px' }}>
            <ul style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {currentRelease.highlights.map(highlight => (
                <li key={highlight} style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', fontSize: '13.5px', color: COLORS.textBody, lineHeight: 1.55 }}>
                  <span style={{ marginTop: '2px' }}>
                    <SVGIcon name="check" color={COLORS.primary} size={16} />
                  </span>
                  {highlight}
                </li>
              ))}
            </ul>
          </div>

          {/* Change history */}
          <h3 style={sharedStyles.sectionTitle}>Change history</h3>
          {olderReleases.length === 0 ? (
            <p style={{ ...sharedStyles.bodyText, color: COLORS.textMuted }}>This is the first release. Earlier versions will appear here after the next update.</p>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {visibleOlderReleases.map(release => (
                <div key={release.version} style={{ ...sharedStyles.card, padding: '16px 20px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', marginBottom: '8px' }}>
                    <span style={{ fontSize: '15px', fontWeight: 700, color: COLORS.textStrong }}>{release.version}</span>
                    <span style={{ ...sharedStyles.pill, backgroundColor: RELEASE_TYPE_COLORS[release.type].bg, color: RELEASE_TYPE_COLORS[release.type].fg }}>{release.type}</span>
                    <span style={{ fontSize: '12px', color: COLORS.textMuted }}>{formatReadableDate(release.date)}</span>
                  </div>
                  <ul style={{ margin: 0, paddingLeft: '20px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    {release.highlights.map(highlight => (
                      <li key={highlight} style={{ ...sharedStyles.bodyText, fontSize: '13px' }}>{highlight}</li>
                    ))}
                  </ul>
                </div>
              ))}

              {olderReleases.length > 3 && (
                <button type="button" onClick={() => setShowAllHistory(prev => !prev)} style={{ ...sharedStyles.secondaryButton, alignSelf: 'flex-start' }}>
                  {showAllHistory ? 'Show fewer' : `Show all ${olderReleases.length} earlier versions`}
                </button>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default VersionReleaseInfo;
