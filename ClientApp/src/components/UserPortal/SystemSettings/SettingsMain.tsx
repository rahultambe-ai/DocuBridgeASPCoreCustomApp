import * as React from 'react';
import { useState } from 'react';
import UserProfile from './UserProfile';
import AppearancePreferences from './AppearancePreferences';
import About from './About';
import HelpSupport from './HelpSupport';
import PrivacyPolicy from './PrivacyPolicy';
import VersionReleaseInfo from './VersionReleaseInfo';
import { COLORS, FONT_FAMILY, NARROW_BREAKPOINT, sharedStyles } from './settingsTheme';
import { IconBadge, SVGIcon, useIsNarrow } from './SettingsShared';

export interface ISettingsProps {
  context: any;
}

type SectionId = 'profile' | 'appearance' | 'about' | 'help' | 'privacy' | 'version';

interface ISectionDefinition {
  id: SectionId;
  title: string;
  description: string;
  iconName: string;
}

// To add a new Settings page: add one entry here and one case in renderSection().
const SECTIONS: ISectionDefinition[] = [
  { id: 'profile', title: 'My Profile', description: 'View your name, official email address and Microsoft account details.', iconName: 'user' },
  { id: 'appearance', title: 'Appearance & Preferences', description: 'Date format, dashboard and table display preferences.', iconName: 'sliders' },
  { id: 'about', title: 'About DMS', description: 'Application overview, purpose and product information.', iconName: 'contrast' },
  { id: 'help', title: 'Help & Support', description: 'Contact IT support, report issues and get assistance with documents or approvals.', iconName: 'headset' },
  { id: 'privacy', title: 'Privacy Policy', description: 'Data privacy, document access, information usage and privacy terms.', iconName: 'shield' },
  { id: 'version', title: 'Version & Release Information', description: 'Current version, release date, latest updates and change history.', iconName: 'infoCircle' }
];

const Settings: React.FC<ISettingsProps> = ({ context }) => {
  const [activeSection, setActiveSection] = useState<SectionId>('profile');
  const [hoveredSection, setHoveredSection] = useState<SectionId | null>(null);
  const isNarrow = useIsNarrow(NARROW_BREAKPOINT);

  const renderSection = (): React.ReactNode => {
    switch (activeSection) {
      case 'profile':
        return <UserProfile context={context} />;
      case 'appearance':
        return <AppearancePreferences context={context} />;
      case 'about':
        return <About context={context} />;
      case 'help':
        return <HelpSupport context={context} />;
      case 'privacy':
        return <PrivacyPolicy />;
      case 'version':
        return <VersionReleaseInfo />;
      default:
        return null;
    }
  };

  return (
    <div style={{ fontFamily: FONT_FAMILY, color: COLORS.textBody }}>
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: isNarrow ? 'minmax(0, 1fr)' : 'minmax(240px, 280px) minmax(0, 1fr)',
          gap: '20px',
          alignItems: 'start'
        }}
      >
        {/* Left column: list of settings pages */}
        <nav style={{ ...sharedStyles.card, padding: '10px' }} aria-label="Settings sections">
          {SECTIONS.map((section, index) => {
            const isActive = section.id === activeSection;
            const isHovered = section.id === hoveredSection;
            const isLast = index === SECTIONS.length - 1;

            return (
              <button
                key={section.id}
                type="button"
                onClick={() => setActiveSection(section.id)}
                onMouseEnter={() => setHoveredSection(section.id)}
                onMouseLeave={() => setHoveredSection(null)}
                aria-current={isActive ? 'page' : undefined}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px',
                  width: '100%',
                  textAlign: 'left',
                  padding: '12px',
                  outline: 'none',
                  border: 'none',
                  borderBottom: isLast || isActive ? '1px solid transparent' : `1px solid ${COLORS.backgroundAlt}`,
                  borderRadius: isActive ? '14px' : '0',
                  backgroundColor: isActive ? COLORS.primarySoft : isHovered ? COLORS.background : 'transparent',
                  cursor: 'pointer',
                  fontFamily: FONT_FAMILY,
                  transition: 'background-color 0.15s ease',
                  boxSizing: 'border-box'
                }}
              >
                <IconBadge iconName={section.iconName} size={40} active={isActive} />
                <span style={{ flex: 1, minWidth: 0 }}>
                  <span style={{ display: 'block', fontSize: '14.5px', fontWeight: 700, color: COLORS.textStrong }}>{section.title}</span>
                  <span style={{ display: 'block', fontSize: '12px', color: COLORS.textMuted, marginTop: '2px', lineHeight: 1.45 }}>{section.description}</span>
                </span>
                <SVGIcon name="chevronRight" color={isActive ? COLORS.primary : COLORS.textFaint} size={18} />
              </button>
            );
          })}
        </nav>

        {/* Right column: the selected page */}
        <section style={{ ...sharedStyles.card, padding: '20px', minHeight: '420px' }}>
          {renderSection()}
        </section>
      </div>
    </div>
  );
};

export default Settings;
