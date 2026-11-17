import * as React from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { IconBadge, NoticeBar } from './SettingsShared';
import { PRIVACY_SECTIONS } from './settingsConfig';
 
const PrivacyPolicy: React.FC = () => {
  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginBottom: '20px' }}>
        {PRIVACY_SECTIONS.map(section => (
          <div key={section.title} style={{ ...sharedStyles.card, padding: '18px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '10px' }}>
              <IconBadge iconName={section.iconName} size={38} />
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: COLORS.textStrong }}>{section.title}</h3>
            </div>
            <ul style={{ margin: 0, paddingLeft: '20px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              {section.points.map(point => (
                <li key={point} style={{ ...sharedStyles.bodyText, fontSize: '13px' }}>{point}</li>
              ))}
            </ul>
          </div>
        ))}
      </div>
 
      <NoticeBar 
        title="Your documents are protected by your existing SharePoint permissions." 
        text="DMS cannot show you a document that you do not already have access to." 
      />
    </div>
  );
};
 
export default PrivacyPolicy;