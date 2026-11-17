import * as React from 'react';
import { photoUrlFor } from '../../../spfx-shim';
import { useState, useEffect, useRef } from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { SVGIcon, MicrosoftLogo, NoticeBar, getContextInfo } from './SettingsShared';

export interface IUserProfileProps {
  context: any;
}

interface IProfileData {
  displayName: string;
  email: string;
  microsoftAccount: string;
  organization: string;
  department: string;
  designation: string;
  photoUrl: string;
}

const EMPTY_VALUE = '—';

// Builds initials such as "MC" from "Mihir Chauhan".
const getInitials = (name: string): string => {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

const UserProfile: React.FC<IUserProfileProps> = ({ context }) => {
  const info = getContextInfo(context);

  // Start with what SharePoint already knows so the page is never empty,
  // then enrich it from Microsoft Graph once the request completes.
  const [profile, setProfile] = useState<IProfileData>({
    displayName: info.userName,
    email: info.userEmail,
    microsoftAccount: info.userEmail,
    organization: '',
    department: '',
    designation: '',
    photoUrl: ''
  });
  const [isPhotoBroken, setIsPhotoBroken] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  // Keeps track of the blob URL so it can be released when the page closes.
  const blobUrlRef = useRef<string>('');

  useEffect(() => {
    if (!context) return;
    let isCancelled = false;

    const loadProfile = async () => {
      let graphClient: any = null;
      try {
        graphClient = await context.msGraphClientFactory.getClient('3');
      } catch (clientError) {
        console.error('Could not create Microsoft Graph client:', clientError);
      }

      if (graphClient) {
        // 1. Profile details
        try {
          const me = await graphClient
            .api('/me')
            .select('displayName,mail,userPrincipalName,companyName,department,jobTitle')
            .get();

          if (!isCancelled) {
            setProfile(prev => ({
              ...prev,
              displayName: me.displayName || prev.displayName,
              email: me.mail || me.userPrincipalName || prev.email,
              microsoftAccount: me.userPrincipalName || me.mail || prev.microsoftAccount,
              organization: me.companyName || prev.organization,
              department: me.department || prev.department,
              designation: me.jobTitle || prev.designation
            }));
          }
        } catch (profileError) {
          console.error('Could not read profile from Microsoft Graph:', profileError);
        }

        // 2. Organization name from the tenant, only when the user has no company name set
        try {
          const orgResponse = await graphClient.api('/organization').select('displayName').get();
          const tenantName: string = orgResponse?.value?.[0]?.displayName || '';
          if (!isCancelled && tenantName) {
            setProfile(prev => ({ ...prev, organization: prev.organization || tenantName }));
          }
        } catch (orgError) {
          // Reading the organization may not be permitted; the field simply stays empty.
        }

        // 3. Profile photo
        try {
          const photoBlob = await graphClient.api('/me/photo/$value').responseType('blob').get();
          if (!isCancelled) {
            const objectUrl = URL.createObjectURL(photoBlob);
            blobUrlRef.current = objectUrl;
            setProfile(prev => ({ ...prev, photoUrl: objectUrl }));
          }
        } catch (photoError) {
          const statusCode = (photoError as any)?.statusCode;
          // 404 means the user has no photo, so initials are shown instead.
          // Any other error falls back to the SharePoint photo endpoint.
          if (!isCancelled && statusCode !== 404 && info.siteUrl && info.userEmail) {
            setProfile(prev => ({
              ...prev,
              photoUrl: photoUrlFor(info.userEmail)
            }));
          }
        }
      } else if (info.siteUrl && info.userEmail) {
        // Graph is not available at all, so use the SharePoint photo endpoint.
        setProfile(prev => ({
          ...prev,
          photoUrl: photoUrlFor(info.userEmail)
        }));
      }

      if (!isCancelled) setIsLoading(false);
    };

    loadProfile();

    return () => {
      isCancelled = true;
      if (blobUrlRef.current) {
        URL.revokeObjectURL(blobUrlRef.current);
        blobUrlRef.current = '';
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [context]);

  const detailRows: { iconName: string; label: string; value: string }[] = [
    { iconName: 'user', label: 'Full Name', value: profile.displayName },
    { iconName: 'mail', label: 'Email Address', value: profile.email },
    { iconName: 'microsoft', label: 'Microsoft Account', value: profile.microsoftAccount },
    { iconName: 'building', label: 'Organization', value: profile.organization },
    { iconName: 'users', label: 'Department', value: profile.department },
    { iconName: 'briefcase', label: 'Designation', value: profile.designation }
  ];

  const showPhoto = profile.photoUrl !== '' && !isPhotoBroken;

  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      <div style={{ border: `1px solid ${COLORS.border}`, borderRadius: '16px', overflow: 'hidden' }}>
        {/* Identity banner */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '22px', padding: '24px', backgroundColor: COLORS.primaryTint, flexWrap: 'wrap' }}>
          {showPhoto ? (
            <img
              src={profile.photoUrl}
              alt={profile.displayName}
              onError={() => setIsPhotoBroken(true)}
              style={{ width: '88px', height: '88px', borderRadius: '50%', objectFit: 'cover', border: `3px solid ${COLORS.surface}`, boxShadow: '0 2px 8px rgba(0,0,0,0.12)' }}
            />
          ) : (
            <div
              style={{
                width: '88px',
                height: '88px',
                borderRadius: '50%',
                backgroundColor: COLORS.primary,
                color: '#fff',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: '30px',
                fontWeight: 700,
                letterSpacing: '1px',
                border: `3px solid ${COLORS.surface}`,
                boxShadow: '0 2px 8px rgba(0,0,0,0.12)'
              }}
            >
              {getInitials(profile.displayName)}
            </div>
          )}

          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: '22px', fontWeight: 700, color: COLORS.textStrong }}>
              {profile.displayName || (isLoading ? 'Loading...' : EMPTY_VALUE)}
            </div>
            <div style={{ fontSize: '14px', color: COLORS.textMuted, marginTop: '2px', wordBreak: 'break-all' }}>{profile.email}</div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '10px', flexWrap: 'wrap' }}>
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', color: COLORS.textBody }}>
                <MicrosoftLogo size={14} />
                Signed in with Microsoft
              </span>
              <span style={{ ...sharedStyles.pill, backgroundColor: COLORS.primarySoft, color: COLORS.primary }}>Office 365</span>
            </div>
          </div>
        </div>

        {/* Detail rows */}
        <div style={{ padding: '4px 20px' }}>
          {detailRows.map((row, index) => (
            <div
              key={row.label}
              style={{
                display: 'grid',
                gridTemplateColumns: 'minmax(140px, 190px) minmax(0, 1fr)',
                alignItems: 'center',
                gap: '12px',
                padding: '14px 0',
                borderBottom: index === detailRows.length - 1 ? 'none' : `1px solid ${COLORS.backgroundAlt}`
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '14px', color: COLORS.textMuted, fontSize: '13.5px' }}>
                <span
                  style={{
                    width: '36px',
                    height: '36px',
                    borderRadius: '10px',
                    backgroundColor: COLORS.primaryTint,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    flexShrink: 0
                  }}
                >
                  {row.iconName === 'microsoft' ? <MicrosoftLogo size={16} /> : <SVGIcon name={row.iconName} color={COLORS.primary} size={18} />}
                </span>
                {row.label}
              </div>
              <div style={{ fontSize: '14px', color: COLORS.textStrong, wordBreak: 'break-word' }}>
                {row.value || EMPTY_VALUE}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ marginTop: '16px' }}>
        <NoticeBar
          title="Your profile information is managed through Microsoft 365 and cannot be edited here."
          text="For changes, please update it from your Microsoft account."
        />
      </div>
    </div>
  );
};

export default UserProfile;
