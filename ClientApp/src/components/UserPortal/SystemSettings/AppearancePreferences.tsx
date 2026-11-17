import * as React from 'react';
import { useState, useEffect } from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { NoticeBar, getContextInfo } from './SettingsShared';
import SettingsSelect from './SettingsSelect';

export interface IAppearancePreferencesProps {
  context: any;
}

export type DateFormat = 'DD/MM/YYYY' | 'MM/DD/YYYY' | 'YYYY-MM-DD';
export type TableDensity = 'comfortable' | 'compact';

export interface IAppearanceSettings {
  dateFormat: DateFormat;
  defaultPage: string;
  tableDensity: TableDensity;
  rowsPerPage: number;
}

export const DEFAULT_APPEARANCE_SETTINGS: IAppearanceSettings = {
  dateFormat: 'DD/MM/YYYY',
  defaultPage: 'Home',
  tableDensity: 'comfortable',
  rowsPerPage: 10
};

// Preferences are stored per user, in the browser only.
const getStorageKey = (email: string) => `dms.appearance.${email || 'anonymous'}`;

// Other modules can call this to read the saved preferences of a user.
export const loadAppearanceSettings = (email: string): IAppearanceSettings => {
  try {
    const raw = window.localStorage.getItem(getStorageKey(email.toLowerCase()));
    if (!raw) return { ...DEFAULT_APPEARANCE_SETTINGS };
    const stored = JSON.parse(raw);
    // Only known keys are read, so any old saved value (such as a removed setting) is ignored.
    const result: IAppearanceSettings = { ...DEFAULT_APPEARANCE_SETTINGS };
    (Object.keys(DEFAULT_APPEARANCE_SETTINGS) as (keyof IAppearanceSettings)[]).forEach(key => {
      if (stored[key] !== undefined) {
        (result as any)[key] = stored[key];
      }
    });
    return result;
  } catch (error) {
    return { ...DEFAULT_APPEARANCE_SETTINGS };
  }
};

// Other modules can call this to show a date the way the user prefers.
export const formatDateWithPreference = (date: Date, format: DateFormat): string => {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear());
  switch (format) {
    case 'MM/DD/YYYY':
      return `${month}/${day}/${year}`;
    case 'YYYY-MM-DD':
      return `${year}-${month}-${day}`;
    default:
      return `${day}/${month}/${year}`;
  }
};

interface IOption<T> {
  value: T;
  label: string;
}

const DATE_FORMAT_OPTIONS: IOption<DateFormat>[] = [
  { value: 'DD/MM/YYYY', label: 'DD/MM/YYYY' },
  { value: 'MM/DD/YYYY', label: 'MM/DD/YYYY' },
  { value: 'YYYY-MM-DD', label: 'YYYY-MM-DD' }
];

// The ids match the menu ids used by the main dashboard.
const DEFAULT_PAGE_OPTIONS: IOption<string>[] = [
  { value: 'Home', label: 'Dashboard' },
  { value: 'MyFiles', label: 'Documents' },
  { value: 'Approvals', label: 'Approvals' },
  { value: 'WorkflowSetting', label: 'Processes' }
];

const DENSITY_OPTIONS: IOption<TableDensity>[] = [
  { value: 'comfortable', label: 'Comfortable' },
  { value: 'compact', label: 'Compact' }
];

const ROWS_OPTIONS: IOption<number>[] = [
  { value: 10, label: '10 rows' },
  { value: 25, label: '25 rows' },
  { value: 50, label: '50 rows' }
];

// One labelled row: description on the left, control on the right.
const PreferenceRow = ({ label, hint, children, isLast = false }: { label: string; hint: string; children: React.ReactNode; isLast?: boolean }) => (
  <div
    style={{
      display: 'grid',
      gridTemplateColumns: 'minmax(180px, 1fr) minmax(0, 1.2fr)',
      alignItems: 'center',
      gap: '16px',
      padding: '16px 0',
      borderBottom: isLast ? 'none' : `1px solid ${COLORS.backgroundAlt}`
    }}
  >
    <div>
      <div style={{ fontSize: '14px', fontWeight: 600, color: COLORS.textStrong }}>{label}</div>
      <div style={{ fontSize: '12px', color: COLORS.textMuted, marginTop: '2px' }}>{hint}</div>
    </div>
    <div>{children}</div>
  </div>
);

// Segmented button group for short lists of choices.
function SegmentedControl<T extends string | number>({ options, value, onChange }: { options: IOption<T>[]; value: T; onChange: (next: T) => void }) {
  return (
    <div style={{ display: 'inline-flex', gap: '4px', backgroundColor: COLORS.backgroundAlt, borderRadius: '10px', padding: '3px', flexWrap: 'wrap' }}>
      {options.map(option => {
        const isActive = option.value === value;
        return (
          <button
            key={String(option.value)}
            type="button"
            onClick={() => onChange(option.value)}
            aria-pressed={isActive}
            style={{
              border: 'none',
              borderRadius: '8px',
              padding: '7px 16px',
              fontSize: '12.5px',
              fontWeight: isActive ? 700 : 600,
              backgroundColor: isActive ? COLORS.primary : 'transparent',
              color: isActive ? '#fff' : COLORS.textMuted,
              cursor: 'pointer',
              fontFamily: FONT_FAMILY
            }}
          >
            {option.label}
          </button>
        );
      })}
    </div>
  );
}

const AppearancePreferences: React.FC<IAppearancePreferencesProps> = ({ context }) => {
  const { userEmail } = getContextInfo(context);

  const [savedSettings, setSavedSettings] = useState<IAppearanceSettings>(() => loadAppearanceSettings(userEmail));
  const [draft, setDraft] = useState<IAppearanceSettings>(savedSettings);
  const [statusMessage, setStatusMessage] = useState<string>('');

  // Reload when the signed-in user changes.
  useEffect(() => {
    const loaded = loadAppearanceSettings(userEmail);
    setSavedSettings(loaded);
    setDraft(loaded);
  }, [userEmail]);

  // Hide the "saved" message after a few seconds.
  useEffect(() => {
    if (!statusMessage) return;
    const timer = setTimeout(() => setStatusMessage(''), 3500);
    return () => clearTimeout(timer);
  }, [statusMessage]);

  const isDirty = JSON.stringify(draft) !== JSON.stringify(savedSettings);

  const updateDraft = <K extends keyof IAppearanceSettings>(key: K, value: IAppearanceSettings[K]) => {
    setDraft(prev => ({ ...prev, [key]: value }));
  };

  const handleSave = () => {
    try {
      window.localStorage.setItem(getStorageKey(userEmail), JSON.stringify(draft));
      setSavedSettings(draft);
      setStatusMessage('Preferences saved.');
    } catch (error) {
      setStatusMessage('Could not save preferences. Browser storage is not available.');
    }
  };

  const handleResetToDefaults = () => {
    setDraft({ ...DEFAULT_APPEARANCE_SETTINGS });
  };

  const todayExample = formatDateWithPreference(new Date(), draft.dateFormat);

  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      <div style={{ ...sharedStyles.card, padding: '4px 20px' }}>
        <PreferenceRow label="Date format" hint={`Example: ${todayExample}`}>
          <div style={{ maxWidth: '280px' }}>
            <SettingsSelect options={DATE_FORMAT_OPTIONS} value={draft.dateFormat} onChange={value => updateDraft('dateFormat', value)} />
          </div>
        </PreferenceRow>

        <PreferenceRow label="Dashboard" hint="The page that opens first when you start DMS.">
          <div style={{ maxWidth: '280px' }}>
            <SettingsSelect options={DEFAULT_PAGE_OPTIONS} value={draft.defaultPage} onChange={value => updateDraft('defaultPage', value)} />
          </div>
        </PreferenceRow>

        <PreferenceRow label="Table density" hint="Row spacing in document and approval tables.">
          <SegmentedControl options={DENSITY_OPTIONS} value={draft.tableDensity} onChange={value => updateDraft('tableDensity', value)} />
        </PreferenceRow>

        <PreferenceRow label="Rows per page" hint="How many rows tables show at once." isLast>
          <div style={{ maxWidth: '280px' }}>
            <SettingsSelect options={ROWS_OPTIONS} value={draft.rowsPerPage} onChange={value => updateDraft('rowsPerPage', value)} />
          </div>
        </PreferenceRow>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginTop: '18px', flexWrap: 'wrap' }}>
        <button
          type="button"
          onClick={handleSave}
          disabled={!isDirty}
          style={{ ...sharedStyles.primaryButton, opacity: isDirty ? 1 : 0.5, cursor: isDirty ? 'pointer' : 'not-allowed' }}
        >
          Save preferences
        </button>
        <button type="button" onClick={handleResetToDefaults} style={sharedStyles.secondaryButton}>
          Reset to defaults
        </button>
        {statusMessage && (
          <span role="status" style={{ fontSize: '12.5px', fontWeight: 600, color: statusMessage.startsWith('Could not') ? COLORS.danger : COLORS.primary }}>
            {statusMessage}
          </span>
        )}
      </div>

      <div style={{ marginTop: '18px' }}>
        <NoticeBar title="Preferences are saved in this browser for your account only." text="They do not change what other users see." />
      </div>
    </div>
  );
};

export default AppearancePreferences;
