import * as React from 'react';
import { useState, useEffect } from 'react';
import { COLORS, FONT_FAMILY, sharedStyles } from './settingsTheme';
import { IconBadge, SVGIcon, getContextInfo } from './SettingsShared';
import SettingsSelect from './SettingsSelect';
import { openOutlookOnWeb, IMailMessage } from './SettingsMail';
import {
  APP_NAME,
  SUPPORT_EMAIL,
  SUPPORT_PHONE,
  FAQ_ITEMS,
  RELEASES,
  ISSUE_TYPES,
  AFFECTED_AREAS,
  ISSUE_PRIORITIES
} from './settingsConfig';

export interface IHelpSupportProps {
  context: any;
}

// Maximum lengths keep the email link short enough for every mail app.
const MAX_SUMMARY_LENGTH = 120;
const MAX_DETAILS_LENGTH = 600;

const toSelectOptions = (values: string[]) => values.map(value => ({ label: value, value }));

const HelpSupport: React.FC<IHelpSupportProps> = ({ context }) => {
  const info = getContextInfo(context);

  const [openFaqIndex, setOpenFaqIndex] = useState<number | null>(0);

  // "Report an issue" form state
  const [isReportOpen, setIsReportOpen] = useState<boolean>(false);
  const [issueType, setIssueType] = useState<string>(ISSUE_TYPES[0]);
  const [affectedArea, setAffectedArea] = useState<string>(AFFECTED_AREAS[0]);
  const [priority, setPriority] = useState<string>(ISSUE_PRIORITIES[1]);
  const [summary, setSummary] = useState<string>('');
  const [details, setDetails] = useState<string>('');
  const [steps, setSteps] = useState<string>('');
  const [showErrors, setShowErrors] = useState<boolean>(false);
  const [statusMessage, setStatusMessage] = useState<string>('');

  // Hide the status message after a few seconds.
  useEffect(() => {
    if (!statusMessage) return;
    const timer = setTimeout(() => setStatusMessage(''), 6000);
    return () => clearTimeout(timer);
  }, [statusMessage]);

  // Technical details that help IT support understand the problem quickly.
  // They are added to the email automatically, so the user never has to type them.
  const buildTechnicalDetails = (): string => {
    const version = RELEASES.length > 0 ? RELEASES[0].version : 'unknown';
    return [
      `Application: ${APP_NAME} ${version}`,
      `User: ${info.userName} (${info.userEmail})`,
      `Site: ${info.siteUrl}`,
      `Tenant ID: ${info.tenantId}`,
      `Browser: ${navigator.userAgent}`,
      `Time: ${new Date().toISOString()}`
    ].join('\r\n');
  };

  // Card 1: plain support request
  const buildSupportMessage = (): IMailMessage => ({
    to: SUPPORT_EMAIL,
    subject: `${APP_NAME} support request`,
    body: [
      'Hello IT Support,',
      '',
      'I need help with:',
      '',
      '',
      'Regards,',
      info.userName
    ].join('\r\n')
  });

  // Card 2: opens the report form
  const handleToggleReport = () => {
    setIsReportOpen(prev => !prev);
    setShowErrors(false);
  };

  const isSummaryMissing = summary.trim() === '';
  const isDetailsMissing = details.trim() === '';

  const resetReportForm = () => {
    setIssueType(ISSUE_TYPES[0]);
    setAffectedArea(AFFECTED_AREAS[0]);
    setPriority(ISSUE_PRIORITIES[1]);
    setSummary('');
    setDetails('');
    setSteps('');
    setShowErrors(false);
  };

  const buildReportMessage = (): IMailMessage => {
    const priorityShort = priority.split(' - ')[0];
    const subject = `${APP_NAME} issue [${priorityShort}] - ${summary.trim()}`;
    const body = [
      'Hello IT Support,',
      '',
      `I would like to report an issue in ${APP_NAME}.`,
      '',
      `Issue type: ${issueType}`,
      `Affected area: ${affectedArea}`,
      `Priority: ${priority}`,
      '',
      `Summary: ${summary.trim()}`,
      '',
      'What happened:',
      details.trim(),
      '',
      'Steps to repeat the problem:',
      steps.trim() || 'Not provided',
      '',
      '--- Technical details (added automatically, please keep) ---',
      buildTechnicalDetails()
    ].join('\r\n');

    return { to: SUPPORT_EMAIL, subject, body };
  };

  // Opens Outlook on the web with the report ready. Stops and shows errors if the form is incomplete.
  const handleSendReport = () => {
    if (isSummaryMissing || isDetailsMissing) {
      setShowErrors(true);
      return;
    }
    openOutlookOnWeb(buildReportMessage());
    setStatusMessage('Outlook is opening with the report ready. Press Send there to submit it.');
    resetReportForm();
    setIsReportOpen(false);
  };

  // Contact details shown inside the "Contact IT Support" card.
  const contactRows: { iconName: string; value: string }[] = [
    { iconName: 'mail', value: SUPPORT_EMAIL },
    { iconName: 'phone', value: SUPPORT_PHONE }
  ];

  // Short list shown inside the "Report an Issue" card.
  const reportPoints: string[] = ['Report bugs or errors', 'Suggest new features', 'Share feedback'];

  // Green text link with an arrow, shown at the bottom of each card.
  const renderCardLink = (label: string, onClick: () => void) => (
    <button type="button" onClick={onClick} style={styles.cardLink}>
      {label}
      <SVGIcon name="arrowRight" color={COLORS.primary} size={18} />
    </button>
  );

  const fieldLabel = (text: string, required?: boolean) => (
    <label style={styles.fieldLabel}>
      {text} {required && <span style={{ color: COLORS.danger }}>*</span>}
    </label>
  );

  return (
    <div style={{ fontFamily: FONT_FAMILY }}>
      <style>{`
        .hs-input:focus { border-color: #004d3d !important; box-shadow: 0 0 0 3px rgba(0, 77, 61, 0.10); }
      `}</style>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '16px', marginBottom: '18px' }}>
        {/* Card 1: Contact IT Support */}
        <div style={styles.supportCard}>
          <div style={styles.cardHeader}>
            <IconBadge iconName="mail" size={52} />
            <div style={{ minWidth: 0 }}>
              <div style={styles.cardTitle}>Contact IT Support</div>
              <div style={styles.cardText}>Reach out to our IT team for help with account access, system issues or general assistance.</div>
            </div>
          </div>

          <div style={styles.cardBody}>
            {contactRows.map(row => (
              <div key={row.iconName} style={styles.contactRow}>
                <span style={styles.smallBadge}>
                  <SVGIcon name={row.iconName} color={COLORS.primary} size={16} />
                </span>
                <span style={styles.contactValue}>{row.value}</span>
              </div>
            ))}
          </div>

          <div style={styles.cardFooter}>
            {renderCardLink('Contact support', () => openOutlookOnWeb(buildSupportMessage()))}
          </div>
        </div>

        {/* Card 2: Report an Issue */}
        <div style={styles.supportCard}>
          <div style={styles.cardHeader}>
            <IconBadge iconName="file" size={52} />
            <div style={{ minWidth: 0 }}>
              <div style={styles.cardTitle}>Report an Issue</div>
              <div style={styles.cardText}>Found a problem or have a suggestion? Let us know and we'll get it resolved.</div>
            </div>
          </div>

          <div style={styles.cardBody}>
            {reportPoints.map(point => (
              <div key={point} style={styles.contactRow}>
                <SVGIcon name="check" color={COLORS.primary} size={18} />
                <span style={styles.contactValue}>{point}</span>
              </div>
            ))}
          </div>

          <div style={styles.cardFooter}>
            {renderCardLink(isReportOpen ? 'Close form' : 'Report issue', handleToggleReport)}
          </div>
        </div>
      </div>

      {statusMessage && (
        <div role="status" style={{ marginBottom: '16px', padding: '12px 14px', borderRadius: '10px', backgroundColor: COLORS.primarySoft, color: COLORS.primary, fontSize: '12.5px', fontWeight: 600 }}>
          {statusMessage}
        </div>
      )}

      {/* Report an issue form */}
      {isReportOpen && (
        <div style={{ ...sharedStyles.card, padding: '20px', marginBottom: '22px' }}>
          <h3 style={{ ...sharedStyles.sectionTitle, marginBottom: '4px' }}>Report an issue</h3>
          <p style={{ ...sharedStyles.bodyText, fontSize: '12.5px', color: COLORS.textMuted, marginBottom: '16px' }}>
            Fill in the details below. Your email app will open with the report ready to send to IT support.
            Your name, site, browser and {APP_NAME} version are added automatically.
          </p>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(190px, 1fr))', gap: '14px', marginBottom: '14px' }}>
            <div style={styles.field}>
              {fieldLabel('Issue type')}
              <SettingsSelect<string> options={toSelectOptions(ISSUE_TYPES)} value={issueType} onChange={setIssueType} />
            </div>
            <div style={styles.field}>
              {fieldLabel('Where did it happen?')}
              <SettingsSelect<string> options={toSelectOptions(AFFECTED_AREAS)} value={affectedArea} onChange={setAffectedArea} />
            </div>
            <div style={styles.field}>
              {fieldLabel('Priority')}
              <SettingsSelect<string> options={toSelectOptions(ISSUE_PRIORITIES)} value={priority} onChange={setPriority} />
            </div>
          </div>

          <div style={{ ...styles.field, marginBottom: '14px' }}>
            {fieldLabel('Short summary', true)}
            <input
              className="hs-input"
              type="text"
              value={summary}
              maxLength={MAX_SUMMARY_LENGTH}
              placeholder="e.g. Approval stays on L1 after approving"
              onChange={event => setSummary(event.target.value)}
              style={{ ...styles.input, borderColor: showErrors && isSummaryMissing ? COLORS.danger : COLORS.borderStrong }}
            />
            {showErrors && isSummaryMissing && <span style={styles.errorText}>Please add a short summary.</span>}
          </div>

          <div style={{ ...styles.field, marginBottom: '14px' }}>
            {fieldLabel('What happened?', true)}
            <textarea
              className="hs-input"
              value={details}
              maxLength={MAX_DETAILS_LENGTH}
              rows={4}
              placeholder="Describe the problem and what you expected to happen. Mention the document or process name if there is one."
              onChange={event => setDetails(event.target.value)}
              style={{ ...styles.input, resize: 'vertical', borderColor: showErrors && isDetailsMissing ? COLORS.danger : COLORS.borderStrong }}
            />
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              {showErrors && isDetailsMissing ? <span style={styles.errorText}>Please describe what happened.</span> : <span />}
              <span style={styles.counterText}>{details.length}/{MAX_DETAILS_LENGTH}</span>
            </div>
          </div>

          <div style={{ ...styles.field, marginBottom: '18px' }}>
            {fieldLabel('Steps to repeat the problem (optional)')}
            <textarea
              className="hs-input"
              value={steps}
              maxLength={MAX_DETAILS_LENGTH}
              rows={3}
              placeholder="1. Open Approvals&#10;2. Click Approve on the document&#10;3. The status does not change"
              onChange={event => setSteps(event.target.value)}
              style={{ ...styles.input, resize: 'vertical' }}
            />
            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <span style={styles.counterText}>{steps.length}/{MAX_DETAILS_LENGTH}</span>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            <button type="button" onClick={handleSendReport} style={sharedStyles.primaryButton}>
              Send report
            </button>
            <button type="button" onClick={handleToggleReport} style={sharedStyles.secondaryButton}>
              Cancel
            </button>
          </div>
        </div>
      )}

      {/* Frequently asked questions */}
      <h3 style={sharedStyles.sectionTitle}>Frequently asked questions</h3>
      <div style={{ ...sharedStyles.card, overflow: 'hidden' }}>
        {FAQ_ITEMS.map((item, index) => {
          const isOpen = openFaqIndex === index;
          return (
            <div key={item.question} style={{ borderBottom: index === FAQ_ITEMS.length - 1 ? 'none' : `1px solid ${COLORS.backgroundAlt}` }}>
              <button
                type="button"
                onClick={() => setOpenFaqIndex(isOpen ? null : index)}
                aria-expanded={isOpen}
                style={{
                  width: '100%',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '12px',
                  padding: '15px 18px',
                  border: 'none',
                  backgroundColor: isOpen ? COLORS.primaryTint : 'transparent',
                  cursor: 'pointer',
                  textAlign: 'left',
                  fontFamily: FONT_FAMILY,
                  fontSize: '13.5px',
                  fontWeight: 600,
                  color: COLORS.textStrong
                }}
              >
                {item.question}
                <span style={{ transform: isOpen ? 'rotate(180deg)' : 'none', transition: 'transform 0.15s ease', display: 'flex' }}>
                  <SVGIcon name="chevronDown" color={COLORS.textMuted} size={18} />
                </span>
              </button>
              {isOpen && (
                <div style={{ padding: '0 18px 16px 18px', backgroundColor: COLORS.primaryTint }}>
                  <p style={sharedStyles.bodyText}>{item.answer}</p>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  field: { display: 'flex', flexDirection: 'column', gap: '4px' },
  fieldLabel: { fontSize: '11.5px', fontWeight: 600, color: '#475569' },
  input: {
    width: '100%',
    boxSizing: 'border-box',
    padding: '7px 10px',
    border: `1px solid ${COLORS.borderStrong}`,
    borderRadius: '6px',
    backgroundColor: COLORS.surface,
    fontSize: '12px',
    color: COLORS.textStrong,
    fontFamily: FONT_FAMILY,
    outline: 'none',
    transition: 'border-color 0.15s ease, box-shadow 0.15s ease'
  },
  errorText: { fontSize: '11px', color: COLORS.danger },
  counterText: { fontSize: '11px', color: COLORS.textFaint },
  contactValue: { color: COLORS.textBody, fontSize: '13.5px' },
  supportCard: { ...sharedStyles.card, padding: '20px', display: 'flex', flexDirection: 'column' },
  cardHeader: { display: 'flex', alignItems: 'flex-start', gap: '16px' },
  cardTitle: { fontSize: '16px', fontWeight: 700, color: COLORS.textStrong },
  cardText: { fontSize: '13px', color: COLORS.textMuted, lineHeight: 1.5, marginTop: '4px' },
  cardBody: { display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '20px', flexGrow: 1 },
  contactRow: { display: 'flex', alignItems: 'center', gap: '14px' },
  smallBadge: { width: '34px', height: '34px', minWidth: '34px', borderRadius: '10px', backgroundColor: COLORS.primaryTint, display: 'flex', alignItems: 'center', justifyContent: 'center' },
  cardFooter: { borderTop: `1px solid ${COLORS.backgroundAlt}`, marginTop: '20px', paddingTop: '14px' },
  cardLink: { display: 'inline-flex', alignItems: 'center', gap: '10px', border: 'none', backgroundColor: 'transparent', padding: 0, cursor: 'pointer', fontSize: '14px', fontWeight: 700, color: COLORS.primary, fontFamily: FONT_FAMILY }
};

export default HelpSupport;
