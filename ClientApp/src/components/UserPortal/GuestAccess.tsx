// Guest Access Management popup — opened from the 3-dot menu of a document row in DocumentList.tsx.
// Same popup pattern as UploadDocuments.tsx / AssignWorkflow.tsx: fixed overlay, box perfectly
// centered, rounded on all 4 sides, and the scroll lives on an inner wrapper so the corners stay round.
import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx } from '../../spfx-shim';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// SharePoint list where every "Generate Access Link" entry is stored.
// https://mkx8.sharepoint.com/Lists/Guest%20Access/AllItems.aspx
// Columns: Title (renamed "GuestAccessID"), DocumentID, VendorName, EmailID, GuestName, Note, RequestorEmail, ProcessOwnerEmail
const GUEST_ACCESS_LIST_TITLE = 'Guest Access';

// Shared audit trail list — every module logs its entries here, one row per action.
// https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';

// Used to find the document's Process Owner email (document -> process folder -> Process Master).
const PROCESS_MASTER_LIST_TITLE = 'Process Master';
const DOCUMENT_LIBRARY_NAME = 'Shared Documents';

export interface IGuestAccessProps {
  isOpen: boolean;
  onClose: () => void;
  /** SPFx context — used to write the entry into the "Guest Access" list. */
  context?: any;
  /** Library item id of the selected document (only used to reset the form per document). */
  documentId: number;
  documentTitle: string;
  /** Document ID of the selected document = the datetime-based ID stored in the library's
   *  FileStatus column (the "Document ID" column, e.g. 20260929142602). Saved into the list's DocumentID column. */
  documentNumber: string;
  /** Called after the entry is saved successfully (popup closes right after). */
  onSaved?: (guestAccessId: string) => void;
}

const NOTE_MAX = 500;
// GuestAccessID = current date/time down to the second: yyyyMMddHHmmss (e.g. 20260929142602)
const generateGuestAccessId = (): string => {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
};

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/* ---------------- small inline icons (same style as DocumentList's SVGIcon) ---------------- */
const Icon = ({ name, color = 'currentColor', size = 14 }: { name: string; color?: string; size?: number }) => {
  const paths: { [key: string]: React.ReactNode } = {
    close: <path d="M6 18L18 6M6 6l12 12" />,
    users: (<g><circle cx="9" cy="8" r="3.2" /><path d="M3 20c0-3.3 2.7-5.5 6-5.5s6 2.2 6 5.5" /><circle cx="17" cy="9" r="2.4" /><path d="M16.5 14.6c2.6.2 4.5 2 4.5 4.9" /></g>),
    mail: (<g><rect x="3" y="5" width="18" height="14" rx="2" /><path d="M3 7l9 6 9-6" /></g>),
    user: (<g><circle cx="12" cy="8" r="3.5" /><path d="M5 20c0-3.6 3.1-6 7-6s7 2.4 7 6" /></g>),
    building: (<g><rect x="5" y="3" width="14" height="18" rx="1.5" /><path d="M9 8h2M13 8h2M9 12h2M13 12h2M10 21v-4h4v4" /></g>),
    file: (<g><path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8z" /><path d="M14 3v5h5M9 13h6M9 17h4" /></g>),
    shield: (<g><path d="M12 3l7 3v5c0 4.5-3 8-7 10-4-2-7-5.5-7-10V6z" /><path d="M9 12l2 2 4-4" /></g>),
    clock: (<g><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></g>),
    link: (<g><path d="M10 14a4 4 0 005.7 0l3-3a4 4 0 00-5.7-5.7l-1 1" /><path d="M14 10a4 4 0 00-5.7 0l-3 3a4 4 0 005.7 5.7l1-1" /></g>)
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }}>
      {paths[name] || null}
    </svg>
  );
};

export const GuestAccess: React.FC<IGuestAccessProps> = ({ isOpen, onClose, context, documentId, documentTitle, documentNumber, onSaved }) => {
  const [vendorName, setVendorName] = useState<string>('');
  const [requestor, setRequestor] = useState<string>('');
  const [email, setEmail] = useState<string>('');
  const [guestName, setGuestName] = useState<string>('');
  const [note, setNote] = useState<string>('');
  const [emailError, setEmailError] = useState<string>('');
  const [saveError, setSaveError] = useState<string>('');
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);

  // Email of the logged-in Microsoft 365 user (auto-captured, used as the Requestor).
  const currentUserEmail: string = context?.pageContext?.user?.email || context?.pageContext?.user?.loginName || '';

  // IMPORTANT: hooks stay above `if (!isOpen) return null;` (same rule as UploadDocuments.tsx).
  // Fresh form every time the popup opens.
  useEffect(() => {
    if (isOpen) {
      setVendorName('');
      setRequestor(currentUserEmail);   // auto-fill Requestor with the logged-in user's email
      setEmail('');
      setGuestName('');
      setNote('');
      setEmailError('');
      setSaveError('');
      setIsSubmitting(false);
    }
  }, [isOpen, documentId, currentUserEmail]);

  // Esc closes the popup.
  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const handleSubmit = async () => {
    const trimmedEmail = email.trim();

    if (!trimmedEmail) {
      setEmailError('Guest email is required.');
      return;
    }
    if (!EMAIL_REGEX.test(trimmedEmail)) {
      setEmailError('Please enter a valid email address.');
      return;
    }
    setEmailError('');

    setIsSubmitting(true);
    setSaveError('');
    try {
      const guestAccessId = generateGuestAccessId();
      const sp = spfi().using(SPFx(context));

      // Process Owner email of the process this document belongs to (empty if it cannot be found).
      // Never blocks saving the Guest Access entry.
      let processOwnerEmail = '';
      let resolvedProcessCode = '';   // Process Code of the document's process (used for the audit log)
      try {
        const webRelativeUrl: string = context?.pageContext?.web?.serverRelativeUrl || '';
        const libraryRelativeUrl = webRelativeUrl && webRelativeUrl !== '/'
          ? `${webRelativeUrl}/${DOCUMENT_LIBRARY_NAME}`
          : `/${DOCUMENT_LIBRARY_NAME}`;
        // The document sits inside a folder named after its Process Code.
        const docItem: any = await sp.web.getList(libraryRelativeUrl).items.getById(documentId).select('FileDirRef')();
        const folderPath: string = docItem?.FileDirRef || '';
        const marker = `/${DOCUMENT_LIBRARY_NAME}/`;
        const markerIndex = folderPath.indexOf(marker);
        const processCode = markerIndex >= 0 ? folderPath.substring(markerIndex + marker.length).split('/')[0] : '';
        resolvedProcessCode = processCode;
        if (processCode) {
          const processItems: any[] = await sp.web.lists.getByTitle(PROCESS_MASTER_LIST_TITLE).items
            .filter(`ProcessCode eq '${processCode.replace(/'/g, "''")}'`)
            .select('ProcessOwnerEmail')
            .top(1)();
          processOwnerEmail = processItems?.[0]?.ProcessOwnerEmail || '';
        }
      } catch (ownerErr) {
        console.warn('Could not resolve the Process Owner email for this document:', ownerErr);
      }

      await sp.web.lists.getByTitle(GUEST_ACCESS_LIST_TITLE).items.add({
        Title: guestAccessId,            // "Title" column — renamed to "GuestAccessID" in the list UI
        DocumentID: documentNumber,      // Document ID of the selected document (library's FileStatus column)
        VendorName: vendorName.trim(),
        RequestorEmail: requestor.trim(),  // auto-captured logged-in user email -> "RequestorEmail" column
        ProcessOwnerEmail: processOwnerEmail,  // Process Owner of the document's process -> "ProcessOwnerEmail" column
        EmailID: trimmedEmail,
        GuestName: guestName.trim(),
        Note: note.trim()
      });

      // One "DMS Audit Logs" entry per Guest Access request (same data as the Guest Access list row).
      // Never blocks the Guest Access save — any failure here is only logged to the console.
      try {
        await sp.web.lists.getByTitle(AUDIT_LOG_LIST_TITLE).items.add({
          AuditID: generateGuestAccessId(),  // yyyyMMddHHmmss — moment the audit entry was written
          Title: resolvedProcessCode || guestAccessId,  // "Title" column (shown as ProcessCode) — the document's Process Code; falls back to the Guest Access ID if it cannot be found
          EntryType: 'Guest Access',
          GuestAccessID: guestAccessId,
          GuestName: guestName.trim(),
          VendorName: vendorName.trim(),
          EmailID: trimmedEmail,
          RequestorEmail: requestor.trim(),
          Note: note.trim()
        });
      } catch (auditErr) {
        console.error('Error writing DMS Audit Logs entry for Guest Access:', auditErr);
      }
      if (onSaved) onSaved(guestAccessId);
      onClose();
    } catch (err) {
      console.error('Error saving entry to Guest Access list:', err);
      setSaveError('Could not save guest access. Please try again.');
    } finally {
      setIsSubmitting(false);
    }
  };

  const features = [
    { icon: 'shield', title: 'Secure Access', text: 'Control what guests can view and do in the system.' },
    { icon: 'clock', title: 'Time Bound', text: 'Set an expiry date for access when needed.' },
    { icon: 'users', title: 'Easy Management', text: 'Revoke or update access at any time.' }
  ];

  return (
    <div style={styles.overlay}>
      <style>{`
        .ga-box { border-color: #cbd5e1; transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .ga-box:hover { border-color: #94a3b8; }
        .ga-box:focus-within { border-color: #005a43; box-shadow: 0 0 0 2px rgba(0, 90, 67, 0.15); }
        .ga-box.ga-error { border-color: #dc2626; }
        .ga-box.ga-error:focus-within { box-shadow: 0 0 0 2px rgba(220, 38, 38, 0.12); }
        .ga-box input::placeholder, .ga-box textarea::placeholder { color: #94a3b8; }
      `}</style>
      <div style={styles.modal} role="dialog" aria-modal="true" aria-label="Guest Access Management">
        <div style={styles.modalScroll}>

          {/* ---------- header ---------- */}
          <div style={styles.headerRow}>
            <div style={styles.headerIconBox}>
              <Icon name="users" color="#005a43" size={20} />
            </div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <h2 style={styles.title}>Guest Access Management</h2>
              <p style={styles.subtitle}>
                Invite external users (guests) to access documents and resources in this system. Manage their permissions and access duration.
              </p>
            </div>
            <button onClick={onClose} style={styles.closeBtn} aria-label="Close">
              <Icon name="close" color="#475569" size={16} />
            </button>
          </div>

          {/* ---------- body ---------- */}
          <div style={styles.bodyRow}>

            {/* left info panel */}
            <div style={styles.leftPanel}>
              <div style={styles.illustrationWrap}>
                <svg width="190" height="130" viewBox="0 0 190 130" fill="none">
                  <ellipse cx="95" cy="66" rx="78" ry="54" fill="#dcf1ea" opacity="0.7" />
                  <circle cx="160" cy="34" r="5" fill="#bfe3d7" />
                  <rect x="42" y="22" width="92" height="82" rx="8" fill="#ffffff" stroke="#e2e8f0" />
                  <circle cx="88" cy="48" r="13" fill="#005a43" />
                  <circle cx="88" cy="44" r="4.5" fill="#ffffff" />
                  <path d="M79 56c1.5-4 5-5.5 9-5.5s7.5 1.5 9 5.5" stroke="#ffffff" strokeWidth="3" strokeLinecap="round" />
                  <rect x="58" y="72" width="60" height="4" rx="2" fill="#cfe5dd" />
                  <rect x="58" y="82" width="46" height="4" rx="2" fill="#cfe5dd" />
                  <rect x="58" y="92" width="30" height="4" rx="2" fill="#e2efe9" />
                  <circle cx="132" cy="96" r="13" fill="#005a43" stroke="#ffffff" strokeWidth="3" />
                  <path d="M132 90v12M126 96h12" stroke="#ffffff" strokeWidth="2.5" strokeLinecap="round" />
                </svg>
              </div>

              <div style={styles.featureList}>
                {features.map((f) => (
                  <div key={f.title} style={styles.featureItem}>
                    <div style={styles.featureIcon}>
                      <Icon name={f.icon} color="#005a43" size={14} />
                    </div>
                    <div style={{ minWidth: 0 }}>
                      <div style={styles.featureTitle}>{f.title}</div>
                      <div style={styles.featureText}>{f.text}</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* right form panel */}
            <div style={styles.rightPanel}>
              <div style={styles.sectionTitle}>Guest Information</div>

              <div style={styles.formGrid}>
                {/* Row 1: Guest Name + Vendor Name; Row 2: Requestor + Email ID */}
                <div style={styles.fieldCol}>
                  <label style={styles.fieldLabel}>Guest Name</label>
                  <div className="ga-box" style={styles.inputBox}>
                    <Icon name="user" color="#94a3b8" size={14} />
                    <input
                      type="text"
                      value={guestName}
                      placeholder="Enter guest's name"
                      onChange={(e) => setGuestName(e.target.value)}
                      style={styles.bareInput}
                    />
                  </div>
                </div>

                <div style={styles.fieldCol}>
                  <label style={styles.fieldLabel}>Vendor Name</label>
                  <div className="ga-box" style={styles.inputBox}>
                    <Icon name="building" color="#94a3b8" size={14} />
                    <input
                      type="text"
                      value={vendorName}
                      placeholder="Enter vendor name"
                      onChange={(e) => setVendorName(e.target.value)}
                      style={styles.bareInput}
                    />
                  </div>
                </div>

                <div style={styles.fieldCol}>
                  <label style={styles.fieldLabel}>Requestor</label>
                  <div className="ga-box" style={{ ...styles.inputBox, backgroundColor: '#f1f5f9', cursor: 'not-allowed' }}>
                    <Icon name="user" color="#94a3b8" size={14} />
                    <input
                      type="text"
                      value={requestor}
                      placeholder="Enter requestor"
                      onChange={(e) => setRequestor(e.target.value)}
                      disabled
                      style={{ ...styles.bareInput, cursor: 'not-allowed', color: '#475569' }}
                    />
                  </div>
                </div>

                <div style={styles.fieldCol}>
                  <label style={styles.fieldLabel}>Email ID <span style={styles.req}>*</span></label>
                  <div className={`ga-box${emailError ? ' ga-error' : ''}`} style={styles.inputBox}>
                    <Icon name="mail" color="#94a3b8" size={14} />
                    <input
                      type="text"
                      value={email}
                      placeholder="Enter guest's email address"
                      onChange={(e) => { setEmail(e.target.value); if (emailError) setEmailError(''); }}
                      style={styles.bareInput}
                    />
                  </div>
                  {emailError && <div style={styles.errorText}>{emailError}</div>}
                </div>
              </div>

              <div style={{ marginTop: '12px' }}>
                <label style={styles.fieldLabel}>Note <span style={styles.optional}>(optional)</span></label>
                <div className="ga-box" style={{ ...styles.inputBox, alignItems: 'flex-start', padding: '8px 10px' }}>
                  <span style={{ marginTop: '1px' }}><Icon name="file" color="#94a3b8" size={14} /></span>
                  <textarea
                    value={note}
                    maxLength={NOTE_MAX}
                    placeholder="Add any additional notes (e.g., reason for access, project details, etc.)"
                    onChange={(e) => setNote(e.target.value)}
                    style={styles.bareTextarea}
                  />
                </div>
                <div style={styles.counter}>{note.length}/{NOTE_MAX}</div>
              </div>

              <div style={styles.divider} />

              <div style={styles.sectionTitle}>Permissions</div>
              <div style={styles.permissionRow}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', minWidth: 0 }}>
                  <Icon name="file" color="#005a43" size={14} />
                  <span style={styles.permissionText} title={documentTitle}>
                    Access to “{documentTitle}”
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* ---------- footer ---------- */}
          {saveError && <div style={{ ...styles.errorText, textAlign: 'right', marginTop: '10px' }}>{saveError}</div>}
          <div style={styles.footer}>
            <button type="button" onClick={onClose} style={styles.cancelBtn} disabled={isSubmitting}>Cancel</button>
            <button
              type="button"
              onClick={handleSubmit}
              disabled={isSubmitting}
              style={{ ...styles.submitBtn, opacity: isSubmitting ? 0.7 : 1, cursor: isSubmitting ? 'not-allowed' : 'pointer' }}
            >
              <Icon name="link" color="#ffffff" size={14} />
              <span>{isSubmitting ? 'Generating...' : 'Generate Access Link'}</span>
            </button>
          </div>

        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: {
    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
    backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)',
    display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
    padding: '20px', boxSizing: 'border-box'
  },
  modal: {
    backgroundColor: '#ffffff', borderRadius: '14px', width: '860px', maxWidth: '96vw',
    maxHeight: '92vh', overflow: 'hidden', boxShadow: '0 20px 50px rgba(0,0,0,0.25)',
    boxSizing: 'border-box', display: 'flex', flexDirection: 'column'
  },
  // Scroll on the inner wrapper so the outer box keeps all 4 corners rounded.
  modalScroll: { overflowY: 'auto', padding: '20px 22px', boxSizing: 'border-box', minHeight: 0 },

  headerRow: { display: 'flex', alignItems: 'flex-start', gap: '12px', marginBottom: '16px' },
  headerIconBox: { width: '40px', height: '40px', borderRadius: '10px', backgroundColor: '#e6f4f1', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '3px 0 0 0', fontSize: '11px', color: '#64748b', lineHeight: 1.5, maxWidth: '560px' },
  closeBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '4px', borderRadius: '6px', display: 'flex', flexShrink: 0 },

  bodyRow: { display: 'flex', gap: '14px', alignItems: 'stretch' },

  leftPanel: { width: '210px', flexShrink: 0, backgroundColor: '#f6fbf9', border: '1px solid #e2efe9', borderRadius: '10px', padding: '14px 14px', boxSizing: 'border-box', display: 'flex', flexDirection: 'column' },
  illustrationWrap: { display: 'flex', justifyContent: 'center', marginBottom: '10px' },
  featureList: { display: 'flex', flexDirection: 'column', gap: '14px' },
  featureItem: { display: 'flex', alignItems: 'flex-start', gap: '10px' },
  featureIcon: { width: '28px', height: '28px', borderRadius: '50%', backgroundColor: '#e6f4f1', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  featureTitle: { fontSize: '12px', fontWeight: 700, color: '#1e293b' },
  featureText: { fontSize: '10.5px', color: '#64748b', lineHeight: 1.45, marginTop: '2px' },

  rightPanel: { flex: 1, minWidth: 0, border: '1px solid #e2e8f0', borderRadius: '10px', padding: '14px 16px', boxSizing: 'border-box', backgroundColor: '#ffffff' },
  sectionTitle: { fontSize: '13px', fontWeight: 700, color: '#0f172a', marginBottom: '10px' },
  formGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px 14px' },
  fieldCol: { minWidth: 0 },
  fieldLabel: { fontSize: '11px', fontWeight: 600, color: '#475569', marginBottom: '5px', display: 'block' },
  req: { color: '#dc2626' },
  optional: { color: '#94a3b8', fontWeight: 500 },

  // Border/focus colours come from the .ga-box classes in the <style> tag (hover/focus/error states).
  inputBox: { display: 'flex', alignItems: 'center', gap: '8px', padding: '7px 10px', border: '1px solid #cbd5e1', borderRadius: '6px', backgroundColor: '#ffffff', boxSizing: 'border-box', minHeight: '32px' },
  bareInput: { flex: 1, minWidth: 0, border: 'none', outline: 'none', background: 'transparent', fontSize: '11px', color: '#1e293b', fontFamily: 'inherit', padding: 0 },
  bareTextarea: { flex: 1, minWidth: 0, border: 'none', outline: 'none', background: 'transparent', fontSize: '11px', color: '#1e293b', fontFamily: 'inherit', padding: 0, resize: 'none', height: '58px' },
  errorText: { fontSize: '10px', color: '#dc2626', marginTop: '4px' },
  counter: { textAlign: 'right', fontSize: '10px', color: '#94a3b8', marginTop: '4px' },

  divider: { height: '1px', backgroundColor: '#f1f5f9', margin: '14px 0' },
  permissionRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px', backgroundColor: '#f1f8f5', border: '1px solid #e2efe9', borderRadius: '6px', padding: '9px 12px' },
  permissionText: { fontSize: '11px', color: '#0d6b52', fontWeight: 500, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },

  footer: { display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '16px' },
  cancelBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', color: '#334155', borderRadius: '6px', padding: '9px 18px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
  submitBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#005a43', border: 'none', color: '#ffffff', borderRadius: '6px', padding: '9px 18px', fontSize: '12px', fontWeight: 600 }
};

export default GuestAccess;
