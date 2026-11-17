import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx, photoUrlFor } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";

export interface IWorkflowEditItem {
  id: number;
  workflowTitle: string;
  processCode: string;
  workflowType: string;
  escalationCriteria: string; // comma-separated
  stepRole: string;
  l1Approver: string;
  l1ApproverEmail: string;
  l2Approver: string;
  l2ApproverEmail: string;
  escalateTo: string;
  escalateToEmail: string;
  finalApprover: string;
  finalApproverEmail: string;
  complianceCheck: string;
  deadlineDays: number;
  notifyApprovers: boolean;
  alertEscalation: boolean;
}

export interface ICreateWorkflowProps {
  isOpen: boolean;
  onClose: () => void;
  context: any;
  onSaved?: () => void;
  editItem?: IWorkflowEditItem | null;
  isViewMode?: boolean;
  defaultProcessCode?: string;
}

interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

// -------------------- Small reusable inputs (same pattern as ProcessManagementForm) --------------------

interface FormSelectOption { label: string; value: string; }

interface FormSelectProps {
  label?: string; required?: boolean; value: string; onChange: (val: string) => void;
  options: FormSelectOption[]; placeholder?: string; icon?: React.ReactNode; disabled?: boolean; searchable?: boolean;
}

const FormSelect: React.FC<FormSelectProps> = ({ label, required, value, onChange, options, placeholder = 'Select', icon, disabled, searchable }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);
  const selected = options.find(opt => opt.value === value);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) { setIsOpen(false); setSearchQuery(''); }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const visibleOptions = searchable && searchQuery
    ? options.filter(opt => opt.label.toLowerCase().includes(searchQuery.toLowerCase()))
    : options;

  return (
    <div style={{ ...styles.field, position: 'relative' }} ref={wrapperRef}>
      {label && <label style={styles.label}>{label} {required && <span style={styles.req}>*</span>}</label>}
      <div
        className="cw-select-box"
        style={{ ...styles.selectBox, borderColor: isOpen ? '#005c42' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 92, 66, 0.10)' : 'none', opacity: disabled ? 0.6 : 1, cursor: disabled ? 'not-allowed' : 'pointer' }}
        onClick={() => !disabled && setIsOpen(prev => !prev)}
      >
        <span style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden', minWidth: 0 }}>
          {icon}
          <span style={{ fontSize: '12.5px', color: (selected || value) ? '#1e293b' : '#94a3b8', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
            {selected ? selected.label : (value || placeholder)}
          </span>
        </span>
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </div>
      {isOpen && !disabled && (
        <div className="cw-dropdown-menu" style={{ ...styles.selectMenu, padding: searchable ? '5px' : styles.selectMenu.padding }}>
          {searchable && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '7px 9px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '7px', marginBottom: '5px' }}>
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>
              <input
                autoFocus
                type="text"
                placeholder="Search process code..."
                value={searchQuery}
                onChange={e => setSearchQuery(e.target.value)}
                style={{ border: 'none', outline: 'none', fontSize: '12px', width: '100%', background: 'transparent', color: '#334155' }}
                onClick={e => e.stopPropagation()}
              />
            </div>
          )}
          {visibleOptions.length === 0 ? (
            <div style={{ fontSize: '11.5px', color: '#94a3b8', padding: '10px', textAlign: 'center' }}>{options.length === 0 ? 'No options available' : 'No matches found'}</div>
          ) : (
            <div style={{ maxHeight: searchable ? '190px' : undefined, overflowY: searchable ? 'auto' : undefined }}>
              {visibleOptions.map(opt => {
                const isSelected = opt.value === value;
                return (
                  <div
                    key={opt.value}
                    className="cw-select-option"
                    style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent', color: isSelected ? '#005c42' : '#334155', fontWeight: isSelected ? 600 : 500 }}
                    onClick={() => { onChange(opt.value); setIsOpen(false); setSearchQuery(''); }}
                  >
                    <span>{opt.label}</span>
                    {isSelected && <span style={{ color: '#005c42', fontWeight: 700 }}>✓</span>}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
};

interface PeopleSelectProps {
  label?: string; required?: boolean; value: string; onChange: (val: string, email?: string) => void;
  disabled?: boolean; placeholder?: string; context?: any;
}

const PeopleSelect: React.FC<PeopleSelectProps> = ({ label, required, value, onChange, disabled, placeholder = 'Select', context }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [msUsers, setMsUsers] = useState<IGraphUser[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        if (!context?.msGraphClientFactory) return;
        const client = await context.msGraphClientFactory.getClient();
        const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(50).get();
        const siteUrl = context.pageContext?.web?.absoluteUrl || '';
        const usersList: IGraphUser[] = response.value.map((u: any) => {
          const email = u.mail || u.userPrincipalName || '';
          return {
            id: u.id,
            displayName: u.displayName || 'No Name',
            mail: email,
            userPrincipalName: u.userPrincipalName || '',
            photoUrl: siteUrl ? photoUrlFor(email) : ''
          };
        });
        setMsUsers(usersList);
      } catch (err) {
        console.error('Graph API error fetching M365 users:', err);
      }
    };
    fetchUsers();
  }, [context]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) setIsOpen(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredUsers = msUsers.filter(u =>
    u.displayName.toLowerCase().includes(searchQuery.toLowerCase()) ||
    u.mail.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const selectedUserObj = msUsers.find(u => u.displayName === value);

  return (
    <div style={{ ...styles.field, position: 'relative' }} ref={wrapperRef}>
      {label && <label style={styles.label}>{label} {required && <span style={styles.req}>*</span>}</label>}
      <div
        className="cw-select-box"
        style={{ ...styles.selectBox, borderColor: isOpen ? '#005c42' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 92, 66, 0.10)' : 'none', opacity: disabled ? 0.6 : 1, cursor: disabled ? 'not-allowed' : 'pointer' }}
        onClick={() => !disabled && setIsOpen(prev => !prev)}
      >
        <span style={{ display: 'flex', alignItems: 'center', gap: '7px', overflow: 'hidden', minWidth: 0 }}>
          <span style={styles.avatarCircle}>
            {selectedUserObj?.photoUrl ? <img src={selectedUserObj.photoUrl} alt="" style={styles.avatarImg} /> : <PersonIcon />}
          </span>
          <span style={{ fontSize: '12.5px', color: value ? '#1e293b' : '#94a3b8', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{value || placeholder}</span>
        </span>
        <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
          <polyline points="6 9 12 15 18 9" />
        </svg>
      </div>
      {isOpen && !disabled && (
        <div className="cw-dropdown-menu" style={{ ...styles.selectMenu, padding: '5px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '7px 9px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '7px', marginBottom: '5px' }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>
            <input
              autoFocus
              type="text"
              placeholder="Search name or email..."
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              style={{ border: 'none', outline: 'none', fontSize: '12px', width: '100%', background: 'transparent', color: '#334155' }}
              onClick={e => e.stopPropagation()}
            />
          </div>
          <div style={{ maxHeight: '190px', overflowY: 'auto' }}>
            {filteredUsers.length === 0 ? (
              <div style={{ fontSize: '11.5px', color: '#94a3b8', padding: '10px', textAlign: 'center' }}>No users found</div>
            ) : (
              filteredUsers.map(u => {
                const isSelected = u.displayName === value;
                return (
                  <div
                    key={u.id}
                    className="cw-select-option"
                    style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent' }}
                    onClick={() => { onChange(u.displayName, u.mail); setIsOpen(false); setSearchQuery(''); }}
                  >
                    <span style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden', minWidth: 0 }}>
                      <span style={styles.avatarCircle}>
                        {u.photoUrl ? <img src={u.photoUrl} alt="" style={styles.avatarImg} /> : <PersonIcon />}
                      </span>
                      <span style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden', minWidth: 0 }}>
                        <span style={{ fontSize: '12px', fontWeight: isSelected ? 600 : 500, color: isSelected ? '#005c42' : '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</span>
                        <span style={{ fontSize: '10px', color: '#94a3b8', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.mail}</span>
                      </span>
                    </span>
                    {isSelected && <span style={{ color: '#005c42', fontWeight: 700, flexShrink: 0 }}>✓</span>}
                  </div>
                );
              })
            )}
          </div>
        </div>
      )}
    </div>
  );
};

// -------------------- Icons --------------------

const PersonIcon = () => (
  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" />
    <circle cx="12" cy="7" r="4" />
  </svg>
);

const WorkflowIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#005c42" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" />
    <rect x="3" y="15" width="6" height="6" rx="1" /><path d="M9 6h6a3 3 0 0 1 3 3v6" /><path d="M9 18h3" />
  </svg>
);

const FolderIcon = () => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
  </svg>
);

const SendIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="22" y1="2" x2="11" y2="13" /><polygon points="22 2 15 22 11 13 2 9 22 2" />
  </svg>
);

const ClockIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#005c42" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="9" /><polyline points="12 7 12 12 15.5 14" />
  </svg>
);

const BellIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#005c42" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9" /><path d="M13.73 21a2 2 0 0 1-3.46 0" />
  </svg>
);

// -------------------- Approval steps config --------------------

const STEPS = [
  { key: 'l1', title: 'Level 1 Approval' },
  { key: 'l2', title: 'Level 2 Approval' },
  { key: 'escalation', title: 'Escalation / Review', optional: true },
  { key: 'final', title: 'Final Approval' }
];

const CreateWorkflow: React.FC<ICreateWorkflowProps> = ({ isOpen, onClose, context, onSaved, editItem, isViewMode = false, defaultProcessCode }) => {
  const [processCodeOptions, setProcessCodeOptions] = useState<FormSelectOption[]>([]);
  const [foldersLoading, setFoldersLoading] = useState<boolean>(false);

  const [workflowTitle, setWorkflowTitle] = useState('');
  const [processCode, setProcessCode] = useState('');
  const [workflowType, setWorkflowType] = useState('');
  const [escalationCriteria, setEscalationCriteria] = useState<string[]>([]);

  const [role, setRole] = useState('');
  const [l1Approver, setL1Approver] = useState('');
  const [l1ApproverEmail, setL1ApproverEmail] = useState('');
  const [l2Approver, setL2Approver] = useState('');
  const [l2ApproverEmail, setL2ApproverEmail] = useState('');
  const [escalateTo, setEscalateTo] = useState('');
  const [escalateToEmail, setEscalateToEmail] = useState('');
  const [finalApprover, setFinalApprover] = useState('');
  const [finalApproverEmail, setFinalApproverEmail] = useState('');
  const [complianceCheck, setComplianceCheck] = useState('');

  const [deadlineDays, setDeadlineDays] = useState('0');
  const [notifyApprovers, setNotifyApprovers] = useState(false);
  const [alertEscalation, setAlertEscalation] = useState(false);

  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState('');

  // Process Name is a dropdown of the process-code folders that already exist
  // in the document library (a folder is auto-created there whenever a
  // process is submitted from the Process Master form).
  useEffect(() => {
    if (!isOpen || !context) return;
    const fetchProcessFolders = async () => {
      setFoldersLoading(true);
      try {
        const sp = spfi().using(SPFx(context));
        const folders = await sp.web.getFolderByServerRelativePath("Shared Documents").folders.select("Name")();
        const opts = folders
          .filter((f: any) => f.Name !== 'Forms')
          .map((f: any) => ({ label: f.Name, value: f.Name }));
        setProcessCodeOptions(opts);
      } catch (err) {
        console.error('Error fetching process-code folders:', err);
      } finally {
        setFoldersLoading(false);
      }
    };
    fetchProcessFolders();
  }, [isOpen, context]);

  // Reset the form each time the modal is (re)opened — pre-fill from editItem
  // when editing an existing workflow, otherwise start blank.
  useEffect(() => {
    if (!isOpen) return;
    if (editItem) {
      setWorkflowTitle(editItem.workflowTitle || '');
      setProcessCode(editItem.processCode || '');
      setWorkflowType(editItem.workflowType || '');
      setEscalationCriteria(editItem.escalationCriteria ? editItem.escalationCriteria.split(',').map(s => s.trim()).filter(Boolean) : []);
      setRole(editItem.stepRole || '');
      setL1Approver(editItem.l1Approver || ''); setL1ApproverEmail(editItem.l1ApproverEmail || '');
      setL2Approver(editItem.l2Approver || ''); setL2ApproverEmail(editItem.l2ApproverEmail || '');
      setEscalateTo(editItem.escalateTo || ''); setEscalateToEmail(editItem.escalateToEmail || '');
      setFinalApprover(editItem.finalApprover || ''); setFinalApproverEmail(editItem.finalApproverEmail || '');
      setComplianceCheck(editItem.complianceCheck || '');
      setDeadlineDays(String(editItem.deadlineDays ?? 0));
      setNotifyApprovers(!!editItem.notifyApprovers);
      setAlertEscalation(!!editItem.alertEscalation);
    } else {
      setWorkflowTitle('');
      // When opened from a specific process's Workflow screen (Process Master
      // → 3-dot menu → Workflow), pre-select that process code. When opened
      // from the general Approval Management → Create Workflow button (no
      // defaultProcessCode passed), leave it blank for manual selection.
      setProcessCode(defaultProcessCode || ''); setWorkflowType(''); setEscalationCriteria([]);
      setRole(''); setL1Approver(''); setL1ApproverEmail(''); setL2Approver(''); setL2ApproverEmail('');
      setEscalateTo(''); setEscalateToEmail(''); setFinalApprover(''); setFinalApproverEmail(''); setComplianceCheck('');
      setDeadlineDays('0'); setNotifyApprovers(false); setAlertEscalation(false);
    }
    setSaveError('');
  }, [isOpen, editItem, defaultProcessCode]);

  if (!isOpen) return null;

  // "Step 4 (Optional): Escalation / Review" and the "Escalate To" approver
  // are only relevant when the user has picked the Escalation Workflow type.
  const isEscalationWorkflow = workflowType === 'Escalation Workflow';
  const visibleSteps = STEPS.filter(s => !s.optional || isEscalationWorkflow);

  // Only one of the 3 escalation-criteria options can be selected at a time —
  // picking one replaces whatever was selected before; picking the same one
  // again clears it.
  const toggleEscalationCriteria = (val: string) => {
    setEscalationCriteria(prev => prev.includes(val) ? [] : [val]);
  };

  // Checked in this exact order — the first missing field is reported so the
  // user always fills them left-to-right / top-to-bottom.
  const getValidationError = (): string => {
    if (processCode.trim() === '') return 'Please select Process Name before saving.';
    if (workflowType.trim() === '') return 'Please select Workflow Type before saving.';
    if (l1Approver.trim() === '') return 'Please select L1 Approver before saving.';
    if (l2Approver.trim() === '') return 'Please select L2 Approver before saving.';
    if (isEscalationWorkflow && escalateTo.trim() === '') return 'Please select Escalate To before saving.';
    if (finalApprover.trim() === '') return 'Please select Final Approver before saving.';
    return '';
  };

  const isValid = getValidationError() === '';

  const handleSave = async () => {
    const validationError = getValidationError();
    if (validationError) {
      setSaveError(validationError);
      return;
    }
    setSaving(true);
    setSaveError('');
    try {
      const sp = spfi().using(SPFx(context));

      // Person/Group columns need SharePoint's own numeric user Id, not the
      // display name or the Azure AD (Graph) id — ensureUser() resolves an
      // email to that SharePoint user record (and adds them to the site's
      // hidden User Information list the first time they're referenced).
      const resolveUserId = async (email: string): Promise<number | null> => {
        if (!email) return null;
        try {
          const user = await sp.web.ensureUser(email);
          return user.Id;
        } catch (e) {
          console.error(`Could not resolve SharePoint user for ${email}:`, e);
          return null;
        }
      };

      const [l1Id, l2Id, escalateId, finalId] = await Promise.all([
        resolveUserId(l1ApproverEmail),
        resolveUserId(l2ApproverEmail),
        resolveUserId(escalateToEmail),
        resolveUserId(finalApproverEmail)
      ]);

      const isEditing = !!editItem;

      // WorkflowID — a datetime-based unique id (yyyyMMddHHmmss), stamped
      // only when the workflow is first created; it's left untouched on
      // update so the id stays stable for the record.
      const generateWorkflowId = (): string => {
        const now = new Date();
        const pad = (n: number) => String(n).padStart(2, '0');
        return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
      };

      // Saves the workflow definition to the "Workflow Configuration" list.
      // Person columns are set via the "<Field>Id" pattern SharePoint expects
      // for single-value Person or Group fields.
      const payload = {
        Title: processCode,
        WorkflowTitle: workflowTitle,
        ...(isEditing ? {} : { WorkflowID: generateWorkflowId() }),
        WorkflowType: workflowType,
        EscalationCriteria: escalationCriteria.join(', '),
        StepRole: role,
        L1ApproverId: l1Id,
        L2ApproverId: l2Id,
        EscalateToId: escalateId,
        FinalApproverId: finalId,
        ComplianceCheck: complianceCheck,
        ApprovalDeadlineDays: Number(deadlineDays) || 0,
        NotifyApproversEachStep: notifyApprovers,
        AlertOnEscalation: alertEscalation
      };

      if (isEditing) {
        await sp.web.lists.getByTitle("Workflow Configuration").items.getById(editItem!.id).update(payload);
      } else {
        await sp.web.lists.getByTitle("Workflow Configuration").items.add(payload);
      }

      // Mirror the action into the shared "DMS Audit Logs" list so there's a
      // trail of who created or updated which workflow, and when.
      // https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
      try {
        const currentUserEmail = context?.pageContext?.user?.email || context?.pageContext?.user?.loginName || '';
        // yyyyMMddHHmmss — exact moment this entry was logged.
        const generateAuditId = (): string => {
          const now = new Date();
          const pad = (n: number) => String(n).padStart(2, '0');
          return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        };
        await sp.web.lists.getByTitle("DMS Audit Logs").items.add({
          Title: processCode,            // "Title" column — renamed to "ProcessID" in the list UI
          EntryType: 'Workflow',
          AuditID: generateAuditId(),
          ApprovalStepName: `Workflow ${isEditing ? 'Updated' : 'Created'} — ${processCode}`,
          WorkflowType: workflowType,
          EscalationCriteria: escalationCriteria.join(', '),
          StepRole: role,
          L1ApproverId: l1Id,
          L2ApproverId: l2Id,
          FinalApproverId: finalId,
          ComplianceCheckApprover: complianceCheck,
          ApprovalDeadlineDays: Number(deadlineDays) || 0,
          NotifyApproversEachStep: notifyApprovers,
          AlertOnEscalationCompliance: alertEscalation,
          ActionBy: currentUserEmail,
          ActionStatus: isEditing ? 'Updated' : 'Created',
          ActionDate: new Date().toISOString()
        });
      } catch (auditErr) {
        // The workflow itself saved fine — don't block on the audit entry,
        // just log it so it's visible during testing.
        console.error('Workflow saved, but audit log entry failed:', auditErr);
      }

      setSaving(false);
      if (onSaved) onSaved();
      onClose();
    } catch (err: any) {
      console.error('Error saving workflow:', err);
      setSaving(false);
      const detail = err?.message ? `: ${err.message}` : '';
      setSaveError(`Could not save workflow${detail}`);
    }
  };

  return (
    <div className="cw-overlay" style={styles.overlay} onClick={onClose}>
      <style>{`
        @keyframes cwFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes cwScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        @keyframes cwSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .cw-overlay { animation: cwFadeIn 0.18s ease; }
        .cw-modal { animation: cwScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
        .cw-dropdown-menu { animation: cwSlideDown 0.15s ease; }
        .cw-select-box:hover { border-color: #94a3b8; }
        .cw-select-option:hover { background-color: #f1f5f9; }
        .cw-btn { transition: transform 0.15s ease, filter 0.15s ease; }
        .cw-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
        .cw-btn:active { transform: translateY(0); filter: brightness(0.97); }
        .cw-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; }
        .cw-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
        .cw-step-chip { transition: transform 0.15s ease; }
        .cw-step-chip:hover { transform: translateY(-1px); }
        .cw-checkbox-label:hover { color: #005c42; }
        .cw-body::-webkit-scrollbar { width: 8px; }
        .cw-body::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 8px; }
        .cw-body::-webkit-scrollbar-track { background-color: transparent; }
      `}</style>

      <div className="cw-modal" style={styles.modalCard} onClick={e => e.stopPropagation()}>
        <div style={styles.header}>
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: '12px' }}>
            <div style={styles.headerIconBox}><WorkflowIcon /></div>
            <div>
              <h2 style={styles.title}>{isViewMode ? 'View Workflow - Approval / Escalation' : editItem ? 'Edit Workflow - Approval / Escalation' : 'Create Workflow - Approval / Escalation'}</h2>
              <p style={styles.subtitle}>{isViewMode ? 'Read-only view of this workflow configuration' : 'Define the workflow process for approvals or escalations'}</p>
            </div>
          </div>
          <button type="button" className="cw-close-btn" style={styles.closeBtn} onClick={onClose} title="Close">✕</button>
        </div>

        <div className="cw-body" style={styles.body}>
          <div style={styles.field}>
            <label style={styles.label}>Workflow Title</label>
            <input
              className="cw-input"
              type="text"
              style={{ ...styles.input, opacity: isViewMode ? 0.7 : 1 }}
              value={workflowTitle}
              onChange={e => setWorkflowTitle(e.target.value)}
              placeholder="Enter workflow title"
              disabled={isViewMode}
            />
          </div>

          <div style={styles.grid2}>
            <FormSelect
              label="Process Code"
              required
              value={processCode}
              onChange={setProcessCode}
              icon={<FolderIcon />}
              placeholder={foldersLoading ? 'Loading process folders...' : 'Select process code folder'}
              options={processCodeOptions}
              searchable
              disabled={isViewMode || !!editItem}
            />
            <FormSelect
              label="Workflow Type"
              required
              value={workflowType}
              onChange={setWorkflowType}
              placeholder="Select Workflow Type"
              options={[
                { label: 'Approval Workflow', value: 'Approval Workflow' },
                { label: 'Escalation Workflow', value: 'Escalation Workflow' },
              ]}
              disabled={isViewMode || !!editItem}
            />
          </div>

          {isEscalationWorkflow && (
            <div style={styles.field}>
              <label style={styles.label}>Escalation Criteria (If applicable) <span style={styles.req}>*</span></label>
              <div style={styles.criteriaRow}>
                {[
                  { label: 'High Value / Financial', value: 'High Value / Financial' },
                  { label: 'Contract / Legal', value: 'Contract / Legal' },
                  { label: 'High-Risk Document', value: 'High-Risk Document' }
                ].map(c => (
                  <label key={c.value} className="cw-checkbox-label" style={{ ...styles.criteriaBox, borderColor: escalationCriteria.includes(c.value) ? '#005c42' : '#e2e8f0', backgroundColor: escalationCriteria.includes(c.value) ? '#f0fdf4' : '#fff', opacity: isViewMode ? 0.7 : 1, cursor: isViewMode ? 'default' : 'pointer' }}>
                    <input type="checkbox" checked={escalationCriteria.includes(c.value)} disabled={isViewMode} onChange={() => toggleEscalationCriteria(c.value)} style={{ accentColor: "#005c42" }} />
                    {c.label}
                  </label>
                ))}
              </div>
            </div>
          )}

          <div style={styles.card}>
            <div style={styles.cardHeader}>
              <div style={styles.cardHeaderIconBox}><WorkflowIcon /></div>
              <span style={styles.cardTitle}>Approval Steps</span>
            </div>

            <div style={styles.stepsRow}>
              {visibleSteps.map((s, idx) => {
                const stepNum = idx + 1;
                const label = s.optional ? `Step ${stepNum} (Optional): ${s.title}` : `Step ${stepNum}: ${s.title}`;
                return (
                  <React.Fragment key={s.key}>
                    <div className="cw-step-chip" style={{ ...styles.stepChip, ...(s.optional ? styles.stepChipOptional : {}) }}>
                      <span style={{ ...styles.stepNumber, backgroundColor: s.optional ? '#dc2626' : '#005c42' }}>{stepNum}</span>
                      <span style={{ ...styles.stepLabel, color: s.optional ? '#b91c1c' : '#0f2b22' }}>{label}</span>
                    </div>
                    {idx < visibleSteps.length - 1 && <span style={styles.stepArrow}>→</span>}
                  </React.Fragment>
                );
              })}
            </div>

            <div style={styles.approversRow}>
              <div style={styles.approverItem}><PeopleSelect value={l1Approver} onChange={(name, email) => { setL1Approver(name); setL1ApproverEmail(email || ''); }} placeholder="Select L1 Approver" context={context} disabled={isViewMode} /></div>
              <div style={styles.approverItem}><PeopleSelect value={l2Approver} onChange={(name, email) => { setL2Approver(name); setL2ApproverEmail(email || ''); }} placeholder="Select L2 Approver" context={context} disabled={isViewMode} /></div>
              {isEscalationWorkflow && (
                <div style={styles.approverItem}><PeopleSelect value={escalateTo} onChange={(name, email) => { setEscalateTo(name); setEscalateToEmail(email || ''); }} placeholder="Escalate To" context={context} disabled={isViewMode} /></div>
              )}
              <div style={styles.approverItem}><PeopleSelect value={finalApprover} onChange={(name, email) => { setFinalApprover(name); setFinalApproverEmail(email || ''); }} placeholder="Select Final Approver" context={context} disabled={isViewMode} /></div>
            </div>
          </div>

          <div style={styles.grid2}>
            <div style={styles.card}>
              <div style={styles.cardHeader}>
                <div style={styles.cardHeaderIconBox}><ClockIcon /></div>
                <span style={styles.cardTitle}>Time Frame for Each Step</span>
              </div>
              <div style={styles.field}>
                <label style={styles.label}>Approval Deadline (in Days) <span style={styles.req}>*</span></label>
                <input
                  className="cw-input"
                  type="number"
                  min={0}
                  style={{ ...styles.input, opacity: isViewMode ? 0.7 : 1 }}
                  value={deadlineDays}
                  onChange={e => setDeadlineDays(e.target.value)}
                  placeholder="Days (e.g., 3)"
                  disabled={isViewMode}
                />
              </div>
            </div>

            <div style={styles.card}>
              <div style={styles.cardHeader}>
                <div style={styles.cardHeaderIconBox}><BellIcon /></div>
                <span style={styles.cardTitle}>Notifications</span>
              </div>
              <label className="cw-checkbox-label" style={{ ...styles.plainCheckboxLabel, cursor: isViewMode ? 'default' : 'pointer' }}>
                <input type="checkbox" checked={notifyApprovers} disabled={isViewMode} onChange={e => setNotifyApprovers(e.target.checked)} style={{ accentColor: "#005c42" }} /> Notify Approvers at Each Step
              </label>
              <label className="cw-checkbox-label" style={{ ...styles.plainCheckboxLabel, cursor: isViewMode ? 'default' : 'pointer' }}>
                <input type="checkbox" checked={alertEscalation} disabled={isViewMode} onChange={e => setAlertEscalation(e.target.checked)} style={{ accentColor: "#005c42" }} /> Alert on Escalation / Compliance
              </label>
            </div>
          </div>
        </div>

        <div style={styles.footer}>
          {saveError && <span style={styles.errorText}>{saveError}</span>}
          <div style={{ display: 'flex', gap: '8px', marginLeft: 'auto' }}>
            {isViewMode ? (
              <button type="button" className="cw-btn" style={styles.saveBtn} onClick={onClose}>Close</button>
            ) : (
              <>
                <button type="button" className="cw-btn" style={styles.cancelBtn} onClick={onClose}>Cancel</button>
                <button type="button" className="cw-btn" style={{ ...styles.saveBtn, opacity: saving || !isValid ? 0.6 : 1, cursor: saving ? 'not-allowed' : 'pointer' }} disabled={saving} onClick={handleSave}>
                  <SendIcon /> {saving ? 'Saving...' : editItem ? 'Update Workflow' : 'Save Workflow'}
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 9999, padding: '28px 16px', boxSizing: 'border-box' },
  modalCard: { backgroundColor: '#fff', borderRadius: '16px', width: '100%', maxWidth: '1080px', maxHeight: '90vh', display: 'flex', flexDirection: 'column', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)', overflow: 'hidden', border: '1px solid #e2e8f0', margin: 'auto' },

  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', padding: '18px 20px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#fff' },
  headerIconBox: { width: '38px', height: '38px', borderRadius: '10px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '3px 0 0 0', fontSize: '12px', color: '#64748b' },
  closeBtn: { background: 'none', border: 'none', fontSize: '15px', cursor: 'pointer', color: '#64748b', padding: '5px 9px', borderRadius: '7px', flexShrink: 0 },

  body: { padding: '18px 20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '16px', backgroundColor: '#f8fafc' },

  grid2: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '14px' },
  grid5: { display: 'grid', gridTemplateColumns: 'repeat(5, minmax(0, 1fr))', gap: '10px', marginTop: '12px' },
  approversRow: { display: 'flex', alignItems: 'stretch', gap: '5px', flexWrap: 'nowrap' as const, marginTop: '12px', width: '100%' },
  approverItem: { flex: '1 1 0', minWidth: 0 },

  field: { display: 'flex', flexDirection: 'column', gap: '5px' },
  label: { fontSize: '12px', fontWeight: 600, color: '#334155' },
  req: { color: '#dc2626' },

  selectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 10px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#fff', height: '36px', boxSizing: 'border-box', transition: 'border-color 0.15s ease, box-shadow 0.15s ease' },
  selectMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '10px', boxShadow: '0 10px 25px rgba(15,23,42,0.12)', zIndex: 500, padding: '5px', maxHeight: '220px', overflowY: 'auto' },
  selectOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 9px', fontSize: '12px', borderRadius: '6px', cursor: 'pointer', margin: '1px 0' },
  avatarCircle: { width: '20px', height: '20px', borderRadius: '50%', backgroundColor: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', flexShrink: 0 },
  avatarImg: { width: '100%', height: '100%', objectFit: 'cover' },

  criteriaRow: { display: 'flex', gap: '10px', flexWrap: 'wrap' },
  criteriaBox: { display: 'flex', alignItems: 'center', gap: '8px', border: '1px solid #e2e8f0', borderRadius: '9px', padding: '10px 14px', fontSize: '12.5px', fontWeight: 600, color: '#334155', cursor: 'pointer', flex: '1 1 200px', transition: 'border-color 0.15s ease, background-color 0.15s ease' },

  card: { backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '12px', padding: '14px 16px' },
  cardHeader: { display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' },
  cardHeaderIconBox: { width: '26px', height: '26px', borderRadius: '7px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  cardTitle: { fontSize: '13.5px', fontWeight: 700, color: '#0f172a' },

  stepsRow: { display: 'flex', alignItems: 'stretch', gap: '5px', flexWrap: 'nowrap' as const, margin: '10px 0', width: '100%' },
  stepChip: { flex: '1 1 0', minWidth: 0, display: 'flex', alignItems: 'center', gap: '6px', border: '1px solid #e2e8f0', borderRadius: '9px', padding: '8px 8px', backgroundColor: '#f8fafc', overflow: 'hidden' },
  stepChipOptional: { backgroundColor: '#fef2f2', borderColor: '#fecaca' },
  stepNumber: { width: '18px', height: '18px', borderRadius: '50%', color: '#fff', fontSize: '10px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  stepLabel: { fontSize: '10.5px', fontWeight: 600, lineHeight: 1.25, overflow: 'hidden' as const, textOverflow: 'ellipsis' as const },
  stepArrow: { fontSize: '12px', color: '#94a3b8', alignSelf: 'center', flexShrink: 0 },

  input: { padding: '8px 10px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '12.5px', outline: 'none', backgroundColor: '#fff', color: '#1e293b', boxSizing: 'border-box', height: '36px' },

  plainCheckboxLabel: { display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', color: '#334155', fontWeight: 500, cursor: 'pointer', padding: '6px 0', transition: 'color 0.15s ease' },

  footer: { display: 'flex', alignItems: 'center', gap: '10px', padding: '14px 20px', borderTop: '1px solid #e2e8f0', backgroundColor: '#fff' },
  errorText: { fontSize: '11.5px', color: '#dc2626', fontWeight: 600 },
  cancelBtn: { padding: '9px 18px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#fff', color: '#475569', fontSize: '12.5px', fontWeight: 700, cursor: 'pointer' },
  saveBtn: { display: 'flex', alignItems: 'center', gap: '7px', padding: '9px 18px', borderRadius: '8px', border: 'none', backgroundColor: '#005c42', color: '#fff', fontSize: '12.5px', fontWeight: 700 }
};

export default CreateWorkflow;