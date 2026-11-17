import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse, photoUrlFor } from '../../spfx-shim';
import AssignmentADGroups from './AssignmentADGroups';

export interface IProcessFormData {
  processName: string;
  processCode: string;
  department: string;
  category: string;
  processOwner: string;
  reviewFrequency: string;
  kraCode: string;
  weightage: number;
  effectiveDate: string;
  nextReviewDate: string;
  description: string;
  status: 'Active' | 'Pending Approval' | 'Draft' | 'Archived';
}

interface IAssignmentGroup {
  id: number;
  role: string;
  entraGroup: string;
  groupType: string;
  description: string;
  assignedUsers: string[];
}

interface ProcessManagementFormProps {
  isOpen: boolean;
  onClose: () => void;
  onSubmit: (formData: IProcessFormData) => void;
  initialData?: IProcessFormData;
  isViewMode?: boolean;
  graphContext?: any;
  selectedId?: number;
}

interface FormSelectOption { label: string; value: string; }

interface FormSelectProps {
  label?: string; required?: boolean; value: string; onChange: (val: string) => void;
  options: FormSelectOption[]; placeholder?: string; width?: string; disabled?: boolean;
}

const FormSelect: React.FC<FormSelectProps> = ({ label, required, value, onChange, options, placeholder = 'Select', width, disabled }) => {
  const [isOpen, setIsOpen] = useState(false);
  const wrapperRef = useRef<HTMLDivElement>(null);
  const selected = options.find((opt) => opt.value === value);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) { setIsOpen(false); }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div style={{ ...styles.field, width, position: 'relative' }} ref={wrapperRef}>
      {label && (<label style={styles.label}>{label} {required && <span style={styles.req}>*</span>}</label>)}
      <div className="fm-select-box" style={{ ...styles.selectBox, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none', opacity: disabled ? 0.6 : 1, cursor: disabled ? 'not-allowed' : 'pointer' }} onClick={() => !disabled && setIsOpen((prev) => !prev)}>
        <span style={{ fontSize: '11px', color: selected ? '#1e293b' : '#94a3b8', fontWeight: 500 }}>{selected ? selected.label : placeholder}</span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </div>
      {isOpen && !disabled && (
        <div className="fm-dropdown-menu" style={styles.selectMenu}>
          {options.map((opt) => {
            const isSelected = opt.value === value;
            return (
              <div key={opt.value} className="fm-select-option" style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent', color: isSelected ? '#004d3d' : '#334155', fontWeight: isSelected ? 600 : 400 }} onClick={() => { onChange(opt.value); setIsOpen(false); }}>
                <span>{opt.label}</span>
                {isSelected && <span style={{ color: '#004d3d', fontWeight: 700 }}>✓</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

export interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

interface PeopleSelectProps {
  label?: string;
  required?: boolean;
  value: string;
  onChange: (val: string) => void;
  disabled?: boolean;
  placeholder?: string;
  graphContext?: any;
}

const PeopleSelect: React.FC<PeopleSelectProps> = ({ label, required, value, onChange, disabled, placeholder = 'Select', graphContext }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [msUsers, setMsUsers] = useState<IGraphUser[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        if (!graphContext?.msGraphClientFactory) return;
        const client = await graphContext.msGraphClientFactory.getClient();
        const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(50).get();
        const siteUrl = graphContext.pageContext?.web?.absoluteUrl || '';

        const usersList: IGraphUser[] = response.value.map((u: any) => {
          const email = u.mail || u.userPrincipalName || '';
          return {
            id: u.id,
            displayName: u.displayName || 'No Name',
            mail: email,
            userPrincipalName: u.userPrincipalName || '',
            photoUrl: siteUrl ? photoUrlFor(email) : '',
          };
        });
        setMsUsers(usersList);
      } catch (err) {
        console.error('Graph API Error for Process Owner:', err);
      }
    };
    fetchUsers();
  }, [graphContext]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) { setIsOpen(false); }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredUsers = msUsers.filter(
    (u) =>
      u.displayName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      u.mail.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const selectedUserObj = msUsers.find((u) => u.displayName === value);

  return (
    <div style={{ ...styles.field, position: 'relative' }} ref={wrapperRef}>
      {label && (<label style={styles.label}>{label} {required && <span style={styles.req}>*</span>}</label>)}
      <div
        className="fm-select-box"
        style={{ ...styles.selectBox, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none', opacity: disabled ? 0.6 : 1, cursor: disabled ? 'not-allowed' : 'pointer' }}
        onClick={() => !disabled && setIsOpen((prev) => !prev)}
      >
        <span style={{ display: 'flex', alignItems: 'center', gap: '6px', overflow: 'hidden', minWidth: 0 }}>
          <span style={styles.avatarCircle}>
            {selectedUserObj?.photoUrl ? (
              <img src={selectedUserObj.photoUrl} alt="" style={styles.avatarImg} />
            ) : (
              <span style={{ fontSize: '10px', color: '#64748b' }}>👤</span>
            )}
          </span>
          <span style={{ fontSize: '11px', color: value ? '#1e293b' : '#94a3b8', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{value || placeholder}</span>
        </span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </div>
      {isOpen && !disabled && (
        <div className="fm-dropdown-menu" style={{ ...styles.selectMenu, padding: '4px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 8px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', marginBottom: '4px' }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
            <input
              autoFocus
              type="text"
              placeholder="Search name or email..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              style={{ border: 'none', outline: 'none', fontSize: '11px', width: '100%', background: 'transparent', color: '#334155' }}
              onClick={(e) => e.stopPropagation()}
            />
          </div>
          <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
            {filteredUsers.length === 0 ? (
              <div style={{ fontSize: '11px', color: '#94a3b8', padding: '10px', textAlign: 'center' }}>No users found</div>
            ) : (
              filteredUsers.map((u) => {
                const isSelected = u.displayName === value;
                return (
                  <div
                    key={u.id}
                    className="fm-select-option"
                    style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent' }}
                    onClick={() => { onChange(u.displayName); setIsOpen(false); setSearchQuery(''); }}
                  >
                    <span style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden', minWidth: 0 }}>
                      <span style={styles.avatarCircle}>
                        {u.photoUrl ? (
                          <img src={u.photoUrl} alt="" style={styles.avatarImg} />
                        ) : (
                          <span style={{ fontSize: '10px', color: '#64748b' }}>👤</span>
                        )}
                      </span>
                      <span style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden', minWidth: 0 }}>
                        <span style={{ fontSize: '11px', fontWeight: isSelected ? 600 : 500, color: isSelected ? '#004d3d' : '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</span>
                        <span style={{ fontSize: '9px', color: '#94a3b8', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.mail}</span>
                      </span>
                    </span>
                    {isSelected && <span style={{ color: '#004d3d', fontWeight: 700, flexShrink: 0 }}>✓</span>}
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

interface PeopleMultiSelectProps {
  value: string[];
  onChange: (val: string[]) => void;
  disabled?: boolean;
  graphContext?: any;
}

const PeopleMultiSelect: React.FC<PeopleMultiSelectProps> = ({ value, onChange, disabled, graphContext }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [msUsers, setMsUsers] = useState<IGraphUser[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const wrapperRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        if (!graphContext?.msGraphClientFactory) return;
        const client = await graphContext.msGraphClientFactory.getClient();
        const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(50).get();
        const siteUrl = graphContext.pageContext?.web?.absoluteUrl || '';

        const usersList: IGraphUser[] = response.value.map((u: any) => {
          const email = u.mail || u.userPrincipalName || '';
          return {
            id: u.id,
            displayName: u.displayName || 'No Name',
            mail: email,
            userPrincipalName: u.userPrincipalName || '',
            photoUrl: siteUrl ? photoUrlFor(email) : '',
          };
        });
        setMsUsers(usersList);
      } catch (err) {
        console.error('Graph API Error for RACI User Picker:', err);
      }
    };
    fetchUsers();
  }, [graphContext]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) { setIsOpen(false); }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredUsers = msUsers.filter(
    (u) =>
      u.displayName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      u.mail.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const toggleUser = (u: IGraphUser) => {
    if (value.includes(u.displayName)) {
      onChange(value.filter((v) => v !== u.displayName));
    } else {
      onChange([...value, u.displayName]);
    }
  };

  return (
    <div style={{ position: 'relative' }} ref={wrapperRef}>
      <button
        type="button"
        className="fm-icon-btn"
        style={styles.userAddIcon}
        title="Add User"
        onClick={() => !disabled && setIsOpen((prev) => !prev)}
      >
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="8.5" cy="7" r="4"></circle><line x1="20" y1="8" x2="20" y2="14"></line><line x1="23" y1="11" x2="17" y2="11"></line></svg>
      </button>
      {isOpen && !disabled && (
        <div className="fm-dropdown-menu" style={{ ...styles.selectMenu, top: 'calc(100% + 4px)', left: 'auto', right: 0, width: '260px', padding: '4px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 8px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', marginBottom: '4px' }}>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
            <input
              autoFocus
              type="text"
              placeholder="Search name or email..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              style={{ border: 'none', outline: 'none', fontSize: '11px', width: '100%', background: 'transparent', color: '#334155' }}
              onClick={(e) => e.stopPropagation()}
            />
          </div>
          <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
            {filteredUsers.length === 0 ? (
              <div style={{ fontSize: '11px', color: '#94a3b8', padding: '10px', textAlign: 'center' }}>No users found</div>
            ) : (
              filteredUsers.map((u) => {
                const isSelected = value.includes(u.displayName);
                return (
                  <div
                    key={u.id}
                    className="fm-select-option"
                    style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent' }}
                    onClick={() => toggleUser(u)}
                  >
                    <span style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden', minWidth: 0 }}>
                      <span style={styles.avatarCircle}>
                        {u.photoUrl ? (
                          <img src={u.photoUrl} alt="" style={styles.avatarImg} />
                        ) : (
                          <span style={{ fontSize: '10px', color: '#64748b' }}>👤</span>
                        )}
                      </span>
                      <span style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden', minWidth: 0 }}>
                        <span style={{ fontSize: '11px', fontWeight: isSelected ? 600 : 500, color: isSelected ? '#004d3d' : '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</span>
                        <span style={{ fontSize: '9px', color: '#94a3b8', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.mail}</span>
                      </span>
                    </span>
                    {isSelected && <span style={{ color: '#004d3d', fontWeight: 700, flexShrink: 0 }}>✓</span>}
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

export const ProcessManagementForm: React.FC<ProcessManagementFormProps> = ({ isOpen, onClose, onSubmit, initialData, isViewMode = false, graphContext, selectedId }) => {
  const [processName, setProcessName] = useState('');
  const [processCode, setProcessCode] = useState('');
  const [department, setDepartment] = useState('');
  const [category, setCategory] = useState('');
  const [processOwner, setProcessOwner] = useState('');
  const [reviewFrequency, setReviewFrequency] = useState('');
  const [kraCode, setKraCode] = useState('');
  const [weightage, setWeightage] = useState('');
  const [effectiveDate, setEffectiveDate] = useState('');
  const [nextReviewDate, setNextReviewDate] = useState('');
  const [description, setDescription] = useState('');

  const [isGeneralInfoSubmitted, setIsGeneralInfoSubmitted] = useState(false);
  const [isSavingGeneralInfo, setIsSavingGeneralInfo] = useState(false);
  const [generalInfoSaveError, setGeneralInfoSaveError] = useState<string | null>(null);

  // Holds the ID of the Process Master item created from this form (create flow),
  // so that a later "Edit" + submit updates the same item instead of creating a new one.
  const [createdItemId, setCreatedItemId] = useState<number | null>(null);
  const currentItemId = selectedId ?? createdItemId;

  const [groups, setGroups] = useState<IAssignmentGroup[]>([]);
  const [isGroupModalOpen, setIsGroupModalOpen] = useState(false);
  const [editingGroup, setEditingGroup] = useState<IAssignmentGroup | null>(null);

  const [responsible, setResponsible] = useState<string[]>([]);
  const [accountable, setAccountable] = useState<string[]>([]);
  const [consulted, setConsulted] = useState<string[]>([]);
  const [informed, setInformed] = useState<string[]>([]);

  const [raciItemId, setRaciItemId] = useState<number | null>(null);
  const [isSavingRaci, setIsSavingRaci] = useState(false);
  const [raciSaveError, setRaciSaveError] = useState<string | null>(null);

  const [approvalRequired, setApprovalRequired] = useState(false);
  const [escalationRequired, setEscalationRequired] = useState(false);
  const [escalationValue, setEscalationValue] = useState('');
  const [escalationUnit, setEscalationUnit] = useState('Days');
  const [sendEmail, setSendEmail] = useState(false);
  const [autoTasks, setAutoTasks] = useState(false);
  const [allowRework, setAllowRework] = useState(false);

  // Held only to resolve the Process Owner's display name back to their email/UPN
  // (see restrictProcessFolderToOwner below).
  const [allGraphUsers, setAllGraphUsers] = useState<IGraphUser[]>([]);

  useEffect(() => {
    const fetchAllUsers = async () => {
      try {
        if (!graphContext?.msGraphClientFactory) return;
        const client = await graphContext.msGraphClientFactory.getClient();
        const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(999).get();
        const usersList: IGraphUser[] = response.value.map((u: any) => ({
          id: u.id,
          displayName: u.displayName || 'No Name',
          mail: u.mail || u.userPrincipalName || '',
          userPrincipalName: u.userPrincipalName || '',
        }));
        setAllGraphUsers(usersList);
      } catch (err) {
        console.error('Graph API Error while pre-loading users for folder permissions:', err);
      }
    };
    fetchAllUsers();
  }, [graphContext]);

  useEffect(() => {
    if (initialData) {
      setProcessName(initialData.processName || '');
      setProcessCode(initialData.processCode || '');
      setDepartment(initialData.department || '');
      setCategory(initialData.category || '');
      setProcessOwner(initialData.processOwner || '');
      setReviewFrequency(initialData.reviewFrequency || '');
      setKraCode(initialData.kraCode || '');
      setWeightage(initialData.weightage !== undefined ? String(initialData.weightage) : '');
      setEffectiveDate(initialData.effectiveDate || '');
      setNextReviewDate(initialData.nextReviewDate || '');
      setDescription(initialData.description || '');
      setIsGeneralInfoSubmitted(true);

      // Fetch the existing RACI record from the RACI Matrix list using the Process Code (edit mode).
      if (initialData.processCode) {
        fetchRaciFromSharePoint(initialData.processCode);
        fetchGroupsFromSharePoint(initialData.processCode);
      }
    } else {
      setProcessName('');
      setProcessCode('');
      setDepartment('');
      setCategory('');
      setProcessOwner('');
      setReviewFrequency('');
      setKraCode('');
      setWeightage('');
      setEffectiveDate('');
      setNextReviewDate('');
      setDescription('');
      setIsGeneralInfoSubmitted(false);

      // Reset the created item ID when a fresh form opens.
      setCreatedItemId(null);

      // Reset the RACI state when a new form opens.
      setRaciItemId(null);
      setResponsible([]);
      setAccountable([]);
      setConsulted([]);
      setInformed([]);
      setRaciSaveError(null);

      // Also reset the Assignment Groups when a new form opens.
      setGroups([]);
    }
  }, [initialData, isOpen]);

  const isGeneralInfoValid = Boolean(
    processName.trim() &&
    processCode.trim() &&
    department.trim() &&
    category.trim() &&
    processOwner.trim() &&
    reviewFrequency.trim() &&
    kraCode.trim() &&
    weightage.trim()
  );

  if (!isOpen) return null;

  const handleClearRaci = () => {
    if (isViewMode) return;
    setResponsible([]);
    setAccountable([]);
    setConsulted([]);
    setInformed([]);
  };

  const handleRemoveTag = (list: string[], setList: React.Dispatch<React.SetStateAction<string[]>>, tagToRemove: string) => {
    if (isViewMode) return;
    setList(list.filter((tag) => tag !== tagToRemove));
  };

  const handleFormSubmit = async (status: 'Active' | 'Pending Approval' | 'Draft') => {
    if (isViewMode) return;

    // Please create and update the RACI Matrix list in real time first. Only then proceed to the next step.
    const raciSaved = await saveRaciToSharePoint();
    if (!raciSaved) return;

    onSubmit({
      processName, processCode, department, category, processOwner, reviewFrequency, kraCode,
      weightage: Number(weightage) || 0, effectiveDate, nextReviewDate, description, status,
    });
  };

  const PROCESS_MASTER_LIST_TITLE = 'Process Master';
  const RACI_LIST_TITLE = 'RACI Matrix';
  const RACI_PROCESS_ID_FIELD = 'ProcessID';
  const ASSIGNMENT_GROUPS_LIST_TITLE = 'Process Assignment Groups';
  const DOCUMENT_LIBRARY_NAME = 'Shared Documents';
  const PROCESS_FOLDER_PERMISSION_LEVEL = 'Edit';

  // Shared audit trail list — every module logs its entries here, one row per action.
  // https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
  const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';

  // yyyyMMddHHmmss — exact moment an audit entry was written.
  const generateAuditId = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  };

  // Writes ONE row into "DMS Audit Logs". Never blocks the caller's own save —
  // any failure here is only logged to console so the real save isn't affected.
  const addDmsAuditLogEntry = async (fields: { [key: string]: any }): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) return;
    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${AUDIT_LOG_LIST_TITLE}')/items`;
      const res: SPHttpClientResponse = await graphContext.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify({ AuditID: generateAuditId(), ...fields })
        }
      );
      if (!res.ok) {
        console.error('DMS Audit Logs entry failed:', await res.text());
      }
    } catch (err) {
      console.error('Error writing DMS Audit Logs entry:', err);
    }
  };

  // Makes the process folder visible ONLY to the Process Owner. Runs right after the folder is
  // created (and on a later General Information save for older folders). It only acts while the
  // folder still inherits the library permissions — if it already has its own permissions
  // (e.g. from the RACI save) they are left untouched.
  const restrictProcessFolderToOwner = async (folderName: string, ownerName: string): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) return;
    const trimmedName = folderName.trim();
    if (!trimmedName || !ownerName.trim()) return;

    let ownerEmail: string = allGraphUsers.find((u) => u.displayName === ownerName)?.mail || '';
    if (!ownerEmail && graphContext.pageContext.user?.displayName === ownerName) {
      ownerEmail = graphContext.pageContext.user?.email || graphContext.pageContext.user?.loginName || '';
    }
    if (!ownerEmail) {
      // Without the owner's email we cannot grant access — leave the folder as it is instead of locking everyone out.
      console.warn('Could not resolve Process Owner email; process folder was not restricted.');
      return;
    }

    const jsonHeaders = { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const webServerRelativeUrl: string = graphContext.pageContext.web.serverRelativeUrl || '';
      const libraryRelativeUrl = webServerRelativeUrl && webServerRelativeUrl !== '/'
        ? `${webServerRelativeUrl}/${DOCUMENT_LIBRARY_NAME}`
        : `/${DOCUMENT_LIBRARY_NAME}`;
      const folderServerRelativeUrl = `${libraryRelativeUrl}/${trimmedName}`;
      const listEndpoint = `${siteUrl}/_api/web/GetList('${encodeURIComponent(libraryRelativeUrl)}')`;

      const itemLookupRes = await graphContext.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/ListItemAllFields?$select=Id,HasUniqueRoleAssignments`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!itemLookupRes.ok) {
        console.warn('Could not resolve process folder item for owner restriction:', await itemLookupRes.text());
        return;
      }
      const itemData = await itemLookupRes.json();
      if (itemData.HasUniqueRoleAssignments) return;
      const itemId = itemData.Id;

      const breakRes = await graphContext.spHttpClient.post(
        `${listEndpoint}/items(${itemId})/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!breakRes.ok) {
        console.warn('Could not break role inheritance on process folder:', await breakRes.text());
        return;
      }

      const roleDefRes = await graphContext.spHttpClient.get(
        `${siteUrl}/_api/web/roledefinitions/getbyname('${PROCESS_FOLDER_PERMISSION_LEVEL}')`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!roleDefRes.ok) {
        console.warn('Could not resolve "Edit" role definition:', await roleDefRes.text());
        return;
      }
      const roleDefId = (await roleDefRes.json()).Id;

      const ensureRes = await graphContext.spHttpClient.post(
        `${siteUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders, body: JSON.stringify({ logonName: ownerEmail }) }
      );
      if (!ensureRes.ok) {
        console.warn(`Could not resolve SharePoint user for ${ownerEmail}:`, await ensureRes.text());
        return;
      }
      const principalId = (await ensureRes.json()).Id;

      const assignRes = await graphContext.spHttpClient.post(
        `${listEndpoint}/items(${itemId})/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!assignRes.ok) {
        console.warn(`Could not grant Process Owner "${ownerEmail}" access to folder "${trimmedName}":`, await assignRes.text());
      }
    } catch (err) {
      console.error('Error restricting process folder to Process Owner:', err);
    }
  };

  const createProcessDocumentFolder = async (folderName: string): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) return;
    const trimmedName = folderName.trim();
    if (!trimmedName) return;

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const webServerRelativeUrl: string = graphContext.pageContext.web.serverRelativeUrl || '';
      const libraryRelativeUrl = webServerRelativeUrl && webServerRelativeUrl !== '/'
        ? `${webServerRelativeUrl}/${DOCUMENT_LIBRARY_NAME}`
        : `/${DOCUMENT_LIBRARY_NAME}`;
      const folderServerRelativeUrl = `${libraryRelativeUrl}/${trimmedName}`;

      const endpoint = `${siteUrl}/_api/web/folders/add('${encodeURIComponent(folderServerRelativeUrl)}')`;

      const response: SPHttpClientResponse = await graphContext.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        // Folder may already exist for this Process Code — safe to ignore, not fatal.
        const errText = await response.text();
        console.warn('Process document folder creation warning:', errText);
      }
    } catch (err) {
      console.error('Error creating process document folder in SharePoint:', err);
    }
  };

  const saveGeneralInfoToSharePoint = async (): Promise<boolean> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) {
      console.error('SharePoint context (graphContext) is not available on ProcessManagementForm.');
      setGeneralInfoSaveError('Failed to connect to SharePoint. Please check the console.');
      return false;
    }

    setIsSavingGeneralInfo(true);
    setGeneralInfoSaveError(null);

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      // Use the item ID from the list (edit flow) or the one captured after creation (in-form Edit).
      const isEditing = Boolean(currentItemId);
      const endpoint = isEditing
        ? `${siteUrl}/_api/web/lists/getByTitle('${PROCESS_MASTER_LIST_TITLE}')/items(${currentItemId})`
        : `${siteUrl}/_api/web/lists/getByTitle('${PROCESS_MASTER_LIST_TITLE}')/items`;

      const bodyData: { [key: string]: any } = {
        Title: processName,
        ProcessCode: processCode,
        Department: department,
        ProcessCategory: category,
        ProcessOwner: processOwner,
        ReviewFrequency: reviewFrequency,
        KRACode: kraCode,
        Weightage: String(Number(weightage) || 0),
        EffectiveDate: effectiveDate ? new Date(effectiveDate).toISOString() : null,
        NextReviewDate: nextReviewDate ? new Date(nextReviewDate).toISOString() : null,
        Description: description,
      };

      // Also store the selected Process Owner's email in the "ProcessOwnerEmail" column
      // (same lookup as restrictProcessFolderToOwner). Skipped if the email cannot be resolved,
      // so an existing value is never overwritten with an empty one.
      let processOwnerEmail: string = allGraphUsers.find((u) => u.displayName === processOwner)?.mail || '';
      if (!processOwnerEmail && graphContext.pageContext.user?.displayName === processOwner) {
        processOwnerEmail = graphContext.pageContext.user?.email || graphContext.pageContext.user?.loginName || '';
      }
      if (processOwnerEmail) {
        bodyData.ProcessOwnerEmail = processOwnerEmail;
      }

      const headers: any = {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      };

      if (isEditing) {
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-Method'] = 'PATCH';
      }

      const response: SPHttpClientResponse = await graphContext.spHttpClient[isEditing ? 'post' : 'post'](
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers,
          body: JSON.stringify(bodyData),
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error('SharePoint "Process Master" save failed:', errText);
        setGeneralInfoSaveError('SharePoint could not be saved. Please check the console.');
        return false;
      }

      if (!isEditing) {
        // Capture the newly created item's ID so the next save is an update (PATCH), not a new create.
        try {
          const created = await response.json();
          if (created?.Id) setCreatedItemId(created.Id);
        } catch (parseErr) {
          console.warn('Could not read the created Process Master item ID:', parseErr);
        }
        await createProcessDocumentFolder(processCode);
      }

      // Folder is visible only to the Process Owner (nobody else) until RACI permissions are applied.
      await restrictProcessFolderToOwner(processCode, processOwner);

      // One "DMS Audit Logs" entry per Process Master save (create or update).
      await addDmsAuditLogEntry({
        Title: processCode,          // "Title" column — renamed to "ProcessID" in the list UI
        EntryType: 'Process',
        ProcessName: processName,
        Department: department,
        ProcessCategory: category,
        ProcessOwner: processOwner,
        ReviewFrequency: reviewFrequency,
        KRACode: kraCode,
        Weightage: String(Number(weightage) || 0),
        EffectiveDate: effectiveDate ? new Date(effectiveDate).toISOString() : null,
        NextReviewDate: nextReviewDate ? new Date(nextReviewDate).toISOString() : null,
        Description: description
      });

      return true;
    } catch (err) {
      console.error('Error saving General Information to SharePoint:', err);
      setGeneralInfoSaveError('An error occurred while saving to SharePoint.');
      return false;
    } finally {
      setIsSavingGeneralInfo(false);
    }
  };

  const fetchRaciFromSharePoint = async (procCode: string): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext || !procCode.trim()) return;

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const safeCode = procCode.trim().replace(/'/g, "''");
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${RACI_LIST_TITLE}')/items?$filter=${RACI_PROCESS_ID_FIELD} eq '${safeCode}'&$top=1`;

      const response: SPHttpClientResponse = await graphContext.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.warn('RACI Matrix fetch warning (the record may not exist):', errText);
        setRaciItemId(null);
        return;
      }

      const data = await response.json();
      const existingItem = data?.value?.[0];

      if (existingItem) {
        setRaciItemId(existingItem.Id);
        setResponsible(existingItem.Responsible ? String(existingItem.Responsible).split(';').map((s: string) => s.trim()).filter(Boolean) : []);
        setAccountable(existingItem.Accountable ? String(existingItem.Accountable).split(';').map((s: string) => s.trim()).filter(Boolean) : []);
        setConsulted(existingItem.Consulted ? String(existingItem.Consulted).split(';').map((s: string) => s.trim()).filter(Boolean) : []);
        setInformed(existingItem.Informed ? String(existingItem.Informed).split(';').map((s: string) => s.trim()).filter(Boolean) : []);
      } else {
        setRaciItemId(null);
      }
    } catch (err) {
      console.error('Error fetching RACI Matrix item from SharePoint:', err);
      setRaciItemId(null);
    }
  };

// Fetch all Entra ID Group rows for this Process Code from the "Process Assignment Groups" library.
// (Title column = Process Code). Expand the AssignedUsers Person field to retrieve display names.
  const fetchGroupsFromSharePoint = async (procCode: string): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext || !procCode.trim()) return;

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const safeCode = procCode.trim().replace(/'/g, "''");
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${ASSIGNMENT_GROUPS_LIST_TITLE}')/items?$select=Id,Title,Role,EntraIDGroupName,GroupType,Description,AssignedUsers&$filter=Title eq '${safeCode}'`;

      const response: SPHttpClientResponse = await graphContext.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.warn('Process Assignment Groups fetch warning:', errText);
        setGroups([]);
        return;
      }

      const data = await response.json();
      const items = data?.value || [];
      const mapped: IAssignmentGroup[] = items.map((item: any) => ({
        id: item.Id,
        role: item.Role || '',
        entraGroup: item.EntraIDGroupName || '',
        groupType: item.GroupType || '',
        description: item.Description || '',
        assignedUsers: item.AssignedUsers ? String(item.AssignedUsers).split(';').map((s: string) => s.trim()).filter(Boolean) : [],
      }));
      setGroups(mapped);
    } catch (err) {
      console.error('Error fetching Process Assignment Groups from SharePoint:', err);
      setGroups([]);
    }
  };

  const deleteGroupFromSharePoint = async (itemId: number): Promise<boolean> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) return false;
    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${ASSIGNMENT_GROUPS_LIST_TITLE}')/items(${itemId})`;
      const response: SPHttpClientResponse = await graphContext.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } }
      );
      if (!response.ok) {
        console.error('SharePoint "Process Assignment Groups" delete failed:', await response.text());
        return false;
      }
      return true;
    } catch (err) {
      console.error('Error deleting Process Assignment Group from SharePoint:', err);
      return false;
    }
  };

  const saveRaciToSharePoint = async (): Promise<boolean> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) {
      console.error('SharePoint context (graphContext) is not available for RACI Matrix.');
      setRaciSaveError('Failed to connect to SharePoint. Please check the console.');
      return false;
    }

    if (!processCode.trim()) {
      setRaciSaveError('The **Process Code** is missing — please save the **General Information** first.');
      return false;
    }

    /*if (responsible.length === 0 || accountable.length === 0) {
      setRaciSaveError('At least one name is mandatory in Responsible (R) and Accountable (A).');
      return false;
    } */

    setIsSavingRaci(true);
    setRaciSaveError(null);

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const isEditingRaci = Boolean(raciItemId);
      const endpoint = isEditingRaci
        ? `${siteUrl}/_api/web/lists/getByTitle('${RACI_LIST_TITLE}')/items(${raciItemId})`
        : `${siteUrl}/_api/web/lists/getByTitle('${RACI_LIST_TITLE}')/items`;

      const bodyData: { [key: string]: any } = {
        Title: processName || processCode,
        [RACI_PROCESS_ID_FIELD]: processCode,
        Responsible: responsible.join('; '),
        Accountable: accountable.join('; '),
        Consulted: consulted.join('; '),
        Informed: informed.join('; '),
      };

      const headers: any = {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      };

      if (isEditingRaci) {
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-Method'] = 'PATCH';
      }

      const response: SPHttpClientResponse = await graphContext.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers,
          body: JSON.stringify(bodyData),
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error('SharePoint "RACI Matrix" save failed:', errText);
        setRaciSaveError('The RACI Matrix could not be saved to SharePoint. Please check the console for the error.');
        return false;
      }

      if (!isEditingRaci) {
        const created = await response.json();
        if (created?.Id) setRaciItemId(created.Id);
      }

      // One "DMS Audit Logs" entry per RACI Matrix save (create or update).
      await addDmsAuditLogEntry({
        Title: processCode,          // "Title" column — renamed to "ProcessID" in the list UI
        EntryType: 'RACI',
        ProcessName: processName || processCode,
        Responsible: responsible.join('; '),
        Accountable: accountable.join('; '),
        Consulted: consulted.join('; '),
        Informed: informed.join('; ')
      });

      return true;
    } catch (err) {
      console.error('Error saving RACI Matrix to SharePoint:', err);
      setRaciSaveError('An error occurred while saving the RACI Matrix.');
      return false;
    } finally {
      setIsSavingRaci(false);
    }
  };

  return (
    <>
    <div className="fm-overlay" style={styles.overlay} onClick={onClose}>
      <style>{`
        @keyframes fmFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes fmScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        @keyframes fmSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .fm-overlay { animation: fmFadeIn 0.18s ease; }
        .fm-modal { animation: fmScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
        .fm-dropdown-menu { animation: fmSlideDown 0.15s ease; }
        .fm-select-box:hover { border-color: #94a3b8; }
        .fm-select-option:hover { background-color: #f1f5f9; }
        .fm-select-option:hover span:first-child { color: inherit; }
        .fm-btn { transition: transform 0.15s ease, filter 0.15s ease, background-color 0.15s ease, box-shadow 0.15s ease; }
        .fm-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
        .fm-btn:active { transform: translateY(0); filter: brightness(0.97); }
        .fm-icon-btn { transition: background-color 0.15s ease, transform 0.15s ease; border-radius: 5px; }
        .fm-icon-btn:hover { background-color: #f1f5f9; transform: scale(1.08); }
        .fm-input, .fm-textarea { transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .fm-input:hover, .fm-textarea:hover { border-color: #94a3b8; }
        .fm-input:focus, .fm-textarea:focus { border-color: #004d3d; box-shadow: 0 0 0 3px rgba(0, 77, 61, 0.10); }
        .fm-row:hover { background-color: #f8fafc; }
        .fm-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; }
        .fm-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
        .fm-checkbox-label:hover { color: #004d3d; }
        .fm-tag-close:hover { color: #dc2626; }
        .fm-body::-webkit-scrollbar { width: 8px; }
        .fm-body::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 8px; }
        .fm-body::-webkit-scrollbar-track { background-color: transparent; }
      `}</style>

      <div className="fm-modal" style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
        <div style={styles.header}>
          <h2 style={styles.title}>
            {isViewMode ? 'View Process' : initialData ? 'Edit Process' : 'Process Master Form'}
          </h2>
          <button className="fm-close-btn" style={styles.closeBtn} onClick={onClose} title="Close Modal">✕</button>
        </div>

        <div className="fm-body" style={styles.body}>
          <div style={styles.card}>
            <div style={styles.cardHeader}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
              <span style={styles.cardTitle}>1. General Information</span>
            </div>
            <div style={styles.grid4}>
              <div style={styles.field}><label style={styles.label}>Process Name <span style={styles.req}>*</span></label><input className="fm-input" type="text" disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.input, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={processName} onChange={(e) => setProcessName(e.target.value)} placeholder="Enter process name" /></div>
              <div style={styles.field}><label style={styles.label}>Process Code <span style={styles.req}>*</span></label><input className="fm-input" type="text" disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.input, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={processCode} onChange={(e) => setProcessCode(e.target.value)} placeholder="Enter process code" /></div>

              <FormSelect label="Department" required value={department} onChange={setDepartment} disabled={isViewMode || isGeneralInfoSubmitted} options={[
                { label: 'Purchase', value: 'Purchase' }, { label: 'Production', value: 'Production' }, { label: 'Quality', value: 'Quality' },
                { label: 'Quality Management', value: 'Quality Management' }, { label: 'Maintenance', value: 'Maintenance' }, { label: 'HR', value: 'HR' },
                { label: 'Finance', value: 'Finance' }, { label: 'Customer Service', value: 'Customer Service' }, { label: 'IT', value: 'IT' },
              ]} />

              <FormSelect label="Process Category" required value={category} onChange={setCategory} disabled={isViewMode || isGeneralInfoSubmitted} options={[
                { label: 'SOP', value: 'SOP' }, { label: 'Policy', value: 'Policy' }, { label: 'Guideline', value: 'Guideline' },
                { label: 'Contract Agreement', value: 'Contract Agreement' }, { label: 'MIS File', value: 'MIS File' }, { label: 'Reference Document', value: 'Reference Document' }, { label: 'Checklist', value: 'Checklist' },
              ]} />

              <PeopleSelect label="Process Owner" required value={processOwner} onChange={setProcessOwner} disabled={isViewMode || isGeneralInfoSubmitted} graphContext={graphContext} placeholder="Select" />

              <FormSelect label="Review Frequency" required value={reviewFrequency} onChange={setReviewFrequency} disabled={isViewMode || isGeneralInfoSubmitted} options={[
                { label: 'Monthly', value: 'Monthly' }, { label: 'Quarterly', value: 'Quarterly' }, { label: 'Semi-Annual', value: 'Semi-Annual' }, { label: 'Annual', value: 'Annual' },
              ]} />

              <FormSelect label="KRA Code" required value={kraCode} onChange={setKraCode} disabled={isViewMode || isGeneralInfoSubmitted} options={[
                { label: 'KRA-001', value: 'KRA-001' }, { label: 'KRA-002', value: 'KRA-002' }, { label: 'KRA-003', value: 'KRA-003' }, { label: 'KRA-004', value: 'KRA-004' },
                { label: 'KRA-005', value: 'KRA-005' }, { label: 'KRA-006', value: 'KRA-006' }, { label: 'KRA-007', value: 'KRA-007' }, { label: 'KRA-008', value: 'KRA-008' },
              ]} />

              <div style={styles.field}><label style={styles.label}>Weightage (%) <span style={styles.req}>*</span></label><input className="fm-input" type="number" min={0} max={100} disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.input, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={weightage} onChange={(e) => setWeightage(e.target.value)} placeholder="0" /></div>
              <div style={styles.field}><label style={styles.label}>Effective Date</label><input className="fm-input" type="date" disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.input, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={effectiveDate} onChange={(e) => setEffectiveDate(e.target.value)} /></div>
              <div style={styles.field}><label style={styles.label}>Next Review Date</label><input className="fm-input" type="date" disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.input, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={nextReviewDate} onChange={(e) => setNextReviewDate(e.target.value)} /></div>

              <div style={{ ...styles.field, gridColumn: 'span 4' }}>
                <label style={styles.label}>Description</label>
                <textarea className="fm-textarea" disabled={isViewMode || isGeneralInfoSubmitted} style={{ ...styles.textarea, backgroundColor: (isViewMode || isGeneralInfoSubmitted) ? '#f1f5f9' : '#ffffff' }} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Enter process description..." />
              </div>
            </div>

            {!isViewMode && !isGeneralInfoSubmitted && (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '6px', marginTop: '10px' }}>
                {generalInfoSaveError && (
                  <span style={{ fontSize: '10px', color: '#dc2626', fontWeight: 600 }}>{generalInfoSaveError}</span>
                )}
                <button
                  className="fm-btn"
                  style={{
                    ...styles.submitBtn,
                    opacity: (isGeneralInfoValid && !isSavingGeneralInfo) ? 1 : 0.6,
                    cursor: (isGeneralInfoValid && !isSavingGeneralInfo) ? 'pointer' : 'not-allowed'
                  }}
                  disabled={!isGeneralInfoValid || isSavingGeneralInfo}
                  onClick={async () => {
                    if (!isGeneralInfoValid || isSavingGeneralInfo) return;
                    const saved = await saveGeneralInfoToSharePoint();
                    if (saved) {
                      setIsGeneralInfoSubmitted(true);
                      if (processCode.trim()) fetchGroupsFromSharePoint(processCode);
                    }
                  }}
                >
                  {isSavingGeneralInfo ? 'Saving...' : 'Submit & Enable Sections'}
                </button>
              </div>
            )}
            {isGeneralInfoSubmitted && !isViewMode && (
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '10px', backgroundColor: '#e6f4ea', padding: '6px 10px', borderRadius: '6px' }}>
                <span style={{ fontSize: '11px', color: '#137333', fontWeight: 600 }}>✓ General Information saved successfully! Lower sections are now enabled.</span>
                <button
                  className="fm-btn"
                  style={{ background: 'none', border: '1px solid #137333', color: '#137333', borderRadius: '4px', padding: '2px 8px', fontSize: '10px', cursor: 'pointer' }}
                  onClick={() => setIsGeneralInfoSubmitted(false)}
                >
                  Edit
                </button>
              </div>
            )}
          </div>

          <div style={{ ...styles.card, opacity: isGeneralInfoSubmitted ? 1 : 0.5, pointerEvents: isGeneralInfoSubmitted ? 'auto' : 'none' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
              <div style={styles.cardHeader}>
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
                <span style={styles.cardTitle}>2. Assignment Groups / Users</span>
              </div>
              {!isViewMode && (
                <button className="fm-btn" style={styles.addGroupBtn} onClick={() => { setEditingGroup(null); setIsGroupModalOpen(true); }}>+ Add Group/User</button>
              )}
            </div>
            {!isGeneralInfoSubmitted && (
              <div style={{ fontSize: '11px', color: '#b45309', backgroundColor: '#fef3c7', padding: '6px 10px', borderRadius: '6px', marginBottom: '8px' }}>
                Please submit General Information above to enable this module.
              </div>
            )}
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Role</th>
                  <th style={styles.th}>Assigned Users</th>
                  <th style={styles.th}>Description</th>
                  {!isViewMode && <th style={{ ...styles.th, textAlign: 'center' }}>Actions</th>}
                </tr>
              </thead>
              <tbody>
                {groups.length === 0 ? (
                  <tr><td colSpan={isViewMode ? 3 : 4} style={{ ...styles.td, textAlign: 'center', color: '#94a3b8', padding: '16px' }}>No groups/Users added yet. Click "+ Add Group/User" to add.</td></tr>
                ) : (
                  groups.map((grp) => (
                    <tr key={grp.id} className="fm-row">
                      <td style={styles.td}>{grp.role}</td>
                      <td style={styles.td}>{(grp.assignedUsers && grp.assignedUsers.length > 0) ? grp.assignedUsers.join(', ') : '-'}</td>
                      <td style={styles.td}>{grp.description}</td>
                      {!isViewMode && (
                        <td style={{ ...styles.td, textAlign: 'center' }}>
                          <button className="fm-icon-btn" style={styles.actionIconBtn} title="Edit" onClick={() => { setEditingGroup(grp); setIsGroupModalOpen(true); }}>
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                          </button>
                          <button className="fm-icon-btn" style={styles.actionIconBtn} title="Delete" onClick={async () => { const ok = await deleteGroupFromSharePoint(grp.id); if (ok && processCode.trim()) fetchGroupsFromSharePoint(processCode); }}>
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#dc2626" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                          </button>
                        </td>
                      )}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          <div style={{ ...styles.card, opacity: isGeneralInfoSubmitted ? 1 : 0.5, pointerEvents: isGeneralInfoSubmitted ? 'auto' : 'none' }}>
            <div style={styles.cardHeader}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
              <span style={styles.cardTitle}>3. RACI Assignment (Process Level)</span>
            </div>
            <div style={styles.raciList}>
              <div style={styles.raciRow}>
                <span style={styles.raciLabel}>Responsible (R) <span style={styles.req}>*</span></span>
                <div style={styles.tagBox}>
                  <div style={styles.tagsFlex}>
                    {responsible.map((item) => (
                      <span key={item} style={styles.tag}>
                        {item} {!isViewMode && <span className="fm-tag-close" style={styles.tagClose} onClick={() => handleRemoveTag(responsible, setResponsible, item)}>✕</span>}
                      </span>
                    ))}
                  </div>
                  {!isViewMode && (
                    <PeopleMultiSelect value={responsible} onChange={setResponsible} disabled={isViewMode} graphContext={graphContext} />
                  )}
                </div>
              </div>

              <div style={styles.raciRow}>
                <span style={styles.raciLabel}>Accountable (A) <span style={styles.req}>*</span></span>
                <div style={styles.tagBox}>
                  <div style={styles.tagsFlex}>
                    {accountable.map((item) => (
                      <span key={item} style={styles.tag}>
                        {item} {!isViewMode && <span className="fm-tag-close" style={styles.tagClose} onClick={() => handleRemoveTag(accountable, setAccountable, item)}>✕</span>}
                      </span>
                    ))}
                  </div>
                  {!isViewMode && (
                    <PeopleMultiSelect value={accountable} onChange={setAccountable} disabled={isViewMode} graphContext={graphContext} />
                  )}
                </div>
              </div>

              <div style={styles.raciRow}>
                <span style={styles.raciLabel}>Consulted (C)</span>
                <div style={styles.tagBox}>
                  <div style={styles.tagsFlex}>
                    {consulted.map((item) => (
                      <span key={item} style={styles.tag}>
                        {item} {!isViewMode && <span className="fm-tag-close" style={styles.tagClose} onClick={() => handleRemoveTag(consulted, setConsulted, item)}>✕</span>}
                      </span>
                    ))}
                  </div>
                  {!isViewMode && (
                    <PeopleMultiSelect value={consulted} onChange={setConsulted} disabled={isViewMode} graphContext={graphContext} />
                  )}
                </div>
              </div>

              <div style={styles.raciRow}>
                <span style={styles.raciLabel}>Informed (I)</span>
                <div style={styles.tagBox}>
                  <div style={styles.tagsFlex}>
                    {informed.map((item) => (
                      <span key={item} style={styles.tag}>
                        {item} {!isViewMode && <span className="fm-tag-close" style={styles.tagClose} onClick={() => handleRemoveTag(informed, setInformed, item)}>✕</span>}
                      </span>
                    ))}
                  </div>
                  {!isViewMode && (
                    <PeopleMultiSelect value={informed} onChange={setInformed} disabled={isViewMode} graphContext={graphContext} />
                  )}
                </div>
              </div>

              {!isViewMode && (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '6px' }}>
                  <span style={{ fontSize: '10px', color: '#dc2626', fontWeight: 600 }}>{raciSaveError || ''}</span>
                  <button className="fm-btn" style={styles.clearAllBtn} onClick={handleClearRaci}>Clear All</button>
                </div>
              )}
            </div>
          </div>

          <div style={{ ...styles.card, opacity: isGeneralInfoSubmitted ? 1 : 0.5, pointerEvents: isGeneralInfoSubmitted ? 'auto' : 'none' }}>
            <div style={styles.cardHeader}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"></circle><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"></path></svg>
              <span style={styles.cardTitle}>4. Workflow Configuration</span>
            </div>
            <div style={styles.workflowRow}>
              <label className="fm-checkbox-label" style={styles.checkboxLabel}><input type="checkbox" disabled={isViewMode} checked={approvalRequired} onChange={(e) => setApprovalRequired(e.target.checked)} /> Approval Required</label>
              <label className="fm-checkbox-label" style={styles.checkboxLabel}><input type="checkbox" disabled={isViewMode} checked={escalationRequired} onChange={(e) => setEscalationRequired(e.target.checked)} /> Escalation Required</label>
              {escalationRequired && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <span style={styles.label}>Escalation After</span>
                  <input className="fm-input" type="text" disabled={isViewMode} style={{ ...styles.input, width: '36px', textAlign: 'center', backgroundColor: isViewMode ? '#f1f5f9' : '#ffffff' }} value={escalationValue} onChange={(e) => setEscalationValue(e.target.value)} />
                  <FormSelect value={escalationUnit} onChange={setEscalationUnit} disabled={isViewMode} width="80px" options={[{ label: 'Days', value: 'Days' }, { label: 'Hours', value: 'Hours' }]} />
                </div>
              )}
              <label className="fm-checkbox-label" style={styles.checkboxLabel}><input type="checkbox" disabled={isViewMode} checked={sendEmail} onChange={(e) => setSendEmail(e.target.checked)} /> Send Email Notification</label>
              <label className="fm-checkbox-label" style={styles.checkboxLabel}><input type="checkbox" disabled={isViewMode} checked={autoTasks} onChange={(e) => setAutoTasks(e.target.checked)} /> Auto Create Tasks</label>
              <label className="fm-checkbox-label" style={styles.checkboxLabel}><input type="checkbox" disabled={isViewMode} checked={allowRework} onChange={(e) => setAllowRework(e.target.checked)} /> Allow Rework</label>
            </div>
          </div>
        </div>

        <div style={styles.footer}>
          <button className="fm-btn" style={styles.cancelBtn} onClick={onClose}>{isViewMode ? 'Close' : 'Cancel'}</button>
          {!isViewMode && (
            <>
              <button className="fm-btn" style={{ ...styles.activateBtn, opacity: isSavingRaci ? 0.6 : 1, cursor: isSavingRaci ? 'not-allowed' : 'pointer' }} disabled={isSavingRaci} onClick={() => handleFormSubmit('Active')}>{isSavingRaci ? 'Saving...' : (initialData ? 'Update Process' : 'Activate Process')}</button>
            </>
          )}
        </div>
      </div>
    </div>

    <AssignmentADGroups
      isOpen={isGroupModalOpen}
      onClose={() => { setIsGroupModalOpen(false); setEditingGroup(null); }}
      onSave={() => { if (processCode.trim()) fetchGroupsFromSharePoint(processCode); }}
      editingGroup={editingGroup}
      graphContext={graphContext}
      processCode={processCode}
    />
    </>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 9999, padding: '32px 16px', boxSizing: 'border-box' },
  modalCard: { backgroundColor: '#ffffff', borderRadius: '16px', width: '100%', maxWidth: '1080px', maxHeight: '88vh', display: 'flex', flexDirection: 'column', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)', overflow: 'hidden', border: '1px solid #e2e8f0', margin: 'auto' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '14px 20px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#ffffff' },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  closeBtn: { background: 'none', border: 'none', fontSize: '16px', cursor: 'pointer', color: '#64748b', padding: '4px 8px', borderRadius: '6px' },
  body: { padding: '16px 20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '12px', backgroundColor: '#f8fafc' },
  card: { backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '12px 14px' },
  cardHeader: { display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' },
  cardTitle: { fontSize: '12px', fontWeight: 700, color: '#0f172a' },
  grid4: { display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px' },
  field: { display: 'flex', flexDirection: 'column', gap: '3px' },
  label: { fontSize: '10px', fontWeight: 600, color: '#475569' },
  req: { color: '#dc2626' },
  input: { padding: '5px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '11px', outline: 'none', backgroundColor: '#ffffff', color: '#1e293b', boxSizing: 'border-box', height: '28px' },
  textarea: { padding: '6px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '11px', outline: 'none', height: '45px', resize: 'vertical', fontFamily: 'inherit', color: '#1e293b' },
  selectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', height: '28px', boxSizing: 'border-box', transition: 'border-color 0.15s ease, box-shadow 0.15s ease' },
  selectMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 8px 20px rgba(0, 0, 0, 0.10)', zIndex: 500, padding: '4px', maxHeight: '220px', overflowY: 'auto' },
  selectOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 8px', fontSize: '11px', borderRadius: '5px', cursor: 'pointer', margin: '1px 0' },
  avatarCircle: { width: '20px', height: '20px', borderRadius: '50%', backgroundColor: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', flexShrink: 0 },
  avatarImg: { width: '100%', height: '100%', objectFit: 'cover' },
  addGroupBtn: { backgroundColor: '#ffffff', border: '1px solid #004d3d', color: '#004d3d', borderRadius: '6px', padding: '3px 10px', fontSize: '10px', fontWeight: 600, cursor: 'pointer' },
  table: { width: '100%', borderCollapse: 'collapse', fontSize: '11px' },
  th: { textAlign: 'left', padding: '6px 8px', backgroundColor: '#f1f5f9', color: '#475569', fontWeight: '600', fontSize: '10px', borderBottom: '1px solid #e2e8f0' },
  td: { padding: '6px 8px', borderBottom: '1px solid #f1f5f9', color: '#334155' },
  actionIconBtn: { background: 'none', border: 'none', cursor: 'pointer', fontSize: '11px', margin: '0 2px', padding: '3px 5px' },
  raciList: { display: 'flex', flexDirection: 'column', gap: '8px' },
  raciRow: { display: 'flex', alignItems: 'center', gap: '8px' },
  raciLabel: { width: '110px', fontSize: '10px', fontWeight: '600', color: '#475569', flexShrink: 0 },
  tagBox: { flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'space-between', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '3px 6px', backgroundColor: '#ffffff', minHeight: '28px' },
  tagsFlex: { display: 'flex', gap: '4px', flexWrap: 'wrap' },
  tag: { backgroundColor: '#f1f5f9', border: '1px solid #e2e8f0', color: '#334155', padding: '1px 6px', borderRadius: '4px', fontSize: '10px', display: 'inline-flex', alignItems: 'center', gap: '4px' },
  tagClose: { cursor: 'pointer', fontSize: '9px', color: '#94a3b8' },
  userAddIcon: { background: 'none', border: 'none', cursor: 'pointer', padding: '3px 5px' },
  clearAllBtn: { background: 'none', border: '1px solid #cbd5e1', borderRadius: '4px', padding: '2px 8px', fontSize: '10px', color: '#64748b', cursor: 'pointer' },
  workflowRow: { display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap', fontSize: '11px' },
  checkboxLabel: { display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#1e293b', fontWeight: '500', transition: 'color 0.15s ease' },
  footer: { display: 'flex', justifyContent: 'flex-end', gap: '8px', padding: '12px 20px', borderTop: '1px solid #e2e8f0', backgroundColor: '#ffffff' },
  cancelBtn: { padding: '6px 14px', borderRadius: '6px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', color: '#475569', fontSize: '11px', fontWeight: '600', cursor: 'pointer' },
  draftBtn: { padding: '6px 14px', borderRadius: '6px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', color: '#004d3d', fontSize: '11px', fontWeight: '600', cursor: 'pointer' },
  submitBtn: { padding: '6px 14px', borderRadius: '6px', border: 'none', backgroundColor: '#004d3d', color: '#ffffff', fontSize: '11px', fontWeight: '600', cursor: 'pointer' },
  activateBtn: { padding: '6px 14px', borderRadius: '6px', border: 'none', backgroundColor: '#004d3d', color: '#ffffff', fontSize: '11px', fontWeight: '600', cursor: 'pointer' },
};

export default ProcessManagementForm;