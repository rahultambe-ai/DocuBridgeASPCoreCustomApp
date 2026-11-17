import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse, photoUrlFor } from '../../spfx-shim';

export interface IAssignmentGroup {
  id: number;
  role: string;
  entraGroup: string;
  groupType: string;
  description: string;
  assignedUsers: string[];
}

export interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

interface AssignmentADGroupsProps {
  isOpen: boolean;
  onClose: () => void;
  /** Called after a successful SharePoint save. Parent should re-fetch the list
   *  from SharePoint so the table always reflects the actual library data. */
  onSave: () => void;
  editingGroup?: IAssignmentGroup | null;
  graphContext?: any;
  /** Process Code of the parent process — saved into the "Title" column of the
   *  "Process Assignment Groups" SharePoint list so groups can be looked up per process. */
  processCode?: string;
}

const ASSIGNMENT_GROUPS_LIST_TITLE = 'Process Assignment Groups';

// Shared audit trail list — every module logs its entries here, one row per action.
// https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';

// yyyyMMddHHmmss — exact moment an audit entry was written.
const generateAuditId = (): string => {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
};

const ROLE_OPTIONS = ['Contributor', 'Viewer'];

// Document library that holds each process's folder (e.g. "Shared Documents/PROC-M-009").
// Same library DocumentList.tsx / ProcessManagementForm.tsx use.
const DOCUMENT_LIBRARY_NAME = 'Shared Documents';

// Role picked in this modal -> SharePoint permission level applied on the process folder.
// Contributor -> "Edit" (manage access shows "Can edit"); Viewer -> "Read" (manage access shows "Can view").
const ROLE_TO_PERMISSION_LEVEL: { [key: string]: string } = {
  Contributor: 'Edit',
  Viewer: 'Read',
};

const FALLBACK_USERS: IGraphUser[] = [
  { id: 'u1', displayName: 'Aarav Mehta', mail: 'aarav.mehta@mkx8.onmicrosoft.com', userPrincipalName: 'aarav.mehta@mkx8.onmicrosoft.com' },
  { id: 'u2', displayName: 'Priya Nair', mail: 'priya.nair@mkx8.onmicrosoft.com', userPrincipalName: 'priya.nair@mkx8.onmicrosoft.com' },
  { id: 'u3', displayName: 'Rohan Kapoor', mail: 'rohan.kapoor@mkx8.onmicrosoft.com', userPrincipalName: 'rohan.kapoor@mkx8.onmicrosoft.com' },
];

const AssignmentADGroups: React.FC<AssignmentADGroupsProps> = ({ isOpen, onClose, onSave, editingGroup, graphContext, processCode }) => {
  const [role, setRole] = useState('');
  const [entraGroup, setEntraGroup] = useState('');
  const [groupType, setGroupType] = useState('');
  const [description, setDescription] = useState('');
  const [assignedUsers, setAssignedUsers] = useState<string[]>([]);

  const [isSaving, setIsSaving] = useState(false);
  const [saveError, setSaveError] = useState<string | null>(null);

  const [roleOpen, setRoleOpen] = useState(false);
  const [usersOpen, setUsersOpen] = useState(false);
  const [userSearch, setUserSearch] = useState('');
  const [msUsers, setMsUsers] = useState<IGraphUser[]>([]);
  const [isLoadingUsers, setIsLoadingUsers] = useState(false);

  const roleRef = useRef<HTMLDivElement>(null);
  const usersRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isOpen) return;
    if (editingGroup) {
      setRole(editingGroup.role);
      setEntraGroup(editingGroup.entraGroup);
      setGroupType(editingGroup.groupType);
      setDescription(editingGroup.description);
      setAssignedUsers(editingGroup.assignedUsers || []);
    } else {
      setRole('');
      setEntraGroup('');
      setGroupType('');
      setDescription('');
      setAssignedUsers([]);
    }
    setUserSearch('');
    setSaveError(null);
  }, [isOpen, editingGroup]);

  useEffect(() => {
    if (!isOpen) return;
    const fetchUsers = async () => {
      setIsLoadingUsers(true);
      try {
        if (!graphContext?.msGraphClientFactory) {
          setMsUsers(FALLBACK_USERS);
          return;
        }
        const client = await graphContext.msGraphClientFactory.getClient();
        const response = await client
          .api('/users')
          .select('id,displayName,mail,userPrincipalName')
          .top(50)
          .get();
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
        console.error('Graph API Error for Assigned Users:', err);
        setMsUsers(FALLBACK_USERS);
      } finally {
        setIsLoadingUsers(false);
      }
    };
    fetchUsers();
  }, [isOpen, graphContext]);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (roleRef.current && !roleRef.current.contains(e.target as Node)) setRoleOpen(false);
      if (usersRef.current && !usersRef.current.contains(e.target as Node)) setUsersOpen(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  if (!isOpen) return null;

  const filteredUsers = msUsers.filter(
    (u) =>
      u.displayName.toLowerCase().includes(userSearch.toLowerCase()) ||
      u.mail.toLowerCase().includes(userSearch.toLowerCase())
  );
  const isEdit = !!editingGroup;
  const isValid = role.trim() !== '';

  const toggleUser = (u: IGraphUser) => {
    setAssignedUsers((prev) =>
      prev.includes(u.displayName) ? prev.filter((v) => v !== u.displayName) : [...prev, u.displayName]
    );
  };

  // Grants the assigned users SharePoint access on this process's document folder
  // (Shared Documents/{processCode}) matching the Role picked above — Contributor -> Edit,
  // Viewer -> Read. Never blocks the Assignment Group save itself; failures are just logged,
  // same as the DMS Audit Logs write below.
  const applyFolderPermissionsForAssignment = async (
    processCodeForFolder: string,
    roleName: string,
    userDisplayNames: string[]
  ): Promise<void> => {
    if (!graphContext?.spHttpClient || !graphContext?.pageContext) return;
    const permissionLevel = ROLE_TO_PERMISSION_LEVEL[roleName];
    if (!permissionLevel) return;

    const emails = userDisplayNames
      .map((name) => msUsers.find((u) => u.displayName === name)?.mail)
      .filter((email): email is string => Boolean(email));
    if (emails.length === 0) return;

    const jsonHeaders = { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' };

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const webServerRelativeUrl: string = graphContext.pageContext.web.serverRelativeUrl || '';
      const libraryRelativeUrl = webServerRelativeUrl && webServerRelativeUrl !== '/'
        ? `${webServerRelativeUrl}/${DOCUMENT_LIBRARY_NAME}`
        : `/${DOCUMENT_LIBRARY_NAME}`;
      const folderServerRelativeUrl = `${libraryRelativeUrl}/${processCodeForFolder}`;
      const listEndpoint = `${siteUrl}/_api/web/GetList('${encodeURIComponent(libraryRelativeUrl)}')`;

      const itemLookupRes = await graphContext.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/ListItemAllFields`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!itemLookupRes.ok) {
        console.warn('Could not resolve process folder item for permissions:', await itemLookupRes.text());
        return;
      }
      const itemData = await itemLookupRes.json();
      const itemId = itemData.Id;

      // Library has "Require content approval" (Draft item security = approvers only), so the
      // folder stays Pending and invisible to everyone but its author/approvers. It is approved
      // here, at assignment time, so only now do the assigned users get to see it.
      try {
        const approveRes = await graphContext.spHttpClient.post(
          `${listEndpoint}/items(${itemId})/validateupdatelistitem`,
          SPHttpClient.configurations.v1,
          {
            headers: jsonHeaders,
            body: JSON.stringify({
              formValues: [{ FieldName: '_ModerationStatus', FieldValue: '0' }],
              bNewDocumentUpdate: false
            })
          }
        );
        if (!approveRes.ok) {
          console.warn('Could not approve process folder:', await approveRes.text());
        }
      } catch (approveErr) {
        console.error('Error approving process folder:', approveErr);
      }

      // copyRoleAssignments=true keeps whatever's already granted (e.g. other assignment
      // groups, RACI access) — this only makes the folder unique the first time it's called.
      const breakRes = await graphContext.spHttpClient.post(
        `${listEndpoint}/items(${itemId})/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!breakRes.ok) {
        console.warn('Could not break role inheritance on process folder:', await breakRes.text());
        return;
      }

      const roleDefRes = await graphContext.spHttpClient.get(
        `${siteUrl}/_api/web/roledefinitions/getbyname('${permissionLevel}')`,
        SPHttpClient.configurations.v1,
        { headers: jsonHeaders }
      );
      if (!roleDefRes.ok) {
        console.warn(`Could not resolve "${permissionLevel}" role definition:`, await roleDefRes.text());
        return;
      }
      const roleDefData = await roleDefRes.json();
      const roleDefId = roleDefData.Id;

      for (const email of emails) {
        try {
          const ensureRes = await graphContext.spHttpClient.post(
            `${siteUrl}/_api/web/ensureuser`,
            SPHttpClient.configurations.v1,
            { headers: jsonHeaders, body: JSON.stringify({ logonName: email }) }
          );
          if (!ensureRes.ok) {
            console.warn(`Could not resolve SharePoint user for ${email}:`, await ensureRes.text());
            continue;
          }
          const ensuredUser = await ensureRes.json();
          const principalId = ensuredUser.Id;

          const assignRes = await graphContext.spHttpClient.post(
            `${listEndpoint}/items(${itemId})/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`,
            SPHttpClient.configurations.v1,
            { headers: jsonHeaders }
          );
          if (!assignRes.ok) {
            console.warn(`Could not grant "${email}" ${permissionLevel} access to folder "${processCodeForFolder}":`, await assignRes.text());
          }
        } catch (userErr) {
          console.error(`Error granting folder access to ${email}:`, userErr);
        }
      }
    } catch (err) {
      console.error('Error applying assignment-group folder permissions:', err);
    }
  };

  const handleSave = async () => {
    if (!isValid || isSaving) return;

    if (!graphContext?.spHttpClient || !graphContext?.pageContext) {
      setSaveError('Could not connect to SharePoint. Check the console.');
      return;
    }
    if (!processCode || !processCode.trim()) {
      setSaveError('Process Code is missing — save General Information first.');
      return;
    }

    setIsSaving(true);
    setSaveError(null);

    try {
      const siteUrl = graphContext.pageContext.web.absoluteUrl;
      const isEditingGroup = Boolean(editingGroup);
      const endpoint = isEditingGroup
        ? `${siteUrl}/_api/web/lists/getByTitle('${ASSIGNMENT_GROUPS_LIST_TITLE}')/items(${editingGroup!.id})`
        : `${siteUrl}/_api/web/lists/getByTitle('${ASSIGNMENT_GROUPS_LIST_TITLE}')/items`;

      const bodyData: { [key: string]: any } = {
        Title: processCode,
        Role: role,
        EntraIDGroupName: entraGroup,
        GroupType: groupType,
        Description: description,
        AssignedUsers: assignedUsers.join('; '),
      };

      const headers: any = {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
      };
      if (isEditingGroup) {
        headers['IF-MATCH'] = '*';
        headers['X-HTTP-Method'] = 'PATCH';
      }

      const response: SPHttpClientResponse = await graphContext.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers, body: JSON.stringify(bodyData) }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error('SharePoint "Process Assignment Groups" save failed:', errText);
        setSaveError("It didn't save in SharePoint. Check the console.");
        return;
      }

      // Grant the assigned users SharePoint access on the process folder matching this Role.
      // Never blocks — the group itself already saved fine above.
      await applyFolderPermissionsForAssignment(processCode, role, assignedUsers);

      // One "DMS Audit Logs" entry per Assignment Group save (create or update).
      // Never blocks — the group itself already saved fine above.
      try {
        const auditEndpoint = `${siteUrl}/_api/web/lists/getByTitle('${AUDIT_LOG_LIST_TITLE}')/items`;
        const auditRes: SPHttpClientResponse = await graphContext.spHttpClient.post(
          auditEndpoint,
          SPHttpClient.configurations.v1,
          {
            headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
            body: JSON.stringify({
              Title: processCode,        // "Title" column — renamed to "ProcessID" in the list UI
              EntryType: 'Assignment Group',
              AuditID: generateAuditId(),
              AssignmentRole: role,
              EntraIDGroupName: entraGroup,
              GroupType: groupType,
              GroupDescription: description,
              AssignedUsers: assignedUsers.join('; ')
            })
          }
        );
        if (!auditRes.ok) {
          console.error('DMS Audit Logs entry failed:', await auditRes.text());
        }
      } catch (auditErr) {
        console.error('Error writing DMS Audit Logs entry:', auditErr);
      }

      onSave();
      onClose();
    } catch (err) {
      console.error('Error saving Assignment Group to SharePoint:', err);
      setSaveError('An error occurred while saving.');
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="aag-overlay" style={styles.overlay} onClick={onClose}>
      <style>{`
        @keyframes aagFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes aagScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        @keyframes aagSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .aag-overlay { animation: aagFadeIn 0.18s ease; }
        .aag-modal { animation: aagScaleIn 0.22s cubic-bezier(0.16, 1, 0.3, 1); }
        .aag-dropdown { animation: aagSlideDown 0.15s ease; }
        .aag-select:hover { border-color: #94a3b8; }
        .aag-option:hover { background-color: #f1f5f9; }
        .aag-btn { transition: transform 0.15s ease, filter 0.15s ease, background-color 0.15s ease, box-shadow 0.15s ease; }
        .aag-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
        .aag-btn:active { transform: translateY(0); filter: brightness(0.97); }
        .aag-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; filter: none; }
        .aag-textarea, .aag-search-input { transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .aag-textarea:focus { border-color: #004d3d; box-shadow: 0 0 0 3px rgba(0, 77, 61, 0.10); }
        .aag-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; }
        .aag-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
        .aag-modal::-webkit-scrollbar { width: 8px; }
        .aag-modal::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 8px; }

        @media (max-width: 680px) {
          .aag-grid2 { grid-template-columns: 1fr !important; }
          .aag-preview-grid { grid-template-columns: repeat(2, 1fr) !important; }
          .aag-footer { flex-direction: column-reverse !important; }
          .aag-footer button { width: 100%; }
        }
        @media (max-width: 420px) {
          .aag-preview-grid { grid-template-columns: 1fr !important; }
        }
      `}</style>

      <div className="aag-modal" style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
        <div style={styles.header}>
          <div>
            <h2 style={styles.title}>{isEdit ? 'Edit Assignment Group/User' : 'Add Assignment Group/User'}</h2>
            <p style={styles.subtitle}>{isEdit ? 'Update this Entra ID Group for this process' : 'Add a new Entra ID Group to this process'}</p>
          </div>
          <button className="aag-close-btn" style={styles.closeBtn} onClick={onClose} title="Close">✕</button>
        </div>

        <div style={styles.body}>
          <div className="aag-grid2" style={styles.grid2}>
            {/* Role */}
            <div style={{ ...styles.field, position: 'relative' }} ref={roleRef}>
              <label style={styles.label}>Role <span style={styles.req}>*</span></label>
              <div
                className="aag-select"
                style={{ ...styles.selectBox, borderColor: roleOpen ? '#004d3d' : '#cbd5e1', boxShadow: roleOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none' }}
                onClick={() => { setRoleOpen((p) => !p); setUsersOpen(false); }}
              >
                <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <ShieldIcon />
                  <span style={{ fontSize: '12px', color: role ? '#1e293b' : '#94a3b8', fontWeight: 500 }}>{role || 'Select Role'}</span>
                </span>
                <ChevronIcon open={roleOpen} />
              </div>
              {roleOpen && (
                <div className="aag-dropdown" style={styles.dropdown}>
                  {ROLE_OPTIONS.map((opt) => (
                    <div
                      key={opt}
                      className="aag-option"
                      style={{ ...styles.option, backgroundColor: role === opt ? '#e6f4ea' : 'transparent', color: role === opt ? '#004d3d' : '#334155', fontWeight: role === opt ? 600 : 400 }}
                      onClick={() => { setRole(opt); setRoleOpen(false); }}
                    >
                      <span>{opt}</span>
                      {role === opt && <span style={{ color: '#004d3d', fontWeight: 700 }}>✓</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Assigned Users */}
            <div style={{ ...styles.field, position: 'relative' }} ref={usersRef}>
              <label style={styles.label}>Assigned Users</label>
              <div
                className="aag-select"
                style={{ ...styles.selectBox, borderColor: usersOpen ? '#004d3d' : '#cbd5e1', boxShadow: usersOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none' }}
                onClick={() => { setUsersOpen((p) => !p); setRoleOpen(false); }}
              >
                <span style={{ display: 'flex', alignItems: 'center', gap: '6px', overflow: 'hidden', minWidth: 0 }}>
                  <span style={styles.avatarCircle}>
                    {assignedUsers.length === 1 && msUsers.find((u) => u.displayName === assignedUsers[0])?.photoUrl ? (
                      <img src={msUsers.find((u) => u.displayName === assignedUsers[0])?.photoUrl} alt="" style={styles.avatarImg} />
                    ) : (
                      <UsersIcon />
                    )}
                  </span>
                  <span style={{ fontSize: '12px', color: assignedUsers.length ? '#1e293b' : '#94a3b8', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {assignedUsers.length === 0
                      ? 'Select Users'
                      : assignedUsers.length === 1
                      ? assignedUsers[0]
                      : `${assignedUsers[0]} +${assignedUsers.length - 1} more`}
                  </span>
                </span>
                <ChevronIcon open={usersOpen} />
              </div>
              {usersOpen && (
                <div className="aag-dropdown" style={{ ...styles.dropdown, padding: '4px' }}>
                  <div style={styles.searchWrap}>
                    <SearchIcon muted />
                    <input
                      autoFocus
                      type="text"
                      className="aag-search-input"
                      placeholder="Search name or email..."
                      value={userSearch}
                      onChange={(e) => setUserSearch(e.target.value)}
                      style={styles.searchInput}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </div>
                  <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
                    {isLoadingUsers ? (
                      <div style={styles.emptyState}>Loading users...</div>
                    ) : filteredUsers.length === 0 ? (
                      <div style={styles.emptyState}>No users found</div>
                    ) : (
                      filteredUsers.map((u) => {
                        const isSelected = assignedUsers.includes(u.displayName);
                        return (
                          <div
                            key={u.id}
                            className="aag-option"
                            style={{ ...styles.option, backgroundColor: isSelected ? '#e6f4ea' : 'transparent' }}
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
                                <span style={{ fontSize: '12px', fontWeight: isSelected ? 600 : 500, color: isSelected ? '#004d3d' : '#1e293b', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.displayName}</span>
                                <span style={{ fontSize: '10px', color: '#94a3b8', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{u.mail}</span>
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
          </div>

          {/* Description */}
          <div style={styles.field}>
            <label style={styles.label}>Description</label>
            <textarea
              className="aag-textarea"
              style={styles.textarea}
              placeholder="Enter description for this group (optional)"
              value={description}
              maxLength={250}
              onChange={(e) => setDescription(e.target.value)}
            />
            <span style={styles.charCount}>{description.length}/250 characters</span>
          </div>

          {/* Preview */}
          <div style={styles.previewBox}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <EyeIcon />
              <span style={styles.previewTitle}>Group/User Information Preview</span>
            </div>
            <p style={styles.previewSubtitle}>Review the group details before adding</p>

            <div className="aag-preview-grid" style={styles.previewGrid}>
              <PreviewItem icon={<ShieldIcon />} label="Role" value={role} />
              <PreviewItem icon={<UsersIcon />} label="Assigned Users" value={assignedUsers.join(', ')} />
              <PreviewItem icon={<DocIcon />} label="Description" value={description} />
            </div>
          </div>
        </div>

        <div className="aag-footer" style={styles.footer}>
          {saveError && <span style={{ fontSize: '10px', color: '#dc2626', fontWeight: 600, marginRight: 'auto', alignSelf: 'center' }}>{saveError}</span>}
          <button className="aag-btn" style={styles.cancelBtn} onClick={onClose} disabled={isSaving}>Cancel</button>
          <button className="aag-btn" style={{ ...styles.saveBtn, opacity: (!isValid || isSaving) ? 0.6 : 1 }} disabled={!isValid || isSaving} onClick={handleSave}>
            <PlusIcon /> {isSaving ? 'Saving...' : isEdit ? 'Update Group/User' : 'Add Group/User'}
          </button>
        </div>
      </div>
    </div>
  );
};

const PreviewItem: React.FC<{ icon: React.ReactNode; label: string; value: string }> = ({ icon, label, value }) => (
  <div style={styles.previewItem}>
    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
      {icon}
      <span style={styles.previewLabel}>{label}</span>
    </div>
    <span style={{ ...styles.previewValue, color: value ? '#1e293b' : '#94a3b8' }}>{value || '-'}</span>
  </div>
);

const ShieldIcon: React.FC = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path></svg>
);
const UsersIcon: React.FC = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
);
const DocIcon: React.FC = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
);
const EyeIcon: React.FC = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
);
const PlusIcon: React.FC = () => (
  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#ffffff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '4px', verticalAlign: '-1px' }}><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
);
const SearchIcon: React.FC<{ muted?: boolean }> = ({ muted }) => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={muted ? '#94a3b8' : '#64748b'} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ flexShrink: 0 }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
);
const ChevronIcon: React.FC<{ open: boolean }> = ({ open }) => (
  <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: open ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
    <polyline points="6 9 12 15 18 9"></polyline>
  </svg>
);

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 10000, padding: '24px 16px', boxSizing: 'border-box' },
  modalCard: { backgroundColor: '#ffffff', borderRadius: '16px', width: '100%', maxWidth: '680px', maxHeight: '90vh', display: 'flex', flexDirection: 'column', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)', overflow: 'auto', border: '1px solid #e2e8f0', margin: 'auto' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', padding: '18px 22px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#ffffff', position: 'sticky', top: 0, zIndex: 2 },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '4px 0 0', fontSize: '11px', color: '#64748b' },
  closeBtn: { background: 'none', border: 'none', fontSize: '15px', cursor: 'pointer', color: '#64748b', padding: '4px 8px', borderRadius: '6px', flexShrink: 0 },
  body: { padding: '20px 22px', display: 'flex', flexDirection: 'column', gap: '16px' },
  grid2: { display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '16px' },
  field: { display: 'flex', flexDirection: 'column', gap: '5px' },
  label: { fontSize: '11px', fontWeight: 600, color: '#334155' },
  req: { color: '#dc2626' },
  selectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '9px 12px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', boxSizing: 'border-box', cursor: 'pointer', transition: 'border-color 0.15s ease, box-shadow 0.15s ease' },
  dropdown: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 8px 20px rgba(0, 0, 0, 0.10)', zIndex: 500, padding: '4px', maxHeight: '230px', overflowY: 'auto' },
  option: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 9px', fontSize: '12px', borderRadius: '5px', cursor: 'pointer', margin: '1px 0' },
  searchWrap: { display: 'flex', alignItems: 'center', gap: '6px', padding: '7px 9px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', marginBottom: '4px' },
  searchInput: { border: 'none', outline: 'none', fontSize: '12px', width: '100%', background: 'transparent', color: '#334155' },
  emptyState: { fontSize: '11px', color: '#94a3b8', padding: '12px', textAlign: 'center' },
  avatarCircle: { width: '18px', height: '18px', borderRadius: '50%', backgroundColor: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', flexShrink: 0 },
  avatarImg: { width: '100%', height: '100%', objectFit: 'cover' },
  textarea: { padding: '9px 12px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '12px', outline: 'none', height: '60px', resize: 'vertical', fontFamily: 'inherit', color: '#1e293b', boxSizing: 'border-box' },
  charCount: { fontSize: '10px', color: '#94a3b8', alignSelf: 'flex-end' },
  previewBox: { backgroundColor: '#f0faf5', border: '1px solid #cdeadb', borderRadius: '10px', padding: '14px 16px' },
  previewTitle: { fontSize: '12px', fontWeight: 700, color: '#0f172a' },
  previewSubtitle: { margin: '2px 0 12px 19px', fontSize: '10px', color: '#64748b' },
  previewGrid: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '12px' },
  previewItem: { display: 'flex', flexDirection: 'column', gap: '6px' },
  previewLabel: { fontSize: '10px', fontWeight: 600, color: '#475569' },
  previewValue: { fontSize: '12px', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  footer: { display: 'flex', justifyContent: 'flex-end', gap: '8px', padding: '14px 22px', borderTop: '1px solid #e2e8f0', backgroundColor: '#ffffff', position: 'sticky', bottom: 0 },
  cancelBtn: { padding: '8px 16px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', color: '#475569', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
  saveBtn: { display: 'flex', alignItems: 'center', padding: '8px 16px', borderRadius: '8px', border: 'none', backgroundColor: '#004d3d', color: '#ffffff', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
};

export default AssignmentADGroups;
