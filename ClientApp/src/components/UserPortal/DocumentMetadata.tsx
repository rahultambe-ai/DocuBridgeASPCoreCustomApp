import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { IDocumentItem } from './DocumentList';
import { SPHttpClient, SPHttpClientResponse, photoUrlFor } from '../../spfx-shim';
import { DocumentTypeField } from './DocumentTypeDropdownlist';

export interface IUserMetadataProps {
  document: IDocumentItem;
  currentUser: string;
  context: any;
  onBack: () => void;
  onSave?: (updatedFields: any) => void;
}

export interface IMicrosoftUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

export interface ILiveVersionEntry {
  id: string;
  versionLabel: string;
  modified: string;
  modifiedBy: string;
  size: string;
  isCurrent: boolean;
}

const SVGIcon = ({ name, color = "currentColor", size = 14 }: { name: string; color?: string; size?: number }) => {
  const icons: { [key: string]: React.ReactNode } = {
    arrowLeft: <path d="M19 12H5M12 19l-7-7 7-7" />,
    check: <path d="M20 6L9 17l-5-5" />,
    chevronDown: <path d="M6 9l6 6 6-6" />,
    search: <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'inline-block', verticalAlign: 'middle' }}>
      {icons[name] || null}
    </svg>
  );
};

interface CustomDropdownProps {
  value: string;
  options: string[];
  onChange: (val: string) => void;
}

const CustomDropdown: React.FC<CustomDropdownProps> = ({ value, options, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div style={customDDStyles.wrapper} ref={dropdownRef}>
      <div 
        onClick={() => setIsOpen(!isOpen)} 
        style={{
          ...customDDStyles.control,
          borderColor: isOpen ? '#005a43' : '#cbd5e1',
          boxShadow: isOpen ? '0 0 0 2px rgba(0, 90, 67, 0.15)' : 'none'
        }}
      >
        <span style={customDDStyles.valueText}>{value}</span>
        <span style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', display: 'flex' }}>
          <SVGIcon name="chevronDown" color="#64748b" size={14} />
        </span>
      </div>

      {isOpen && (
        <div style={customDDStyles.menu}>
          {options.map((opt, idx) => {
            const isSelected = opt === value;
            return (
              <div
                key={idx}
                onClick={() => {
                  onChange(opt);
                  setIsOpen(false);
                }}
                style={{
                  ...customDDStyles.option,
                  backgroundColor: isSelected ? '#f0fdf4' : '#ffffff',
                  color: isSelected ? '#005a43' : '#334155',
                  fontWeight: isSelected ? 600 : 500
                }}
              >
                <span>{opt}</span>
                {isSelected && <SVGIcon name="check" color="#005a43" size={14} />}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

interface OwnerDropdownProps {
  owner: string;
  onChange: (ownerName: string) => void;
  context: any;
}

const OwnerDropdown: React.FC<OwnerDropdownProps> = ({ owner, onChange, context }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [msUsers, setMsUsers] = useState<IMicrosoftUser[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const fetchUsers = async () => {
      try {
        if (!context?.msGraphClientFactory) return;
        const client = await context.msGraphClientFactory.getClient();
        const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(50).get();
        const siteUrl = context.pageContext?.web?.absoluteUrl || '';

        const usersList: IMicrosoftUser[] = response.value.map((u: any) => {
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
        console.error("Graph API Error for Owner:", err);
      }
    };
    fetchUsers();

    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [context]);

  const filteredUsers = msUsers.filter(u =>
    u.displayName.toLowerCase().includes(searchQuery.toLowerCase()) ||
    u.mail.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const selectedUserObj = msUsers.find(u => u.displayName === owner);

  return (
    <div style={customDDStyles.wrapper} ref={dropdownRef}>
      <div 
        onClick={() => setIsOpen(!isOpen)} 
        style={{
          ...customDDStyles.control,
          borderColor: isOpen ? '#005a43' : '#cbd5e1',
          boxShadow: isOpen ? '0 0 0 2px rgba(0, 90, 67, 0.15)' : 'none'
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
          <div style={customDDStyles.avatarCircle}>
            {selectedUserObj?.photoUrl ? (
              <img src={selectedUserObj.photoUrl} alt="" style={customDDStyles.avatarImg} />
            ) : (
              <span style={{ fontSize: '10px', color: '#64748b' }}>👤</span>
            )}
          </div>
          <span style={customDDStyles.valueText}>{owner || 'Select Owner'}</span>
        </div>
        <span style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', display: 'flex', flexShrink: 0 }}>
          <SVGIcon name="chevronDown" color="#64748b" size={14} />
        </span>
      </div>

      {isOpen && (
        <div style={{ ...customDDStyles.menu, padding: '0' }}>
          <div style={customDDStyles.searchContainer}>
            <SVGIcon name="search" color="#94a3b8" size={13} />
            <input
              type="text"
              placeholder="Search user..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              style={customDDStyles.searchInput}
              autoFocus
            />
          </div>
          <div style={{ maxHeight: '180px', overflowY: 'auto', padding: '4px' }}>
            {filteredUsers.length === 0 ? (
              <div style={{ padding: '8px', fontSize: '11px', color: '#64748b', textAlign: 'center' }}>No users found</div>
            ) : (
              filteredUsers.map((user) => {
                const isSelected = user.displayName === owner;
                return (
                  <div
                    key={user.id}
                    onClick={() => {
                      onChange(user.displayName);
                      setIsOpen(false);
                      setSearchQuery('');
                    }}
                    style={{
                      ...customDDStyles.option,
                      backgroundColor: isSelected ? '#f0fdf4' : '#ffffff',
                      color: isSelected ? '#005a43' : '#334155'
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
                      <div style={customDDStyles.avatarCircle}>
                        {user.photoUrl ? (
                          <img src={user.photoUrl} alt="" style={customDDStyles.avatarImg} />
                        ) : (
                          <span style={{ fontSize: '10px', color: '#64748b' }}>👤</span>
                        )}
                      </div>
                      <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                        <span style={{ fontSize: '12px', fontWeight: isSelected ? 600 : 500, color: isSelected ? '#005a43' : '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{user.displayName}</span>
                        <span style={{ fontSize: '10px', color: '#64748b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{user.mail}</span>
                      </div>
                    </div>
                    {isSelected && <SVGIcon name="check" color="#005a43" size={14} />}
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

export const UserMetadata: React.FC<IUserMetadataProps> = ({ document, currentUser, context, onBack, onSave }) => {
  const [activeTab, setActiveTab] = useState<'History' | 'Activity'>('History');
  const [documentType, setDocumentType] = useState<string>(document.docType || 'SOP');
  const [documentStatus, setDocumentStatus] = useState<string>(document.status || 'Active');
  
  const [owner, setOwner] = useState<string>(document.owner || currentUser);
  const [lastModifiedBy, setLastModifiedBy] = useState<string>(document.modifiedBy || currentUser);

  const [currentVersion, setCurrentVersion] = useState<string>(document.version || 'v1.0');
  const [isSaving, setIsSaving] = useState<boolean>(false);

  const [versionHistory, setVersionHistory] = useState<ILiveVersionEntry[]>([]);
  const [loadingVersions, setLoadingVersions] = useState<boolean>(false);

  const docTags = (document as any).tags as string[] | undefined;
  const [tags, setTags] = useState<string[]>(docTags && docTags.length > 0 ? docTags : []);
  const [newTagInput, setNewTagInput] = useState<string>('');
  const [isAddingTag, setIsAddingTag] = useState<boolean>(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const docStatusOptions = ['Active','Pending Approval','Draft','In Review','Archived'];

  const docAny = document as any;
  const fileName = docAny.name || docAny.title || 'Document.pdf';
  const fileExtension = fileName.includes('.') ? `.${fileName.split('.').pop()}` : '.pdf';
  
  const formatFileSize = (size: any) => {
    if (!size) return '0 KB';
    const num = Number(size);
    if (isNaN(num)) return size;
    if (num < 1024 * 1024) {
      return `${(num / 1024).toFixed(0)} KB`;
    }
    return `${(num / (1024 * 1024)).toFixed(2)} MB`;
  };
  const fileSize = formatFileSize(docAny.size);

  // Live-fetches SharePoint's own built-in version info for this file — the current
  // version label (SP.File UIVersionLabel) plus the library's version history
  // (SP.FileVersion collection). Nothing here is written/maintained by our code;
  // it's exactly what the SharePoint document library is tracking automatically,
  // so it always reflects the latest version right after any change to the file.
  const fetchLiveVersionInfo = async (): Promise<string> => {
    if (!context?.spHttpClient) return currentVersion;
    const serverRelativeUrl = docAny.serverRelativeUrl;
    if (!serverRelativeUrl) return currentVersion;

    setLoadingVersions(true);
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const safeRelUrl = String(serverRelativeUrl).replace(/'/g, "''");

      // 1. Current live version of the file as it stands right now in the library.
      const fileEndpoint = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(safeRelUrl)}')?$select=UIVersionLabel,TimeLastModified,Length&$expand=ModifiedBy`;
      const fileRes: SPHttpClientResponse = await context.spHttpClient.get(
        fileEndpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      let liveLabel = currentVersion;
      const entries: ILiveVersionEntry[] = [];

      if (fileRes.ok) {
        const fileData = await fileRes.json();
        liveLabel = fileData.UIVersionLabel ? `v${fileData.UIVersionLabel}` : currentVersion;
        entries.push({
          id: 'current',
          versionLabel: liveLabel,
          modified: fileData.TimeLastModified ? new Date(fileData.TimeLastModified).toLocaleString() : '',
          modifiedBy: fileData.ModifiedBy?.Title || lastModifiedBy,
          size: formatFileSize(fileData.Length),
          isCurrent: true
        });
        setCurrentVersion(liveLabel);
      } else {
        console.error("Failed to fetch live file version:", await fileRes.text());
      }

      // 2. Older versions kept by the library's version history.
      const versionsEndpoint = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(safeRelUrl)}')/versions?$select=VersionLabel,Created,Size,CreatedBy/Title&$expand=CreatedBy&$orderby=Created desc`;
      const verRes: SPHttpClientResponse = await context.spHttpClient.get(
        versionsEndpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (verRes.ok) {
        const verData = await verRes.json();
        const olderEntries: ILiveVersionEntry[] = (verData.value || [])
          .map((raw: any, idx: number) => ({
            id: `v-${idx}`,
            versionLabel: `v${raw.VersionLabel}`,
            modified: raw.Created ? new Date(raw.Created).toLocaleString() : '',
            modifiedBy: raw.CreatedBy?.Title || '',
            size: formatFileSize(raw.Size),
            isCurrent: false
          }))
          .filter((entry: ILiveVersionEntry) => entry.versionLabel !== liveLabel);
        entries.push(...olderEntries);
      } else {
        console.error("Failed to fetch live version history:", await verRes.text());
      }

      setVersionHistory(entries);
      return liveLabel;
    } catch (err) {
      console.error("Error fetching live SharePoint version info:", err);
      return currentVersion;
    } finally {
      setLoadingVersions(false);
    }
  };

  useEffect(() => {
    fetchLiveVersionInfo();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [docAny.serverRelativeUrl]);

  const handleAddTag = () => {
    if (newTagInput.trim() && !tags.includes(newTagInput.trim())) {
      setTags([...tags, newTagInput.trim()]);
      setNewTagInput('');
      setIsAddingTag(false);
    }
  };

  const handleRemoveTag = (tagToRemove: string) => {
    setTags(tags.filter(t => t !== tagToRemove));
  };

  // When the user picks a Document Type, auto-create a tag with that same
  // value. If a type was already selected before, its auto tag is swapped
  // out for the new one so tags don't pile up on every re-selection; any
  // other tag the user added manually is left untouched.
  const handleDocumentTypeChange = (value: string) => {
    setDocumentType(value);
    setTags((prevTags) => {
      const withoutOldTypeTag = documentType ? prevTags.filter((t) => t !== documentType) : prevTags;
      if (!value || withoutOldTypeTag.includes(value)) return withoutOldTypeTag;
      return [...withoutOldTypeTag, value];
    });
  };

  // The `id` passed down from the parent component isn't always the real SharePoint
  // list item ID (e.g. it can be a row/index placeholder). To avoid writing to the
  // wrong item (or a non-existent one), resolve the true item ID by matching the
  // file name (FileLeafRef) directly against the library before saving.
  const resolveItemId = async (libraryName: string): Promise<number | null> => {
    try {
      if (!context?.spHttpClient) return null;
      const siteUrl = context.pageContext.web.absoluteUrl;
      const safeFileName = fileName.replace(/'/g, "''");
      const filterUrl = `${siteUrl}/_api/web/lists/getByTitle('${libraryName}')/items?$filter=FileLeafRef eq '${encodeURIComponent(safeFileName)}'&$select=Id&$top=1`;

      const res: SPHttpClientResponse = await context.spHttpClient.get(
        filterUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!res.ok) {
        console.error("Failed to resolve item ID by file name:", await res.text());
        return null;
      }

      const data = await res.json();
      if (data?.value?.length > 0) {
        return data.value[0].Id;
      }
      return null;
    } catch (err) {
      console.error("Error resolving item ID by file name:", err);
      return null;
    }
  };

  // Shared audit trail list — every module logs its entries here, one row per action.
  // https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
  const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';

  // yyyyMMddHHmmss — exact moment an audit entry was written.
  const generateAuditId = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  };

  // Writes ONE row into "DMS Audit Logs" for this metadata save. Never blocks the
  // caller — any failure here is only logged to console, the real save already succeeded.
  const addDmsAuditLogEntry = async (): Promise<void> => {
    if (!context?.spHttpClient) return;
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${AUDIT_LOG_LIST_TITLE}')/items`;
      const res: SPHttpClientResponse = await context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify({
            Title: '',                 // "Title" column — renamed to "ProcessID" in the list UI (not available on this screen)
            EntryType: 'Document',
            AuditID: generateAuditId(),
            DocumentID: docAny.docNumber || '',
            DocumentName: fileName,
            FileName: fileName,
            DocumentType: documentType,
            DocumentStatus: documentStatus,
            DocumentOwner: owner,
            Tags: tags.join(', ')
          })
        }
      );
      if (!res.ok) {
        console.error('DMS Audit Logs entry failed:', await res.text());
      }
    } catch (err) {
      console.error('Error writing DMS Audit Logs entry:', err);
    }
  };

  const updateSharePointListItem = async (): Promise<boolean> => {
    try {
      if (!context?.spHttpClient) {
        console.error("SharePoint HttpClient unavailable.");
        return false;
      }

      const siteUrl = context.pageContext.web.absoluteUrl;
      const libraryName = "Documents";

      // The parent dashboard now fetches the real SharePoint list item ID
      // (via ListItemAllFields), so it's the primary source of truth here.
      // Only fall back to resolving by file name if it's ever missing.
      const suppliedId = docAny.id ?? docAny.Id ?? docAny.ID ?? docAny.ItemId ?? docAny.itemId;
      const itemId = (suppliedId !== undefined && suppliedId !== null)
        ? suppliedId
        : await resolveItemId(libraryName);

      if (itemId === undefined || itemId === null) {
        console.error("Could not determine a valid SharePoint item ID for this document.", { fileName, suppliedId });
        return false;
      }

      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${libraryName}')/items(${itemId})`;

      // Exact Mapped Internal Field Names provided by user.
      // 'FileStatus' (Document ID) is only included the FIRST time this document is
      // saved (when it has no value yet in SharePoint). Once it has a real value we
      // never re-send it, since doing so risks colliding with SharePoint's unique-value
      // validation on that column and failing the ENTIRE item update.
      const bodyData: { [key: string]: any } = {
        'DocumentType': documentType,     // Internal Name for Document Type
        'DocumentStatus': documentStatus, // Internal Name for Document Status
        'DocumentTags': tags.join(', ')   // Internal Name for Document Tags
        // Version is no longer written here — SharePoint's own library versioning
        // auto-increments it whenever this item/file is updated, and we live-fetch
        // that value instead (see fetchLiveVersionInfo).
      };

      const docNumberIsSaved: boolean = (document as any).docNumberIsSaved === true;
      if (!docNumberIsSaved && docAny.docNumber) {
        bodyData['FileStatus'] = docAny.docNumber; // Internal Name for Document ID
      }

      const response: SPHttpClientResponse = await context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(bodyData)
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.error("SharePoint Save Failed:", errText);
        return false;
      }

      return true;
    } catch (err) {
      console.error("Error updating SharePoint Item:", err);
      return false;
    }
  };

  const handleSave = async () => {
    setIsSaving(true);

    // Save metadata directly to SharePoint Library Item
    const saved = await updateSharePointListItem();

    setIsSaving(false);

    if (!saved) {
      setToastMessage("Save failed. Please check the console for details.");
      setTimeout(() => setToastMessage(null), 4000);
      return;
    }

    setLastModifiedBy(currentUser);

    // One "DMS Audit Logs" entry per metadata save.
    await addDmsAuditLogEntry();

    // SharePoint auto-increments the document's own version when the item is
    // updated — live-fetch it (and the refreshed history) instead of computing
    // a version number ourselves.
    const freshVersion = await fetchLiveVersionInfo();

    const updatedData = {
      documentType,
      documentStatus,
      owner,
      tags,
      version: freshVersion,
      lastModifiedBy: currentUser
    };

    if (onSave) {
      onSave(updatedData);
    }

    setToastMessage("Data successfully saved to SharePoint!");
    setTimeout(() => setToastMessage(null), 3000);
  };

  return (
    <div style={styles.container}>
      {toastMessage && (
        <div style={styles.toast}>
          <SVGIcon name="check" color="#fff" size={16} />
          <span>{toastMessage}</span>
        </div>
      )}

      <div style={styles.header}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <button onClick={onBack} style={styles.backButton}>
            <SVGIcon name="arrowLeft" color="#334155" size={18} />
          </button>
          <div>
            <h1 style={styles.title}>Document Details</h1>
            <p style={styles.subtitle}>View and manage document metadata and approval workflow for <strong>{fileName}</strong>.</p>
          </div>
        </div>
      </div>

      <div style={styles.mainLayout}>
        <div style={styles.leftSection}>
          <div style={styles.card}>
            <h2 style={styles.sectionHeading}>Document Information</h2>
            
            <div style={styles.gridRow}>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Document ID</label>
                <input type="text" readOnly value={document.docNumber || 'PUR_D003'} style={styles.readOnlyInput} />
              </div>
              <div style={styles.fieldGroup}>
                <DocumentTypeField
                  value={documentType}
                  onChange={handleDocumentTypeChange}
                  context={context}
                />
              </div>
            </div>

            <div style={styles.gridRow}>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Created On</label>
                <input type="text" readOnly value={document.modified || new Date().toLocaleDateString()} style={styles.readOnlyInput} />
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Created By</label>
                <input type="text" readOnly value={document.modifiedBy || currentUser} style={styles.readOnlyInput} />
              </div>
            </div>

            <div style={styles.gridRow}>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Last Modified On</label>
                <input type="text" readOnly value={new Date().toLocaleDateString()} style={styles.readOnlyInput} />
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Last Modified By</label>
                <input type="text" readOnly value={lastModifiedBy} style={{ ...styles.readOnlyInput, fontWeight: 700, color: '#005a43' }} />
              </div>
            </div>

            <div style={styles.gridRow}>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Document Status</label>
                <CustomDropdown
                  value={documentStatus}
                  options={docStatusOptions}
                  onChange={(val) => setDocumentStatus(val)}
                />
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Owner (Selected User)</label>
                <OwnerDropdown
                  owner={owner}
                  onChange={(newOwner) => setOwner(newOwner)}
                  context={context}
                />
              </div>
            </div>

            <div style={{ marginTop: '14px' }}>
              <label style={styles.label}>Tags</label>
              <div style={styles.tagsContainer}>
                {tags.map((tag, idx) => (
                  <span key={idx} style={styles.tagPill}>
                    {tag}
                    <button onClick={() => handleRemoveTag(tag)} style={styles.tagRemoveBtn}>×</button>
                  </span>
                ))}
                {isAddingTag ? (
                  <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                    <input
                      type="text"
                      placeholder="Enter tag..."
                      value={newTagInput}
                      onChange={(e) => setNewTagInput(e.target.value)}
                      style={styles.newTagInput}
                      autoFocus
                    />
                    <button onClick={handleAddTag} style={styles.addTagSubmitBtn}>Add</button>
                    <button onClick={() => setIsAddingTag(false)} style={styles.tagRemoveBtn}>Cancel</button>
                  </div>
                ) : (
                  <button onClick={() => setIsAddingTag(true)} style={styles.addTagBtn}>+ Add Tag</button>
                )}
              </div>
            </div>
          </div>

          <div style={{ ...styles.card, marginTop: '16px' }}>
            <h2 style={styles.sectionHeading}>System Fields (Auto Generated)</h2>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '16px' }}>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>File Name</label>
                <span style={{ fontSize: '12px', color: '#334155', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={fileName}>{fileName}</span>
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>File Size</label>
                <span style={{ fontSize: '12px', color: '#334155', fontWeight: 600 }}>{fileSize}</span>
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>File Type</label>
                <span style={{ fontSize: '12px', color: '#334155', fontWeight: 600 }}>{fileExtension}</span>
              </div>
              <div style={styles.fieldGroup}>
                <label style={styles.label}>Version</label>
                <span style={{ fontSize: '12px', color: '#005a43', fontWeight: 700 }}>{currentVersion}</span>
              </div>
            </div>
          </div>

          <div style={styles.footerActions}>
            <button onClick={handleSave} disabled={isSaving} style={{ ...styles.saveBtn, opacity: isSaving ? 0.7 : 1 }}>
              <SVGIcon name="check" color="#fff" size={14} />
              {isSaving ? "Saving..." : "Save & Update Version"}
            </button>
            <button onClick={onBack} style={styles.cancelBtn}>Cancel</button>
          </div>
        </div>

        <div style={styles.rightSection}>
          <div style={styles.card}>
            <div style={styles.tabsRow}>
              <button
                onClick={() => setActiveTab('History')}
                style={{ ...styles.tabBtn, borderBottom: activeTab === 'History' ? '2px solid #005a43' : '2px solid transparent', color: activeTab === 'History' ? '#005a43' : '#64748b' }}
              >
                History
              </button>
              <button
                onClick={() => setActiveTab('Activity')}
                style={{ ...styles.tabBtn, borderBottom: activeTab === 'Activity' ? '2px solid #005a43' : '2px solid transparent', color: activeTab === 'Activity' ? '#005a43' : '#64748b' }}
              >
                Activity
              </button>
            </div>

            {activeTab === 'History' ? (
              <div style={{ marginTop: '14px' }}>
                <h3 style={styles.sidebarSubHeading}>Approval History</h3>
                <div style={styles.historyItem}>
                  <div style={{ ...styles.historyDot, backgroundColor: '#15803d' }}>✓</div>
                  <div>
                    <span style={styles.historyText}>Assigned/Reviewed by <strong>{owner}</strong></span>
                    <span style={styles.historyDate}>Modified on {document.modified || new Date().toLocaleDateString()}</span>
                  </div>
                </div>

                <h3 style={{ ...styles.sidebarSubHeading, marginTop: '20px' }}>Version History</h3>
                {loadingVersions ? (
                  <div style={{ fontSize: '11px', color: '#64748b' }}>Fetching live version history from SharePoint…</div>
                ) : versionHistory.length === 0 ? (
                  <div style={{ fontSize: '11px', color: '#64748b' }}>No version history found for this document.</div>
                ) : (
                  versionHistory.map((v) => (
                    <div key={v.id} style={styles.historyItem}>
                      <div style={v.isCurrent ? { ...styles.versionCircle, backgroundColor: '#005a43' } : styles.versionCircle}></div>
                      <div>
                        <span style={styles.historyText}>
                          <strong>{v.versionLabel}</strong>{v.isCurrent ? ' (Current)' : ''} — {v.size}{v.modifiedBy ? ` by ${v.modifiedBy}` : ''}
                        </span>
                        <span style={styles.historyDate}>{v.modified}</span>
                      </div>
                    </div>
                  ))
                )}
              </div>
            ) : (
              <div style={{ marginTop: '14px', fontSize: '12px', color: '#475569' }}>
                <p>• Document routed through approval pipeline.</p>
                <p>• Last saved by {lastModifiedBy}.</p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const customDDStyles: { [key: string]: React.CSSProperties } = {
  wrapper: { position: 'relative', width: '100%' },
  control: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 10px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', fontSize: '12px', color: '#334155', cursor: 'pointer', transition: 'all 0.15s ease-in-out' },
  valueText: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  menu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)', zIndex: 999, overflow: 'hidden', padding: '4px' },
  option: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 10px', borderRadius: '6px', fontSize: '12px', cursor: 'pointer', transition: 'background-color 0.1s ease', marginBottom: '2px' },
  searchContainer: { display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 10px', backgroundColor: '#f8fafc', borderBottom: '1px solid #e2e8f0' },
  searchInput: { width: '100%', border: 'none', backgroundColor: 'transparent', outline: 'none', fontSize: '11px', color: '#334155' },
  avatarCircle: { width: '22px', height: '22px', borderRadius: '50%', backgroundColor: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', flexShrink: 0 },
  avatarImg: { width: '100%', height: '100%', objectFit: 'cover' }
};

const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '20px', backgroundColor: '#f8fafc', minHeight: '100vh', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', boxSizing: 'border-box' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', maxWidth: '1300px', margin: '0 auto 20px auto' },
  backButton: { background: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '8px', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' },
  title: { fontSize: '20px', fontWeight: 700, color: '#0f172a', margin: 0 },
  subtitle: { fontSize: '12px', color: '#64748b', margin: '2px 0 0 0' },
  mainLayout: { display: 'grid', gridTemplateColumns: '1fr 340px', gap: '20px', maxWidth: '1300px', margin: '0 auto' },
  leftSection: { display: 'flex', flexDirection: 'column' },
  rightSection: { display: 'flex', flexDirection: 'column', gap: '16px' },
  card: { backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '20px', boxSizing: 'border-box' },
  sectionHeading: { fontSize: '15px', fontWeight: 700, color: '#1e293b', margin: '0 0 16px 0' },
  gridRow: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px', marginBottom: '14px' },
  fieldGroup: { display: 'flex', flexDirection: 'column', gap: '5px' },
  label: { fontSize: '11px', fontWeight: 600, color: '#475569' },
  readOnlyInput: { padding: '8px 10px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', fontSize: '12px', color: '#334155', outline: 'none' },
  tagsContainer: { display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '4px', alignItems: 'center' },
  tagPill: { backgroundColor: '#e2e8f0', color: '#334155', padding: '4px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '6px' },
  tagRemoveBtn: { background: 'none', border: 'none', cursor: 'pointer', fontSize: '12px', color: '#64748b', padding: 0 },
  addTagBtn: { backgroundColor: '#f1f5f9', border: '1px dashed #cbd5e1', color: '#cbd5e1', padding: '4px 10px', borderRadius: '4px', fontSize: '11px', fontWeight: 600, cursor: 'pointer' },
  newTagInput: { padding: '4px 8px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '11px', outline: 'none' },
  addTagSubmitBtn: { backgroundColor: '#005a43', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', fontSize: '11px', cursor: 'pointer' },
  footerActions: { display: 'flex', gap: '10px', marginTop: '16px' },
  saveBtn: { backgroundColor: '#005a43', color: '#ffffff', border: 'none', borderRadius: '6px', padding: '9px 16px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' },
  cancelBtn: { backgroundColor: '#ffffff', color: '#334155', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '9px 16px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
  tabsRow: { display: 'flex', borderBottom: '1px solid #e2e8f0', gap: '20px' },
  tabBtn: { background: 'none', border: 'none', padding: '0 0 8px 0', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
  sidebarSubHeading: { fontSize: '12px', fontWeight: 700, color: '#334155', margin: '0 0 10px 0' },
  historyItem: { display: 'flex', gap: '10px', alignItems: 'flex-start', marginBottom: '10px' },
  historyDot: { width: '18px', height: '18px', borderRadius: '50%', color: '#fff', fontSize: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  versionCircle: { width: '10px', height: '10px', borderRadius: '50%', border: '2px solid #005a43', marginTop: '3px', flexShrink: 0 },
  historyText: { fontSize: '11px', color: '#1e293b', display: 'block' },
  historyDate: { fontSize: '10px', color: '#64748b', display: 'block' },
  toast: { position: 'fixed', top: '20px', right: '20px', backgroundColor: '#005a43', color: '#fff', padding: '10px 16px', borderRadius: '6px', fontSize: '12px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px', zIndex: 9999 }
};

export default UserMetadata;