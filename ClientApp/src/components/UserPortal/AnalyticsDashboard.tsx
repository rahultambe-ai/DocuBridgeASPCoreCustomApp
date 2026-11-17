import * as React from 'react';
import { apiFetch, photoUrlFor, fileOpenUrl } from '../../spfx-shim';
import { useState, useEffect, useRef } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import DocumentList from './DocumentList';
import ProcessManagement from './ProcessManagement';
import DocumentApproval from './ApprovalWorkflow';
import LicenseExpiryPopup from './LicenseExpiryPopup';
import LicenseExpiryBanner from './LicenseExpiryBanner';
import Settings from './SystemSettings/SettingsMain';

export interface IAnalyticsDashboardProps {
  context: any;
}

export interface IDocumentItem {
  id: number;
  name: string;
  serverRelativeUrl: string;
  modified: string;
  modifiedTime: number;
  modifiedBy: string;
  size: number;
  extension: string;
  docType?: string;
  processCode?: string;
  status: 'Not in DMS' | 'Draft' | 'Active' | 'Approved' | 'Pending Approval' | 'In Review' | 'Archived';
  // Raw value of the "DocumentCurrentStatus" choice column: Draft, Send for L1,
  // L1 Approved, Send for L2, L2 Approved, Send for L3, L3 Approved. This is what
  // actually tells us which approver's desk a document is sitting on right now.
  currentStage?: string;
  isFavorite?: boolean;
}

// One row from the "Workflow Configuration" list, trimmed down to just the
// process code + the approver emails needed to know who a document is
// assigned to at each stage.
export interface IWorkflowAssignment {
  processCode: string;
  l1Email: string;
  l2Email: string;
  escalateEmail: string;
  finalEmail: string;
}

export interface IProcessSummaryItem {
  id: number;
  processCode: string;
  processName: string;
  department: string;
  status: string;
}

export interface IMicrosoftUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

const SVGIcon = ({ name, color = "currentColor", size = 20 }: { name: string; color?: string; size?: number }) => {
  const icons: { [key: string]: React.ReactNode } = {
    hamburger: (
      <>
        <line x1="4" y1="6" x2="20" y2="6" />
        <line x1="4" y1="12" x2="20" y2="12" />
        <line x1="4" y1="18" x2="20" y2="18" />
      </>
    ),
    dashboard: (
      <>
        <rect x="3.5" y="3.5" width="7" height="7" rx="1.5" />
        <rect x="13.5" y="3.5" width="7" height="7" rx="1.5" />
        <rect x="13.5" y="13.5" width="7" height="7" rx="1.5" />
        <rect x="3.5" y="13.5" width="7" height="7" rx="1.5" />
      </>
    ),
    documents: (
      <>
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <polyline points="14 2 14 8 20 8" />
      </>
    ),
    approvals: (
      <>
        <circle cx="12" cy="12" r="9" />
        <polyline points="9 12 11.5 14.5 15 9.5" />
      </>
    ),
    processes: (
      <>
        <rect x="3.5" y="3.5" width="17" height="17" rx="2.5" />
        <path d="M7.5 12.5h4.5v4.5H7.5z" fill={color} stroke="none" />
      </>
    ),
    settings: (
      <>
        <circle cx="12" cy="12" r="3" />
        <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33h.01a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51h.01a1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82v.01a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
      </>
    ),
    search: (
      <>
        <circle cx="11" cy="11" r="8" />
        <line x1="21" y1="21" x2="16.65" y2="16.65" />
      </>
    ),
    folder: (
      <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
    ),
    users: (
      <>
        <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
        <circle cx="9" cy="7" r="4" />
        <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
        <path d="M16 3.13a4 4 0 0 1 0 7.75" />
      </>
    ),
    checkSquare: (
      <>
        <polyline points="9 11 12 14 22 4" />
        <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
      </>
    ),
    file: (
      <>
        <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z" />
        <polyline points="13 2 13 9 20 9" />
      </>
    ),
    plusFolder: (
      <>
        <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
        <line x1="12" y1="11" x2="12" y2="17" />
        <line x1="9" y1="14" x2="15" y2="14" />
      </>
    ),
    close: (
      <>
        <line x1="18" y1="6" x2="6" y2="18" />
        <line x1="6" y1="6" x2="18" y2="18" />
      </>
    ),
    moreVertical: (
      <>
        <circle cx="12" cy="12" r="1" />
        <circle cx="12" cy="5" r="1" />
        <circle cx="12" cy="19" r="1" />
      </>
    ),
    check: (
      <polyline points="20 6 9 17 4 12" />
    ),
    alertTriangle: (
      <>
        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
        <line x1="12" y1="9" x2="12" y2="13" />
        <line x1="12" y1="17" x2="12.01" y2="17" />
      </>
    ),
    clock: (
      <>
        <circle cx="12" cy="12" r="9" />
        <polyline points="12 7 12 12 15.5 14" />
      </>
    ),
    plus: (
      <>
        <line x1="12" y1="5" x2="12" y2="19" />
        <line x1="5" y1="12" x2="19" y2="12" />
      </>
    )
  };

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke={color}
      strokeWidth="2.2"
      strokeLinecap="round"
      strokeLinejoin="round"
      style={{ display: 'inline-block', verticalAlign: 'middle' }}
    >
      {icons[name] || null}
    </svg>
  );
};

// ---- Licence check (30 / 90 day trial) ----
// The app asks the Licence API how many trial days are left. When daysRemaining
// is 0 or less, the "Trial Period Expired" popup blocks the whole app.
// tenantId and siteId are NOT typed here: both are read from the SharePoint
// context at runtime (tenantId = Microsoft 365 tenant, siteId = this site's Site ID).
const LICENSE_API_URL = '/api/license';

const LICENSE_RECHECK_MS = 30 * 60 * 1000; // re-check every 30 minutes while the app stays open

const BLOCK_WHEN_LICENSE_API_UNREACHABLE = false;

const SETTINGS_ADMIN_EMAILS: string[] = [
  'mihir@mkx8.onmicrosoft.com',
  'dhvani@mkx8.onmicrosoft.com',
  'third@yourdomain.com',
  'fourth@yourdomain.com'
].map(e => e.trim().toLowerCase());

export const AnalyticsDashboard: React.FC<IAnalyticsDashboardProps> = ({ context }) => {
  const [currentUser, setCurrentUser] = useState<{ displayName: string; email: string; jobTitle: string; userPhotoUrl: string }>({
    displayName: 'Loading...',
    email: '',
    jobTitle: 'User',
    userPhotoUrl: ''
  });

  const [documents, setDocuments] = useState<IDocumentItem[]>([]);
  const [filteredDocuments, setFilteredDocuments] = useState<IDocumentItem[]>([]);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(true);
  const [activeMenu, setActiveMenu] = useState<string>('Home');
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);

  const [processes, setProcesses] = useState<IProcessSummaryItem[]>([]);
  const [processesLoading, setProcessesLoading] = useState<boolean>(true);

  // Keyed by normalized process code (trimmed + uppercased) -> approver emails
  // for that process, read live from the "Workflow Configuration" list.
  const [workflowAssignments, setWorkflowAssignments] = useState<Record<string, IWorkflowAssignment>>({});

  const [isFolderModalOpen, setIsFolderModalOpen] = useState<boolean>(false);
  const [folderName, setFolderName] = useState<string>('');
  const [isCreatingFolder, setIsCreatingFolder] = useState<boolean>(false);

  const [openDropdownId, setOpenDropdownId] = useState<number | null>(null);
  const [toastMessage, setToastMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  const [msUsers, setMsUsers] = useState<IMicrosoftUser[]>([]);
  const [isShareModalOpen, setIsShareModalOpen] = useState<boolean>(false);
  const [selectedDocForShare, setSelectedDocForShare] = useState<IDocumentItem | null>(null);
  
  const [userSearchQuery, setUserSearchQuery] = useState<string>('');
  const [selectedMsUser, setSelectedMsUser] = useState<IMicrosoftUser | null>(null);
  const [isDropdownOpen, setIsDropdownOpen] = useState<boolean>(false);
  const [isSharing, setIsSharing] = useState<boolean>(false);

  const dropdownRef = useRef<HTMLDivElement>(null);

  // true => trial is over, LicenseExpiryPopup blocks the app
  const [isLicenseExpired, setIsLicenseExpired] = useState<boolean>(false);
  // trial days left (from Licence API) - shown in the banner under the header
  const [licenseDaysLeft, setLicenseDaysLeft] = useState<number | null>(null);

  useEffect(() => {
    if (context) {
      console.log('Tenant ID:', context.pageContext?.aadInfo?.tenantId?.toString());
      console.log('Site ID:', context.pageContext?.site?.id?.toString());
      console.log('Web ID:', context.pageContext?.web?.id?.toString());
      console.log('Graph Site ID:', `${window.location.hostname},${context.pageContext?.site?.id?.toString()},${context.pageContext?.web?.id?.toString()}`);
      fetchUserProfile();
      fetchSharePointDocuments();
      fetchMicrosoft365Users();
      fetchProcessSummary();
      fetchWorkflowAssignments();
    }

    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);

    // Keep document statuses live: re-pull from SharePoint every 30 minutes so
    // status changes (made here or by anyone else editing the library) show up
    // without needing a manual page refresh.
    const liveRefreshInterval = setInterval(() => {
      if (context) {
        fetchSharePointDocuments();
        fetchWorkflowAssignments();
      }
    }, 30 * 60 * 1000);

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      clearInterval(liveRefreshInterval);
    };
  }, [context]);

  // Licence check: calls the Licence API with this tenant + site and reads daysRemaining.
  useEffect(() => {
    if (!context) return;
    let cancelled = false;

    const checkLicense = async () => {
      try {
        const tenantId = (context.pageContext?.aadInfo?.tenantId?.toString() || '').toUpperCase();
        const siteId = (context.pageContext?.site?.id?.toString() || '').toUpperCase();
        const url = `${LICENSE_API_URL}?tenantId=${encodeURIComponent(tenantId)}&siteId=${encodeURIComponent(siteId)}`;

        // ngrok-skip-browser-warning is only needed while the API runs on a free ngrok URL.
        const response = await apiFetch(url);
        if (!response.ok) throw new Error(`Licence API returned ${response.status}`);

        const data = await response.json();
        const daysRemaining = Number(data?.daysRemaining);
        if (isNaN(daysRemaining)) throw new Error('Licence API response has no valid daysRemaining');

        if (!cancelled) {
          setLicenseDaysLeft(daysRemaining);
          setIsLicenseExpired(daysRemaining <= 0);
        }
      } catch (err) {
        console.error('Licence check failed:', err);
        if (!cancelled) setIsLicenseExpired(BLOCK_WHEN_LICENSE_API_UNREACHABLE);
      }
    };

    checkLicense();
    const licenseTimer = setInterval(checkLicense, LICENSE_RECHECK_MS);
    return () => {
      cancelled = true;
      clearInterval(licenseTimer);
    };
  }, [context]);

  useEffect(() => {
    if (toastMessage) {
      const timer = setTimeout(() => setToastMessage(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [toastMessage]);

  useEffect(() => {
    if (!searchTerm.trim()) {
      setFilteredDocuments(documents);
    } else {
      const lower = searchTerm.toLowerCase();
      setFilteredDocuments(
        documents.filter(doc =>
          doc.name.toLowerCase().includes(lower) ||
          doc.modifiedBy.toLowerCase().includes(lower) ||
          (doc.currentStage || '').toLowerCase().includes(lower)
        )
      );
    }
  }, [searchTerm, documents]);

  const fetchUserProfile = async () => {
    try {
      // The signed-in person comes from the portal login (/api/config). SharePoint only sees the app itself,
      // so sp.web.currentUser() would return the app principal and must not be used here.
      const siteUrl = context.pageContext?.web?.absoluteUrl || '';
      const email = context.pageContext?.user?.email || '';
      const displayName = context.pageContext?.user?.displayName || 'User';

      let photoUrl = '';
      try {
        const client = await context.msGraphClientFactory.getClient();
        const photoResponse = await client.api('/me/photo/$value').responseType('blob').get();
        photoUrl = URL.createObjectURL(photoResponse);
      } catch (graphErr) {
        photoUrl = siteUrl ? photoUrlFor(email) : '';
      }

      setCurrentUser({
        displayName: displayName,
        email: email,
        jobTitle: 'Team Member',
        userPhotoUrl: photoUrl
      });
    } catch (err) {
      const email = context.pageContext?.user?.email || '';
      const siteUrl = context.pageContext?.web?.absoluteUrl || '';
      setCurrentUser({
        displayName: context.pageContext?.user?.displayName || 'User',
        email: email,
        jobTitle: 'Team Member',
        userPhotoUrl: siteUrl ? photoUrlFor(email) : ''
      });
    }
  };

  const fetchSharePointDocuments = async () => {
    setLoading(true);
    try {
      const sp = spfi().using(SPFx(context));

      const DOCUMENT_LIBRARY_ROOT = "Shared Documents";

      // When a process is submitted from the Process Master form, a folder named
      // after that process's Process Code is auto-created inside the document
      // library, and the process's documents live inside that folder. So the
      // "process code" for a file is simply the name of the top-level folder it
      // sits in directly under the library root (not the file's own field).
      const extractProcessCodeFromPath = (relativeUrl: string): string => {
        const parts = relativeUrl.split('/').filter(Boolean);
        const rootIdx = parts.findIndex(p => p.toLowerCase() === DOCUMENT_LIBRARY_ROOT.toLowerCase());
        if (rootIdx === -1) return '';
        // parts[rootIdx] = library root, parts[rootIdx + 1] = top-level folder (process code) if present
        if (rootIdx + 2 < parts.length) {
          return decodeURIComponent(parts[rootIdx + 1]);
        }
        return ''; // file sits directly in the library root — no process folder
      };

      // Recursively walk a folder and all its subfolders, collecting every file along the way
      const getAllFilesRecursively = async (folderServerRelativeUrl: string): Promise<any[]> => {
        const folder = sp.web.getFolderByServerRelativePath(folderServerRelativeUrl);

        const [files, subFolders] = await Promise.all([
          folder.files
            .select(
              "Title",
              "Name",
              "TimeLastModified",
              "ServerRelativeUrl",
              "Length",
              "Author/Title",
              "ListItemAllFields/Id",
              "ListItemAllFields/DocumentStatus",
              "ListItemAllFields/DocumentType",
              "ListItemAllFields/DocumentCurrentStatus"
            )
            .expand("Author", "ListItemAllFields")(),
          folder.folders.select("Name", "ServerRelativeUrl")()
        ]);

        let allFiles = [...files];

        for (const sub of subFolders) {
          // Skip SharePoint's internal "Forms" system folder
          if (sub.Name === "Forms") continue;
          const nestedFiles = await getAllFilesRecursively(sub.ServerRelativeUrl);
          allFiles = allFiles.concat(nestedFiles);
        }

        return allFiles;
      };

      const files = await getAllFilesRecursively(DOCUMENT_LIBRARY_ROOT);

      const mappedDocs: IDocumentItem[] = files.map((file: any, index: number) => {
        const listItem = file.ListItemAllFields || {};
        const modifiedDate = new Date(file.TimeLastModified);
        const liveStatus = (listItem.DocumentStatus as IDocumentItem['status']) || 'Pending Approval';

        return {
          id: listItem.Id ?? index + 1,
          name: file.Name,
          serverRelativeUrl: file.ServerRelativeUrl,
          modified: modifiedDate.toLocaleDateString(),
          modifiedTime: modifiedDate.getTime(),
          modifiedBy: file.Author ? file.Author.Title : 'System',
          size: file.Length || 0,
          extension: file.Name.split('.').pop()?.toLowerCase() || '',
          docType: listItem.DocumentType || '',
          processCode: extractProcessCodeFromPath(file.ServerRelativeUrl),
          status: liveStatus,
          currentStage: listItem.DocumentCurrentStatus || '',
          isFavorite: false
        };
      });

      mappedDocs.sort((a, b) => b.modifiedTime - a.modifiedTime);

      setDocuments(mappedDocs);
      setFilteredDocuments(mappedDocs);
    } catch (error) {
      console.error("Error fetching library files:", error);
    } finally {
      setLoading(false);
    }
  };

  const fetchProcessSummary = async () => {
    setProcessesLoading(true);
    try {
      const sp = spfi().using(SPFx(context));
      const items = await sp.web.lists.getByTitle("Process Master").items
        .select("Id", "Title", "ProcessCode", "Department")
        .orderBy("Modified", false)
        .top(50)();

      const mapped: IProcessSummaryItem[] = items.map((item: any) => ({
        id: item.Id,
        processCode: item.ProcessCode || `PR-${item.Id}`,
        processName: item.Title || 'Untitled Process',
        department: item.Department || 'General',
        status: 'Active'
      }));

      setProcesses(mapped);
    } catch (error) {
      console.error("Error fetching Process Master items:", error);
    } finally {
      setProcessesLoading(false);
    }
  };

  // Reads the "Workflow Configuration" list — the same list the Approvals page
  // uses — to find out, for every process code, who the L1/L2/Escalate To/Final
  // approver actually is. This is what lets us show a document to only the
  // person whose turn it currently is, instead of to everyone.
  const fetchWorkflowAssignments = async () => {
    try {
      const sp = spfi().using(SPFx(context));
      const items = await sp.web.lists.getByTitle("Workflow Configuration").items
        .select(
          "Id", "Title",
          "L1Approver/EMail", "L2Approver/EMail", "EscalateTo/EMail", "FinalApprover/EMail"
        )
        .expand("L1Approver", "L2Approver", "EscalateTo", "FinalApprover")
        .top(200)();

      const emailOf = (val: any): string => (val && typeof val === 'object') ? (val.EMail || '') : '';

      const map: Record<string, IWorkflowAssignment> = {};
      items.forEach((item: any) => {
        const processCode = (item.Title || '').trim();
        if (!processCode) return;
        map[processCode.toUpperCase()] = {
          processCode,
          l1Email: emailOf(item.L1Approver),
          l2Email: emailOf(item.L2Approver),
          escalateEmail: emailOf(item.EscalateTo),
          finalEmail: emailOf(item.FinalApprover)
        };
      });

      setWorkflowAssignments(map);
    } catch (error) {
      console.error("Error fetching Workflow Configuration items:", error);
    }
  };

  const fetchMicrosoft365Users = async () => {
    try {
      const client = await context.msGraphClientFactory.getClient();
      const response = await client.api('/users').select('id,displayName,mail,userPrincipalName').top(999).get();
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
      if (usersList.length > 0) {
        setSelectedMsUser(usersList[0]);
      }
    } catch (err) {
      try {
        const sp = spfi().using(SPFx(context));
        const siteUsers = await sp.web.siteUsers();
        const siteUrl = context.pageContext?.web?.absoluteUrl || '';
        const mapped: IMicrosoftUser[] = siteUsers
          .filter((u: any) => u.Email)
          .map((u: any) => ({
            id: String(u.Id),
            displayName: u.Title,
            mail: u.Email,
            userPrincipalName: u.Email,
            photoUrl: siteUrl ? photoUrlFor(u.Email) : ''
          }));
        setMsUsers(mapped);
        if (mapped.length > 0) setSelectedMsUser(mapped[0]);
      } catch (fallbackErr) {
        console.error("Fallback error:", fallbackErr);
      }
    }
  };

  const handleCreateFolder = async () => {
    if (!folderName.trim()) {
      setToastMessage({ type: 'error', text: 'Please enter a folder name.' });
      return;
    }

    setIsCreatingFolder(true);
    try {
      const sp = spfi().using(SPFx(context));
      const library = sp.web.getFolderByServerRelativePath("Shared Documents");
      await library.folders.addUsingPath(folderName);

      setToastMessage({ type: 'success', text: `Folder "${folderName}" created successfully!` });
      setFolderName('');
      setIsFolderModalOpen(false);
      fetchSharePointDocuments();
    } catch (err) {
      setToastMessage({ type: 'error', text: 'Failed to create folder.' });
    } finally {
      setIsCreatingFolder(false);
    }
  };

  const updateDocumentStatus = async (docId: number, newStatus: IDocumentItem['status']) => {
    setOpenDropdownId(null);

    const targetDoc = documents.find(d => d.id === docId);
    if (!targetDoc) return;

    const previousStatus = targetDoc.status;

    // Optimistic UI update
    setDocuments(prev => prev.map(doc => doc.id === docId ? { ...doc, status: newStatus } : doc));

    try {
      const sp = spfi().using(SPFx(context));
      const fileItem = await sp.web.getFileByServerRelativePath(targetDoc.serverRelativeUrl).getItem();
      await fileItem.update({ DocumentStatus: newStatus });
      setToastMessage({ type: 'success', text: `"${targetDoc.name}" status set to ${newStatus}` });

      // One "DMS Audit Logs" entry per approval/status action. Never blocks the
      // status update above — any failure here is only logged to console.
      try {
        const auditList = sp.web.lists.getByTitle('DMS Audit Logs');
        const pad = (n: number) => String(n).padStart(2, '0');
        const now = new Date();
        const auditId = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
        const currentUserEmail = context?.pageContext?.user?.email || context?.pageContext?.user?.loginName || '';
        await auditList.items.add({
          Title: targetDoc.processCode || '', // "Title" column — renamed to "ProcessID" in the list UI
          EntryType: 'Approval Step',
          AuditID: auditId,
          ApprovalStepName: `Document status changed to ${newStatus}`,
          DocumentName: targetDoc.name,
          ActionBy: currentUserEmail,
          ActionStatus: newStatus,
          ActionDate: new Date().toISOString()
        });
      } catch (auditErr) {
        console.error('DMS Audit Logs entry failed for status change:', auditErr);
      }
    } catch (error) {
      console.error("Error updating document status in SharePoint:", error);
      // Revert local change since SharePoint write failed
      setDocuments(prev => prev.map(doc => doc.id === docId ? { ...doc, status: previousStatus } : doc));
      setToastMessage({ type: 'error', text: `Failed to update status for "${targetDoc.name}"` });
    }
  };

  const handleSendShare = async () => {
    if (!selectedDocForShare || !selectedMsUser) {
      setToastMessage({ type: 'error', text: 'Please select a Microsoft user to share with.' });
      return;
    }

    setIsSharing(true);
    try {
      const client = await context.msGraphClientFactory.getClient();
      const absoluteWebUrl = context.pageContext?.web?.absoluteUrl || '';
      const docUrl = `${absoluteWebUrl}/${selectedDocForShare.serverRelativeUrl}`;

      const chatBody = {
        chatType: "oneOnOne",
        members: [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${currentUser.email}')`
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${selectedMsUser.id}')`
          }
        ]
      };

      let chatId = "";
      try {
        const chatResponse = await client.api('/chats').post(chatBody);
        chatId = chatResponse.id;
      } catch (chatCreateErr) {
        const existingChats = await client.api('/me/chats').expand('members').get();
        const targetChat = existingChats.value.find((c: any) => 
          c.chatType === 'oneOnOne' && c.members.some((m: any) => m.userId === selectedMsUser.id)
        );
        if (targetChat) chatId = targetChat.id;
      }

      if (chatId) {
        const messagePayload = {
          body: {
            contentType: "html",
            content: `Hello <b>${selectedMsUser.displayName}</b>! <b>${currentUser.displayName}</b> has shared a document with you: <br/><br/>📄 <b><a href="${docUrl}" target="_blank">${selectedDocForShare.name}</a></b><br/><br/>Click the link above to view or download it from SharePoint.`
          }
        };
        await client.api(`/chats/${chatId}/messages`).post(messagePayload);
        setToastMessage({ type: 'success', text: `Document shared with ${selectedMsUser.displayName}!` });
      }
    } catch (err: any) {
      setToastMessage({ type: 'error', text: 'Error sending notification.' });
    } finally {
      setIsSharing(false);
      setIsShareModalOpen(false);
      setSelectedDocForShare(null);
    }
  };

  const filteredMsUsers = msUsers.filter(u =>
    u.displayName.toLowerCase().includes(userSearchQuery.toLowerCase()) ||
    u.mail.toLowerCase().includes(userSearchQuery.toLowerCase())
  );

  // ESCALATION_DAYS: a pending doc (see below) still sitting past this many
  // days (measured from its real last-modified time) is surfaced as escalated.
  const ESCALATION_DAYS = 5;
  const nowTime = Date.now();
  const daysSince = (t: number) => (nowTime - t) / (1000 * 60 * 60 * 24);

  // "Pending" is driven by the library's DocumentCurrentStatus column: every document that has
  // a current status shows up EXCEPT 'L3 Approved' (L3 is the final approval, so it is done).
  const pendingDocs = documents.filter(d => !!d.currentStage && d.currentStage !== 'L3 Approved');
  const escalatedDocs = pendingDocs.filter(d => daysSince(d.modifiedTime) >= ESCALATION_DAYS);

  // ---- Real per-user assignment, driven by Workflow Configuration + DocumentCurrentStatus ----
  const normalizeEmail = (e?: string) => (e || '').trim().toLowerCase();
  const myEmail = normalizeEmail(currentUser.email);

  const getAssignment = (doc: IDocumentItem): IWorkflowAssignment | undefined =>
    workflowAssignments[(doc.processCode || '').trim().toUpperCase()];

  // A document is fully approved once its DocumentCurrentStatus reaches "L3 Approved"
  // (final approval).
  const isDocFullyApproved = (doc: IDocumentItem) => doc.currentStage === 'L3 Approved';

  // Whose desk is this document on right now? Overdue items go to the
  // escalation contact instead of whichever approver's stage it's stuck at.
  const getCurrentApproverEmail = (doc: IDocumentItem): string => {
    const assignment = getAssignment(doc);
    if (!assignment || !doc.currentStage || isDocFullyApproved(doc)) return '';

    if (daysSince(doc.modifiedTime) >= ESCALATION_DAYS && assignment.escalateEmail) {
      return assignment.escalateEmail;
    }

    const stage = (doc.currentStage || '').trim();
    if (stage === 'L1 Approved' || stage === 'Send for L2') return assignment.l2Email;
    if (stage === 'L2 Approved' || stage === 'Send for L3') return assignment.finalEmail;
    // Default / "Send for L1" / Draft / blank — sits with the L1 approver first.
    return assignment.l1Email;
  };

  const isMyDeskDoc = (doc: IDocumentItem) => myEmail !== '' && normalizeEmail(getCurrentApproverEmail(doc)) === myEmail;

  // Documents currently awaiting MY action, oldest first (most urgent on top).
  const myPendingApprovals = documents
    .filter(d => isMyDeskDoc(d))
    .sort((a, b) => a.modifiedTime - b.modifiedTime);

  // All documents that have finished the approval flow, library-wide and real-time
  // (currentStage === 'L3 Approved') — not scoped to my own
  // approval chain, so this matches what you actually see as "Approved" in the list.
  const completedApprovals = documents.filter(d => isDocFullyApproved(d));

  // Friendly label for what stage a document is at / who it's waiting on.
  const getStageLabel = (doc: IDocumentItem): string => {
    if (isDocFullyApproved(doc)) return 'Approved';
    if (daysSince(doc.modifiedTime) >= ESCALATION_DAYS) return 'Escalated';
    const stage = (doc.currentStage || '').trim();
    if (stage === 'L1 Approved' || stage === 'Send for L2') return 'Awaiting L2 Approval';
    if (stage === 'L2 Approved' || stage === 'Send for L3') return 'Awaiting Final Approval';
    return 'Awaiting L1 Approval';
  };

  const isSettingsAdmin = myEmail !== '' && SETTINGS_ADMIN_EMAILS.indexOf(myEmail) !== -1;

  const menuItems = [
    { id: 'Home', label: 'Dashboard', iconName: 'dashboard' },
    { id: 'MyFiles', label: 'Documents', iconName: 'documents' },
    { id: 'Approvals', label: 'Approvals', iconName: 'approvals' },
    { id: 'WorkflowSetting', label: 'Processes', iconName: 'processes' },
    
    ...(isSettingsAdmin
      ? [{ id: 'Settings', label: 'Settings', iconName: 'settings' }]
      : [])
  ];

  const getStatusStyle = (status: string) => {
    switch (status) {
      case 'Active':
      case 'Approved':
      case 'L1 Approved':
      case 'L2 Approved':
      case 'L3 Approved':
        return { backgroundColor: '#e6f4ea', color: '#005c42' };
      case 'Send for L1':
      case 'Send for L2':
      case 'Send for L3':
        return { backgroundColor: '#fef3c7', color: '#d97706' };
      case 'L1 Rejected':
      case 'L2 Rejected':
      case 'L3 Rejected':
        return { backgroundColor: '#fee2e2', color: '#b91c1c' };
      case 'In Review':
        return { backgroundColor: '#e0f2fe', color: '#0284c7' };
      case 'Pending Approval':
        return { backgroundColor: '#fef3c7', color: '#d97706' };
      case 'Draft':
        return { backgroundColor: '#f3f4f6', color: '#4b5563' };
      case 'Archived':
        return { backgroundColor: '#e2e8f0', color: '#475569' };
      case 'Not in DMS':
        return { backgroundColor: '#ffedd5', color: '#c2410c' };
      default:
        return { backgroundColor: '#f3f4f6', color: '#4b5563' };
    }
  };

  // Real monthly document activity (last 6 months) for the Approval Trends chart.
  const monthlyTrend = React.useMemo(() => {
    const months: { key: string; label: string; count: number }[] = [];
    const ref = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(ref.getFullYear(), ref.getMonth() - i, 1);
      months.push({ key: `${d.getFullYear()}-${d.getMonth()}`, label: d.toLocaleString('default', { month: 'short' }), count: 0 });
    }
    documents.forEach(doc => {
      const d = new Date(doc.modifiedTime);
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      const bucket = months.find(m => m.key === key);
      if (bucket) bucket.count += 1;
    });
    const max = Math.max(1, ...months.map(m => m.count));
    return { months, max };
  }, [documents]);

  // Real status distribution (for the donut) — actual counts, no fabricated categories.
  const statusBreakdown = React.useMemo(() => {
    // Driven by DocumentCurrentStatus (the choice column), not DocumentStatus.
    const order: { status: string; color: string }[] = [
      { status: 'Draft', color: '#94a3b8' },
      { status: 'Send for L1', color: '#d97706' },
      { status: 'L1 Approved', color: '#65a30d' },
      { status: 'L1 Rejected', color: '#dc2626' },
      { status: 'Send for L2', color: '#0284c7' },
      { status: 'L2 Approved', color: '#0d9488' },
      { status: 'L2 Rejected', color: '#e11d48' },
      { status: 'Send for L3', color: '#7c3aed' },
      { status: 'L3 Approved', color: '#059669' },
      { status: 'L3 Rejected', color: '#b91c1c' },
      { status: 'No Status', color: '#cbd5e1' }
    ];
    const total = documents.length || 1;
    let cursor = 0;
    const segments = order
      .map(o => ({ ...o, count: documents.filter(d => (d.currentStage || 'No Status') === o.status).length }))
      .filter(o => o.count > 0)
      .map(o => {
        const pct = (o.count / total) * 100;
        const seg = { ...o, pct, start: cursor };
        cursor += pct;
        return seg;
      });
    return { segments, total: documents.length };
  }, [documents]);

  return (
    <div style={styles.dashboardWrapper}>
      {/* Licence expired popup — blocks the whole app when the trial days are over */}
      <LicenseExpiryPopup isOpen={isLicenseExpired} />

      {toastMessage && (
        <div style={{
          ...styles.toastNotification,
          backgroundColor: toastMessage.type === 'success' ? '#005c42' : '#ef4444'
        }}>
          <span style={{ fontWeight: 600, fontSize: '13px' }}>{toastMessage.text}</span>
          <button onClick={() => setToastMessage(null)} style={styles.toastCloseBtn}>×</button>
        </div>
      )}

      {/* Top Header */}
      <header style={styles.topHeader}>
        <button
          onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
          style={styles.hamburgerBtn}
          title="Toggle Navigation Menu"
        >
          <SVGIcon name="hamburger" color="#384c66" size={18} />
        </button>

        <div style={styles.searchBarWrapper}>
          <div style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }}>
            <SVGIcon name="search" color="#94a3b8" size={16} />
          </div>
          <input
            type="text"
            placeholder="Search documents across all views (e.g. filename, status)..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchInput}
          />
        </div>
      </header>

      {/* Trial days banner - shows how many licence days are left (hidden when expired / not loaded) */}
      <LicenseExpiryBanner daysRemaining={licenseDaysLeft} />

      <div style={styles.mainContentLayout}>
        <aside style={{
          ...styles.sidebar,
          width: isSidebarCollapsed ? '64px' : '170px',
          minWidth: isSidebarCollapsed ? '64px' : '170px',
          maxWidth: isSidebarCollapsed ? '64px' : '170px',
          padding: isSidebarCollapsed ? '16px 8px' : '16px 12px'
        }}>
          <nav style={styles.navSection}>
            {menuItems.map(item => {
              const isActive = activeMenu === item.id;
              const activeColor = '#005c42';
              const inactiveColor = '#384c66';
              return (
                <button
                  key={item.id}
                  onClick={() => setActiveMenu(item.id)}
                  style={{
                    ...styles.navItem,
                    ...(isActive ? styles.activeNavItem : {}),
                    justifyContent: isSidebarCollapsed ? 'center' : 'flex-start',
                    padding: isSidebarCollapsed ? '10px' : '12px 16px'
                  }}
                  title={isSidebarCollapsed ? item.label : undefined}
                >
                  <SVGIcon
                    name={item.iconName}
                    color={isActive ? activeColor : inactiveColor}
                    size={20}
                  />
                  {!isSidebarCollapsed && (
                    <span style={{ 
                      marginLeft: '12px', 
                      textAlign: 'left', 
                      fontWeight: isActive ? 700 : 600, 
                      fontSize: '13px',
                      color: isActive ? activeColor : inactiveColor,
                      flexGrow: 1 
                    }}>
                      {item.label}
                    </span>
                  )}
                </button>
              );
            })}
          </nav>
        </aside>

        <main style={styles.contentBody}>
          <div style={styles.scrollableContentInner}>
            {activeMenu === 'Home' && (
              <>
                <div style={styles.welcomeBanner}>
                  <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 700 }}>
                    Welcome back, {currentUser.displayName.split(' ')[0]}! 👋
                  </h2>
                  <p style={{ margin: '4px 0 0 0', color: '#666', fontSize: '13px' }}>
                    Here's what's happening in your workspace today.
                  </p>
                </div>

                {/* Top stat cards — every number below is computed live from the fetched library/list data */}
                <div style={styles.statsGrid4Col}>
                  <div style={styles.statCard} onClick={() => setActiveMenu('MyFiles')} role="button">
                    <div style={styles.statIconBox}><SVGIcon name="folder" color="#005c42" /></div>
                    <div>
                      <span style={styles.statLabel}>My Documents</span>
                      <h3 style={styles.statNumber}>{loading ? '—' : documents.length}</h3>
                      <span style={{ fontSize: '10px', color: '#94a3b8' }}>Total documents</span>
                    </div>
                  </div>

                  <div style={styles.statCard} onClick={() => setActiveMenu('Approvals')} role="button">
                    <div style={{ ...styles.statIconBox, backgroundColor: '#fef3c7' }}><SVGIcon name="checkSquare" color="#d97706" /></div>
                    <div>
                      <span style={styles.statLabel}>Pending Approvals</span>
                      <h3 style={styles.statNumber}>{loading ? '—' : pendingDocs.length}</h3>
                      <span style={{ fontSize: '10px', color: '#94a3b8' }}>Awaiting your action</span>
                    </div>
                  </div>

                  <div style={styles.statCard} onClick={() => setActiveMenu('Approvals')} role="button">
                    <div style={{ ...styles.statIconBox, backgroundColor: '#fee2e2' }}><SVGIcon name="alertTriangle" color="#dc2626" /></div>
                    <div>
                      <span style={styles.statLabel}>Escalated Requests</span>
                      <h3 style={styles.statNumber}>{loading ? '—' : escalatedDocs.length}</h3>
                      <span style={styles.statSubTextNegative}>Pending {ESCALATION_DAYS}+ days</span>
                    </div>
                  </div>

                  <div style={styles.statCard} onClick={() => setActiveMenu('Approvals')} role="button">
                    <div style={{ ...styles.statIconBox, backgroundColor: '#e6f4ea' }}><SVGIcon name="approvals" color="#005c42" /></div>
                    <div>
                      <span style={styles.statLabel}>Completed Approvals</span>
                      <h3 style={styles.statNumber}>{loading ? '—' : completedApprovals.length}</h3>
                      <span style={styles.statSubTextPositive}>Successfully completed</span>
                    </div>
                  </div>
                </div>

                {/* My Documents / My Approvals / Workflow Management */}
                <div style={styles.sectionsGrid3Col}>
                  <div style={styles.panelCard}>
                    <div style={styles.panelHeader}>
                      <h4 style={styles.panelTitle}>My Documents</h4>
                      <span onClick={() => setActiveMenu('MyFiles')} style={styles.viewAllLink}>View All</span>
                    </div>

                    <div style={{ overflowX: 'auto', overflowY: 'auto', maxHeight: '260px' }}>
                      <table style={styles.miniTable}>
                        <thead>
                          <tr>
                            <th style={styles.miniTh}>Document Name</th>
                            <th style={styles.miniTh}>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {loading ? (
                            <tr><td style={styles.miniTd} colSpan={2}>Fetching documents...</td></tr>
                          ) : filteredDocuments.length === 0 ? (
                            <tr><td style={styles.miniTd} colSpan={2}>No documents found.</td></tr>
                          ) : (
                            filteredDocuments.map(doc => {
                              const pillStyle = getStatusStyle(doc.currentStage || '');
                              return (
                                <tr key={doc.id}>
                                  <td style={styles.miniTd}>
                                    <a href={fileOpenUrl(doc.serverRelativeUrl)} target="_blank" rel="noreferrer" style={{ color: '#1e293b', textDecoration: 'none', fontWeight: 600 }}>
                                      {doc.name}
                                    </a>
                                  </td>
                                  <td style={styles.miniTd}>
                                    <span style={{ ...styles.statusPill, backgroundColor: pillStyle.backgroundColor, color: pillStyle.color }}>{doc.currentStage || '—'}</span>
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <div style={styles.panelCard}>
                    <div style={styles.panelHeader}>
                      <h4 style={styles.panelTitle}>My Pending Approvals</h4>
                      <span style={styles.countBadgeSm}>{myPendingApprovals.length} on your desk</span>
                    </div>

                    <div style={{ ...styles.listContainer, maxHeight: '260px', overflowY: 'auto' }}>
                      {loading ? (
                        <p style={{ fontSize: '12px', color: '#666' }}>Loading approvals...</p>
                      ) : myPendingApprovals.length === 0 ? (
                        <p style={{ fontSize: '12px', color: '#888' }}>Nothing is waiting on you right now.</p>
                      ) : (
                        myPendingApprovals.map(item => {
                          const overdue = daysSince(item.modifiedTime) >= ESCALATION_DAYS;
                          return (
                            <a
                              key={item.id}
                              href={fileOpenUrl(item.serverRelativeUrl)}
                              target="_blank"
                              rel="noreferrer"
                              style={{ ...styles.approvalRow, textDecoration: 'none', color: 'inherit' }}
                            >
                              <div style={{ flexGrow: 1, overflow: 'hidden', minWidth: 0, paddingRight: '8px' }}>
                                <span style={{ fontWeight: 600, fontSize: '13px', display: 'block', textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden' }}>{item.name}</span>
                                <span style={{ fontSize: '11px', color: '#6b7280', display: 'block' }}>{item.processCode ? `${item.processCode} • ` : ''}{item.modifiedBy}</span>
                              </div>
                              <span style={{
                                fontSize: '10.5px',
                                fontWeight: 700,
                                padding: '4px 9px',
                                borderRadius: '6px',
                                flexShrink: 0,
                                backgroundColor: overdue ? '#fee2e2' : '#fef3c7',
                                color: overdue ? '#dc2626' : '#d97706'
                              }}>
                                {getStageLabel(item)}
                              </span>
                            </a>
                          );
                        })
                      )}
                    </div>
                  </div>

                  <div style={styles.panelCard}>
                    <div style={styles.panelHeader}>
                      <h4 style={styles.panelTitle}>Workflow Management</h4>
                    </div>

                    <div style={{ ...styles.listContainer, marginTop: '12px', maxHeight: '260px', overflowY: 'auto' }}>
                      {loading ? (
                        <p style={{ fontSize: '12px', color: '#666' }}>Loading pending approvals...</p>
                      ) : pendingDocs.length === 0 ? (
                        <p style={{ fontSize: '12px', color: '#888' }}>Nothing pending right now.</p>
                      ) : (
                        pendingDocs.map(doc => {
                          const pillStyle = getStatusStyle(doc.currentStage || '');
                          return (
                            <div
                              key={doc.id}
                              onClick={() => { setSearchTerm(doc.name); setActiveMenu('Approvals'); }}
                              style={styles.workflowPendingRow}
                              role="button"
                            >
                              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px' }}>
                                <span style={styles.workflowProcessTag}>{doc.processCode || 'N/A'}</span>
                                <span style={{ ...styles.statusPill, backgroundColor: pillStyle.backgroundColor, color: pillStyle.color }}>{doc.currentStage}</span>
                              </div>
                              <span style={{ fontWeight: 600, fontSize: '13px', display: 'block', marginTop: '6px', textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden' }}>{doc.name}</span>
                              <span style={{ fontSize: '11px', color: '#6b7280' }}>{doc.modifiedBy} • {doc.modified}</span>
                            </div>
                          );
                        })
                      )}
                    </div>

                    <span onClick={() => setActiveMenu('Approvals')} style={{ ...styles.viewAllLink, display: 'block', marginTop: '10px' }}>View All Pending Approvals →</span>
                  </div>
                </div>

                {/* Reports & Analytics + Escalated Requests — built from real document data, no fabricated numbers */}
                <div style={styles.sectionsGrid2Col}>
                  <div style={styles.panelCard}>
                    <h4 style={styles.panelTitle}>Reports &amp; Analytics</h4>

                    <div style={{ display: 'flex', gap: '20px', marginTop: '14px', flexWrap: 'wrap' }}>
                      <div style={{ flex: '1 1 160px' }}>
                        <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>Document Activity (6 mo)</span>
                        <div style={styles.trendChartRow}>
                          {monthlyTrend.months.map(m => (
                            <div key={m.key} style={styles.trendBarWrap} title={`${m.label}: ${m.count} document(s)`}>
                              <div style={{ ...styles.trendBar, height: `${(m.count / monthlyTrend.max) * 70 + (m.count > 0 ? 4 : 0)}px` }} />
                              <span style={styles.trendBarLabel}>{m.label}</span>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div style={{ flex: '1 1 160px' }}>
                        <span style={{ fontSize: '11px', color: '#64748b', fontWeight: 600 }}>Status Breakdown</span>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '14px', marginTop: '8px' }}>
                          <div
                            style={{
                              width: '84px',
                              height: '84px',
                              borderRadius: '50%',
                              flexShrink: 0,
                              background: statusBreakdown.total === 0
                                ? '#e2e8f0'
                                : `conic-gradient(${statusBreakdown.segments.map(s => `${s.color} ${s.start}% ${s.start + s.pct}%`).join(', ')})`,
                              position: 'relative'
                            }}
                          >
                            <div style={{ position: 'absolute', inset: '14px', backgroundColor: '#fff', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px', fontWeight: 700, color: '#334155' }}>
                              {statusBreakdown.total}
                            </div>
                          </div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                            {statusBreakdown.segments.length === 0 ? (
                              <span style={{ fontSize: '11px', color: '#94a3b8' }}>No documents yet</span>
                            ) : (
                              statusBreakdown.segments.map(s => (
                                <span key={s.status} style={{ fontSize: '10.5px', color: '#475569', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                  <span style={{ width: '8px', height: '8px', borderRadius: '2px', backgroundColor: s.color, display: 'inline-block' }} />
                                  {s.status} ({s.count})
                                </span>
                              ))
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div style={styles.panelCard}>
                    <div style={styles.panelHeader}>
                      <h4 style={styles.panelTitle}>Escalated Requests</h4>
                      <span onClick={() => setActiveMenu('Approvals')} style={styles.viewAllLink}>View All</span>
                    </div>

                    <div style={styles.listContainer}>
                      {loading ? (
                        <p style={{ fontSize: '12px', color: '#666' }}>Checking pending items...</p>
                      ) : escalatedDocs.length === 0 ? (
                        <p style={{ fontSize: '12px', color: '#888' }}>Nothing escalated — all pending items are within {ESCALATION_DAYS} days.</p>
                      ) : (
                        escalatedDocs.slice(0, 4).map(doc => (
                          <div key={doc.id} style={styles.escalatedRow}>
                            <SVGIcon name="alertTriangle" color="#dc2626" size={18} />
                            <div style={{ flexGrow: 1, overflow: 'hidden', marginLeft: '10px' }}>
                              <span style={{ fontWeight: 600, fontSize: '13px', display: 'block', textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden' }}>{doc.name}</span>
                              <span style={{ fontSize: '11px', color: '#dc2626' }}>{doc.currentStage} for {Math.floor(daysSince(doc.modifiedTime))} day(s) — attention needed</span>
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
              </>
            )}

            {activeMenu === 'MyFiles' && (
              <DocumentList context={context} onNavigateToApprovals={() => setActiveMenu('Approvals')} />
            )}

            {activeMenu === 'Approvals' && (
              <DocumentApproval
                context={context}
                documents={documents}
                filteredDocuments={filteredDocuments}
                searchTerm={searchTerm}
                openDropdownId={openDropdownId}
                setOpenDropdownId={setOpenDropdownId}
                updateDocumentStatus={updateDocumentStatus}
                getStatusStyle={getStatusStyle}
                processes={processes}
                processesLoading={processesLoading}
              />
            )}

            {activeMenu === 'WorkflowSetting' && (
              <ProcessManagement graphContext={context} />
            )}

            {activeMenu === 'Settings' && isSettingsAdmin && (
              <Settings context={context} />
            )}
          </div>
        </main>
      </div>

      {isFolderModalOpen && (
        <div style={styles.modalOverlay}>
          <div style={styles.enhancedFolderModalContainer}>
            <div style={styles.enhancedFolderModalHeader}>
              <div style={styles.folderModalIconWrapper}>
                <SVGIcon name="plusFolder" color="#005c42" size={24} />
              </div>
              <div>
                <h3 style={styles.modalTitle}>Create New Folder</h3>
                <span style={{ fontSize: '12px', color: '#64748b' }}>Add a new directory inside Shared Documents</span>
              </div>
              <button onClick={() => { setIsFolderModalOpen(false); setFolderName(''); }} style={styles.modalCloseBtn}>
                <SVGIcon name="close" color="#64748b" size={18} />
              </button>
            </div>

            <div style={styles.enhancedFolderModalBody}>
              <label style={styles.modalLabel}>Folder Name</label>
              <div style={{ position: 'relative' }}>
                <input
                  type="text"
                  placeholder="e.g. Project Financials 2026"
                  value={folderName}
                  onChange={(e) => setFolderName(e.target.value)}
                  style={styles.enhancedFolderInput}
                  autoFocus
                />
              </div>
            </div>

            <div style={styles.modalFooter}>
              <button onClick={() => { setIsFolderModalOpen(false); setFolderName(''); }} style={styles.cancelBtn} disabled={isCreatingFolder}>Cancel</button>
              <button onClick={handleCreateFolder} style={styles.createBtn} disabled={isCreatingFolder}>
                {isCreatingFolder ? 'Creating Folder...' : 'Create Folder'}
              </button>
            </div>
          </div>
        </div>
      )}

      {isShareModalOpen && selectedDocForShare && (
        <div style={styles.modalOverlay}>
          <div style={styles.expandedModalContainer}>
            <div style={styles.modalHeader}>
              <div>
                <h3 style={styles.modalTitle}>Share Document via Teams</h3>
                <span style={{ fontSize: '12px', color: '#64748b' }}>Sends direct live notification to Microsoft Teams chat</span>
              </div>
              <button onClick={() => setIsShareModalOpen(false)} style={styles.modalCloseBtn}>
                <SVGIcon name="close" color="#64748b" size={18} />
              </button>
            </div>

            <div style={styles.expandedModalBody}>
              <div style={{ marginBottom: '16px', padding: '12px 14px', backgroundColor: '#f8fafc', borderRadius: '8px', border: '1px solid #e2e8f0' }}>
                <span style={{ fontSize: '11px', color: '#64748b', display: 'block', textTransform: 'uppercase', fontWeight: 600, marginBottom: '4px' }}>Selected Document:</span>
                <span style={{ fontSize: '13px', fontWeight: 600, color: '#1e293b', wordBreak: 'break-all' }}>📄 {selectedDocForShare.name}</span>
              </div>

              <label style={styles.modalLabel}>Search & Select Microsoft 365 User</label>
              
              <div style={{ position: 'relative', marginBottom: '16px' }} ref={dropdownRef}>
                <div
                  onClick={() => setIsDropdownOpen(!isDropdownOpen)}
                  style={styles.comboboxControl}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', overflow: 'hidden' }}>
                    <div style={styles.userItemAvatarBox}>
                      {selectedMsUser?.photoUrl ? (
                        <img 
                          src={selectedMsUser.photoUrl} 
                          alt="" 
                          style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }}
                          onError={(e) => { (e.target as HTMLElement).style.display = 'none'; }}
                        />
                      ) : null}
                      <span style={{ fontSize: '10px', fontWeight: 700, color: '#005c42', display: selectedMsUser?.photoUrl ? 'none' : 'block' }}>
                        {selectedMsUser?.displayName ? selectedMsUser.displayName.charAt(0).toUpperCase() : 'U'}
                      </span>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                      <span style={{ fontSize: '13px', fontWeight: 600, color: '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        {selectedMsUser ? selectedMsUser.displayName : 'Select user...'}
                      </span>
                      <span style={{ fontSize: '11px', color: '#64748b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                        {selectedMsUser ? selectedMsUser.mail : ''}
                      </span>
                    </div>
                  </div>
                  <span style={{ fontSize: '12px', color: '#64748b', marginLeft: '8px' }}>▼</span>
                </div>

                {isDropdownOpen && (
                  <div style={styles.comboboxDropdown}>
                    <div style={{ padding: '10px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#fff' }}>
                      <div style={{ position: 'relative' }}>
                        <div style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)' }}>
                          <SVGIcon name="search" color="#94a3b8" size={14} />
                        </div>
                        <input
                          type="text"
                          placeholder="Type name or email to filter..."
                          value={userSearchQuery}
                          onChange={(e) => setUserSearchQuery(e.target.value)}
                          style={styles.comboboxSearchInput}
                          autoFocus
                        />
                      </div>
                    </div>

                    <div style={styles.comboboxList}>
                      {filteredMsUsers.length === 0 ? (
                        <div style={{ padding: '20px', textAlign: 'center', fontSize: '13px', color: '#94a3b8' }}>No Microsoft users found</div>
                      ) : (
                        filteredMsUsers.map((user) => {
                          const isSelected = selectedMsUser?.id === user.id;
                          return (
                            <div
                              key={user.id}
                              onClick={() => {
                                setSelectedMsUser(user);
                                setIsDropdownOpen(false);
                                setUserSearchQuery('');
                              }}
                              style={isSelected ? { ...styles.comboboxOption, ...styles.comboboxOptionSelected } : styles.comboboxOption}
                            >
                              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', overflow: 'hidden' }}>
                                <div style={styles.userItemAvatarBox}>
                                  {user.photoUrl ? (
                                    <img 
                                      src={user.photoUrl} 
                                      alt="" 
                                      style={{ width: '100%', height: '100%', borderRadius: '50%', objectFit: 'cover' }}
                                      onError={(e) => { (e.target as HTMLElement).style.display = 'none'; }}
                                    />
                                  ) : null}
                                  <span style={{ fontSize: '10px', fontWeight: 700, color: '#005c42', display: user.photoUrl ? 'none' : 'block' }}>
                                    {user.displayName ? user.displayName.charAt(0).toUpperCase() : 'U'}
                                  </span>
                                </div>
                                <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                                  <span style={{ fontWeight: 600, fontSize: '13px', color: '#1e293b' }}>{user.displayName}</span>
                                  <span style={{ fontSize: '11px', color: '#64748b' }}>{user.mail}</span>
                                </div>
                              </div>
                              {isSelected && <SVGIcon name="check" color="#005c42" size={16} />}
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div style={styles.modalFooter}>
              <button onClick={() => setIsShareModalOpen(false)} style={styles.cancelBtn} disabled={isSharing}>Cancel & Close</button>
              <button onClick={handleSendShare} style={styles.createBtn} disabled={isSharing}>
                {isSharing ? 'Sending Teams Notification...' : 'Send Teams Notification'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  dashboardWrapper: { fontFamily: '"Segoe UI", Arial, sans-serif', backgroundColor: '#f8fafc', height: '100vh', color: '#334155', position: 'relative', display: 'flex', flexDirection: 'column', overflow: 'hidden', boxSizing: 'border-box' },
  toastNotification: { position: 'fixed', top: '20px', right: '20px', color: '#fff', padding: '12px 20px', borderRadius: '8px', boxShadow: '0 10px 25px rgba(0,0,0,0.15)', display: 'flex', alignItems: 'center', gap: '16px', zIndex: 99999 },
  toastCloseBtn: { background: 'none', border: 'none', color: '#fff', fontSize: '18px', fontWeight: 700, cursor: 'pointer' },
  topHeader: { position: 'relative', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '12px 24px', backgroundColor: '#fff', borderBottom: '1px solid #e2e8f0', height: '60px', minHeight: '60px', boxSizing: 'border-box' },
  hamburgerBtn: { position: 'absolute', left: '24px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '8px', width: '36px', height: '36px', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', outline: 'none', flexShrink: 0 },
  searchBarWrapper: { position: 'relative', width: '100%', maxWidth: '480px', margin: '0 auto' },
  searchInput: { width: '100%', padding: '8px 12px 8px 36px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', outline: 'none', backgroundColor: '#f1f5f9', fontFamily: '"Segoe UI", Arial, sans-serif', boxSizing: 'border-box' },
  mainContentLayout: { display: 'flex', flexGrow: 1, height: 'calc(100vh - 60px)', overflow: 'hidden', boxSizing: 'border-box' },
  sidebar: { height: '100%', flexShrink: 0, backgroundColor: '#fff', borderRight: '1px solid #e2e8f0', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflowY: 'auto', transition: 'all 0.2s ease-in-out' },
  navSection: { display: 'flex', flexDirection: 'column', gap: '6px' },
  navItem: { display: 'flex', alignItems: 'center', border: 'none', backgroundColor: 'transparent', borderRadius: '14px', fontSize: '14px', color: '#384c66', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif', width: '100%', boxSizing: 'border-box', transition: 'all 0.15s ease' },
  activeNavItem: { backgroundColor: '#e6f4ea', color: '#005c42', fontWeight: 700 },
  contentBody: { flexGrow: 1, height: '100%', overflowY: 'auto', backgroundColor: '#f8fafc', boxSizing: 'border-box' },
  scrollableContentInner: { padding: '24px', minHeight: '100%', boxSizing: 'border-box' },
  welcomeBanner: { marginBottom: '20px' },
  statsGrid3Col: { display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: '16px', marginBottom: '24px' },
  statsGrid4Col: { display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '16px', marginBottom: '20px' },
  sectionsGrid3Col: { display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: '16px', marginBottom: '20px', boxSizing: 'border-box' },
  miniTable: { width: '100%', borderCollapse: 'collapse', fontSize: '12px' },
  miniTh: { textAlign: 'left', padding: '6px 8px', fontSize: '10.5px', color: '#64748b', fontWeight: 700, borderBottom: '1px solid #e2e8f0', whiteSpace: 'nowrap' },
  miniTd: { padding: '8px', borderBottom: '1px solid #f1f5f9', fontSize: '12px', maxWidth: '220px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  tabRow: { display: 'flex', gap: '4px', backgroundColor: '#f1f5f9', borderRadius: '8px', padding: '3px', marginBottom: '10px' },
  tabItem: { flex: 1, border: 'none', backgroundColor: 'transparent', color: '#64748b', fontSize: '11.5px', fontWeight: 600, padding: '6px 8px', borderRadius: '6px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif' },
  tabItemActive: { flex: 1, border: 'none', backgroundColor: '#005c42', color: '#fff', fontSize: '11.5px', fontWeight: 700, padding: '6px 8px', borderRadius: '6px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif' },
  workflowCreateBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#005c42', color: '#fff', border: 'none', borderRadius: '8px', padding: '8px 12px', fontSize: '12px', fontWeight: 700, cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif' },
  workflowRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #f1f5f9', gap: '8px' },
  workflowPendingRow: { padding: '10px 0', borderBottom: '1px solid #f1f5f9', cursor: 'pointer' },
  workflowProcessTag: { fontSize: '10px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '3px 8px', borderRadius: '7px', whiteSpace: 'nowrap' as const, display: 'inline-block' },
  escalatedRow: { display: 'flex', alignItems: 'center', padding: '10px 0', borderBottom: '1px solid #f1f5f9' },
  trendChartRow: { display: 'flex', alignItems: 'flex-end', gap: '8px', height: '96px', marginTop: '10px', paddingLeft: '2px' },
  trendBarWrap: { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '4px', flex: 1 },
  trendBar: { width: '100%', maxWidth: '18px', backgroundColor: '#005c42', borderRadius: '3px 3px 0 0', minHeight: '2px' },
  trendBarLabel: { fontSize: '9.5px', color: '#94a3b8' },
  statCard: { backgroundColor: '#fff', padding: '16px', borderRadius: '12px', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '12px', cursor: 'pointer', boxSizing: 'border-box' },
  statIconBox: { width: '42px', height: '42px', borderRadius: '8px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  statLabel: { fontSize: '11px', color: '#64748b', display: 'block' },
  statNumber: { margin: '2px 0', fontSize: '18px', fontWeight: 700 },
  statSubTextPositive: { fontSize: '10px', color: '#005c42', fontWeight: 600 },
  statSubTextNegative: { fontSize: '10px', color: '#dc2626', fontWeight: 600 },
  sectionsGrid2Col: { display: 'grid', gridTemplateColumns: 'repeat(2, minmax(0, 1fr))', gap: '16px', boxSizing: 'border-box' },
  panelCard: { backgroundColor: '#fff', borderRadius: '12px', border: '1px solid #e2e8f0', padding: '16px', boxSizing: 'border-box', minWidth: 0 },
  panelHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' },
  panelTitle: { margin: 0, fontSize: '14px', fontWeight: 700 },
  viewAllLink: { fontSize: '11px', color: '#005c42', fontWeight: 600, cursor: 'pointer', textDecoration: 'none' },
  countBadgeSm: { fontSize: '10.5px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '3px 9px', borderRadius: '999px', whiteSpace: 'nowrap' as const, flexShrink: 0 },
  listContainer: { display: 'flex', flexDirection: 'column', gap: '10px' },
  docRow: { display: 'flex', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid #f1f5f9' },
  approvalRow: { padding: '10px 0', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxSizing: 'border-box', width: '100%' },
  approvalCenterRow: { padding: '12px 14px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', backgroundColor: '#fff', boxSizing: 'border-box' },
  approveBtn: { backgroundColor: '#005c42', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 10px', fontSize: '11px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif', fontWeight: 600, flexShrink: 0 },
  rejectBtn: { backgroundColor: '#dc2626', color: '#fff', border: 'none', borderRadius: '4px', padding: '4px 10px', fontSize: '11px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif', fontWeight: 600, flexShrink: 0 },
  threeDotBtn: { background: 'none', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '4px 6px', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', backgroundColor: '#f8fafc' },
  dropdownMenu: { position: 'absolute', right: 0, top: '100%', marginTop: '4px', backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.1)', zIndex: 100, width: '140px', overflow: 'hidden', display: 'flex', flexDirection: 'column' },
  statusPill: { fontSize: '10px', fontWeight: 600, padding: '2px 8px', borderRadius: '12px', textTransform: 'uppercase', flexShrink: 0 },
  modalOverlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0, 0, 0, 0.4)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999 },
  enhancedFolderModalContainer: { backgroundColor: '#fff', borderRadius: '14px', width: '440px', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)', overflow: 'hidden' },
  enhancedFolderModalHeader: { display: 'flex', alignItems: 'center', gap: '14px', padding: '20px 24px 16px 24px', borderBottom: '1px solid #f1f5f9' },
  folderModalIconWrapper: { width: '44px', height: '44px', borderRadius: '10px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  enhancedFolderModalBody: { padding: '20px 24px' },
  enhancedFolderInput: { width: '100%', padding: '10px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '13px', outline: 'none', backgroundColor: '#fff', fontFamily: '"Segoe UI", Arial, sans-serif', boxSizing: 'border-box' },
  expandedModalContainer: { backgroundColor: '#fff', borderRadius: '12px', width: '520px', height: '400px', maxHeight: '120vh', boxShadow: '0 10px 25px rgba(0,0,0,0.15)', display: 'flex', flexDirection: 'column', overflow: 'hidden' },
  modalHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '20px 24px 12px 24px', flexShrink: 0 },
  modalTitle: { margin: 0, fontSize: '18px', fontWeight: 700, color: '#1e293b' },
  modalCloseBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '4px', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '4px' },
  expandedModalBody: { padding: '8px 24px 24px 24px', overflowY: 'auto', flexGrow: 1 },
  modalLabel: { display: 'block', fontSize: '12px', fontWeight: 600, color: '#64748b', marginBottom: '8px' },
  modalFooter: { display: 'flex', justifyContent: 'flex-end', gap: '10px', padding: '16px 24px', backgroundColor: '#f8fafc', borderTop: '1px solid #f1f5f9', flexShrink: 0 },
  cancelBtn: { padding: '8px 16px', border: '1px solid #cbd5e1', backgroundColor: '#fff', borderRadius: '8px', cursor: 'pointer', fontSize: '12px', fontWeight: 600, fontFamily: '"Segoe UI", Arial, sans-serif', color: '#334155' },
  createBtn: { padding: '8px 16px', border: 'none', backgroundColor: '#005c42', color: '#fff', borderRadius: '8px', cursor: 'pointer', fontSize: '12px', fontWeight: 600, fontFamily: '"Segoe UI", Arial, sans-serif' },
  comboboxControl: { display: 'flex', alignItems: 'center', padding: '10px 14px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#fff', cursor: 'pointer', justifyContent: 'space-between' },
  comboboxDropdown: { position: 'absolute', top: '100%', left: 0, right: 0, marginTop: '4px', backgroundColor: '#fff', border: '1px solid #cbd5e1', borderRadius: '8px', boxShadow: '0 8px 20px rgba(0,0,0,0.12)', zIndex: 1000, overflow: 'hidden' },
  comboboxSearchInput: { width: '100%', padding: '8px 8px 8px 30px', borderRadius: '6px', border: '1px solid #e2e8f0', fontSize: '12px', outline: 'none', boxSizing: 'border-box', backgroundColor: '#f8fafc' },
  comboboxList: { maxHeight: '300px', overflowY: 'auto' },
  comboboxOption: { padding: '10px 14px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer', borderBottom: '1px solid #f1f5f9' },
  comboboxOptionSelected: { backgroundColor: '#e6f4ea' },
  userItemAvatarBox: { width: '28px', height: '28px', borderRadius: '50%', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: '0', overflow: 'hidden' }
};

export default AnalyticsDashboard;