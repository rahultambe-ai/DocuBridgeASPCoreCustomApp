import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx, fileOpenUrl } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { UserMetadata } from './DocumentMetadata';
import ProcessManagement from './ProcessManagement';
import UploadDocumentModal, { IUploadMetadata } from './UploadDocuments';
import AssignWorkflow from './AssignWorkflow';
import GuestAccess from './GuestAccess';

export interface IUserDashboardProps {
  context: any;
  onCreateProcess?: () => void;
  onViewAllProcesses?: () => void;
  onSelectProcess?: (process: IProcessNavItem) => void;
  /** Navigates to the Approval Management screen (e.g. "+ Create Workflow" clicked
   * from the Upload Document modal's "no workflow configured" notice). */
  onNavigateToApprovals?: () => void;
}

export interface IProcessNavItem {
  id: number;
  processCode: string;
  processName: string;
}

export interface IDocumentItem {
  id: number;
  title: string;
  subTitle?: string;
  docNumber: string;
  docType: string;
  version: string;
  status: 'Not in DMS' | 'Draft' | 'Active' | 'Pending Approval' | 'In Review' | 'Archived';
  /** "DocumentCurrentStatus" column (Draft, Send for L1, L1 Approved, ...) — this is what the STATUS column in the list shows. */
  currentStatus?: string;
  /** "WorkflowID" column — the workflow assigned to this document (blank = no workflow yet). */
  workflowId?: string;
  fileExtension: 'doc' | 'xls' | 'pdf' | 'img' | 'other';
  serverRelativeUrl: string;
  size: number;
  modified: string;
  modifiedTime: number;
  modifiedBy: string;
  owner: string;
  tags: string[];
  docNumberIsSaved: boolean;
}

export interface IMicrosoftUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  photoUrl?: string;
}

const SVGIcon = ({ name, color = "currentColor", size = 14 }: { name: string; color?: string; size?: number }) => {
  const icons: { [key: string]: React.ReactNode } = {
    search: <path d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />,
    plus: <path d="M12 4v16m8-8H4" />,
    moreVertical: <path d="M12 13a1 1 0 100-2 1 1 0 000 2zm0 5a1 1 0 100-2 1 1 0 000 2zm0-10a1 1 0 100-2 1 1 0 000 2z" />,
    close: <path d="M6 18L18 6M6 6l12 12" />,
    chevronLeft: <path d="M15 18l-6-6 6-6" />,
    chevronRight: <path d="M9 18l6-6-6-6" />,
    chevronDown: <path d="M6 9l6 6 6-6" />,
    flag: <path d="M4 21V4a1 1 0 011-1h11l-2 4 2 4H6a1 1 0 00-1 1v9" />,
    arrowRight: <path d="M5 12h14M13 6l6 6-6 6" />
  };

  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }}>
      {icons[name] || null}
    </svg>
  );
};

const FileTypeBadge = ({ type }: { type: 'doc' | 'xls' | 'pdf' | 'img' | 'other' }) => {
  const badgeStyles: { [key: string]: { label: string } } = {
    doc: { label: 'DOC' },
    xls: { label: 'XLS' },
    pdf: { label: 'PDF' },
    img: { label: 'IMG' },
    other: { label: 'FILE' }
  };

  const current = badgeStyles[type] || badgeStyles.other;

  return (
    <div style={{ width: '28px', height: '28px', borderRadius: '4px', border: '1px solid #cbd5e1', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '8px', fontWeight: 700, color: '#475569', backgroundColor: '#f8fafc', flexShrink: 0 }}>
      {current.label}
    </div>
  );
};

export const UserDashboard: React.FC<IUserDashboardProps> = ({ context, onCreateProcess, onViewAllProcesses, onSelectProcess, onNavigateToApprovals }) => {
  const [currentUser, setCurrentUser] = useState<{ displayName: string; email: string }>({ displayName: 'User', email: '' });
  const [processNavItems, setProcessNavItems] = useState<IProcessNavItem[]>([]);
  const [loadingProcessNav, setLoadingProcessNav] = useState<boolean>(true);
  const [showProcessMaster, setShowProcessMaster] = useState<boolean>(false);
  const [selectedFolderCode, setSelectedFolderCode] = useState<string | null>(null);
  const [documents, setDocuments] = useState<IDocumentItem[]>([]);
  const [filteredDocuments, setFilteredDocuments] = useState<IDocumentItem[]>([]);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(true);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isUploadModalOpen, setIsUploadModalOpen] = useState<boolean>(false);
  const [openDropdownId, setOpenDropdownId] = useState<number | null>(null);
  // Document whose "Assign Workflow" popup is open.
  const [assignTarget, setAssignTarget] = useState<IDocumentItem | null>(null);
  // Document whose "Guest Access" popup is open.
  const [guestAccessTarget, setGuestAccessTarget] = useState<IDocumentItem | null>(null);
  const [toastMessage, setToastMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  // Role of the logged-in user (from "Process Assignment Groups") for the selected process folder.
  // Contributor = can upload + edit, Viewer = read-only. null = not listed (e.g. admin/owner) -> not restricted.
  const [folderRole, setFolderRole] = useState<'Contributor' | 'Viewer' | null>(null);

  const [selectedDocForMetadata, setSelectedDocForMetadata] = useState<IDocumentItem | null>(null);

  const [currentPage, setCurrentPage] = useState<number>(1);
  const [itemsPerPage, setItemsPerPage] = useState<number>(10);
  const [isItemsPerPageOpen, setIsItemsPerPageOpen] = useState<boolean>(false);

  const actionMenuRef = useRef<HTMLDivElement>(null);
  const itemsPerPageRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (context) {
      fetchUserProfile().then((userName) => {
        fetchSharePointDocuments(userName);
      });
      fetchProcessNavigationItems();
    }

    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Node;
      if (actionMenuRef.current && !actionMenuRef.current.contains(target)) {
        setOpenDropdownId(null);
      }
      if (itemsPerPageRef.current && !itemsPerPageRef.current.contains(target)) {
        setIsItemsPerPageOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [context]);

  useEffect(() => {
    if (toastMessage) {
      const timer = setTimeout(() => setToastMessage(null), 4000);
      return () => clearTimeout(timer);
    }
  }, [toastMessage]);

  useEffect(() => {
    let cancelled = false;
    const fetchFolderRole = async () => {
      if (!context || !selectedFolderCode) {
        setFolderRole(null);
        return;
      }
      try {
        const sp = spfi().using(SPFx(context));
        const safeCode = selectedFolderCode.replace(/'/g, "''");
        const rows: any[] = await sp.web.lists.getByTitle('Process Assignment Groups').items
          .select('Role', 'AssignedUsers')
          .filter(`Title eq '${safeCode}'`)();

        const me = currentUser.displayName.trim().toLowerCase();
        let role: 'Contributor' | 'Viewer' | null = null;
        for (let i = 0; i < rows.length; i++) {
          const names = String(rows[i].AssignedUsers || '').split(';').map((n: string) => n.trim().toLowerCase());
          if (names.indexOf(me) === -1) continue;
          if (rows[i].Role === 'Contributor') { role = 'Contributor'; break; }
          if (rows[i].Role === 'Viewer') role = 'Viewer';
        }
        if (!cancelled) setFolderRole(role);
      } catch (error) {
        console.error('Error fetching assignment role for process folder:', error);
        if (!cancelled) setFolderRole(null);
      }
    };
    fetchFolderRole();
    return () => { cancelled = true; };
  }, [context, selectedFolderCode, currentUser.displayName]);

  const isViewer = folderRole === 'Viewer';

  useEffect(() => {
    setCurrentPage(1);
    if (!searchTerm.trim()) {
      setFilteredDocuments(documents);
    } else {
      const lower = searchTerm.toLowerCase();
      setFilteredDocuments(
        documents.filter(doc =>
          doc.title.toLowerCase().includes(lower) ||
          doc.docNumber.toLowerCase().includes(lower) ||
          doc.docType.toLowerCase().includes(lower) ||
          (doc.currentStatus || '').toLowerCase().includes(lower) ||
          (doc.subTitle && doc.subTitle.toLowerCase().includes(lower))
        )
      );
    }
  }, [searchTerm, documents]);

  const fetchUserProfile = async (): Promise<string> => {
    // The signed-in person comes from the portal login (/api/config). SharePoint only sees the app itself,
    // so sp.web.currentUser() would return the app principal and must not be used here.
    const name = context.pageContext?.user?.displayName || 'User';
    setCurrentUser({
      displayName: name,
      email: context.pageContext?.user?.email || ''
    });
    return name;
  };

  const getLibraryRelativePath = (folderCode?: string | null): string => {
    return folderCode ? `Shared Documents/${folderCode}` : "Shared Documents";
  };

  const fetchSharePointDocuments = async (currentUserName?: string, folderCode?: string | null) => {
    setLoading(true);
    try {
      const sp = spfi().using(SPFx(context));
      const library = sp.web.getFolderByServerRelativePath(getLibraryRelativePath(folderCode));

      const files = await library.files
        .select(
          "Title",
          "Name",
          "TimeLastModified",
          "ServerRelativeUrl",
          "Length",
          "UIVersionLabel",
          "Author/Title",
          "ListItemAllFields/Id",
          "ListItemAllFields/FileStatus",
          "ListItemAllFields/DocumentType",
          "ListItemAllFields/DocumentStatus",
          "ListItemAllFields/DocumentCurrentStatus",
          "ListItemAllFields/WorkflowID",
          "ListItemAllFields/DocumentTags"
        )
        .expand("Author", "ListItemAllFields")();

      const types = ['SOP', 'MIS File', 'Reference Document', 'Checklist', 'SOP'];
      const activeUserName = currentUserName || currentUser.displayName;

      const mappedDocs: IDocumentItem[] = files.map((file: any, index: number) => {
        const ext = file.Name.split('.').pop()?.toLowerCase() || '';
        let fileType: 'doc' | 'xls' | 'pdf' | 'img' | 'other' = 'other';
        if (ext.includes('doc')) fileType = 'doc';
        else if (ext.includes('xls') || ext.includes('csv')) fileType = 'xls';
        else if (ext.includes('pdf')) fileType = 'pdf';
        else if (['png', 'jpg', 'jpeg', 'gif', 'svg', 'webp'].includes(ext)) fileType = 'img';

        const seqNumber = String(index + 1).padStart(3, '0');
        const modifiedDate = new Date(file.TimeLastModified);
        const authorName = file.Author ? file.Author.Title : activeUserName;

        const listItem = file.ListItemAllFields || {};

        const realId: number = listItem.Id ?? (index + 1);
        const docNumberIsSaved: boolean = !!listItem.FileStatus;
        const liveDocNumber: string = listItem.FileStatus || `PUR_D${seqNumber}`;
        const liveDocType: string = listItem.DocumentType || types[index % types.length];
        const liveStatus = (listItem.DocumentStatus as IDocumentItem['status']) || 'Pending Approval';
        // Live version comes straight from SharePoint's own built-in file versioning
        // (UIVersionLabel), not a manually-maintained column. It updates automatically
        // whenever the file/list item is modified in the library.
        const liveVersion: string = file.UIVersionLabel ? `v${file.UIVersionLabel}` : 'v1.0';
        const liveTags: string[] = listItem.DocumentTags
          ? String(listItem.DocumentTags).split(',').map((t: string) => t.trim()).filter(Boolean)
          : [];

        return {
          id: realId,
          title: file.Name.replace(/\.[^/.]+$/, ""),
          subTitle: 'KRA PUR-KRA01',
          docNumber: liveDocNumber,
          docType: liveDocType,
          version: liveVersion,
          status: liveStatus,
          currentStatus: listItem.DocumentCurrentStatus || '',
          workflowId: listItem.WorkflowID || '',
          fileExtension: fileType,
          serverRelativeUrl: file.ServerRelativeUrl,
          size: file.Length || 0,
          modified: modifiedDate.toLocaleDateString(),
          modifiedTime: modifiedDate.getTime(),
          modifiedBy: authorName,
          owner: authorName,
          tags: liveTags,
          docNumberIsSaved: docNumberIsSaved
        };
      });

      mappedDocs.sort((a, b) => b.modifiedTime - a.modifiedTime);

      setDocuments(mappedDocs);
      setFilteredDocuments(mappedDocs);
    } catch (error) {
      console.error("Error fetching files:", error);
    } finally {
      setLoading(false);
    }
  };

  const fetchProcessNavigationItems = async () => {
    setLoadingProcessNav(true);
    try {
      const sp = spfi().using(SPFx(context));

      // 1. Only real folders that actually exist inside the document library represent valid processes
      const libraryFolders: any[] = await sp.web.getFolderByServerRelativePath("Shared Documents").folders
        .select("Name")();

      const realFolderNames = new Set<string>(
        libraryFolders
          .map((f: any) => f.Name as string)
          .filter((name: string) => !!name && name.toLowerCase() !== 'forms')
      );

      if (realFolderNames.size === 0) {
        setProcessNavItems([]);
        return;
      }

      // 2. Resolve a friendly process name for each folder using the Process Master list (matched by Process Code)
      const list = sp.web.lists.getByTitle('Process Master');
      const items: any[] = await list.items
        .select("Id", "Title", "ProcessCode")
        .orderBy("Id", true)();

      const mappedItems: IProcessNavItem[] = items
        .filter((item) => realFolderNames.has(item.ProcessCode))
        .map((item) => ({
          id: item.Id,
          processCode: item.ProcessCode,
          processName: item.Title || item.ProcessCode
        }));

      const matchedCodes = new Set(mappedItems.map((m) => m.processCode));
      let fallbackId = -1;
      realFolderNames.forEach((folderName) => {
        if (!matchedCodes.has(folderName)) {
          mappedItems.push({ id: fallbackId--, processCode: folderName, processName: folderName });
        }
      });

      setProcessNavItems(mappedItems);

      // Always default to the most recently created process folder (last item, since
      // Process Master is fetched ordered by Id ascending) so the user lands on it
      // straight away instead of having to click it manually every time. Also refetch
      // the documents for that folder — just setting the selected code alone left the
      // doc list stuck on whatever the initial unfiltered fetch returned.
      if (mappedItems.length > 0) {
        const latestFolderCode = mappedItems[mappedItems.length - 1].processCode;
        setSelectedFolderCode(latestFolderCode);
        fetchSharePointDocuments(undefined, latestFolderCode);
      }
    } catch (error) {
      console.error("Error fetching process folders:", error);
      setProcessNavItems([]);
    } finally {
      setLoadingProcessNav(false);
    }
  };

  const handleGoToProcessMaster = () => {
    setShowProcessMaster(true);
    setOpenDropdownId(null);
  };

  const handleBackFromProcessMaster = () => {
    setShowProcessMaster(false);
    fetchProcessNavigationItems();
  };

  const handleSelectFolder = (processCode: string) => {
    const newSelection = selectedFolderCode === processCode ? null : processCode;
    setSelectedFolderCode(newSelection);
    setCurrentPage(1);
    fetchSharePointDocuments(undefined, newSelection);
  };

  // Same library used by DocumentMetadata.tsx's own SharePoint save call.
  const DOCUMENT_LIBRARY_TITLE = 'Documents';

  // Shared audit trail list — every module (Process/Document/Workflow/Approval Step/
  // RACI/Assignment Group) logs its entries here, one row per action.
  // https://mkx8.sharepoint.com/Lists/DMS%20Audit%20Logs/AllItems.aspx
  const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';

  // Logs ONE entry per uploaded file into the "DMS Audit Logs" list, right after
  // the file itself lands in the document library. AuditID carries the exact
  // yyyyMMddHHmmss moment the entry was written; DocumentID is the same
  // datetime-based id that's written onto the document library item's FileStatus column.
  const addDocumentAuditEntry = async (
    fileName: string,
    documentId: string,
    processCode: string | null,
    metadata?: IUploadMetadata
  ) => {
    try {
      const sp = spfi().using(SPFx(context));
      const auditList = sp.web.lists.getByTitle(AUDIT_LOG_LIST_TITLE);
      const matchedProcess = processNavItems.find(p => p.processCode === processCode);
      await auditList.items.add({
        Title: processCode || '',      // "Title" column — renamed to "ProcessID" in the list UI
        EntryType: 'Document',
        AuditID: generateDocumentId(), // yyyyMMddHHmmss — exact moment this entry was logged
        ProcessName: matchedProcess ? matchedProcess.processName : '',
        DocumentID: documentId,
        DocumentName: fileName,
        FileName: fileName,
        DocumentType: metadata?.documentType || '',
        DocumentStatus: metadata?.documentStatus || '',
        Tags: (metadata?.tags || []).join(', ')
      });
    } catch (err) {
      console.error('Error adding entry to DMS Audit Logs list for', fileName, err);
    }
  };

  // After a file lands in the library, SharePoint creates its list item automatically —
  // there's no id to write metadata onto until it exists, so we look the new item up by
  // file name (FileLeafRef), exactly the way DocumentMetadata.tsx resolves an item id.
  //
  // IMPORTANT: this writes the Document ID (FileStatus) AND the metadata (DocumentType /
  // DocumentStatus / DocumentTags), when supplied, in a SINGLE .update() call. SharePoint
  // bumps the minor version on every separate .update(), so doing this as two calls was
  // pushing every upload straight to v0.3 (v0.1 initial upload, v0.2 metadata, v0.3 Document
  // ID). One combined call keeps every upload at v0.2, no matter whether the Metadata panel
  // was touched or left at its defaults.
  // Returns the generated Document ID so it can also be logged to DMS Audit Logs.
  const finalizeUploadedFile = async (fileName: string, metadata?: IUploadMetadata): Promise<string> => {
    const documentId = generateDocumentId();
    try {
      const sp = spfi().using(SPFx(context));
      const list = sp.web.lists.getByTitle(DOCUMENT_LIBRARY_TITLE);
      const safeFileName = fileName.replace(/'/g, "''");

      const matches: any[] = await list.items
        .filter(`FileLeafRef eq '${safeFileName}'`)
        .select('Id')
        .top(1)();

      if (!matches || matches.length === 0) {
        console.error('Could not resolve the new item to write Document ID/metadata onto:', fileName);
        return documentId;
      }

      await list.items.getById(matches[0].Id).update({
        FileStatus: documentId,
        // Every freshly uploaded document auto-enters the approval chain at L1 —
        // always set, independent of the Metadata panel, so the Power Automate
        // flow watching this column notifies that process's L1 Approver.
        // (Upload screen sends 'Draft' when Select Workflow is None.)
        DocumentCurrentStatus: metadata?.documentCurrentStatus || 'Send for L1',
        // WorkflowID of the workflow selected on the Upload screen (copied from that row of
        // the Workflow Configuration list) — same single update call, so no extra version.
        ...(metadata?.workflowId ? { WorkflowID: metadata.workflowId } : {}),
        ...(metadata ? {
          DocumentType: metadata.documentType,
          DocumentStatus: metadata.documentStatus,
          DocumentTags: metadata.tags.join(', ')
        } : {})
      });
    } catch (err) {
      console.error('Error writing Document ID/metadata to SharePoint for', fileName, err);
    }
    return documentId;
  };

  // Generates a Document ID from the current date/time down to the second,
  // e.g. 31-Aug-2026 14:30:05 -> "20260831143005".
  const generateDocumentId = (): string => {
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
  };

  // Same SharePoint upload logic as before — used by both the hidden file input
  // and the new "Upload documents" popup, so behavior stays identical either way.
  // When metadata is supplied (from the Upload popup's Metadata panel), it's written
  // onto each newly created item's DocumentType/DocumentStatus/DocumentTags columns
  // right after the upload — the same columns DocumentMetadata.tsx edits later.
  const uploadFilesToFolder = async (files: File[], metadata?: IUploadMetadata) => {
    if (!files || files.length === 0) return;

    setIsUploading(true);
    try {
      const sp = spfi().using(SPFx(context));
      const targetFolder = sp.web.getFolderByServerRelativePath(getLibraryRelativePath(selectedFolderCode));

      for (let i = 0; i < files.length; i++) {
        await targetFolder.files.addUsingPath(files[i].name, files[i], { Overwrite: true });
        // Metadata (if provided) and the datetime-based Document ID are written onto
        // FileStatus/DocumentType/DocumentStatus/DocumentTags together in ONE update
        // call, so the item only picks up one extra version (v0.1 -> v0.2) instead of
        // stacking a separate version for metadata and another for the Document ID.
        const documentId = await finalizeUploadedFile(files[i].name, metadata);
        await addDocumentAuditEntry(files[i].name, documentId, selectedFolderCode, metadata);
      }

      const fileWord = files.length > 1 ? `${files.length} documents` : 'Document';
      setToastMessage({
        type: 'success',
        text: selectedFolderCode
          ? `${fileWord} uploaded to ${selectedFolderCode} successfully!`
          : `${fileWord} uploaded successfully!`
      });
      setCurrentPage(1);
      fetchSharePointDocuments(undefined, selectedFolderCode);
      setIsUploadModalOpen(false);
    } catch {
      setToastMessage({ type: 'error', text: 'Error uploading files.' });
    } finally {
      setIsUploading(false);
    }
  };

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;
    await uploadFilesToFolder(Array.from(files));
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Called by the UploadDocument popup once the user picks files, sets Document Type /
  // Status / Tags, and clicks "Submit for Approval". Upload uses the exact same logic
  // as the plain file input above, and the metadata gets written to the same SharePoint
  // columns DocumentMetadata.tsx uses, so it shows up live everywhere else immediately.
  const handleModalUpload = (files: File[], metadata: IUploadMetadata) => {
    uploadFilesToFolder(files, metadata);
  };

  const handleOpenDocumentInNewTab = (serverRelativeUrl: string) => {
    // Streamed through this app so no Microsoft sign-in is ever shown.
    window.open(fileOpenUrl(serverRelativeUrl), '_blank', 'noopener,noreferrer');
  };

  const handleOpenMetadata = (doc: IDocumentItem) => {
    setSelectedDocForMetadata(doc);
    setOpenDropdownId(null);
  };

  const handleOpenAssignWorkflow = (doc: IDocumentItem) => {
    setAssignTarget(doc);
    setOpenDropdownId(null);
  };

  const handleOpenGuestAccess = (doc: IDocumentItem) => {
    setGuestAccessTarget(doc);
    setOpenDropdownId(null);
  };

  // Process code of a document = the folder it lives in under the library
  // (…/Documents/PROC-M-003/file.docx); falls back to the selected process.
  const getDocProcessCode = (doc: IDocumentItem): string => {
    const parts = (doc.serverRelativeUrl || '').split('/');
    const idx = parts.indexOf(DOCUMENT_LIBRARY_TITLE);
    if (idx >= 0 && parts.length > idx + 2) {
      try { return decodeURIComponent(parts[idx + 1]); } catch (e) { return parts[idx + 1]; }
    }
    return selectedFolderCode || '';
  };

  if (showProcessMaster) {
    return (
      <div style={styles.pageContainer}>
        <div style={styles.contentWrapper}>
          <button onClick={handleBackFromProcessMaster} style={styles.backToDocsBtn}>
            <SVGIcon name="chevronLeft" color="#005a43" size={14} />
            <span>Back to Documents</span>
          </button>
          <ProcessManagement graphContext={context} />
        </div>
      </div>
    );
  }

  if (selectedDocForMetadata) {
    const metadataScreen = (
      <UserMetadata
        document={selectedDocForMetadata}
        currentUser={currentUser.displayName}
        context={context}
        onBack={() => setSelectedDocForMetadata(null)}
        onSave={(updatedFields) => {
          setDocuments(prev => prev.map(d => {
            if (d.id === selectedDocForMetadata.id || d.serverRelativeUrl === selectedDocForMetadata.serverRelativeUrl) {
              const updatedDoc = {
                ...d,
                docType: updatedFields.documentType,
                status: updatedFields.documentStatus,
                owner: updatedFields.owner,
                tags: updatedFields.tags || d.tags,
                modifiedBy: currentUser.displayName,
                version: updatedFields.version,
                modified: new Date().toLocaleDateString(),
                docNumberIsSaved: true
              };

              setSelectedDocForMetadata(updatedDoc);
              return updatedDoc;
            }
            return d;
          }));
        }}
      />
    );

    // Viewer: same screen, but nothing on it can be clicked or edited — only the Back bar below works.
    if (isViewer) {
      return (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '10px 16px', backgroundColor: '#fef3c7', borderBottom: '1px solid #fde68a', fontSize: '12px', color: '#92400e', fontWeight: 600 }}>
            <button onClick={() => setSelectedDocForMetadata(null)} style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', background: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '4px 10px', fontSize: '12px', fontWeight: 700, color: '#005a43', cursor: 'pointer' }}>
              <SVGIcon name="chevronLeft" color="#005a43" size={12} />
              <span>Back</span>
            </button>
            <span>Read-only — you have Viewer access to this process, so you cannot edit this document.</span>
          </div>
          <div style={{ pointerEvents: 'none', userSelect: 'none' }} {...({ inert: '' } as any)}>
            {metadataScreen}
          </div>
        </div>
      );
    }

    return metadataScreen;
  }

  const totalItems = filteredDocuments.length;
  const totalPages = Math.ceil(totalItems / itemsPerPage) || 1;
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = Math.min(startIndex + itemsPerPage, totalItems);
  const currentPagedDocs = filteredDocuments.slice(startIndex, endIndex);

  // STATUS column now shows the library's "DocumentCurrentStatus" choice column
  // (Draft, Send for L1, L1 Approved, Send for L2, L2 Approved, Send for L3, L3 Approved).
  const renderStatusBadge = (status?: string) => {
    switch (status) {
      case 'Draft':
        return <span style={styles.statusDraft}><span style={{ ...styles.dot, backgroundColor: '#94a3b8' }}></span>Draft</span>;
      case 'Send for L1':
      case 'Send for L2':
      case 'Send for L3':
        return <span style={styles.statusPending}><span style={{ ...styles.dot, backgroundColor: '#b45309' }}></span>{status}</span>;
      case 'L1 Approved':
      case 'L2 Approved':
      case 'L3 Approved':
        return <span style={styles.statusActive}><span style={{ ...styles.dot, backgroundColor: '#059669' }}></span>{status}</span>;
      case 'L1 Rejected':
      case 'L2 Rejected':
      case 'L3 Rejected':
        return <span style={{ ...styles.statusDraft, backgroundColor: '#fee2e2', color: '#b91c1c' }}><span style={{ ...styles.dot, backgroundColor: '#dc2626' }}></span>{status}</span>;
      case undefined:
      case '':
        return <span style={{ color: '#94a3b8' }}>—</span>;
      default:
        return <span>{status}</span>;
    }
  };

  return (
    <div style={styles.pageContainer}>
      <input type="file" ref={fileInputRef} onChange={handleFileUpload} multiple style={{ display: 'none' }} />

      {/* --- PROCESS NAVIGATION SIDEBAR --- */}
      <div style={styles.processSidebar}>
        <div style={styles.processSidebarHeader}>Process Navigation</div>

        <div style={styles.processNavList}>
          {loadingProcessNav ? (
            <div style={styles.processNavEmptyText}>Loading...</div>
          ) : processNavItems.length === 0 ? (
            <div style={styles.processNavEmptyText}>No processes found</div>
          ) : (
            processNavItems.map((item) => {
              const isSelected = selectedFolderCode === item.processCode;
              return (
                <div
                  key={item.id}
                  style={{ ...styles.processNavItem, ...(isSelected ? styles.processNavItemActive : {}) }}
                  onClick={() => { handleSelectFolder(item.processCode); if (onSelectProcess) onSelectProcess(item); }}
                >
                  <SVGIcon name="flag" color={isSelected ? '#ffffff' : '#005a43'} size={13} />
                  <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    <span style={{ ...styles.processNavCode, color: isSelected ? '#ffffff' : styles.processNavCode.color }}>{item.processCode}</span>
                    <span style={{ ...styles.processNavName, color: isSelected ? '#e6f4f1' : styles.processNavName.color }} title={item.processName}>{item.processName}</span>
                  </div>
                </div>
              );
            })
          )}
        </div>

        <div
          style={styles.viewAllProcessesLink}
          onClick={() => { handleGoToProcessMaster(); if (onViewAllProcesses) onViewAllProcesses(); }}
        >
          <span>View all processes</span>
          <SVGIcon name="arrowRight" color="#005a43" size={12} />
        </div>
      </div>

      {/* --- MAIN CONTENT AREA --- */}
      <div style={{ ...styles.contentWrapper, paddingTop: 0 }}>
        {toastMessage && (
          <div style={{ ...styles.toastNotification, backgroundColor: toastMessage.type === 'success' ? '#005a43' : '#ef4444' }}>
            <span>{toastMessage.text}</span>
            <button onClick={() => setToastMessage(null)} style={styles.toastCloseBtn}>×</button>
          </div>
        )}

        <div style={styles.searchActionRow}>
          <div style={styles.searchBox}>
            <SVGIcon name="search" color="#94a3b8" size={14} />
            <input type="text" placeholder="Search by title, number, type, tags..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} style={styles.searchInput} />
          </div>
          <button
            style={isViewer ? { ...styles.createBtn, opacity: 0.5, cursor: 'not-allowed' } : styles.createBtn}
            title={isViewer ? 'You have Viewer access to this process — only Contributors can upload documents.' : undefined}
            onClick={() => {
              if (isViewer) return;
              if (!selectedFolderCode) {
                setToastMessage({ type: 'error', text: 'Please select a process folder first to upload documents.' });
                return;
              }
              setIsUploadModalOpen(true);
            }}
            disabled={isUploading || isViewer}
          >
            <SVGIcon name="plus" color="#fff" size={14} />
            <span>{isUploading ? 'Uploading...' : 'Upload documents'}</span>
          </button>
        </div>

        <div style={styles.tableCard}>
          <div style={styles.tableHeaderRow}>
            <div style={{ flex: '4' }}>DOCUMENT</div>
            <div style={{ flex: '1.5', textAlign: 'center' }}>DOC NUMBER</div>
            <div style={{ flex: '1.8' }}>TYPE</div>
            <div style={{ flex: '1' }}>VERSION</div>
            <div style={{ flex: '1.5', textAlign: 'left' }}>STATUS</div>
            <div style={{ width: '32px' }}></div>
          </div>

          <div style={styles.tableBody}>
            {loading ? (
              <div style={{ padding: '20px', textAlign: 'center', color: '#64748b', fontSize: '12px' }}>Loading documents...</div>
            ) : currentPagedDocs.length === 0 ? (
              <div style={{ padding: '20px', textAlign: 'center', color: '#64748b', fontSize: '12px' }}>No documents found</div>
            ) : (
              currentPagedDocs.map((doc, index) => {
                const isNearBottom = index >= currentPagedDocs.length - 2 && currentPagedDocs.length > 2;

                return (
                  <div key={doc.id} style={styles.tableRow}>
                    <div style={{ flex: '4', display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <FileTypeBadge type={doc.fileExtension} />
                      <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                        <span onClick={() => handleOpenDocumentInNewTab(doc.serverRelativeUrl)} style={styles.docTitle} title={doc.title}>
                          {doc.title}
                        </span>
                        <span style={styles.docSubtitle}>{doc.subTitle}</span>
                      </div>
                    </div>

                    <div style={{ flex: '1.5', display: 'flex', justifyContent: 'center' }}>
                      <span style={styles.docNumberPill}>{doc.docNumber}</span>
                    </div>

                    <div style={{ flex: '1.8', fontSize: '12px', color: '#475569', fontWeight: 500 }}>
                      {doc.docType}
                    </div>

                    <div style={{ flex: '1', fontSize: '12px', color: '#475569', fontWeight: 500 }}>
                      {doc.version}
                    </div>

                    <div style={{ flex: '1.5', display: 'flex', alignItems: 'center', justifyContent: 'flex-start' }}>
                      {renderStatusBadge(doc.currentStatus)}
                    </div>

                    <div style={{ width: '32px', position: 'relative' }} ref={openDropdownId === doc.id ? actionMenuRef : null}>
                      <button onClick={() => setOpenDropdownId(openDropdownId === doc.id ? null : doc.id)} style={styles.actionIconButton}>
                        <SVGIcon name="moreVertical" color="#94a3b8" size={15} />
                      </button>

                      {openDropdownId === doc.id && (
                        <div
                          style={{
                            ...styles.actionMenu,
                            ...(isNearBottom
                              ? { bottom: '100%', top: 'auto', marginBottom: '4px' }
                              : { top: '100%', bottom: 'auto', marginTop: '4px' })
                          }}
                        >
                          <button onClick={() => handleOpenMetadata(doc)} style={{ ...styles.actionMenuItem, fontWeight: 700, color: '#005a43' }}>Document Details</button>
                          {!doc.workflowId && (
                            <button onClick={() => handleOpenAssignWorkflow(doc)} style={{ ...styles.actionMenuItem, fontWeight: 700, color: '#005a43' }}>Assign Workflow</button>
                          )}
                          <button onClick={() => handleOpenGuestAccess(doc)} style={{ ...styles.actionMenuItem, fontWeight: 700, color: '#005a43' }}>Guest Access</button>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })
            )}
          </div>

          {!loading && totalItems > 0 && (
            <div style={styles.paginationContainer}>
              <div style={styles.paginationInfo}>
                Showing {totalItems === 0 ? 0 : startIndex + 1} to {endIndex} of {totalItems} items
              </div>

              <div style={styles.paginationControls}>
                <button
                  onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                  disabled={currentPage === 1}
                  style={{ ...styles.pageNavBtn, opacity: currentPage === 1 ? 0.4 : 1, cursor: currentPage === 1 ? 'not-allowed' : 'pointer' }}
                >
                  <SVGIcon name="chevronLeft" color="#475569" size={14} />
                </button>

                {Array.from({ length: totalPages }, (_, i) => i + 1).map((pageNum) => {
                  const isSelected = pageNum === currentPage;
                  return (
                    <button
                      key={pageNum}
                      onClick={() => setCurrentPage(pageNum)}
                      style={{
                        ...styles.pageNumberBtn,
                        backgroundColor: isSelected ? '#005a43' : '#ffffff',
                        color: isSelected ? '#ffffff' : '#334155',
                        borderColor: isSelected ? '#005a43' : '#e2e8f0',
                        fontWeight: isSelected ? 700 : 500
                      }}
                    >
                      {pageNum}
                    </button>
                  );
                })}

                <button
                  onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                  disabled={currentPage === totalPages}
                  style={{ ...styles.pageNavBtn, opacity: currentPage === totalPages ? 0.4 : 1, cursor: currentPage === totalPages ? 'not-allowed' : 'pointer' }}
                >
                  <SVGIcon name="chevronRight" color="#475569" size={14} />
                </button>
              </div>

              <div style={styles.itemsPerPageWrapper} ref={itemsPerPageRef}>
                <span style={styles.itemsPerPageLabel}>Items per page:</span>
                <div style={{ position: 'relative' }}>
                  <div
                    onClick={() => setIsItemsPerPageOpen(!isItemsPerPageOpen)}
                    style={styles.customSelectBox}
                  >
                    <span>{itemsPerPage}</span>
                    <SVGIcon name="chevronDown" color="#64748b" size={12} />
                  </div>

                  {isItemsPerPageOpen && (
                    <div style={styles.customSelectMenu}>
                      {[5, 10, 20, 50].map((num) => {
                        const isSelected = itemsPerPage === num;
                        return (
                          <div
                            key={num}
                            onClick={() => {
                              setItemsPerPage(num);
                              setCurrentPage(1);
                              setIsItemsPerPageOpen(false);
                            }}
                            style={{
                              ...styles.customSelectOption,
                              backgroundColor: isSelected ? '#005a43' : '#ffffff',
                              color: isSelected ? '#ffffff' : '#334155',
                              fontWeight: isSelected ? 700 : 500
                            }}
                          >
                            {num}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      <UploadDocumentModal
        isOpen={isUploadModalOpen}
        onClose={() => setIsUploadModalOpen(false)}
        context={context}
        selectedFolderCode={selectedFolderCode}
        currentUserName={currentUser.displayName}
        isUploading={isUploading}
        onUpload={handleModalUpload}
        onCreateWorkflow={onNavigateToApprovals}
      />

      <AssignWorkflow
        isOpen={!!assignTarget}
        context={context}
        documentId={assignTarget ? assignTarget.id : 0}
        documentTitle={assignTarget ? assignTarget.title : ''}
        processCode={assignTarget ? getDocProcessCode(assignTarget) : ''}
        onClose={() => setAssignTarget(null)}
        onAssigned={(workflowTitle) => {
          setAssignTarget(null);
          setToastMessage({ type: 'success', text: `Workflow "${workflowTitle}" assigned successfully!` });
          fetchSharePointDocuments(undefined, selectedFolderCode);
        }}
      />

      <GuestAccess
        isOpen={!!guestAccessTarget}
        context={context}
        documentId={guestAccessTarget ? guestAccessTarget.id : 0}
        documentTitle={guestAccessTarget ? guestAccessTarget.title : ''}
        documentNumber={guestAccessTarget ? guestAccessTarget.docNumber : ''}
        onSaved={() => setToastMessage({ type: 'success', text: 'Guest access information saved successfully.' })}
        onClose={() => setGuestAccessTarget(null)}
      />
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  pageContainer: {
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
    backgroundColor: '#f8fafc',
    minHeight: '100vh',
    display: 'flex',
    boxSizing: 'border-box'
  },
  
  contentWrapper: {
    flexGrow: 1,
    padding: '16px 20px',
    overflowX: 'hidden'
  },

  processSidebar: {
    width: '200px',
    flexShrink: 0,
    alignSelf: 'flex-start',
    height: '560px',
    backgroundColor: '#ffffff',
    border: '1px solid #e2e8f0',
    borderRadius: '6px',
    margin: '0 0 16px 0',
    padding: '16px 12px',
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    overflow: 'hidden'
  },
  processSidebarHeader: { fontSize: '12px', fontWeight: 700, color: '#334155', marginBottom: '10px', flexShrink: 0 },
  createProcessBtn: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '6px',
    width: '100%',
    padding: '7px 10px',
    backgroundColor: '#ffffff',
    border: '1px solid #005a43',
    borderRadius: '5px',
    color: '#005a43',
    fontSize: '11px',
    fontWeight: 600,
    cursor: 'pointer',
    marginBottom: '12px',
    boxSizing: 'border-box',
    flexShrink: 0
  },
  processNavList: { display: 'flex', flexDirection: 'column', gap: '4px', flexGrow: 1, overflowY: 'auto', minHeight: 0 },
  processNavItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '6px 6px',
    borderRadius: '5px',
    cursor: 'pointer'
  },
  processNavItemActive: {
    backgroundColor: '#005a43'
  },
  processNavCode: { fontSize: '10px', fontWeight: 700, color: '#334155', lineHeight: 1.3 },
  processNavName: { fontSize: '10px', color: '#64748b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '140px' },
  processNavEmptyText: { fontSize: '11px', color: '#94a3b8', padding: '6px' },
  viewAllProcessesLink: {
    display: 'flex',
    alignItems: 'center',
    gap: '5px',
    fontSize: '11px',
    fontWeight: 600,
    color: '#005a43',
    cursor: 'pointer',
    marginTop: '10px',
    paddingTop: '10px',
    borderTop: '1px solid #f1f5f9',
    flexShrink: 0
  },
  backToDocsBtn: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    background: 'none',
    border: 'none',
    color: '#005a43',
    fontSize: '12px',
    fontWeight: 600,
    cursor: 'pointer',
    padding: '4px 0',
    marginBottom: '10px'
  },

  searchActionRow: { display: 'flex', gap: '10px', marginBottom: '14px', alignItems: 'center' },
  folderContextBar: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    backgroundColor: '#e6f4f1',
    border: '1px solid #ccece6',
    borderRadius: '5px',
    padding: '6px 12px',
    fontSize: '11px',
    color: '#0d6b52',
    marginBottom: '14px'
  },
  clearFolderLink: { fontWeight: 700, cursor: 'pointer', textDecoration: 'underline', color: '#005a43' },
  searchBox: { flexGrow: 1, display: 'flex', alignItems: 'center', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '5px', padding: '6px 10px', gap: '8px' },
  searchInput: { border: 'none', outline: 'none', width: '100%', fontSize: '12px', color: '#334155' },
  createBtn: { backgroundColor: '#005a43', color: '#ffffff', border: 'none', borderRadius: '5px', padding: '7px 14px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px', whiteSpace: 'nowrap' },
  
  tableCard: { backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '6px', overflow: 'visible', position: 'relative' },
  tableHeaderRow: { display: 'flex', alignItems: 'center', padding: '8px 12px', borderBottom: '1px solid #e2e8f0', fontSize: '10px', fontWeight: 700, color: '#94a3b8', letterSpacing: '0.4px', backgroundColor: '#ffffff', borderTopLeftRadius: '6px', borderTopRightRadius: '6px' },
  tableBody: { display: 'flex', flexDirection: 'column' },
  tableRow: { display: 'flex', alignItems: 'center', padding: '8px 12px', borderBottom: '1px solid #f1f5f9', backgroundColor: '#ffffff' },
  
  docTitle: { fontSize: '12px', fontWeight: 700, color: '#1e293b', textDecoration: 'none', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', cursor: 'pointer' },
  docSubtitle: { fontSize: '10px', fontWeight: 600, color: '#94a3b8', marginTop: '1px' },
  docNumberPill: { backgroundColor: '#e6f4f1', color: '#0d9488', border: '1px solid #ccece6', borderRadius: '4px', padding: '2px 6px', fontSize: '10px', fontWeight: 700 },
  
  statusNotInDms: { backgroundColor: '#ffedd5', color: '#c2410c', borderRadius: '12px', padding: '3px 10px', fontSize: '11px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap' },
  statusDraft: { backgroundColor: '#f1f5f9', color: '#475569', borderRadius: '12px', padding: '3px 10px', fontSize: '11px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap' },
  statusActive: { backgroundColor: '#dcfce7', color: '#15803d', borderRadius: '12px', padding: '3px 10px', fontSize: '11px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap' },
  statusPending: { backgroundColor: '#fef3c7', color: '#b45309', borderRadius: '12px', padding: '3px 10px', fontSize: '11px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap' },
  dot: { width: '4px', height: '4px', borderRadius: '50%', marginRight: '5px' },
  
  actionIconButton: { background: 'none', border: 'none', cursor: 'pointer', padding: '4px 6px', borderRadius: '4px' },
  actionMenu: { position: 'absolute', right: '0px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', boxShadow: '0 8px 20px rgba(0,0,0,0.12)', zIndex: 999, width: '160px', overflow: 'hidden', padding: '4px 0' },
  actionMenuItem: { width: '100%', padding: '7px 12px', border: 'none', backgroundColor: 'transparent', textAlign: 'left', fontSize: '11px', color: '#334155', cursor: 'pointer' },
  
  paginationContainer: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 14px', borderTop: '1px solid #e2e8f0', backgroundColor: '#ffffff', borderBottomLeftRadius: '6px', borderBottomRightRadius: '6px' },
  paginationInfo: { fontSize: '12px', color: '#64748b', fontWeight: 500 },
  paginationControls: { display: 'flex', alignItems: 'center', gap: '6px' },
  pageNavBtn: { width: '28px', height: '28px', border: '1px solid #e2e8f0', backgroundColor: '#ffffff', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'center' },
  pageNumberBtn: { width: '28px', height: '28px', border: '1px solid', borderRadius: '6px', fontSize: '11px', cursor: 'pointer', transition: 'all 0.15s ease' },
  
  itemsPerPageWrapper: { display: 'flex', alignItems: 'center', gap: '8px' },
  itemsPerPageLabel: { fontSize: '12px', color: '#64748b' },
  customSelectBox: { display: 'flex', alignItems: 'center', gap: '8px', padding: '4px 10px', border: '1px solid #cbd5e1', borderRadius: '6px', backgroundColor: '#ffffff', fontSize: '12px', color: '#334155', cursor: 'pointer', minWidth: '45px', justifyContent: 'space-between' },
  customSelectMenu: { position: 'absolute', bottom: '110%', right: 0, backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', boxShadow: '0 -4px 16px rgba(0,0,0,0.1)', zIndex: 999, overflow: 'hidden', minWidth: '65px' },
  customSelectOption: { padding: '6px 12px', fontSize: '12px', cursor: 'pointer', textAlign: 'center', transition: 'all 0.1s ease' },

  toastNotification: { position: 'fixed', top: '14px', right: '14px', backgroundColor: '#005a43', color: '#ffffff', padding: '8px 14px', borderRadius: '5px', fontSize: '11px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '10px', zIndex: 9999 },
  toastCloseBtn: { background: 'none', border: 'none', color: '#ffffff', fontSize: '12px', cursor: 'pointer' },
  
  modalOverlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 },
  modalContainer: { backgroundColor: '#ffffff', borderRadius: '12px', width: '440px', maxWidth: '92vw', boxShadow: '0 10px 25px rgba(0,0,0,0.15)', overflow: 'hidden', display: 'flex', flexDirection: 'column' },
  modalHeader: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', padding: '16px 18px 8px 18px' },
  modalTitle: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  modalSubtitle: { margin: '2px 0 0 0', fontSize: '11px', color: '#64748b' },
  modalCloseBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '2px', borderRadius: '4px' },
  modalBody: { padding: '0 18px 14px 18px', display: 'flex', flexDirection: 'column' },
  selectedDocCard: { backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', padding: '8px 12px' },
  selectedDocLabel: { fontSize: '9px', fontWeight: 700, color: '#64748b', letterSpacing: '0.4px' },
  selectedDocInfo: { display: 'flex', alignItems: 'center', gap: '6px', marginTop: '3px' },
  selectedDocName: { fontSize: '12px', fontWeight: 700, color: '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  fieldLabel: { fontSize: '11px', fontWeight: 600, color: '#475569', marginBottom: '4px', display: 'block' },
  avatarCircleSmall: { width: '24px', height: '24px', borderRadius: '50%', backgroundColor: '#e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'center', overflow: 'hidden', flexShrink: 0 },
  avatarImg: { width: '100%', height: '100%', objectFit: 'cover' },
  userNameText: { fontSize: '11px', fontWeight: 600, color: '#1e293b', lineHeight: 1.1 },
  userEmailText: { fontSize: '9px', color: '#64748b' },
  comboboxControl: { padding: '6px 10px', border: '1px solid #cbd5e1', borderRadius: '6px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer', backgroundColor: '#ffffff' },
  
  comboboxDropdown: { position: 'relative', width: '100%', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', boxShadow: '0 4px 12px rgba(0,0,0,0.08)', zIndex: 100, overflow: 'hidden', marginTop: '4px' },
  dropdownSearchContainer: { display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 10px', backgroundColor: '#f8fafc', borderBottom: '1px solid #e2e8f0' },
  comboboxSearchInput: { width: '100%', border: 'none', backgroundColor: 'transparent', outline: 'none', fontSize: '11px', color: '#334155' },
  userListScrollable: { maxHeight: '130px', overflowY: 'auto' },
  comboboxOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 10px', cursor: 'pointer', borderBottom: '1px solid #f1f5f9' },
  
  modalFooter: { display: 'flex', justifyContent: 'flex-end', gap: '8px', padding: '12px 18px', backgroundColor: '#ffffff', borderTop: '1px solid #f1f5f9' },
  cancelBtn: { padding: '7px 14px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', borderRadius: '6px', fontSize: '11px', fontWeight: 600, color: '#334155', cursor: 'pointer' },
  sendTeamsBtn: { padding: '7px 16px', border: 'none', backgroundColor: '#005a43', borderRadius: '6px', fontSize: '11px', fontWeight: 600, color: '#ffffff', cursor: 'pointer' }
};

export default UserDashboard;

// Document list