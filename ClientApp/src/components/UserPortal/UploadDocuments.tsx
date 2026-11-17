import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse, photoUrlFor } from '../../spfx-shim';
import { DocumentTypeField } from './DocumentTypeDropdownlist';

export interface IRaciEntry {
  role: 'Responsible' | 'Accountable' | 'Consulted' | 'Informed';
  name: string;
  names?: string[];
  subtitle: string;
}

export type ApprovalStepStatus = 'Completed' | 'Pending' | 'Rejected';

export interface IApprovalStep {
  id: string;
  title: string;
  subtitle: string;
  status: ApprovalStepStatus;
  timestamp?: string;
  approverName?: string;
}

export interface IWorkflowConfigEntry {
  id: number;
  /** Auto-generated datetime id (yyyyMMddHHmmss) stamped by CreateWorkflow.tsx in the "WorkflowID" column. */
  workflowId?: string;
  workflowTitle: string;
  workflowType: string;
  escalationCriteria: string;
  stepRole: string;
  l1Approver: string;
  l2Approver: string;
  escalateTo: string;
  finalApprover: string;
  complianceCheck: string;
  deadlineDays: number;
}

export interface IUploadMetadata {
  documentType: string;
  documentStatus: string;
  tags: string[];
  /**
   * Internal column "DocumentCurrentStatus" (Choice: Draft, Send for L1, L1 Approved,
   * Send for L2, L2 Approved, Send for L3, L3 Approved) — separate from documentStatus
   * above. Always sent as 'Send for L1' on every upload (not user-editable here) so the
   * Power Automate flow watching this column for 'Send for L1' fires and notifies
   * whoever is set as L1 Approver in that process's Workflow Configuration record.
   */
  documentCurrentStatus: string;
  /** "WorkflowID" of the workflow picked in "Select Workflow" — written to the library's WorkflowID column. */
  workflowId?: string;
}

// local "uploading" progress indicator before it's confirmed into the real files list.
export interface IStagingFile {
  id: string;
  file: File;
  progress: number;
  status: 'uploading' | 'completed';
}

export interface IUploadDocumentModalProps {
  isOpen: boolean;
  onClose: () => void;
  context: any;
  selectedFolderCode?: string | null;
  processName?: string;
  currentUserName: string;
  isUploading?: boolean;
  onUpload: (files: File[], metadata: IUploadMetadata) => void;
  onCreateWorkflow?: () => void;
}

const ACCEPTED_EXTENSIONS = ['pdf', 'docx', 'xlsx', 'pptx', 'tiff', 'tif', 'doc', 'xls', 'ppt'];
const MAX_FILE_SIZE_MB = 100;

const STATUS_DOT_COLORS: { [key: string]: string } = {
  'Active': '#15803d',
  'Pending Approval': '#b45309',
  'Draft': '#64748b',
  'In Review': '#2563eb',
  'Archived': '#64748b'
};

const RACI_LIST_TITLE = 'RACI Matrix';
const RACI_PROCESS_ID_FIELD = 'ProcessID';
const WORKFLOW_CONFIG_LIST_TITLE = 'Workflow Configuration';

const Icon = ({ name, color = 'currentColor', size = 14 }: { name: string; color?: string; size?: number }) => {
  const paths: { [key: string]: React.ReactNode } = {
    cloud: <path d="M7 18a4 4 0 01-.6-7.96A5 5 0 0116.9 8.05 3.5 3.5 0 0117.5 15H7zM12 12v6m0-6l-2.5 2.5M12 12l2.5 2.5" />,
    close: <path d="M6 18L18 6M6 6l12 12" />,
    check: <path d="M20 6L9 17l-5-5" />,
    refresh: <path d="M23 4v6h-6M1 20v-6h6M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15" />,
    send: <path d="M22 2L11 13M22 2l-7 20-4-9-9-4 20-7z" />,
    chevronRight: <path d="M9 18l6-6-6-6" />,
    chevronDown: <path d="M6 9l6 6 6-6" />,
    user: <><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" /><circle cx="12" cy="7" r="4" /></>,
    edit: <><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" /></>,
    file: <><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" /><path d="M14 2v6h6" /></>,
    plus: <><path d="M12 5v14" /><path d="M5 12h14" /></>,
    arrowRight: <><path d="M5 12h14" /><path d="M12 5l7 7-7 7" /></>
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }}>
      {paths[name] || null}
    </svg>
  );
};

const formatBytes = (bytes: number): string => {
  if (!bytes) return '0 KB';
  const mb = bytes / (1024 * 1024);
  if (mb >= 1) return `${mb.toFixed(1)} MB`;
  return `${(bytes / 1024).toFixed(0)} KB`;
};

const PersonAvatar = ({ name, photoUrl, style }: { name: string; photoUrl?: string; style: React.CSSProperties }) => {
  const [imgFailed, setImgFailed] = useState(false);
  if (photoUrl && !imgFailed) {
    return (
      <img
        src={photoUrl}
        alt={name}
        onError={() => setImgFailed(true)}
        style={{ width: style.width, height: style.height, borderRadius: '50%', objectFit: 'cover', flexShrink: 0, backgroundColor: '#e2e8f0' }}
      />
    );
  }
  return <div style={style}>{name.charAt(0).toUpperCase()}</div>;
};

const emptyMetadata: IUploadMetadata = {
  documentType: '',
  documentStatus: 'Draft',
  tags: [],
  documentCurrentStatus: 'Send for L1'
};

interface CompactDropdownOption {
  id: number;
  label: string;
}

interface CompactDropdownProps {
  value: number | null;
  options: CompactDropdownOption[];
  onChange: (id: number | null) => void;
}

const CompactDropdown: React.FC<CompactDropdownProps> = ({ value, options, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const selectedOption = options.find(o => o.id === value);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div style={styles.headerDdWrapper} ref={dropdownRef}>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          ...styles.headerDdControl,
          borderColor: isOpen ? '#005a43' : '#cbd5e1',
          boxShadow: isOpen ? '0 0 0 2px rgba(0, 90, 67, 0.15)' : 'none'
        }}
      >
        <span style={styles.ddValueText}>{selectedOption?.label ?? 'None'}</span>
        <span style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', display: 'flex' }}>
          <Icon name="chevronDown" color="#64748b" size={12} />
        </span>
      </div>

      {isOpen && (
        <div style={styles.headerDdMenu}>
          <div
            onClick={() => { onChange(null); setIsOpen(false); }}
            style={{
              ...styles.ddOption,
              backgroundColor: value === null ? '#f0fdf4' : '#ffffff',
              color: value === null ? '#005a43' : '#334155',
              fontWeight: value === null ? 600 : 500
            }}
          >
            <span>None</span>
            {value === null && <Icon name="check" color="#005a43" size={14} />}
          </div>
          {options.map((opt) => {
            const isSelected = opt.id === value;
            return (
              <div
                key={opt.id}
                onClick={() => { onChange(opt.id); setIsOpen(false); }}
                style={{
                  ...styles.ddOption,
                  backgroundColor: isSelected ? '#f0fdf4' : '#ffffff',
                  color: isSelected ? '#005a43' : '#334155',
                  fontWeight: isSelected ? 600 : 500
                }}
              >
                <span>{opt.label}</span>
                {isSelected && <Icon name="check" color="#005a43" size={14} />}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

  // COMPONENT

export const UploadDocumentModal: React.FC<IUploadDocumentModalProps> = ({
  isOpen,
  onClose,
  context,
  selectedFolderCode,
  processName,
  currentUserName,
  isUploading = false,
  onUpload,
  onCreateWorkflow
}) => {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [fileError, setFileError] = useState<string>('');

  // Files currently going through the "Browse Files" progress indicator, before
  // they're attached into selectedFiles below.
  const [stagingFiles, setStagingFiles] = useState<IStagingFile[]>([]);
  const stagingIntervalsRef = useRef<{ [id: string]: ReturnType<typeof setInterval> }>({});

  const [documentType, setDocumentType] = useState<string>(emptyMetadata.documentType);
  const [documentStatus, setDocumentStatus] = useState<string>(emptyMetadata.documentStatus);

  const [tags, setTags] = useState<string[]>([]);
  const [newTagInput, setNewTagInput] = useState<string>('');
  const [isAddingTag, setIsAddingTag] = useState<boolean>(false);

  const [raci, setRaci] = useState<IRaciEntry[]>([
    { role: 'Responsible', name: currentUserName, names: [currentUserName], subtitle: 'Document Owner' },
    { role: 'Accountable', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Department Head' },
    { role: 'Consulted', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Legal / Compliance' },
    { role: 'Informed', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Stakeholders' }
  ]);
  const [isRaciLoading, setIsRaciLoading] = useState<boolean>(false);
  const [personPhotos, setPersonPhotos] = useState<{ [name: string]: string }>({});
  const [workflows, setWorkflows] = useState<IWorkflowConfigEntry[]>([]);
  const [isWorkflowLoading, setIsWorkflowLoading] = useState<boolean>(false);
  const [selectedWorkflowId, setSelectedWorkflowId] = useState<number | null>(null);

  const defaultRaci = (): IRaciEntry[] => ([
    { role: 'Responsible', name: currentUserName, names: [currentUserName], subtitle: 'Document Owner' },
    { role: 'Accountable', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Department Head' },
    { role: 'Consulted', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Legal / Compliance' },
    { role: 'Informed', name: 'Not Assigned', names: ['Not Assigned'], subtitle: 'Stakeholders' }
  ]);

  const resolvePersonPhotos = async (names: string[]) => {
    if (!context?.spHttpClient || !context?.pageContext) return;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const toResolve = Array.from(new Set(names)).filter(
      n => n && n !== 'Not Assigned' && !personPhotos[n]
    );
    if (!toResolve.length) return;

    await Promise.all(
      toResolve.map(async (personName) => {
        try {
          const safeName = personName.replace(/'/g, "''");
          const endpoint = `${siteUrl}/_api/web/siteusers?$filter=Title eq '${safeName}'&$select=Email,LoginName&$top=1`;
          const response: SPHttpClientResponse = await context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1,
            { headers: { 'Accept': 'application/json;odata=nometadata' } }
          );
          if (!response.ok) return;
          const data = await response.json();
          const user = data?.value?.[0];
          const loginName: string = user?.LoginName || '';
          const upn: string = user?.Email || (loginName.includes('|') ? loginName.split('|').pop() : loginName) || '';
          if (!upn) return;
          const photoUrl = photoUrlFor(upn);
          setPersonPhotos(prev => ({ ...prev, [personName]: photoUrl }));
        } catch (err) {
          console.warn('Could not resolve M365 photo for', personName, err);
        }
      })
    );
  };

  const fetchRaciForProcess = async (processCode: string) => {
    if (!context?.spHttpClient || !context?.pageContext || !processCode.trim()) {
      setRaci(defaultRaci());
      resolvePersonPhotos([currentUserName]);
      return;
    }

    setIsRaciLoading(true);
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const safeCode = processCode.trim().replace(/'/g, "''");
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${RACI_LIST_TITLE}')/items?$filter=${RACI_PROCESS_ID_FIELD} eq '${safeCode}'&$top=1`;

      const response: SPHttpClientResponse = await context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.warn('RACI Matrix fetch warning for', processCode, errText);
        setRaci(defaultRaci());
        resolvePersonPhotos([currentUserName]);
        return;
      }

      const data = await response.json();
      const item = data?.value?.[0];

      const splitNames = (val: any): string[] =>
        val ? String(val).split(';').map((s: string) => s.trim()).filter(Boolean) : [];

      const responsibleNames = splitNames(item?.Responsible);
      const accountableNames = splitNames(item?.Accountable);
      const consultedNames = splitNames(item?.Consulted);
      const informedNames = splitNames(item?.Informed);

      setRaci([
        { role: 'Responsible', name: responsibleNames.length ? responsibleNames.join(', ') : 'Not Assigned', names: responsibleNames.length ? responsibleNames : ['Not Assigned'], subtitle: 'Document Owner' },
        { role: 'Accountable', name: accountableNames.length ? accountableNames.join(', ') : 'Not Assigned', names: accountableNames.length ? accountableNames : ['Not Assigned'], subtitle: 'Department Head' },
        { role: 'Consulted', name: consultedNames.length ? consultedNames.join(', ') : 'Not Assigned', names: consultedNames.length ? consultedNames : ['Not Assigned'], subtitle: 'Legal / Compliance' },
        { role: 'Informed', name: informedNames.length ? informedNames.join(', ') : 'Not Assigned', names: informedNames.length ? informedNames : ['Not Assigned'], subtitle: 'Stakeholders' }
      ]);

      resolvePersonPhotos([...responsibleNames, ...accountableNames, ...consultedNames, ...informedNames]);
    } catch (err) {
      console.error('Error fetching RACI Matrix for process', processCode, err);
      setRaci(defaultRaci());
      resolvePersonPhotos([currentUserName]);
    } finally {
      setIsRaciLoading(false);
    }
  };

  const fetchWorkflowsForProcess = async (processCode: string) => {
    if (!context?.spHttpClient || !context?.pageContext || !processCode.trim()) {
      setWorkflows([]);
      setSelectedWorkflowId(null);
      return;
    }

    setIsWorkflowLoading(true);
    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      const safeCode = processCode.trim().replace(/'/g, "''");
      const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${WORKFLOW_CONFIG_LIST_TITLE}')/items`
        + `?$filter=Title eq '${safeCode}'`
        + `&$select=Id,Title,WorkflowID,WorkflowTitle,WorkflowType,EscalationCriteria,StepRole,ComplianceCheck,ApprovalDeadlineDays,L1Approver/Title,L2Approver/Title,EscalateTo/Title,FinalApprover/Title`
        + `&$expand=L1Approver,L2Approver,EscalateTo,FinalApprover&$top=50`;

      const response: SPHttpClientResponse = await context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        const errText = await response.text();
        console.warn('Workflow Configuration fetch warning for', processCode, errText);
        setWorkflows([]);
        setSelectedWorkflowId(null);
        return;
      }

      const data = await response.json();
      const nameOf = (val: any): string => (val && typeof val === 'object') ? (val.Title || '') : '';

      const mapped: IWorkflowConfigEntry[] = (data.value || []).map((item: any) => ({
        id: item.Id,
        workflowId: item.WorkflowID || '',
        workflowTitle: item.WorkflowTitle || item.WorkflowType || 'Untitled Workflow',
        workflowType: item.WorkflowType || 'Approval Workflow',
        escalationCriteria: item.EscalationCriteria || '',
        stepRole: item.StepRole || '',
        l1Approver: nameOf(item.L1Approver),
        l2Approver: nameOf(item.L2Approver),
        escalateTo: nameOf(item.EscalateTo),
        finalApprover: nameOf(item.FinalApprover),
        complianceCheck: item.ComplianceCheck || '',
        deadlineDays: item.ApprovalDeadlineDays || 0
      }));

      setWorkflows(mapped);
      setSelectedWorkflowId(null);

      const approverNames = mapped
        .reduce((acc: string[], w) => acc.concat([w.l1Approver, w.l2Approver, w.escalateTo, w.finalApprover]), [])
        .filter(Boolean);
      if (approverNames.length) resolvePersonPhotos(approverNames);
    } catch (err) {
      console.error('Error fetching Workflow Configuration for process', processCode, err);
      setWorkflows([]);
      setSelectedWorkflowId(null);
    } finally {
      setIsWorkflowLoading(false);
    }
  };

  const activeWorkflow = workflows.find(w => w.id === selectedWorkflowId) || null;
  const approvalSteps: IApprovalStep[] = activeWorkflow ? [
    { id: 'draft', title: 'Draft', subtitle: `Created by ${currentUserName}`, status: 'Completed', timestamp: new Date().toLocaleString([], { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }), approverName: currentUserName },
    { id: 'l1', title: 'L1 Approval', subtitle: activeWorkflow.l1Approver || 'Pending Assignment', status: 'Pending', approverName: activeWorkflow.l1Approver },
    { id: 'l2', title: 'L2 Approval', subtitle: activeWorkflow.l2Approver || 'Pending Assignment', status: 'Pending', approverName: activeWorkflow.l2Approver },
    ...(activeWorkflow.workflowType === 'Escalation Workflow'
      ? [{ id: 'esc', title: 'Escalation / Compliance', subtitle: activeWorkflow.escalateTo || 'Compliance Officer', status: 'Pending' as ApprovalStepStatus, approverName: activeWorkflow.escalateTo }]
      : []),
    { id: 'final', title: 'Final Approval', subtitle: activeWorkflow.finalApprover || 'Final Approver', status: 'Pending', approverName: activeWorkflow.finalApprover }
  ] : [
    { id: 'draft', title: 'Draft', subtitle: `Created by ${currentUserName}`, status: 'Completed', timestamp: new Date().toLocaleString([], { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }), approverName: currentUserName }
  ];

  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    return () => {
      Object.values(stagingIntervalsRef.current).forEach(clearInterval);
      stagingIntervalsRef.current = {};
    };
  }, []);

  useEffect(() => {
    if (isOpen) {
      Object.values(stagingIntervalsRef.current).forEach(clearInterval);
      stagingIntervalsRef.current = {};
      setStagingFiles([]);
      setSelectedFiles([]);
      setTags([]);
      setNewTagInput('');
      setIsAddingTag(false);
      setFileError('');
      setDocumentType(emptyMetadata.documentType);
      setDocumentStatus(emptyMetadata.documentStatus);

      if (selectedFolderCode) {
        fetchRaciForProcess(selectedFolderCode);
        fetchWorkflowsForProcess(selectedFolderCode);
      } else {
        setRaci(defaultRaci());
        resolvePersonPhotos([currentUserName]);
        setWorkflows([]);
        setSelectedWorkflowId(null);
      }
    }
  }, [isOpen, selectedFolderCode]);

  if (!isOpen) return null;

  /* ---------------- file handling ---------------- */

  const startStagingProgress = (id: string) => {
    const interval = setInterval(() => {
      setStagingFiles(prev => prev.map(sf => {
        if (sf.id !== id || sf.status === 'completed') return sf;
        const next = Math.min(100, sf.progress + (12 + Math.random() * 18));
        if (next >= 100) {
          clearInterval(stagingIntervalsRef.current[id]);
          delete stagingIntervalsRef.current[id];
          return { ...sf, progress: 100, status: 'completed' };
        }
        return { ...sf, progress: next };
      }));
    }, 220);
    stagingIntervalsRef.current[id] = interval;
  };

  const applyFiles = (incoming: FileList | File[]) => {
    const incomingArray = Array.from(incoming);
    const validFiles: File[] = [];
    let error = '';

    incomingArray.forEach((file) => {
      const ext = file.name.split('.').pop()?.toLowerCase() || '';
      if (!ACCEPTED_EXTENSIONS.includes(ext)) {
        error = 'Unsupported file type. Allowed: PDF, DOCX, XLSX, PPTX, TIFF.';
        return;
      }
      if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) {
        error = `One or more files exceed the ${MAX_FILE_SIZE_MB} MB limit.`;
        return;
      }
      validFiles.push(file);
    });

    setFileError(error);
    if (validFiles.length === 0) return;

    // avoid duplicate name+size pairs against both already-attached and currently-staging files
    const existingKeys = new Set([
      ...selectedFiles.map(f => `${f.name}_${f.size}`),
      ...stagingFiles.map(sf => `${sf.file.name}_${sf.file.size}`)
    ]);
    const deduped = validFiles.filter(f => !existingKeys.has(`${f.name}_${f.size}`));
    if (deduped.length === 0) return;

    const newStaging: IStagingFile[] = deduped.map(file => ({
      id: `${file.name}_${file.size}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      file,
      progress: 0,
      status: 'uploading'
    }));

    setStagingFiles(prev => [...prev, ...newStaging]);
    newStaging.forEach(sf => startStagingProgress(sf.id));
  };

  // "Discard" — cancels the in-progress staging batch without attaching anything.
  const handleDiscardStaging = () => {
    Object.values(stagingIntervalsRef.current).forEach(clearInterval);
    stagingIntervalsRef.current = {};
    setStagingFiles([]);
  };

  const handleAttachStaging = () => {
    const readyFiles = stagingFiles.filter(sf => sf.status === 'completed').map(sf => sf.file);
    if (readyFiles.length === 0) return;
    setSelectedFiles(prev => [...prev, ...readyFiles]);
    setStagingFiles([]);
  };

  const handleBrowseFiles = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) applyFiles(files);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) applyFiles(files);
  };

  const handleRemoveFile = (fileToRemove: File) => {
    setSelectedFiles(prev => prev.filter(f => f !== fileToRemove));
    setFileError('');
  };

  /* ---------------- tags (same add/remove pattern as DocumentMetadata.tsx) ---------------- */

  const handleAddTag = () => {
    const value = newTagInput.trim();
    if (value && !tags.includes(value)) {
      setTags([...tags, value]);
      setNewTagInput('');
      setIsAddingTag(false);
    }
  };

  const handleRemoveTag = (tagToRemove: string) => setTags(tags.filter(t => t !== tagToRemove));
  const handleDocumentTypeChange = (value: string) => {
    setDocumentType(value);
    setTags((prevTags) => {
      const withoutOldTypeTag = documentType ? prevTags.filter((t) => t !== documentType) : prevTags;
      if (!value || withoutOldTypeTag.includes(value)) return withoutOldTypeTag;
      return [...withoutOldTypeTag, value];
    });
  };

  /* ---------------- RACI ---------------- */
  /* ---------------- submit ---------------- */

  const handleClose = () => {
    Object.values(stagingIntervalsRef.current).forEach(clearInterval);
    stagingIntervalsRef.current = {};
    setStagingFiles([]);
    setSelectedFiles([]);
    setTags([]);
    setNewTagInput('');
    setIsAddingTag(false);
    setFileError('');
    setDocumentType(emptyMetadata.documentType);
    setDocumentStatus(emptyMetadata.documentStatus);
    onClose();
  };

  const handleSubmit = () => {
    if (selectedFiles.length === 0) {
      setFileError('Please select at least one file before submitting.');
      return;
    }

    // Power Automate flow listens for to notify that process's L1 Approver.
    onUpload(selectedFiles, {
      documentType,
      documentStatus,
      tags,
      // Workflow selected -> 'Send for L1' (Power Automate flow fires); "None" -> 'Draft'
      documentCurrentStatus: activeWorkflow ? 'Send for L1' : 'Draft',
      // WorkflowID of the workflow currently selected in the "Select Workflow" dropdown
      workflowId: activeWorkflow?.workflowId || ''
    });
  };

  /* ---------------- render ---------------- */
  // Same open pattern as ProcessManagementForm.tsx: plain fixed overlay + CSS
  // keyframe fade/scale-in animation, no portal.

  return (
    <div className="ud-overlay" style={styles.overlay} onMouseDown={(e) => { if (e.target === e.currentTarget) handleClose(); }}>
      <style>{`
        @keyframes udFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes udScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        .ud-overlay { animation: udFadeIn 0.18s ease; }
        .ud-modal { animation: udScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
        .ud-btn { transition: transform 0.15s ease, filter 0.15s ease, background-color 0.15s ease, box-shadow 0.15s ease; }
        .ud-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
        .ud-btn:active { transform: translateY(0); filter: brightness(0.97); }
        .ud-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; border-radius: 5px; }
        .ud-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
        .ud-drop-btn { transition: transform 0.15s ease, filter 0.15s ease; }
        .ud-drop-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
      `}</style>
      <div className="ud-modal" style={styles.modal} onClick={(e) => e.stopPropagation()}>
      <div className="ud-modal-scroll" style={styles.modalScroll}>
        {/* Header */}
        <div style={styles.headerRow}>
          <div style={styles.headerCol}>
            <span style={styles.headerLabel}>Process</span>
            <span style={styles.headerValue}>{processName || selectedFolderCode || 'General'}</span>
          </div>
          <div style={styles.headerCol}>
            <span style={styles.headerLabel}>Owner</span>
            <span style={styles.headerValue}><Icon name="user" color="#005a43" size={12} /> {currentUserName}</span>
          </div>
          <div style={styles.headerCol}>
            <span style={styles.headerLabel}>Select Workflow</span>
            {isWorkflowLoading ? (
              <span style={styles.headerValue}>Loading…</span>
            ) : workflows.length === 0 ? (
              <span style={{ ...styles.headerValue, color: '#94a3b8', fontWeight: 600 }}>No workflow configured</span>
            ) : (
              <CompactDropdown
                value={selectedWorkflowId}
                options={workflows.map(w => ({ id: w.id, label: w.workflowTitle }))}
                onChange={(id) => {
                  setSelectedWorkflowId(id);
                  // None -> Draft, a workflow selected -> Pending Approval
                  setDocumentStatus(id === null ? 'Draft' : 'Pending Approval');
                }}
              />
            )}
          </div>
          <div style={styles.headerCol}>
            <span style={styles.headerLabel}>Status</span>
            <span style={{ ...styles.headerValue, color: STATUS_DOT_COLORS[documentStatus] || '#475569' }}>
              <span style={{ width: '6px', height: '6px', borderRadius: '50%', backgroundColor: STATUS_DOT_COLORS[documentStatus] || '#94a3b8', display: 'inline-block', marginRight: '5px' }} />
              {documentStatus}
            </span>
          </div>
          <button className="ud-close-btn" onClick={handleClose} style={styles.closeBtn}>
            <Icon name="close" color="#64748b" size={18} />
          </button>
        </div>

        {selectedFolderCode && (
          <div style={styles.folderBanner}>
            Uploading to folder: <strong>{selectedFolderCode}</strong>
          </div>
        )}

        {/* Step indicator */}
        <div style={styles.stepRow}>
          {[{ n: 1, label: 'Upload Document' }, { n: 2, label: 'Metadata' }, { n: 3, label: 'Approval Workflow' }].map((s, idx) => (
            <React.Fragment key={s.n}>
              <div style={styles.stepItem}>
                <span style={styles.stepCircle}>{s.n}</span>
                <span style={styles.stepLabel}>{s.label}</span>
              </div>
              {idx < 2 && <div style={styles.stepLine} />}
            </React.Fragment>
          ))}
        </div>

        {/* Three panels */}
        <div style={styles.panelsRow}>
          {/* Panel 1: Upload */}
          <div style={styles.panel}>
            <div style={styles.panelTitle}>1. Upload Document</div>
            <div style={styles.panelSubtitle}>Select the files to upload</div>

            <input type="file" ref={fileInputRef} onChange={handleBrowseFiles} style={{ display: 'none' }} accept={ACCEPTED_EXTENSIONS.map(e => `.${e}`).join(',')} multiple />

            {stagingFiles.length > 0 ? (
              /* Progress indicator shown right after Browse Files / drop, before files are attached */
              <div style={styles.stagingCard}>
                <div style={styles.stagingHeaderRow}>
                  <span style={styles.stagingHeaderText}>
                    {stagingFiles.every(sf => sf.status === 'completed')
                      ? `${stagingFiles.length} file${stagingFiles.length > 1 ? 's' : ''} ready`
                      : `${stagingFiles.length} file${stagingFiles.length > 1 ? 's' : ''} uploading...`}
                  </span>
                  <button onClick={handleDiscardStaging} style={styles.stagingCloseBtn} title="Discard">
                    <Icon name="close" color="#94a3b8" size={14} />
                  </button>
                </div>

                <div style={styles.stagingList}>
                  {stagingFiles.map((sf) => (
                    <div key={sf.id} style={styles.stagingItem}>
                      <div style={styles.stagingItemTopRow}>
                        <div style={styles.stagingItemIcon}><Icon name="file" color="#005a43" size={14} /></div>
                        <div style={{ flexGrow: 1, minWidth: 0 }}>
                          <div style={styles.stagingItemName} title={sf.file.name}>{sf.file.name}</div>
                          <div style={styles.stagingItemMeta}>
                            {sf.status === 'completed' ? 'Upload completed' : 'Uploading'} · {formatBytes(sf.file.size)}
                          </div>
                        </div>
                        {sf.status === 'completed'
                          ? <Icon name="check" color="#15803d" size={16} />
                          : <span style={styles.stagingPercent}>{Math.round(sf.progress)}%</span>}
                      </div>
                      <div style={styles.stagingProgressTrack}>
                        <div style={{ ...styles.stagingProgressFill, width: `${sf.progress}%`, backgroundColor: sf.status === 'completed' ? '#15803d' : '#005a43' }} />
                      </div>
                    </div>
                  ))}
                </div>

                <div style={styles.stagingFooterRow}>
                  <button style={styles.stagingDiscardBtn} onClick={handleDiscardStaging}>Discard</button>
                  <button
                    style={{
                      ...styles.stagingAttachBtn,
                      opacity: stagingFiles.every(sf => sf.status === 'completed') ? 1 : 0.55,
                      cursor: stagingFiles.every(sf => sf.status === 'completed') ? 'pointer' : 'not-allowed'
                    }}
                    onClick={handleAttachStaging}
                    disabled={!stagingFiles.every(sf => sf.status === 'completed')}
                  >
                    Attach files
                  </button>
                </div>
              </div>
            ) : (
              <div
                style={{ ...styles.dropZone, ...(isDragging ? styles.dropZoneActive : {}) }}
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={handleDrop}
              >
                <Icon name="cloud" color="#005a43" size={28} />
                <div style={styles.dropZoneText}>Drag &amp; drop your files here</div>
                <div style={styles.dropZoneOr}>or</div>
                <button className="ud-drop-btn" style={styles.browseBtn} onClick={() => fileInputRef.current?.click()}>Browse Files</button>
                <div style={styles.dropZoneHint}>PDF, DOCX, XLSX, PPTX, TIFF, {MAX_FILE_SIZE_MB} MB max per file · multiple files allowed</div>
              </div>
            )}

            {fileError && <div style={styles.errorText}>{fileError}</div>}

            {selectedFiles.length > 0 && (
              <>
                <div style={styles.uploadedLabel}>Uploaded Files ({selectedFiles.length})</div>
                <div style={styles.uploadedFilesList}>
                  {selectedFiles.map((file, idx) => (
                    <div key={`${file.name}_${file.size}_${idx}`} style={styles.uploadedFileCard}>
                      <div style={styles.uploadedFileIcon}><Icon name="file" color="#dc2626" size={16} /></div>
                      <div style={{ flexGrow: 1, overflow: 'hidden' }}>
                        <div style={styles.uploadedFileName}>{file.name}</div>
                        <div style={styles.uploadedFileSize}>{formatBytes(file.size)}</div>
                      </div>
                      <button onClick={() => handleRemoveFile(file)} style={styles.removeFileBtn} title="Remove file">
                        <Icon name="close" color="#94a3b8" size={13} />
                      </button>
                      <Icon name="check" color="#15803d" size={15} />
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>

          {/* Panel 2: Metadata — same fields SharePoint actually stores (DocumentType,
              DocumentStatus, DocumentTags), same UI/behavior as the Document Details
              screen (DocumentMetadata.tsx), so this is real metadata, not a mock. */}
          <div style={styles.panel}>
            <div style={styles.panelTitle}>2. Metadata</div>
            <div style={styles.panelSubtitle}>Set the document fields stored in SharePoint</div>

            <DocumentTypeField
              value={documentType}
              onChange={handleDocumentTypeChange}
              context={context}
            />

            <div style={styles.fieldBlock}>
              <label style={styles.fieldLabel}>Tags</label>
              <div style={styles.tagsContainer}>
                {tags.map((tag) => (
                  <span key={tag} style={styles.tagPill}>
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
                      onKeyDown={(e) => { if (e.key === 'Enter') handleAddTag(); }}
                      style={styles.newTagInput}
                      autoFocus
                    />
                    <button onClick={handleAddTag} style={styles.addTagSubmitBtn}>Add</button>
                    <button onClick={() => { setIsAddingTag(false); setNewTagInput(''); }} style={styles.tagRemoveBtn}>Cancel</button>
                  </div>
                ) : (
                  <button onClick={() => setIsAddingTag(true)} style={styles.addTagBtn}>+ Add Tag</button>
                )}
              </div>
            </div>

            <div style={styles.systemFieldsBlock}>
              <div style={styles.fieldLabel}>System Fields (Auto Generated)</div>
              <div style={styles.systemFieldsGrid}>
                <div style={styles.fieldGroupSm}>
                  <label style={styles.subLabel}>File Name</label>
                  <span style={styles.systemFieldValue} title={selectedFiles[0]?.name}>
                    {selectedFiles.length === 0 ? '—' : selectedFiles.length === 1 ? selectedFiles[0].name : `${selectedFiles[0].name} +${selectedFiles.length - 1} more`}
                  </span>
                </div>
                <div style={styles.fieldGroupSm}>
                  <label style={styles.subLabel}>File Size</label>
                  <span style={styles.systemFieldValue}>{selectedFiles.length === 0 ? '—' : formatBytes(selectedFiles[0].size)}</span>
                </div>
                <div style={styles.fieldGroupSm}>
                  <label style={styles.subLabel}>File Type</label>
                  <span style={styles.systemFieldValue}>{selectedFiles.length === 0 ? '—' : `.${selectedFiles[0].name.split('.').pop()}`}</span>
                </div>
                <div style={styles.fieldGroupSm}>
                  <label style={styles.subLabel}>Version</label>
                  <span style={{ ...styles.systemFieldValue, color: '#005a43', fontWeight: 700 }}>{selectedFiles.length === 0 ? '—' : 'v0.1'}</span>
                </div>
              </div>
            </div>
          </div>

          {/* Panel 3: Approval workflow */}
          <div style={styles.panel}>
            <div style={styles.panelTitle}>3. Approval Workflow</div>
            <div style={styles.panelSubtitle}>
              {isWorkflowLoading ? 'Loading workflow…' : activeWorkflow ? 'Live from Workflow Configuration' : 'Track the approval process'}
            </div>

            {!isWorkflowLoading && workflows.length === 0 ? (
              <div style={styles.noWorkflowBox}>
                <div>
                  No workflow is configured for this process yet. Documents will still upload, but
                  won't have an approval chain until one is created in Approval Management.
                </div>
                <button
                  className="ud-btn"
                  style={styles.createWorkflowBtn}
                  onClick={() => onCreateWorkflow && onCreateWorkflow()}
                  type="button"
                >
                  <Icon name="plus" color="#ffffff" size={11} />
                  Create Workflow
                </button>
              </div>
            ) : (
              <div style={styles.timeline}>
                {approvalSteps.map((step, idx) => (
                  <div key={step.id} style={styles.timelineRow}>
                    <div style={styles.timelineMarkerCol}>
                      {step.approverName ? (
                        <PersonAvatar name={step.approverName} photoUrl={personPhotos[step.approverName]} style={styles.timelineAvatar} />
                      ) : (
                        <div style={styles.timelineDot}><Icon name="user" color="#005a43" size={11} /></div>
                      )}
                      {idx < approvalSteps.length - 1 && <div style={styles.timelineConnector} />}
                    </div>
                    <div style={styles.timelineContent}>
                      <div style={styles.timelineTop}>
                        <span style={styles.timelineTitle}>{step.title}</span>
                      </div>
                      <div style={styles.timelineSubtitle}>{step.subtitle}</div>
                      {step.timestamp && <div style={styles.timelineTimestamp}>{step.timestamp}</div>}
                    </div>
                  </div>
                ))}
              </div>
            )}

            {activeWorkflow && activeWorkflow.deadlineDays > 0 && (
              <div style={styles.workflowDeadlineNote}>
                Approval deadline: <strong>{activeWorkflow.deadlineDays} day(s)</strong> per step
              </div>
            )}
          </div>
        </div>

        {/* RACI Matrix */}
        <div style={styles.raciCard}>
          <div style={styles.raciHeaderRow}>
            <div>
              <div style={styles.raciTitle}>RACI Matrix</div>
              <div style={styles.raciSubtitle}>{isRaciLoading ? 'Loading roles & responsibilities...' : 'Roles and responsibilities for this process'}</div>
            </div>
          </div>

          <div style={styles.raciGrid}>
            {raci.map(entry => (
              <div key={entry.role} style={styles.raciEntry}>
                <div style={styles.raciRoleRow}>
                  <Icon name="user" color="#005a43" size={12} />
                  <span style={styles.raciRoleLabel}>{entry.role}</span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  {(entry.names && entry.names.length ? entry.names : [entry.name]).map((personName, personIdx) => (
                    <div key={personIdx} style={styles.raciPersonRow}>
                      <PersonAvatar name={personName} photoUrl={personPhotos[personName]} style={styles.raciAvatar} />
                      <div style={{ overflow: 'hidden' }}>
                        <div style={styles.raciName}>{personName}</div>
                      </div>
                    </div>
                  ))}
                  <div style={styles.raciMeta}>{entry.subtitle}</div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer actions */}
        <div style={styles.footerRow}>
          <button className="ud-btn" style={styles.cancelBtn} onClick={handleClose} disabled={isUploading}>
            Cancel
          </button>
          <button className="ud-btn" style={styles.submitBtn} onClick={handleSubmit} disabled={isUploading}>
            <Icon name="send" color="#ffffff" size={13} />
            {isUploading ? 'Submitting...' : 'Submit'}
          </button>
        </div>
      </div>
      </div>
    </div>
  );
};

  // STYLES

const styles: { [key: string]: React.CSSProperties } = {
  overlay: {
    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
    backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)',
    display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999,
    fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
    padding: '20px'
  },
  modal: {
    backgroundColor: '#ffffff', borderRadius: '10px', width: '980px', maxWidth: '96vw',
    maxHeight: '92vh', overflow: 'hidden', boxShadow: '0 20px 50px rgba(0,0,0,0.25)',
    boxSizing: 'border-box', display: 'flex', flexDirection: 'column'
  },

  modalScroll: {
    overflowY: 'auto', padding: '18px 22px', boxSizing: 'border-box', minHeight: 0
  },

  headerRow: { display: 'flex', alignItems: 'flex-start', gap: '28px', paddingBottom: '14px', borderBottom: '1px solid #f1f5f9', position: 'relative' },
  headerCol: { display: 'flex', flexDirection: 'column', gap: '3px' },
  headerLabel: { fontSize: '10px', fontWeight: 600, color: '#94a3b8', letterSpacing: '0.3px' },
  headerValue: { fontSize: '12px', fontWeight: 700, color: '#1e293b', display: 'flex', alignItems: 'center', gap: '4px' },
  closeBtn: { position: 'absolute', right: 0, top: 0, background: 'none', border: 'none', cursor: 'pointer', padding: '2px' },

  folderBanner: { backgroundColor: '#e6f4f1', border: '1px solid #ccece6', color: '#0d6b52', fontSize: '11px', borderRadius: '5px', padding: '6px 10px', marginTop: '10px' },

  stepRow: { display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '18px 0' },
  stepItem: { display: 'flex', alignItems: 'center', gap: '8px' },
  stepCircle: { width: '24px', height: '24px', borderRadius: '50%', backgroundColor: '#005a43', color: '#ffffff', fontSize: '12px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  stepLabel: { fontSize: '12px', fontWeight: 600, color: '#1e293b' },
  stepLine: { width: '70px', height: '2px', backgroundColor: '#005a43', margin: '0 14px' },

  panelsRow: { display: 'flex', gap: '14px', alignItems: 'stretch' },
  panel: { flex: 1, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '14px', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', minWidth: 0 },
  panelTitle: { fontSize: '13px', fontWeight: 700, color: '#0f172a', marginBottom: '2px' },
  panelSubtitle: { fontSize: '10px', color: '#94a3b8', marginBottom: '12px' },

  dropZone: { border: '2px dashed #cbd5e1', borderRadius: '8px', padding: '18px 10px', display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', backgroundColor: '#fafcfc', transition: 'all 0.15s ease' },
  dropZoneActive: { borderColor: '#005a43', backgroundColor: '#e6f4f1' },
  dropZoneText: { fontSize: '12px', fontWeight: 600, color: '#334155', marginTop: '8px' },
  dropZoneOr: { fontSize: '10px', color: '#94a3b8', margin: '4px 0' },
  browseBtn: { backgroundColor: '#005a43', color: '#ffffff', border: 'none', borderRadius: '5px', padding: '7px 16px', fontSize: '11px', fontWeight: 600, cursor: 'pointer' },
  dropZoneHint: { fontSize: '9px', color: '#94a3b8', marginTop: '10px' },
  errorText: { fontSize: '10px', color: '#dc2626', marginTop: '6px' },

  // Staging / progress indicator shown right after Browse Files or drag-drop.
  stagingCard: { border: '1px solid #e2e8f0', borderRadius: '8px', backgroundColor: '#ffffff', padding: '10px', display: 'flex', flexDirection: 'column', gap: '8px' },
  stagingHeaderRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between' },
  stagingHeaderText: { fontSize: '11px', fontWeight: 700, color: '#1e293b' },
  stagingCloseBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '2px' },
  stagingList: { display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '180px', overflowY: 'auto' },
  stagingItem: { border: '1px solid #f1f5f9', borderRadius: '6px', padding: '8px', backgroundColor: '#f8fafc' },
  stagingItemTopRow: { display: 'flex', alignItems: 'center', gap: '8px' },
  stagingItemIcon: { width: '24px', height: '24px', borderRadius: '4px', backgroundColor: '#e6f4f1', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  stagingItemName: { fontSize: '11px', fontWeight: 600, color: '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  stagingItemMeta: { fontSize: '9px', color: '#94a3b8', marginTop: '1px' },
  stagingPercent: { fontSize: '10px', fontWeight: 700, color: '#005a43', flexShrink: 0 },
  stagingProgressTrack: { height: '5px', borderRadius: '3px', backgroundColor: '#e2e8f0', marginTop: '6px', overflow: 'hidden' },
  stagingProgressFill: { height: '100%', borderRadius: '3px', transition: 'width 0.2s ease' },
  stagingFooterRow: { display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '2px' },
  stagingDiscardBtn: { backgroundColor: '#ffffff', border: '1px solid #cbd5e1', color: '#334155', borderRadius: '5px', padding: '6px 14px', fontSize: '11px', fontWeight: 600, cursor: 'pointer' },
  stagingAttachBtn: { backgroundColor: '#005a43', border: 'none', color: '#ffffff', borderRadius: '5px', padding: '6px 16px', fontSize: '11px', fontWeight: 600 },

  uploadedLabel: { fontSize: '10px', fontWeight: 700, color: '#64748b', marginTop: '12px', marginBottom: '6px' },
  uploadedFilesList: { display: 'flex', flexDirection: 'column', gap: '6px' },
  uploadedFileCard: { display: 'flex', alignItems: 'center', gap: '8px', border: '1px solid #e2e8f0', borderRadius: '6px', padding: '8px 10px', backgroundColor: '#f8fafc' },
  uploadedFileIcon: { width: '26px', height: '26px', borderRadius: '4px', backgroundColor: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  uploadedFileName: { fontSize: '11px', fontWeight: 600, color: '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  uploadedFileSize: { fontSize: '9px', color: '#94a3b8' },
  removeFileBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '2px' },

  fieldBlock: { marginTop: '12px' },
  fieldLabel: { fontSize: '10px', fontWeight: 600, color: '#475569', marginBottom: '4px', display: 'block' },
  textInput: { width: '100%', padding: '7px 9px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#1e293b', boxSizing: 'border-box', outline: 'none' },
  selectInput: { width: '100%', padding: '7px 9px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#1e293b', boxSizing: 'border-box', outline: 'none', backgroundColor: '#ffffff' },
  textarea: { width: '100%', padding: '7px 9px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#1e293b', boxSizing: 'border-box', outline: 'none', resize: 'vertical', fontFamily: 'inherit' },

  tagsBox: { display: 'flex', flexWrap: 'wrap', gap: '5px', alignItems: 'center', border: '1px solid #cbd5e1', borderRadius: '5px', padding: '5px 7px' },
  tagChip: { backgroundColor: '#e6f4f1', color: '#0d6b52', fontSize: '10px', fontWeight: 600, borderRadius: '4px', padding: '2px 6px', display: 'flex', alignItems: 'center', gap: '4px' },
  tagRemove: { cursor: 'pointer', fontWeight: 700, color: '#0d6b52' },
  tagInput: { border: 'none', outline: 'none', fontSize: '11px', flexGrow: 1, minWidth: '60px' },

  // Custom dropdown (Document Type / Document Status) — same visual language as DocumentMetadata.tsx
  ddWrapper: { position: 'relative', width: '100%' },
  ddControl: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 9px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#334155', cursor: 'pointer', transition: 'all 0.15s ease-in-out', boxSizing: 'border-box' },
  ddValueText: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  ddMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1)', zIndex: 999, overflow: 'hidden', padding: '4px' },
  ddOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 10px', borderRadius: '6px', fontSize: '11px', cursor: 'pointer', marginBottom: '2px' },

  // Compact variant of the dropdown above, sized to sit inline in the header row (Workflow Type)
  headerDdWrapper: { position: 'relative', display: 'inline-block', width: '168px' },
  headerDdControl: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', padding: '4px 8px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', fontWeight: 700, color: '#1e293b', cursor: 'pointer', transition: 'all 0.15s ease-in-out', boxSizing: 'border-box', width: '100%' },
  headerDdMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1)', zIndex: 999, overflow: 'hidden', padding: '4px' },

  // Tags (real add/remove pattern, same as DocumentMetadata.tsx)
  tagsContainer: { display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '4px', alignItems: 'center' },
  tagPill: { backgroundColor: '#e6f4f1', color: '#0d6b52', padding: '3px 8px', borderRadius: '4px', fontSize: '10px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '6px' },
  tagRemoveBtn: { background: 'none', border: 'none', cursor: 'pointer', fontSize: '11px', color: '#64748b', padding: 0 },
  addTagBtn: { backgroundColor: '#f8fafc', border: '1px dashed #cbd5e1', color: '#64748b', padding: '3px 9px', borderRadius: '4px', fontSize: '10px', fontWeight: 600, cursor: 'pointer' },
  newTagInput: { padding: '4px 7px', border: '1px solid #cbd5e1', borderRadius: '4px', fontSize: '10px', outline: 'none' },
  addTagSubmitBtn: { backgroundColor: '#005a43', color: '#fff', border: 'none', padding: '4px 8px', borderRadius: '4px', fontSize: '10px', cursor: 'pointer' },

  // System fields (auto generated) — mirrors DocumentMetadata.tsx's read-only strip
  systemFieldsBlock: { marginTop: '16px', paddingTop: '12px', borderTop: '1px dashed #e2e8f0' },
  systemFieldsGrid: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', marginTop: '8px' },
  fieldGroupSm: { display: 'flex', flexDirection: 'column', gap: '3px', minWidth: 0 },
  subLabel: { fontSize: '9px', fontWeight: 600, color: '#94a3b8' },
  systemFieldValue: { fontSize: '11px', color: '#334155', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },

  timeline: { display: 'flex', flexDirection: 'column' },
  timelineRow: { display: 'flex', gap: '10px' },
  timelineMarkerCol: { display: 'flex', flexDirection: 'column', alignItems: 'center' },
  timelineDot: { width: '22px', height: '22px', borderRadius: '50%', border: '2px solid #005a43', backgroundColor: '#ffffff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  timelineDotDone: { backgroundColor: '#005a43' },
  timelineAvatar: { width: '22px', height: '22px', borderRadius: '50%', backgroundColor: '#005a43', color: '#ffffff', fontSize: '10px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  timelineConnector: { width: '2px', flexGrow: 1, backgroundColor: '#cbd5e1', minHeight: '24px', margin: '2px 0' },
  timelineContent: { paddingBottom: '14px', flexGrow: 1, minWidth: 0 },
  timelineTop: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' },
  timelineTitle: { fontSize: '11px', fontWeight: 700, color: '#1e293b' },
  timelineSubtitle: { fontSize: '10px', color: '#64748b', marginTop: '1px' },
  timelineTimestamp: { fontSize: '9px', color: '#94a3b8', marginTop: '1px' },
  statusPill: { fontSize: '9px', fontWeight: 700, borderRadius: '10px', padding: '2px 8px', whiteSpace: 'nowrap' },

  viewWorkflowLink: { display: 'flex', alignItems: 'center', gap: '5px', fontSize: '11px', fontWeight: 600, color: '#005a43', cursor: 'pointer', marginTop: '4px' },
  noWorkflowBox: { fontSize: '11px', color: '#94a3b8', backgroundColor: '#f8fafc', border: '1px dashed #e2e8f0', borderRadius: '6px', padding: '12px', lineHeight: 1.5, display: 'flex', flexDirection: 'column', gap: '10px', alignItems: 'flex-start' },
  createWorkflowBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#005a43', border: 'none', color: '#ffffff', borderRadius: '5px', padding: '7px 12px', fontSize: '11px', fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' },
  workflowDeadlineNote: { fontSize: '10.5px', color: '#64748b', marginTop: '8px', paddingTop: '8px', borderTop: '1px dashed #e2e8f0' },

  raciCard: { border: '1px solid #e2e8f0', borderRadius: '8px', padding: '14px', marginTop: '16px' },
  raciHeaderRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '12px' },
  raciTitle: { fontSize: '13px', fontWeight: 700, color: '#0f172a' },
  raciSubtitle: { fontSize: '10px', color: '#94a3b8', marginTop: '2px' },
  editRaciBtn: { display: 'flex', alignItems: 'center', gap: '5px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', color: '#005a43', borderRadius: '5px', padding: '5px 10px', fontSize: '10px', fontWeight: 600, cursor: 'pointer' },
  raciGrid: { display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '10px' },
  raciEntry: { border: '1px solid #f1f5f9', borderRadius: '6px', padding: '10px', minWidth: 0 },
  raciRoleRow: { display: 'flex', alignItems: 'center', gap: '5px', marginBottom: '8px' },
  raciRoleLabel: { fontSize: '10px', fontWeight: 700, color: '#334155' },
  raciPersonRow: { display: 'flex', alignItems: 'center', gap: '8px' },
  raciAvatar: { width: '26px', height: '26px', borderRadius: '50%', backgroundColor: '#005a43', color: '#ffffff', fontSize: '11px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  raciName: { fontSize: '11px', fontWeight: 700, color: '#1e293b', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  raciNameInput: { fontSize: '11px', fontWeight: 700, color: '#1e293b', border: '1px solid #cbd5e1', borderRadius: '4px', padding: '2px 4px', width: '100%', boxSizing: 'border-box' },
  raciMeta: { fontSize: '9px', color: '#94a3b8', marginTop: '1px' },

  footerRow: { display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '18px', paddingTop: '14px', borderTop: '1px solid #f1f5f9' },
  cancelBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', color: '#334155', borderRadius: '6px', padding: '9px 16px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' },
  submitBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#005a43', border: 'none', color: '#ffffff', borderRadius: '6px', padding: '9px 18px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' }
};

export default UploadDocumentModal;