import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx, fileOpenUrl } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import type { IDocumentItem, IProcessSummaryItem } from './AnalyticsDashboard';
import CreateWorkflow from './CreateWorkflow';

export interface IDocumentApprovalProps {
  context: any;
  documents: IDocumentItem[];
  filteredDocuments: IDocumentItem[];
  searchTerm: string;
  openDropdownId: number | null;
  setOpenDropdownId: (id: number | null) => void;
  updateDocumentStatus: (docId: number, newStatus: IDocumentItem['status']) => void;
  getStatusStyle: (status: string) => { backgroundColor: string; color: string };
  processes: IProcessSummaryItem[];
  processesLoading: boolean;
}

interface IRaciRow {
  id: number;
  processName: string;
  responsible: string;
  accountable: string;
  consulted: string;
  informed: string;
}

interface IWorkflowRow {
  id: number;
  workflowTitle: string;
  processCode: string;
  workflowType: string;
  escalationCriteria: string;
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

// -------------------- Icons --------------------

const FileIcon = ({ size = 18 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="#005c42" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
  </svg>
);

const CheckIcon = ({ color = "#005c42", size = 18 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="9" />
    <polyline points="9 12 11.5 14.5 15 9.5" />
  </svg>
);

const MoreIcon = () => (
  <svg width="4" height="14" viewBox="0 0 4 14" fill="#64748b" xmlns="http://www.w3.org/2000/svg">
    <circle cx="2" cy="2" r="1.5" />
    <circle cx="2" cy="7" r="1.5" />
    <circle cx="2" cy="12" r="1.5" />
  </svg>
);

const ClockIcon = ({ color = "#d97706", size = 20 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="9" />
    <polyline points="12 7 12 12 15.5 14" />
  </svg>
);

const AlertIcon = ({ color = "#dc2626", size = 20 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
    <line x1="12" y1="9" x2="12" y2="13" />
    <line x1="12" y1="17" x2="12.01" y2="17" />
  </svg>
);

const ShieldIcon = ({ color = "#d97706", size = 20 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
  </svg>
);

const UsersIcon = ({ color = "#005c42", size = 16 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M23 21v-2a4 4 0 0 0-3-3.87" />
    <path d="M16 3.13a4 4 0 0 1 0 7.75" />
  </svg>
);

const InfoIcon = ({ size = 13, color = "#94a3b8" }: { size?: number; color?: string }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="10" />
    <line x1="12" y1="16" x2="12" y2="12" />
    <line x1="12" y1="8" x2="12.01" y2="8" />
  </svg>
);

const EyeIcon = ({ color = "#475569", size = 12 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" />
    <circle cx="12" cy="12" r="3" />
  </svg>
);

const PencilIcon = ({ color = "#475569", size = 12 }: { color?: string; size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
    <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
  </svg>
);

const PlusIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
    <line x1="12" y1="5" x2="12" y2="19" />
    <line x1="5" y1="12" x2="19" y2="12" />
  </svg>
);


const STAGES: { status: IDocumentItem['status']; label: string; owner: string }[] = [
  { status: 'Draft', label: 'Draft', owner: 'Initiator' },
  { status: 'Pending Approval', label: 'Pending Approval', owner: 'Process Owner' },
  { status: 'In Review', label: 'In Review', owner: 'Compliance / HOD' },
  { status: 'Active', label: 'Approved', owner: 'Published' }
];

// How many days a document can sit in Pending Approval / In Review before it
// is treated as escalated. Kept in sync with AnalyticsDashboard's own value.
const ESCALATION_DAYS = 5;

// "My Approvals" is driven by the library's "DocumentCurrentStatus" column: every document
// that has a current status shows up here EXCEPT 'L3 Approved' (L3 is the final approval,
// so those documents are fully done).
const FINAL_APPROVED_STATUS = 'L3 Approved';

const DocumentApproval: React.FC<IDocumentApprovalProps> = ({
  context,
  documents,
  filteredDocuments,
  searchTerm,
  openDropdownId,
  setOpenDropdownId,
  updateDocumentStatus,
  getStatusStyle,
  processes,
  processesLoading
}) => {
  const [hoveredRowId, setHoveredRowId] = useState<number | null>(null);
  const [activeTab, setActiveTab] = useState<'My Approvals' | 'All Requests' | 'RACI Matrix' | 'Approval Rules' | 'All Workflows'>('My Approvals');

  const [raciRows, setRaciRows] = useState<IRaciRow[]>([]);
  const [raciLoading, setRaciLoading] = useState<boolean>(false);
  const [raciError, setRaciError] = useState<string>('');

  const [workflowRows, setWorkflowRows] = useState<IWorkflowRow[]>([]);
  const [workflowLoading, setWorkflowLoading] = useState<boolean>(false);
  const [workflowError, setWorkflowError] = useState<string>('');

  const [toastMessage, setToastMessage] = useState<string>('');

  useEffect(() => {
    if (!toastMessage) return;
    const t = setTimeout(() => setToastMessage(''), 3500);
    return () => clearTimeout(t);
  }, [toastMessage]);

  // ---- Derived, live document buckets (recomputed on every render as documents refresh) ----
  const nowTime = Date.now();
  const daysSince = (t: number) => (nowTime - t) / (1000 * 60 * 60 * 24);

  const inReviewDocs = documents.filter(d => d.status === 'In Review');
  const actionableDocs = documents.filter(d => !!d.currentStage && d.currentStage !== FINAL_APPROVED_STATUS);
  const escalatedDocs = actionableDocs.filter(d => daysSince(d.modifiedTime) >= ESCALATION_DAYS);
  const completedDocs = documents.filter(d => d.currentStage === 'L3 Approved');

  const sortedActionable = [...actionableDocs].sort((a, b) => a.modifiedTime - b.modifiedTime);

  // ---- Fetch RACI Matrix from SharePoint, live, while the tab is open ----
  const fetchRaciMatrix = async () => {
    if (!context) return;
    setRaciLoading(true);
    setRaciError('');
    try {
      const sp = spfi().using(SPFx(context));
      let rawItems: any[] = [];

      try {
        // Preferred read: assumes Responsible/Accountable/Consulted/Informed are
        // People fields on the "RACI Matrix" list. If your list uses different
        // internal column names, adjust the field names below to match.
        rawItems = await sp.web.lists.getByTitle("RACI Matrix").items
          .select(
            "Id", "Title", "ProcessName",
            "Responsible/Title", "Accountable/Title", "Consulted/Title", "Informed/Title"
          )
          .expand("Responsible", "Accountable", "Consulted", "Informed")
          .top(200)();
      } catch (expandErr) {
        // Fallback for lists where those columns are plain text instead of
        // People fields — read the raw items so the table still renders.
        rawItems = await sp.web.lists.getByTitle("RACI Matrix").items.top(200)();
      }

      const textOf = (val: any): string => {
        if (val === undefined || val === null || val === '') return '—';
        if (Array.isArray(val)) return val.map((v: any) => (v && v.Title) ? v.Title : v).join(', ') || '—';
        if (typeof val === 'object') return val.Title || '—';
        return String(val);
      };

      const mapped: IRaciRow[] = rawItems.map((item: any) => ({
        id: item.Id,
        processName: item.ProcessName || item.Title || 'Untitled',
        responsible: textOf(item.Responsible),
        accountable: textOf(item.Accountable),
        consulted: textOf(item.Consulted),
        informed: textOf(item.Informed)
      }));

      setRaciRows(mapped);
    } catch (err) {
      console.error("Error fetching RACI Matrix list:", err);
      setRaciError('Could not load the RACI Matrix list. Check that a list named "RACI Matrix" exists and the column names match.');
    } finally {
      setRaciLoading(false);
    }
  };

  useEffect(() => {
    if (activeTab !== 'RACI Matrix') return;
    fetchRaciMatrix();
    const interval = setInterval(fetchRaciMatrix, 30000);
    return () => clearInterval(interval);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, context]);

  // ---- Fetch Workflow Configuration from SharePoint, live, while the tab is open ----
  const fetchWorkflows = async () => {
    if (!context) return;
    setWorkflowLoading(true);
    setWorkflowError('');
    try {
      const sp = spfi().using(SPFx(context));
      let rawItems: any[] = [];

      try {
        rawItems = await sp.web.lists.getByTitle("Workflow Configuration").items
          .select(
            "Id", "Title", "WorkflowTitle", "WorkflowType", "EscalationCriteria", "StepRole",
            "L1Approver/Title", "L1Approver/EMail", "L2Approver/Title", "L2Approver/EMail",
            "EscalateTo/Title", "EscalateTo/EMail", "FinalApprover/Title", "FinalApprover/EMail",
            "ComplianceCheck", "ApprovalDeadlineDays", "NotifyApproversEachStep", "AlertOnEscalation"
          )
          .expand("L1Approver", "L2Approver", "EscalateTo", "FinalApprover")
          .top(200)();
      } catch (expandErr) {
        // Fallback in case the Person columns aren't expandable for some reason
        rawItems = await sp.web.lists.getByTitle("Workflow Configuration").items.top(200)();
      }

      const nameOf = (val: any): string => {
        if (val === undefined || val === null || val === '') return '—';
        if (typeof val === 'object') return val.Title || '—';
        return String(val);
      };
      const emailOf = (val: any): string => (val && typeof val === 'object') ? (val.EMail || '') : '';

      const mapped: IWorkflowRow[] = rawItems.map((item: any) => ({
        id: item.Id,
        workflowTitle: item.WorkflowTitle || '',
        processCode: item.Title || 'Untitled',
        workflowType: item.WorkflowType || '—',
        escalationCriteria: item.EscalationCriteria || '—',
        stepRole: item.StepRole || '—',
        l1Approver: nameOf(item.L1Approver),
        l1ApproverEmail: emailOf(item.L1Approver),
        l2Approver: nameOf(item.L2Approver),
        l2ApproverEmail: emailOf(item.L2Approver),
        escalateTo: nameOf(item.EscalateTo),
        escalateToEmail: emailOf(item.EscalateTo),
        finalApprover: nameOf(item.FinalApprover),
        finalApproverEmail: emailOf(item.FinalApprover),
        complianceCheck: item.ComplianceCheck || '—',
        deadlineDays: item.ApprovalDeadlineDays || 0,
        notifyApprovers: !!item.NotifyApproversEachStep,
        alertEscalation: !!item.AlertOnEscalation
      }));

      setWorkflowRows(mapped);
    } catch (err) {
      console.error("Error fetching Workflow Configuration list:", err);
      setWorkflowError('Could not load workflows. Check that a "Workflow Configuration" list exists with matching columns.');
    } finally {
      setWorkflowLoading(false);
    }
  };

  useEffect(() => {
    if (activeTab !== 'All Workflows') return;
    fetchWorkflows();
    const interval = setInterval(fetchWorkflows, 30000);
    return () => clearInterval(interval);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, context]);

  const [isCreateWorkflowOpen, setIsCreateWorkflowOpen] = useState(false);
  const [editingWorkflow, setEditingWorkflow] = useState<IWorkflowRow | null>(null);
  const [openWorkflowActionId, setOpenWorkflowActionId] = useState<number | null>(null);
  const [workflowModalMode, setWorkflowModalMode] = useState<'create' | 'edit' | 'view'>('create');

  const formatDate = (ms: number) => new Date(ms).toLocaleDateString();
  const dueDateFor = (doc: IDocumentItem) => doc.modifiedTime + ESCALATION_DAYS * 24 * 60 * 60 * 1000;
  const isOverdue = (doc: IDocumentItem) => daysSince(doc.modifiedTime) >= ESCALATION_DAYS;

  // -------------------- Small building blocks --------------------

  const StatCard: React.FC<{ icon: React.ReactNode; iconBg: string; label: string; value: number | string; sub: string }> = ({ icon, iconBg, label, value, sub }) => (
    <div style={styles.statCard}>
      <div style={{ ...styles.statIconBox, backgroundColor: iconBg }}>{icon}</div>
      <div style={{ minWidth: 0 }}>
        <span style={styles.statLabel}>{label}</span>
        <h3 style={styles.statNumber}>{value}</h3>
        <span style={styles.statSub}>{sub}</span>
      </div>
    </div>
  );

  const ActionRow: React.FC<{ doc: IDocumentItem }> = ({ doc }) => {
    const overdue = isOverdue(doc);
    const cs = doc.currentStage || '';
    const pillStyle = cs.indexOf('Rejected') !== -1
      ? { backgroundColor: '#fee2e2', color: '#b91c1c' }
      : cs.indexOf('Send for') !== -1
        ? { backgroundColor: '#fef3c7', color: '#b45309' }
        : cs.indexOf('Approved') !== -1
          ? { backgroundColor: '#dcfce7', color: '#15803d' }
          : { backgroundColor: '#f1f5f9', color: '#475569' };
    return (
      <div style={styles.approvalTableRow}>
        <div style={{ flex: '0.9', minWidth: 0 }}>
          <span style={styles.processTag}>{doc.processCode || 'N/A'}</span>
        </div>
        <div style={{ flex: '1.5', minWidth: 0, display: 'flex', alignItems: 'center', gap: '10px' }}>
          <div style={{ minWidth: 0 }}>
            <a href={fileOpenUrl(doc.serverRelativeUrl)} target="_blank" rel="noreferrer" style={styles.docName}>{doc.name}</a>
            <div style={styles.docMeta}>{doc.modifiedBy}</div>
          </div>
        </div>
        <div style={{ flex: '0.9' }}>
          <span style={{ ...styles.statusPill, backgroundColor: pillStyle.backgroundColor, color: pillStyle.color }}>{doc.currentStage}</span>
        </div>
        <div style={{ flex: '0.7', fontSize: '11.5px', color: overdue ? '#dc2626' : '#64748b', fontWeight: overdue ? 700 : 500 }}>
          {formatDate(dueDateFor(doc))}{overdue ? ' (Overdue)' : ''}
        </div>
        <div style={{ flex: '0.7', fontSize: '11.5px', color: '#64748b' }}>{doc.modified}</div>
      </div>
    );
  };

  // -------------------- Tab: My Approvals --------------------

  const MyApprovalsTab = (
    <>
      <div style={styles.statsGrid4Col}>
        <StatCard icon={<ClockIcon />} iconBg="#fef3c7" label="Pending Approvals" value={actionableDocs.length} sub="Require your action" />
        <StatCard icon={<AlertIcon />} iconBg="#fee2e2" label="Escalated Requests" value={escalatedDocs.length} sub="Needs immediate attention" />
        <StatCard icon={<ShieldIcon />} iconBg="#fef3c7" label="In Compliance Review" value={inReviewDocs.length} sub="Under compliance check" />
        <StatCard icon={<CheckIcon color="#059669" />} iconBg="#e6f4ea" label="Completed Approvals" value={completedDocs.length} sub="Successfully completed" />
      </div>

      <div style={styles.panelCard}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '10px', marginBottom: '14px' }}>
          <div>
            <h3 style={styles.title}>Approval Workflow</h3>
            <p style={styles.subtitle}>How a document moves from draft to approval</p>
          </div>
        </div>
        <div style={styles.stageFlowRow}>
          {STAGES.map((stage, idx) => {
            const isFinal = stage.status === 'Active';
            return (
              <React.Fragment key={stage.status}>
                <div style={styles.stageCard}>
                  <div style={{ ...styles.stageIconBox, backgroundColor: isFinal ? '#059669' : '#e6f4ea' }}>
                    {isFinal ? <CheckIcon color="#fff" size={16} /> : stage.status === 'In Review' ? <UsersIcon size={14} /> : <FileIcon size={14} />}
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                    <span style={styles.stageCardLabel}>{stage.label}</span>
                    <span style={styles.stageCardOwner}>{stage.owner}</span>
                  </div>
                </div>
                {idx < STAGES.length - 1 && <span style={styles.stageArrow}>→</span>}
              </React.Fragment>
            );
          })}
        </div>
      </div>

      <div style={styles.twoColGrid}>
        <div style={{ ...styles.panelCard, flex: '2.3 1 480px', minWidth: 0 }}>
          <div style={styles.panelHeaderRow}>
            <h3 style={styles.title}>Pending Approvals</h3>
            <span style={styles.countBadge}>{sortedActionable.length} awaiting action</span>
          </div>
          {sortedActionable.length === 0 ? (
            <p style={styles.emptyState}>Nothing pending — all documents are up to date.</p>
          ) : (
            <div>
              <div style={styles.approvalTableHeader}>
                <span style={{ flex: '0.9' }}>Process</span>
                <span style={{ flex: '1.5' }}>Document</span>
                <span style={{ flex: '0.9' }}>Stage</span>
                <span style={{ flex: '0.7' }}>Due</span>
                <span style={{ flex: '0.7' }}>Requested On</span>
              </div>
              {sortedActionable.map(doc => <ActionRow key={doc.id} doc={doc} />)}
            </div>
          )}
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: '0.9 1 260px', minWidth: 0 }}>
          <div style={styles.panelCard}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '10px' }}>
              <h3 style={{ ...styles.title, fontSize: '13px' }}>RACI Overview</h3>
              <InfoIcon />
            </div>
            {[
              { letter: 'R', color: '#0284c7', label: 'Responsible', desc: 'Performs the task' },
              { letter: 'A', color: '#059669', label: 'Accountable', desc: 'Accountable for outcome' },
              { letter: 'C', color: '#d97706', label: 'Consulted', desc: 'Provides input / expertise' },
              { letter: 'I', color: '#7c3aed', label: 'Informed', desc: 'Kept informed of progress' }
            ].map(r => (
              <div key={r.letter} style={styles.raciLegendRow}>
                <span style={{ ...styles.raciLetterBox, backgroundColor: r.color }}>{r.letter}</span>
                <div style={{ minWidth: 0 }}>
                  <div style={styles.raciLegendLabel}>{r.label}</div>
                  <div style={styles.raciLegendDesc}>{r.desc}</div>
                </div>
              </div>
            ))}
          </div>

          <div style={styles.panelCard}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
              <AlertIcon size={16} />
              <h3 style={{ ...styles.title, fontSize: '13px' }}>Escalated Requests</h3>
            </div>
            {escalatedDocs.length === 0 ? (
              <p style={styles.emptyState}>Nothing escalated — all pending items are within {ESCALATION_DAYS} days.</p>
            ) : (
              escalatedDocs.slice(0, 5).map(doc => (
                <div key={doc.id} style={styles.miniListRow}>
                  <a href={fileOpenUrl(doc.serverRelativeUrl)} target="_blank" rel="noreferrer" style={styles.miniListName}>{doc.name}</a>
                  <span style={styles.escalatedBadge}>Overdue</span>
                </div>
              ))
            )}
          </div>

          <div style={styles.panelCard}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
              <ShieldIcon size={16} />
              <h3 style={{ ...styles.title, fontSize: '13px' }}>In Compliance Review</h3>
            </div>
            {inReviewDocs.length === 0 ? (
              <p style={styles.emptyState}>Nothing currently under compliance review.</p>
            ) : (
              inReviewDocs.slice(0, 5).map(doc => (
                <div key={doc.id} style={styles.miniListRow}>
                  <a href={fileOpenUrl(doc.serverRelativeUrl)} target="_blank" rel="noreferrer" style={styles.miniListName}>{doc.name}</a>
                  <span style={styles.reviewBadge}>In Review</span>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </>
  );

  // -------------------- Tab: All Requests --------------------

  const AllRequestsTab = (
    <div style={styles.panelCard}>
      <div style={styles.panelHeaderRow}>
        <div>
          <h3 style={styles.title}>All Document Requests</h3>
          <p style={styles.subtitle}>{searchTerm ? `Filtered results for "${searchTerm}"` : 'Every document tracked in the library, live from SharePoint'}</p>
        </div>
        <span style={styles.countBadge}>Showing: {filteredDocuments.length} / {documents.length}</span>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '14px' }}>
        {filteredDocuments.length === 0 ? (
          <p style={styles.emptyState}>No matching documents found.</p>
        ) : (
          filteredDocuments.map(doc => {
            const pillStyle = getStatusStyle(doc.currentStage || '');
            const isHovered = hoveredRowId === doc.id;
            return (
              <div
                key={doc.id}
                onMouseEnter={() => setHoveredRowId(doc.id)}
                onMouseLeave={() => setHoveredRowId(null)}
                style={{
                  ...styles.row,
                  boxShadow: isHovered ? '0 6px 16px rgba(0, 92, 66, 0.08)' : '0 1px 2px rgba(15, 23, 42, 0.03)',
                  borderColor: isHovered ? '#bbf0d4' : '#e2e8f0'
                }}
              >
                <div style={styles.rowIconBox}><FileIcon /></div>
                <div style={{ flexGrow: 1, marginLeft: '12px', overflow: 'hidden' }}>
                  <a href={fileOpenUrl(doc.serverRelativeUrl)} target="_blank" rel="noreferrer" style={styles.docName}>{doc.name}</a>
                  <span style={styles.docMeta}>Uploaded by {doc.modifiedBy} • {doc.modified}</span>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '14px', flexShrink: 0 }}>
                  <span style={{ ...styles.statusPill, backgroundColor: pillStyle.backgroundColor, color: pillStyle.color }}>{doc.currentStage || '—'}</span>
                </div>
              </div>
            );
          })
        )}
      </div>
    </div>
  );

  // -------------------- Tab: RACI Matrix --------------------

  const RaciMatrixTab = (
    <div style={styles.panelCard}>
      <div style={styles.panelHeaderRow}>
        <div>
          <h3 style={styles.title}>RACI Matrix</h3>
          <p style={styles.subtitle}>Live from the "RACI Matrix" SharePoint list</p>
        </div>
        <span style={styles.countBadge}>{raciRows.length} entries</span>
      </div>

      {raciLoading && raciRows.length === 0 ? (
        <p style={styles.emptyState}>Loading RACI Matrix…</p>
      ) : raciError ? (
        <p style={{ ...styles.emptyState, color: '#dc2626' }}>{raciError}</p>
      ) : raciRows.length === 0 ? (
        <p style={styles.emptyState}>No RACI entries found.</p>
      ) : (
        <div style={{ overflowX: 'auto' as const }}>
          <table style={styles.raciTable}>
            <thead>
              <tr>
                <th style={styles.raciTh}>Process</th>
                <th style={styles.raciTh}>Responsible</th>
                <th style={styles.raciTh}>Accountable</th>
                <th style={styles.raciTh}>Consulted</th>
                <th style={styles.raciTh}>Informed</th>
              </tr>
            </thead>
            <tbody>
              {raciRows.map(row => (
                <tr key={row.id}>
                  <td style={{ ...styles.raciTd, fontWeight: 600, color: '#1e293b' }}>{row.processName}</td>
                  <td style={styles.raciTd}>{row.responsible}</td>
                  <td style={styles.raciTd}>{row.accountable}</td>
                  <td style={styles.raciTd}>{row.consulted}</td>
                  <td style={styles.raciTd}>{row.informed}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  // -------------------- Tab: All Workflows --------------------

  const AllWorkflowsTab = (
    <div style={styles.panelCard}>
      <div style={styles.panelHeaderRow}>
        <div>
          <h3 style={styles.title}>All Workflows</h3>
          <p style={styles.subtitle}>Live from the "Workflow Configuration" SharePoint list</p>
        </div>
        <span style={styles.countBadge}>{workflowRows.length} workflows</span>
      </div>

      {workflowLoading && workflowRows.length === 0 ? (
        <p style={styles.emptyState}>Loading workflows…</p>
      ) : workflowError ? (
        <p style={{ ...styles.emptyState, color: '#dc2626' }}>{workflowError}</p>
      ) : workflowRows.length === 0 ? (
        <p style={styles.emptyState}>No workflows created yet.</p>
      ) : (
        <div style={{ overflowX: 'auto' as const }}>
          <table style={styles.raciTable}>
            <thead>
              <tr>
                <th style={styles.raciTh}>Process Code</th>
                <th style={styles.raciTh}>Type</th>
                <th style={styles.raciTh}>Escalation Criteria</th>
                <th style={styles.raciTh}>Role</th>
                <th style={styles.raciTh}>L1 Approver</th>
                <th style={styles.raciTh}>L2 Approver</th>
                <th style={styles.raciTh}>Escalate To</th>
                <th style={styles.raciTh}>Final Approver</th>
                <th style={styles.raciTh}>Compliance</th>
                <th style={styles.raciTh}>Deadline</th>
                <th style={{ ...styles.raciTh, textAlign: 'right' as const }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {workflowRows.map(row => (
                <tr key={row.id}>
                  <td style={{ ...styles.raciTd, fontWeight: 600, color: '#1e293b' }}>{row.processCode}</td>
                  <td style={styles.raciTd}>{row.workflowType}</td>
                  <td style={styles.raciTd}>{row.escalationCriteria}</td>
                  <td style={styles.raciTd}>{row.stepRole}</td>
                  <td style={styles.raciTd}>{row.l1Approver}</td>
                  <td style={styles.raciTd}>{row.l2Approver}</td>
                  <td style={styles.raciTd}>{row.escalateTo}</td>
                  <td style={styles.raciTd}>{row.finalApprover}</td>
                  <td style={styles.raciTd}>
                    <span style={{ ...styles.statusPill, backgroundColor: row.complianceCheck === 'Required' ? '#fef3c7' : '#e6f4ea', color: row.complianceCheck === 'Required' ? '#92400e' : '#005c42' }}>{row.complianceCheck}</span>
                  </td>
                  <td style={styles.raciTd}>{row.deadlineDays} day(s)</td>
                  <td style={{ ...styles.raciTd, position: 'relative' as const, textAlign: 'right' as const }}>
                    <button
                      type="button"
                      className="aw-action-btn"
                      onClick={() => setOpenWorkflowActionId(openWorkflowActionId === row.id ? null : row.id)}
                      style={styles.threeDotBtn}
                      title="Actions"
                    >
                      <MoreIcon />
                    </button>
                    {openWorkflowActionId === row.id && (
                      <div style={styles.actionsDropdown}>
                        <div
                          className="aw-menu-item"
                          style={styles.actionMenuItem}
                          onClick={() => {
                            setEditingWorkflow(row);
                            setWorkflowModalMode('view');
                            setOpenWorkflowActionId(null);
                            setIsCreateWorkflowOpen(true);
                          }}
                        >
                          <EyeIcon />
                          <span>View</span>
                        </div>
                        <div
                          className="aw-menu-item"
                          style={styles.actionMenuItem}
                          onClick={() => {
                            setEditingWorkflow(row);
                            setWorkflowModalMode('edit');
                            setOpenWorkflowActionId(null);
                            setIsCreateWorkflowOpen(true);
                          }}
                        >
                          <PencilIcon />
                          <span>Edit</span>
                        </div>
                      </div>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );

  // -------------------- Tab: Approval Rules --------------------

  const ApprovalRulesTab = (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
      <div style={styles.panelCard}>
        <h3 style={styles.title}>Escalation Policy</h3>
        <p style={styles.subtitle}>
          Documents left in <b>Pending Approval</b> or <b>In Review</b> for {ESCALATION_DAYS}+ days are
          automatically surfaced under Escalated Requests until actioned.
        </p>
      </div>

      <div style={styles.panelCard}>
        <div style={styles.panelHeaderRow}>
          <div>
            <h3 style={styles.title}>Process Master</h3>
            <p style={styles.subtitle}>Live from the "Process Master" SharePoint list</p>
          </div>
          <span style={styles.countBadge}>{processes.length} processes</span>
        </div>
        {processesLoading && processes.length === 0 ? (
          <p style={styles.emptyState}>Loading processes…</p>
        ) : processes.length === 0 ? (
          <p style={styles.emptyState}>No processes found in Process Master.</p>
        ) : (
          <div style={{ overflowX: 'auto' as const }}>
            <table style={styles.raciTable}>
              <thead>
                <tr>
                  <th style={styles.raciTh}>Code</th>
                  <th style={styles.raciTh}>Process Name</th>
                  <th style={styles.raciTh}>Department</th>
                  <th style={styles.raciTh}>Status</th>
                </tr>
              </thead>
              <tbody>
                {processes.map(p => (
                  <tr key={p.id}>
                    <td style={{ ...styles.raciTd, fontWeight: 600, color: '#1e293b' }}>{p.processCode}</td>
                    <td style={styles.raciTd}>{p.processName}</td>
                    <td style={styles.raciTd}>{p.department}</td>
                    <td style={styles.raciTd}>
                      <span style={{ ...styles.statusPill, backgroundColor: '#e6f4ea', color: '#005c42' }}>{p.status}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );

  const TABS: { key: typeof activeTab; label: string }[] = [
    { key: 'My Approvals', label: 'My Approvals' },
    { key: 'All Requests', label: 'All Requests' },
    { key: 'All Workflows', label: 'All Workflows' },
    { key: 'RACI Matrix', label: 'RACI Matrix' },
    { key: 'Approval Rules', label: 'Approval Rules' }
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
      <style>{`
        .aw-action-btn:hover { background-color: #e2e8f0; }
        .aw-menu-item:hover { background-color: #f1f5f9; color: #004d3d; }
      `}</style>
      {toastMessage && (
        <div style={styles.toast}>
          <CheckIcon color="#fff" size={15} /> {toastMessage}
        </div>
      )}

      <div style={styles.pageHeaderRow}>
        <div>
          <h2 style={styles.pageTitle}>Approval Management</h2>
          <p style={styles.pageSubtitle}>Manage and Track Approvals Efficiently</p>
        </div>
        <button type="button" style={styles.createWorkflowBtn} onClick={() => { setEditingWorkflow(null); setWorkflowModalMode('create'); setIsCreateWorkflowOpen(true); }}>
          <PlusIcon /> Create Workflow
        </button>
      </div>

      <div style={styles.tabRow}>
        {TABS.map(tab => (
          <button
            type="button"
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            style={activeTab === tab.key ? styles.tabItemActive : styles.tabItem}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {activeTab === 'My Approvals' && MyApprovalsTab}
      {activeTab === 'All Requests' && AllRequestsTab}
      {activeTab === 'All Workflows' && AllWorkflowsTab}
      {activeTab === 'RACI Matrix' && RaciMatrixTab}
      {activeTab === 'Approval Rules' && ApprovalRulesTab}

      <CreateWorkflow
        isOpen={isCreateWorkflowOpen}
        onClose={() => { setIsCreateWorkflowOpen(false); setEditingWorkflow(null); setWorkflowModalMode('create'); }}
        context={context}
        editItem={editingWorkflow}
        isViewMode={workflowModalMode === 'view'}
        onSaved={() => {
          setToastMessage(editingWorkflow ? 'Workflow updated successfully!' : 'Data successfully saved to SharePoint!');
          setEditingWorkflow(null);
          setWorkflowModalMode('create');
          fetchWorkflows();
        }}
      />
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  pageHeaderRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '12px' },
  toast: { position: 'fixed', top: '20px', right: '20px', backgroundColor: '#059669', color: '#fff', padding: '12px 18px', borderRadius: '10px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 600, boxShadow: '0 10px 25px rgba(0,0,0,0.15)', zIndex: 10000 },
  pageTitle: { margin: 0, fontSize: '18px', fontWeight: 700, color: '#0f2b22' },
  pageSubtitle: { margin: '4px 0 0 0', color: '#64748b', fontSize: '12px' },
  createWorkflowBtn: { display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#005c42', color: '#fff', border: 'none', borderRadius: '8px', padding: '8px 14px', fontSize: '12px', fontWeight: 700, cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif' },

  tabRow: { display: 'flex', gap: '4px', backgroundColor: '#f1f5f9', borderRadius: '8px', padding: '3px', width: 'fit-content', flexWrap: 'wrap' },
  tabItem: { border: 'none', backgroundColor: 'transparent', color: '#64748b', fontSize: '11.5px', fontWeight: 600, padding: '6px 14px', borderRadius: '6px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif' },
  tabItemActive: { border: 'none', backgroundColor: '#fff', color: '#005c42', fontSize: '11.5px', fontWeight: 700, padding: '6px 14px', borderRadius: '6px', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif', boxShadow: '0 1px 3px rgba(15, 23, 42, 0.1)' },

  statsGrid4Col: { display: 'grid', gridTemplateColumns: 'repeat(4, minmax(0, 1fr))', gap: '16px' },
  statCard: { backgroundColor: '#fff', padding: '16px', borderRadius: '12px', border: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '12px', boxSizing: 'border-box' },
  statIconBox: { width: '42px', height: '42px', borderRadius: '8px', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  statLabel: { fontSize: '11px', color: '#64748b', display: 'block' },
  statNumber: { margin: '2px 0', fontSize: '18px', fontWeight: 700, color: '#0f2b22' },
  statSub: { fontSize: '10px', color: '#94a3b8' },

  panelCard: { backgroundColor: '#fff', borderRadius: '12px', border: '1px solid #e2e8f0', padding: '16px', boxSizing: 'border-box' },
  panelHeaderRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '10px', marginBottom: '12px' },
  title: { margin: 0, fontSize: '14px', fontWeight: 700, color: '#0f2b22' },
  subtitle: { margin: '4px 0 0 0', color: '#64748b', fontSize: '11.5px' },
  countBadge: { fontSize: '11px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '5px 11px', borderRadius: '8px', flexShrink: 0, whiteSpace: 'nowrap' },
  emptyState: { fontSize: '12px', color: '#94a3b8', padding: '20px', textAlign: 'center' },

  stageFlowRow: { display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', marginTop: '4px' },
  stageCard: { flex: '1 1 150px', display: 'flex', flexDirection: 'row', alignItems: 'center', gap: '9px', border: '1px solid #e2e8f0', borderRadius: '10px', padding: '8px 10px', backgroundColor: '#f8fafc', textAlign: 'left' },
  stageIconBox: { width: '26px', height: '26px', borderRadius: '7px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  stageCardLabel: { fontSize: '11px', fontWeight: 700, color: '#1e293b', lineHeight: 1.2, display: 'block', marginBottom: '3px' },
  stageCardOwner: { fontSize: '9.5px', color: '#94a3b8', lineHeight: 1.2 },
  stageCardCount: { fontSize: '13px', fontWeight: 700, color: '#005c42' },
  stageArrow: { fontSize: '13px', color: '#cbd5e1', flexShrink: 0 },

  twoColGrid: { display: 'flex', gap: '16px', alignItems: 'flex-start', flexWrap: 'wrap' as const },

  approvalTableHeader: { display: 'flex', gap: '10px', padding: '0 10px 8px 10px', borderBottom: '1px solid #e2e8f0', fontSize: '10px', fontWeight: 700, color: '#94a3b8', textTransform: 'uppercase', letterSpacing: '0.3px' },
  approvalTableRow: { display: 'flex', gap: '10px', alignItems: 'center', padding: '10px', borderBottom: '1px solid #f1f5f9' },
  rowIconBoxSm: { width: '28px', height: '28px', borderRadius: '7px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },


  raciLegendRow: { display: 'flex', alignItems: 'center', gap: '10px', padding: '7px 0', borderBottom: '1px solid #f1f5f9' },
  raciLetterBox: { width: '24px', height: '24px', borderRadius: '6px', color: '#fff', fontSize: '11px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  raciLegendLabel: { fontSize: '11.5px', fontWeight: 700, color: '#1e293b' },
  raciLegendDesc: { fontSize: '10px', color: '#94a3b8' },

  miniListRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px', padding: '7px 0', borderBottom: '1px solid #f1f5f9' },
  miniListName: { fontSize: '11.5px', fontWeight: 600, color: '#1e293b', textDecoration: 'none', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flexGrow: 1, minWidth: 0 },
  escalatedBadge: { fontSize: '9.5px', fontWeight: 700, color: '#dc2626', backgroundColor: '#fee2e2', padding: '3px 8px', borderRadius: '10px', flexShrink: 0 },
  reviewBadge: { fontSize: '9.5px', fontWeight: 700, color: '#0284c7', backgroundColor: '#e0f2fe', padding: '3px 8px', borderRadius: '10px', flexShrink: 0 },

  raciTable: { width: '100%', borderCollapse: 'collapse', fontSize: '12px' },
  raciTh: { textAlign: 'left', padding: '7px 10px', fontSize: '10px', color: '#94a3b8', fontWeight: 700, borderBottom: '1px solid #e2e8f0', textTransform: 'uppercase', whiteSpace: 'nowrap' },
  raciTd: { padding: '9px 10px', borderBottom: '1px solid #f1f5f9', color: '#475569', fontSize: '12px' },

  row: {
    display: 'flex',
    alignItems: 'center',
    padding: '12px 14px',
    borderRadius: '10px',
    border: '1px solid #e2e8f0',
    backgroundColor: '#fff',
    transition: 'box-shadow 0.15s ease, border-color 0.15s ease',
    boxSizing: 'border-box'
  },
  rowIconBox: { width: '34px', height: '34px', borderRadius: '9px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  docName: { textDecoration: 'none', fontWeight: 600, fontSize: '12.5px', color: '#1e293b', display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  docMeta: { fontSize: '11px', color: '#64748b' },
  statusPill: { fontSize: '9.5px', fontWeight: 700, padding: '3px 9px', borderRadius: '10px', textTransform: 'uppercase', flexShrink: 0, letterSpacing: '0.2px' },
  threeDotBtn: { background: 'transparent', border: 'none', borderRadius: '4px', padding: '4px 6px', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', transition: 'background-color 0.15s ease' },
  processTag: { fontSize: '10px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '3px 8px', borderRadius: '7px', whiteSpace: 'nowrap' as const, display: 'inline-block' },
  dropdownLinkAction: {
    display: 'flex', alignItems: 'center', gap: '9px', width: '100%', padding: '9px 13px',
    borderBottom: '1px solid #f1f5f9', textAlign: 'left' as const, fontSize: '11.5px', fontWeight: 600,
    color: '#334155', cursor: 'pointer', fontFamily: '"Segoe UI", Arial, sans-serif',
    textDecoration: 'none', boxSizing: 'border-box' as const
  },
  dropdownMenu: {
    position: 'absolute',
    right: 0,
    top: 'calc(100% + 6px)',
    backgroundColor: '#fff',
    border: '1px solid #e2e8f0',
    borderRadius: '10px',
    boxShadow: '0 10px 25px rgba(15, 23, 42, 0.12)',
    zIndex: 100,
    width: '140px',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column'
  },
  actionsDropdown: {
    position: 'absolute',
    right: 0,
    top: 'calc(100% + 4px)',
    backgroundColor: '#ffffff',
    border: '1px solid #e2e8f0',
    borderRadius: '6px',
    boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)',
    zIndex: 200,
    padding: '4px',
    minWidth: '90px',
    textAlign: 'left'
  },
  actionMenuItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    padding: '6px 8px',
    fontSize: '11px',
    borderRadius: '4px',
    cursor: 'pointer',
    color: '#334155'
  }
};

export default DocumentApproval;
