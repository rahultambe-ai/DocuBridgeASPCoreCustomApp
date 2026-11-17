import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import CreateWorkflow from './CreateWorkflow';

export interface IProcessWiseWorkflowProps {
  isOpen: boolean;
  onClose: () => void;
  context: any;
  processCode: string;
}

interface IProcessWorkflowRow {
  id: number;
  workflowId: string;
  workflowTitle: string;
  processCode: string;
  workflowType: string;
  escalationCriteria: string;
  l1Approver: string;
  l2Approver: string;
  escalateTo: string;
  finalApprover: string;
  deadlineDays: number;
  notifyApprovers: boolean;
  alertEscalation: boolean;
}

// style as CreateWorkflow

const WorkflowIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#005c42" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" />
    <rect x="3" y="15" width="6" height="6" rx="1" /><path d="M9 6h6a3 3 0 0 1 3 3v6" /><path d="M9 18h3" />
  </svg>
);

const PlusIcon = () => (
  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round">
    <line x1="12" y1="5" x2="12" y2="19" />
    <line x1="5" y1="12" x2="19" y2="12" />
  </svg>
);

const ProcessWiseWorkflow: React.FC<IProcessWiseWorkflowProps> = ({ isOpen, onClose, context, processCode }) => {
  const [rows, setRows] = useState<IProcessWorkflowRow[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [loadError, setLoadError] = useState<string>('');
  const [isCreateOpen, setIsCreateOpen] = useState<boolean>(false);

  const fetchProcessWorkflows = async () => {
    if (!context || !processCode) return;
    setLoading(true);
    setLoadError('');
    try {
      const sp = spfi().using(SPFx(context));
      const safeCode = processCode.replace(/'/g, "''");
      let rawItems: any[] = [];

      try {
        rawItems = await sp.web.lists.getByTitle("Workflow Configuration").items
          .select(
            "Id", "Title", "WorkflowID", "WorkflowTitle", "WorkflowType", "EscalationCriteria",
            "L1Approver/Title", "L2Approver/Title", "EscalateTo/Title", "FinalApprover/Title",
            "ApprovalDeadlineDays", "NotifyApproversEachStep", "AlertOnEscalation"
          )
          .expand("L1Approver", "L2Approver", "EscalateTo", "FinalApprover")
          .filter(`Title eq '${safeCode}'`)
          .top(200)();
      } catch (expandErr) {
        rawItems = await sp.web.lists.getByTitle("Workflow Configuration").items
          .filter(`Title eq '${safeCode}'`)
          .top(200)();
      }

      const nameOf = (val: any): string => {
        if (val === undefined || val === null || val === '') return '—';
        if (typeof val === 'object') return val.Title || '—';
        return String(val);
      };

      const mapped: IProcessWorkflowRow[] = rawItems.map((item: any) => ({
        id: item.Id,
        workflowId: item.WorkflowID || '—',
        workflowTitle: item.WorkflowTitle || '—',
        processCode: item.Title || processCode,
        workflowType: item.WorkflowType || '—',
        escalationCriteria: item.EscalationCriteria || '—',
        l1Approver: nameOf(item.L1Approver),
        l2Approver: nameOf(item.L2Approver),
        escalateTo: nameOf(item.EscalateTo),
        finalApprover: nameOf(item.FinalApprover),
        deadlineDays: item.ApprovalDeadlineDays || 0,
        notifyApprovers: !!item.NotifyApproversEachStep,
        alertEscalation: !!item.AlertOnEscalation
      }));

      setRows(mapped);
    } catch (err) {
      console.error('Error fetching workflows for process', processCode, err);
      setLoadError('Could not load workflows for this process.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (isOpen) fetchProcessWorkflows();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, processCode]);

  if (!isOpen) return null;

  return (
    <>
      <div className="pww-overlay" style={styles.overlay} onClick={onClose}>
        <style>{`
          @keyframes pwwFadeIn { from { opacity: 0; } to { opacity: 1; } }
          @keyframes pwwScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
          .pww-overlay { animation: pwwFadeIn 0.18s ease; }
          .pww-modal { animation: pwwScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
          .pww-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; }
          .pww-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
          .pww-create-btn { transition: transform 0.15s ease, filter 0.15s ease; }
          .pww-create-btn:hover { transform: translateY(-1px); filter: brightness(1.06); }
          .pww-body::-webkit-scrollbar { height: 8px; width: 8px; }
          .pww-body::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 8px; }
          .pww-body::-webkit-scrollbar-track { background-color: transparent; }
        `}</style>

        <div className="pww-modal" style={styles.modalCard} onClick={e => e.stopPropagation()}>
          <div style={styles.header}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: '12px' }}>
              <div style={styles.headerIconBox}><WorkflowIcon /></div>
              <div>
                <h2 style={styles.title}>Workflows — {processCode}</h2>
                <p style={styles.subtitle}>All approval / escalation workflows configured for this process</p>
              </div>
            </div>
            <button type="button" className="pww-close-btn" style={styles.closeBtn} onClick={onClose} title="Close">✕</button>
          </div>

          <div className="pww-body" style={styles.body}>
            <div style={styles.toolbar}>
              <span style={styles.countBadge}>{rows.length} workflow{rows.length === 1 ? '' : 's'}</span>
              <button type="button" className="pww-create-btn" style={styles.createBtn} onClick={() => setIsCreateOpen(true)}>
                <PlusIcon /> Create Workflow
              </button>
            </div>

            {loadError && <div style={styles.errorBox}>{loadError}</div>}

            <div style={styles.tableScroll}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Process Code</th>
                    <th style={styles.th}>WorkflowID</th>
                    <th style={styles.th}>Workflow Title</th>
                    <th style={styles.th}>Workflow Type</th>
                    <th style={styles.th}>Escalation Criteria</th>
                    <th style={styles.th}>L1 Approver</th>
                    <th style={styles.th}>L2 Approver</th>
                    <th style={styles.th}>Escalate To</th>
                    <th style={styles.th}>Final Approver</th>
                    <th style={styles.th}>Deadline (Days)</th>
                    <th style={styles.th}>Notify Approvers</th>
                    <th style={styles.th}>Alert on Escalation</th>
                  </tr>
                </thead>
                <tbody>
                  {loading ? (
                    <tr><td style={{ ...styles.td, textAlign: 'center', color: '#64748b' }} colSpan={12}>Loading workflows...</td></tr>
                  ) : rows.length === 0 ? (
                    <tr><td style={{ ...styles.td, textAlign: 'center', color: '#94a3b8' }} colSpan={12}>No workflows found for this process yet.</td></tr>
                  ) : rows.map(r => (
                    <tr key={r.id} style={styles.tr}>
                      <td style={styles.td}><span style={styles.codeTag}>{r.processCode}</span></td>
                      <td style={styles.td}>{r.workflowId}</td>
                      <td style={styles.td}>{r.workflowTitle}</td>
                      <td style={styles.td}>{r.workflowType}</td>
                      <td style={styles.td}>{r.escalationCriteria}</td>
                      <td style={styles.td}>{r.l1Approver}</td>
                      <td style={styles.td}>{r.l2Approver}</td>
                      <td style={styles.td}>{r.escalateTo}</td>
                      <td style={styles.td}>{r.finalApprover}</td>
                      <td style={styles.td}>{r.deadlineDays}</td>
                      <td style={styles.td}>{r.notifyApprovers ? 'Yes' : 'No'}</td>
                      <td style={styles.td}>{r.alertEscalation ? 'Yes' : 'No'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div style={styles.footer}>
            <button type="button" className="pww-close-btn" style={styles.closeFooterBtn} onClick={onClose}>Close</button>
          </div>
        </div>
      </div>

      <CreateWorkflow
        isOpen={isCreateOpen}
        onClose={() => setIsCreateOpen(false)}
        context={context}
        defaultProcessCode={processCode}
        onSaved={() => {
          setIsCreateOpen(false);
          fetchProcessWorkflows();
        }}
      />
    </>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 9998, padding: '28px 16px', boxSizing: 'border-box' },
  modalCard: { backgroundColor: '#fff', borderRadius: '16px', width: '100%', maxWidth: '1240px', maxHeight: '90vh', display: 'flex', flexDirection: 'column', boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)', overflow: 'hidden', border: '1px solid #e2e8f0', margin: 'auto' },

  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', padding: '18px 20px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#fff' },
  headerIconBox: { width: '38px', height: '38px', borderRadius: '10px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '3px 0 0 0', fontSize: '12px', color: '#64748b' },
  closeBtn: { background: 'none', border: 'none', fontSize: '15px', cursor: 'pointer', color: '#64748b', padding: '5px 9px', borderRadius: '7px', flexShrink: 0 },

  body: { padding: '18px 20px', overflow: 'auto', display: 'flex', flexDirection: 'column', gap: '14px', backgroundColor: '#f8fafc' },

  toolbar: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '10px' },
  countBadge: { fontSize: '11px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '5px 11px', borderRadius: '8px', flexShrink: 0, whiteSpace: 'nowrap' },
  createBtn: { display: 'flex', alignItems: 'center', gap: '7px', padding: '8px 16px', borderRadius: '8px', border: 'none', backgroundColor: '#005c42', color: '#fff', fontSize: '12.5px', fontWeight: 700, cursor: 'pointer' },

  errorBox: { fontSize: '11.5px', color: '#dc2626', fontWeight: 600, backgroundColor: '#fef2f2', border: '1px solid #fecaca', borderRadius: '8px', padding: '8px 12px' },

  tableScroll: { backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '10px', overflow: 'auto' },
  table: { width: '100%', borderCollapse: 'collapse', textAlign: 'left' },
  th: { padding: '9px 12px', fontSize: '10px', fontWeight: 700, color: '#64748b', letterSpacing: '0.04em', whiteSpace: 'nowrap', borderBottom: '1px solid #e2e8f0', backgroundColor: '#fff', textTransform: 'uppercase' },
  tr: { borderBottom: '1px solid #f1f5f9' },
  td: { padding: '9px 12px', fontSize: '12px', color: '#334155', whiteSpace: 'nowrap' },
  codeTag: { fontSize: '10px', fontWeight: 700, color: '#005c42', backgroundColor: '#e6f4ea', padding: '3px 8px', borderRadius: '7px', whiteSpace: 'nowrap' as const, display: 'inline-block' },

  footer: { display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: '10px', padding: '14px 20px', borderTop: '1px solid #e2e8f0', backgroundColor: '#fff' },
  closeFooterBtn: { padding: '9px 18px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#fff', color: '#475569', fontSize: '12.5px', fontWeight: 700, cursor: 'pointer' }
};

export default ProcessWiseWorkflow;
