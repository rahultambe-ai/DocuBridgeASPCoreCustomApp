import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx, photoUrlFor } from '../../spfx-shim';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { SPHttpClient, SPHttpClientResponse } from '../../spfx-shim';

// Popup opened from the document list's 3-dot menu ("Assign Workflow" for a Draft document
// that has no workflow yet). The workflows offered here are the rows of the "Workflow
// Configuration" list that belong to the document's process — the same list the Upload
// screen's "Select Workflow" reads.

const WORKFLOW_CONFIG_LIST_TITLE = 'Workflow Configuration';
const DOCUMENT_LIBRARY_TITLE = 'Documents';

interface IWorkflowOption {
  id: number;
  workflowId: string;
  workflowTitle: string;
  workflowType: string;
  l1Approver: string;
  l2Approver: string;
  escalateTo: string;
  finalApprover: string;
  deadlineDays: number;
}

export interface IAssignWorkflowProps {
  isOpen: boolean;
  context: any;
  /** Library list item Id of the document. */
  documentId: number;
  documentTitle: string;
  processCode: string;
  onClose: () => void;
  /** Called after the document was successfully updated in SharePoint. */
  onAssigned: (workflowTitle: string) => void;
}

const initialsOf = (name: string): string => {
  const parts = (name || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

// Shows the approver's real Microsoft 365 profile picture (SharePoint's userphoto.aspx).
// While the photo is still resolving, or if the person has none / the image fails to load,
// it quietly falls back to the green initials circle so the UI never breaks.
const PersonAvatar = ({ name, photoUrl }: { name: string; photoUrl?: string }) => {
  const [imgFailed, setImgFailed] = useState(false);
  useEffect(() => { setImgFailed(false); }, [photoUrl]);
  if (photoUrl && !imgFailed) {
    return (
      <img
        src={photoUrl}
        alt={name}
        onError={() => setImgFailed(true)}
        style={{ width: '28px', height: '28px', borderRadius: '50%', objectFit: 'cover', flexShrink: 0, backgroundColor: '#e2e8f0' }}
      />
    );
  }
  return <span style={styles.stepAvatar}>{initialsOf(name)}</span>;
};

const AssignWorkflow: React.FC<IAssignWorkflowProps> = ({
  isOpen,
  context,
  documentId,
  documentTitle,
  processCode,
  onClose,
  onAssigned
}) => {
  const [workflows, setWorkflows] = useState<IWorkflowOption[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [loadError, setLoadError] = useState<string>('');
  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [saving, setSaving] = useState<boolean>(false);
  const [saveError, setSaveError] = useState<string>('');
  const [personPhotos, setPersonPhotos] = useState<{ [name: string]: string }>({});

  // Looks each approver up in the site's user list (same lookup the Upload screen uses) and
  // builds the live profile-photo URL from their email / login name.
  const resolvePersonPhotos = async (names: string[]) => {
    if (!context?.spHttpClient || !context?.pageContext) return;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const toResolve = Array.from(new Set(names)).filter(n => n && n !== 'Not Assigned');
    if (!toResolve.length) return;

    await Promise.all(
      toResolve.map(async (personName) => {
        try {
          const safeName = personName.replace(/'/g, "''");
          const endpoint = `${siteUrl}/_api/web/siteusers?$filter=Title eq '${safeName}'&$select=Email,LoginName&$top=1`;
          const response: SPHttpClientResponse = await context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1,
            { headers: { Accept: 'application/json;odata=nometadata' } }
          );
          if (!response.ok) return;
          const data = await response.json();
          const user = data?.value?.[0];
          const loginName: string = user?.LoginName || '';
          const upn: string = user?.Email || (loginName.indexOf('|') !== -1 ? (loginName.split('|').pop() || '') : loginName) || '';
          if (!upn) return;
          const photoUrl = photoUrlFor(upn);
          setPersonPhotos(prev => ({ ...prev, [personName]: photoUrl }));
        } catch (err) {
          // That person just keeps the initials avatar.
          console.warn('AssignWorkflow: could not resolve photo for', personName, err);
        }
      })
    );
  };

  // Load the workflows of this document's process every time the popup opens.
  useEffect(() => {
    if (!isOpen) return;

    setSaveError('');
    setLoadError('');
    setSaving(false);
    setWorkflows([]);
    setSelectedId(null);

    if (!context?.spHttpClient || !context?.pageContext || !processCode.trim()) {
      setLoadError('Could not work out which process this document belongs to.');
      return;
    }

    let cancelled = false;
    const load = async () => {
      setLoading(true);
      try {
        const siteUrl = context.pageContext.web.absoluteUrl;
        const safeCode = processCode.trim().replace(/'/g, "''");
        const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${WORKFLOW_CONFIG_LIST_TITLE}')/items`
          + `?$filter=Title eq '${safeCode}'`
          + `&$select=Id,Title,WorkflowID,WorkflowTitle,WorkflowType,ApprovalDeadlineDays,L1Approver/Title,L2Approver/Title,EscalateTo/Title,FinalApprover/Title`
          + `&$expand=L1Approver,L2Approver,EscalateTo,FinalApprover&$top=50`;

        const response: SPHttpClientResponse = await context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          { headers: { Accept: 'application/json;odata=nometadata' } }
        );

        if (!response.ok) {
          throw new Error(await response.text());
        }

        const data = await response.json();
        const nameOf = (val: any): string => (val && typeof val === 'object') ? (val.Title || '') : '';

        // A row with no WorkflowID can't be stamped on a document, so it isn't offered.
        const mapped: IWorkflowOption[] = (data.value || [])
          .filter((item: any) => !!item.WorkflowID)
          .map((item: any) => ({
            id: item.Id,
            workflowId: item.WorkflowID,
            workflowTitle: item.WorkflowTitle || item.WorkflowType || 'Untitled Workflow',
            workflowType: item.WorkflowType || 'Approval Workflow',
            l1Approver: nameOf(item.L1Approver),
            l2Approver: nameOf(item.L2Approver),
            escalateTo: nameOf(item.EscalateTo),
            finalApprover: nameOf(item.FinalApprover),
            deadlineDays: item.ApprovalDeadlineDays || 0
          }));

        if (cancelled) return;
        setWorkflows(mapped);

        // Real profile photos for every approver across the workflows found.
        const approverNames = mapped
          .reduce((acc: string[], w) => acc.concat([w.l1Approver, w.l2Approver, w.escalateTo, w.finalApprover]), [])
          .filter(Boolean);
        resolvePersonPhotos(approverNames);
      } catch (err) {
        console.error('AssignWorkflow: could not load workflows for', processCode, err);
        if (!cancelled) setLoadError('Could not load workflows for this process. Please try again.');
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    load();
    return () => { cancelled = true; };
  }, [isOpen, processCode, context]);

  // Esc closes the popup (unless a save is in flight).
  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !saving) onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [isOpen, saving, onClose]);

  if (!isOpen) return null;

  const selected = workflows.find(w => w.id === selectedId) || null;
  const canSubmit = !!selected && !saving;

  const handleAssign = async () => {
    if (!selected) return;
    setSaving(true);
    setSaveError('');
    try {
      const sp = spfi().using(SPFx(context));
      // Same values the Upload screen writes when a workflow is picked: the document enters
      // the approval chain at L1 (this is what the Power Automate flow watches for).
      await sp.web.lists.getByTitle(DOCUMENT_LIBRARY_TITLE).items.getById(documentId).update({
        WorkflowID: selected.workflowId,
        DocumentStatus: 'Pending Approval',
        DocumentCurrentStatus: 'Send for L1'
      });
      onAssigned(selected.workflowTitle);
    } catch (err) {
      console.error('AssignWorkflow: update failed for document', documentId, err);
      setSaveError('Could not save the workflow on this document. Please try again.');
      setSaving(false);
    }
  };

  const stepsFor = (w: IWorkflowOption) => [
    { key: 'l1', title: 'L1 Approval', person: w.l1Approver || 'Pending Assignment' },
    { key: 'l2', title: 'L2 Approval', person: w.l2Approver || 'Pending Assignment' },
    ...(w.workflowType === 'Escalation Workflow'
      ? [{ key: 'esc', title: 'Escalation / Compliance', person: w.escalateTo || 'Compliance Officer' }]
      : []),
    { key: 'final', title: 'Final Approval', person: w.finalApprover || 'Final Approver' }
  ];

  return (
    <div style={styles.overlay} onClick={() => { if (!saving) onClose(); }}>
      <div style={styles.modal} onClick={(e) => e.stopPropagation()} role="dialog" aria-modal="true">
        {/* Header */}
        <div style={styles.header}>
          <div style={styles.headerIcon}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ffffff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <circle cx="6" cy="6" r="2.5" /><circle cx="18" cy="12" r="2.5" /><circle cx="6" cy="18" r="2.5" />
              <path d="M8.5 6h4a3 3 0 0 1 3 3v.5" /><path d="M8.5 18h4a3 3 0 0 0 3-3v-.5" />
            </svg>
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <h3 style={styles.title}>Assign Workflow</h3>
            <p style={styles.subtitle}>This document is still a Draft. Pick a workflow to send it for approval.</p>
          </div>
          <button type="button" onClick={() => { if (!saving) onClose(); }} style={styles.closeBtn} aria-label="Close">×</button>
        </div>

        {/* Document + process summary */}
        <div style={styles.docStrip}>
          <div style={{ minWidth: 0 }}>
            <span style={styles.stripLabel}>Document</span>
            <span style={styles.stripValue} title={documentTitle}>{documentTitle}</span>
          </div>
          <div style={{ flexShrink: 0 }}>
            <span style={styles.stripLabel}>Process</span>
            <span style={styles.processTag}>{processCode || 'N/A'}</span>
          </div>
        </div>

        {/* Body */}
        <div style={styles.body}>
          <span style={styles.sectionLabel}>Select Workflow</span>

          {loading ? (
            <p style={styles.hint}>Loading workflows…</p>
          ) : loadError ? (
            <p style={styles.errorText}>{loadError}</p>
          ) : workflows.length === 0 ? (
            <div style={styles.emptyBox}>
              <strong style={{ color: '#005a43' }}>No workflow found for {processCode}.</strong>
              <span style={{ fontSize: '12px', color: '#64748b' }}>
                Create a workflow for this process in Approvals → Create Workflow, then come back here.
              </span>
            </div>
          ) : (
            <div style={styles.workflowList}>
              {workflows.map(w => {
                const isSelected = w.id === selectedId;
                return (
                  <div
                    key={w.id}
                    role="radio"
                    aria-checked={isSelected}
                    tabIndex={0}
                    onClick={() => setSelectedId(w.id)}
                    onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelectedId(w.id); } }}
                    style={{
                      ...styles.workflowCard,
                      borderColor: isSelected ? '#005a43' : '#e2e8f0',
                      backgroundColor: isSelected ? '#f0fdf4' : '#ffffff',
                      boxShadow: isSelected ? '0 0 0 3px rgba(0, 90, 67, 0.12)' : 'none'
                    }}
                  >
                    <span style={{ ...styles.radio, borderColor: isSelected ? '#005a43' : '#cbd5e1' }}>
                      {isSelected && <span style={styles.radioDot} />}
                    </span>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <span style={styles.workflowName}>{w.workflowTitle}</span>
                      <span style={styles.workflowMeta}>
                        {w.workflowType}{w.deadlineDays ? ` • ${w.deadlineDays} day(s) per step` : ''}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Approval flow preview of the selected workflow */}
          {selected && (
            <div style={styles.flowBox}>
              <span style={styles.sectionLabel}>Approval flow</span>
              <div style={{ marginTop: '8px' }}>
                {stepsFor(selected).map((step, idx, arr) => (
                  <div key={step.key} style={{ display: 'flex', gap: '12px' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
                      <PersonAvatar name={step.person} photoUrl={personPhotos[step.person]} />
                      {idx < arr.length - 1 && <span style={styles.stepLine} />}
                    </div>
                    <div style={{ paddingBottom: idx < arr.length - 1 ? '10px' : 0, minWidth: 0 }}>
                      <span style={styles.stepTitle}>{step.title}</span>
                      <span style={styles.stepPerson}>{step.person}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {selected && (
            <p style={styles.note}>The document will move from Draft to Send for L1 and go to the L1 approver.</p>
          )}
          {saveError && <p style={styles.errorText}>{saveError}</p>}
        </div>

        {/* Footer */}
        <div style={styles.footer}>
          <button type="button" onClick={onClose} disabled={saving} style={{ ...styles.cancelBtn, opacity: saving ? 0.6 : 1 }}>Cancel</button>
          <button
            type="button"
            onClick={handleAssign}
            disabled={!canSubmit}
            style={{ ...styles.primaryBtn, opacity: canSubmit ? 1 : 0.5, cursor: canSubmit ? 'pointer' : 'not-allowed' }}
          >
            {saving ? 'Saving…' : 'Assign Workflow'}
          </button>
        </div>
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  overlay: { position: 'fixed', top: 0, right: 0, bottom: 0, left: 0, backgroundColor: 'rgba(15, 23, 42, 0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10000, padding: '16px', fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, sans-serif' },
  modal: { width: '100%', maxWidth: '520px', maxHeight: '90vh', backgroundColor: '#ffffff', borderRadius: '14px', boxShadow: '0 20px 50px rgba(0, 0, 0, 0.25)', display: 'flex', flexDirection: 'column', overflow: 'hidden' },
  header: { display: 'flex', alignItems: 'center', gap: '12px', padding: '18px 20px', borderBottom: '1px solid #e2e8f0' },
  headerIcon: { width: '38px', height: '38px', borderRadius: '10px', backgroundColor: '#005a43', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '2px 0 0', fontSize: '12px', color: '#64748b' },
  closeBtn: { border: 'none', background: 'transparent', fontSize: '24px', lineHeight: 1, color: '#94a3b8', cursor: 'pointer', padding: '2px 6px' },
  docStrip: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '16px', padding: '12px 20px', backgroundColor: '#f0fdf4', borderBottom: '1px solid #bbf0d4' },
  stripLabel: { display: 'block', fontSize: '10px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.4px', marginBottom: '2px' },
  stripValue: { display: 'block', fontSize: '13px', fontWeight: 700, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  processTag: { display: 'inline-block', fontSize: '11px', fontWeight: 700, color: '#005a43', backgroundColor: '#dcfce7', padding: '2px 8px', borderRadius: '6px' },
  body: { padding: '16px 20px', overflowY: 'auto', flex: 1 },
  sectionLabel: { display: 'block', fontSize: '11px', fontWeight: 700, color: '#475569', textTransform: 'uppercase', letterSpacing: '0.4px', marginBottom: '8px' },
  workflowList: { display: 'flex', flexDirection: 'column', gap: '8px' },
  workflowCard: { display: 'flex', alignItems: 'center', gap: '12px', padding: '11px 14px', border: '1.5px solid #e2e8f0', borderRadius: '10px', cursor: 'pointer', transition: 'border-color 0.15s, background-color 0.15s' },
  radio: { width: '18px', height: '18px', borderRadius: '50%', border: '2px solid #cbd5e1', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, boxSizing: 'border-box' },
  radioDot: { width: '8px', height: '8px', borderRadius: '50%', backgroundColor: '#005a43' },
  workflowName: { display: 'block', fontSize: '13px', fontWeight: 700, color: '#0f172a', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' },
  workflowMeta: { display: 'block', fontSize: '11px', color: '#64748b', marginTop: '2px' },
  flowBox: { marginTop: '16px', padding: '14px', border: '1px solid #e2e8f0', borderRadius: '10px', backgroundColor: '#f8fafc' },
  stepAvatar: { width: '28px', height: '28px', borderRadius: '50%', backgroundColor: '#005a43', color: '#ffffff', fontSize: '10px', fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 },
  stepLine: { width: '2px', flex: 1, minHeight: '12px', backgroundColor: '#bbf0d4', margin: '3px 0' },
  stepTitle: { display: 'block', fontSize: '12.5px', fontWeight: 700, color: '#0f172a' },
  stepPerson: { display: 'block', fontSize: '11.5px', color: '#64748b' },
  note: { margin: '12px 0 0', fontSize: '11.5px', color: '#005a43', backgroundColor: '#f0fdf4', border: '1px solid #bbf0d4', borderRadius: '8px', padding: '8px 10px' },
  hint: { margin: '4px 0 0', fontSize: '12px', color: '#64748b' },
  errorText: { margin: '10px 0 0', fontSize: '12px', color: '#dc2626', fontWeight: 600 },
  emptyBox: { display: 'flex', flexDirection: 'column', gap: '4px', padding: '14px', border: '1px dashed #bbf0d4', borderRadius: '10px', backgroundColor: '#f0fdf4', fontSize: '13px' },
  footer: { display: 'flex', justifyContent: 'flex-end', gap: '10px', padding: '14px 20px', borderTop: '1px solid #e2e8f0', backgroundColor: '#ffffff' },
  cancelBtn: { padding: '8px 18px', fontSize: '13px', fontWeight: 600, color: '#334155', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '8px', cursor: 'pointer' },
  primaryBtn: { padding: '8px 20px', fontSize: '13px', fontWeight: 700, color: '#ffffff', backgroundColor: '#005a43', border: 'none', borderRadius: '8px' }
};

export default AssignWorkflow;
