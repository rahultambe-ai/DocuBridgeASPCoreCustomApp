import * as React from 'react';
import { useState, useEffect, useMemo, useRef } from 'react';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

// -----------------------------------------------------------------------------
// DMS Audit Log popup
// Opened from the "Audit Logs" button in the top-right of the dashboard
// (see AnalyticsDashboard.tsx). Reads LIVE from the "DMS Audit Logs" SharePoint
// list every time the popup is opened. Popup look = same card as the
// "Process Master Form" popup (centered, 16px rounded corners, dimmed backdrop).
//
// Table: exactly the columns of the "DMS Audit Logs" list (see COLUMNS below),
// scrolls both horizontally and vertically; only the header row stays pinned at
// the top, every column (AuditID included) scrolls sideways normally.
// Filters: custom green-theme dropdowns (same look as the FormSelect dropdown in
// ProcessManagementForm.tsx) and a custom green-theme date range calendar.
//
// NEW: "Columns" button (SharePoint-style Show / hide columns). The user can hide
// the fields he does not need and add them back any time. Columns are grouped
// (General, Process, RACI ...), searchable, with Show all / Reset. The choice is
// remembered in the browser (localStorage) so it is still there next time.
//
// NEW: "Export" button (CSV / Excel). Exports the filtered rows and the visible
// columns exactly as shown in the table.
// -----------------------------------------------------------------------------

export interface IDMSAuditLogProps {
    isOpen: boolean;
    onClose: () => void;
    context: any;
}

type ColType = 'text' | 'long' | 'date' | 'datetime' | 'bool' | 'person' | 'chip';

interface IColumn {
    key: string;      // header label + key in row.values (the list's column name)
    source?: string;  // internal SharePoint field name when it differs from key
    type: ColType;
    width: number;    // min width in px
}

// The columns shown in the table, in this order.
// "ProcessCode" is the list's Title column (renamed to ProcessID in the list UI).
// First columns: Created (SharePoint's Created date/time), AuditID, ProcessCode, EntryType (shown as a coloured chip).
// The last six columns (GuestAccessID ... Note) are the Guest Access columns of the list.
const COLUMNS: IColumn[] = [
    { key: 'Created', type: 'datetime', width: 160 },
    { key: 'AuditID', type: 'text', width: 140 },
    { key: 'ProcessCode', source: 'Title', type: 'text', width: 130 },
    { key: 'EntryType', type: 'chip', width: 150 },
    { key: 'ProcessName', type: 'text', width: 170 },
    { key: 'Department', type: 'text', width: 130 },
    { key: 'ProcessCategory', type: 'text', width: 140 },
    { key: 'ProcessOwner', type: 'text', width: 140 },
    { key: 'ReviewFrequency', type: 'text', width: 130 },
    { key: 'KRACode', type: 'text', width: 100 },
    { key: 'Weightage', type: 'text', width: 100 },
    { key: 'EffectiveDate', type: 'date', width: 120 },
    { key: 'NextReviewDate', type: 'date', width: 120 },
    { key: 'Description', type: 'long', width: 240 },
    { key: 'Responsible', type: 'text', width: 130 },
    { key: 'Accountable', type: 'text', width: 130 },
    { key: 'Consulted', type: 'text', width: 130 },
    { key: 'Informed', type: 'text', width: 130 },
    { key: 'AssignmentRole', type: 'text', width: 130 },
    { key: 'EntraIDGroupName', type: 'text', width: 160 },
    { key: 'GroupType', type: 'text', width: 110 },
    { key: 'AssignedUsers', type: 'long', width: 220 },
    { key: 'GroupDescription', type: 'long', width: 220 },
    { key: 'DocumentID', type: 'text', width: 120 },
    { key: 'DocumentType', type: 'text', width: 130 },
    { key: 'DocumentStatus', type: 'text', width: 130 },
    { key: 'Tags', type: 'text', width: 130 },
    { key: 'FileName', type: 'text', width: 180 },
    { key: 'FileSize', type: 'text', width: 90 },
    { key: 'FileType', type: 'text', width: 90 },
    { key: 'DocumentOwner', type: 'text', width: 140 },
    { key: 'DocumentName', type: 'text', width: 180 },
    { key: 'DocumentVersion', type: 'text', width: 120 },
    { key: 'WorkflowType', type: 'text', width: 130 },
    { key: 'EscalationCriteria', type: 'long', width: 220 },
    { key: 'ApprovalDeadlineDays', type: 'text', width: 150 },
    { key: 'NotifyApproversEachStep', type: 'bool', width: 180 },
    { key: 'AlertOnEscalationCompliance', type: 'bool', width: 200 },
    { key: 'StepRole', type: 'text', width: 110 },
    { key: 'L1Approver', type: 'person', width: 140 },
    { key: 'L2Approver', type: 'person', width: 140 },
    { key: 'FinalApprover', type: 'person', width: 140 },
    { key: 'ComplianceCheckApprover', type: 'text', width: 190 },
    { key: 'ApprovalStepName', type: 'long', width: 220 },
    { key: 'ActionBy', type: 'text', width: 180 },
    { key: 'ActionStatus', type: 'text', width: 120 },
    { key: 'ActionDate', type: 'datetime', width: 160 },
    { key: 'GuestAccessID', type: 'text', width: 150 },
    { key: 'GuestName', type: 'text', width: 140 },
    { key: 'VendorName', type: 'text', width: 140 },
    { key: 'EmailID', type: 'text', width: 200 },
    { key: 'RequestorEmail', type: 'text', width: 200 },
    { key: 'Note', type: 'long', width: 240 }
];

// Groups used only inside the "Columns" panel (so 52 columns are easy to find).
// Every column key of COLUMNS must appear in exactly one group.
const COLUMN_GROUPS: { name: string; keys: string[] }[] = [
    { name: 'General', keys: ['Created', 'AuditID', 'ProcessCode', 'EntryType'] },
    { name: 'Process', keys: ['ProcessName', 'Department', 'ProcessCategory', 'ProcessOwner', 'ReviewFrequency', 'KRACode', 'Weightage', 'EffectiveDate', 'NextReviewDate', 'Description'] },
    { name: 'RACI', keys: ['Responsible', 'Accountable', 'Consulted', 'Informed'] },
    { name: 'Assignment Group', keys: ['AssignmentRole', 'EntraIDGroupName', 'GroupType', 'AssignedUsers', 'GroupDescription'] },
    { name: 'Document', keys: ['DocumentID', 'DocumentType', 'DocumentStatus', 'Tags', 'FileName', 'FileSize', 'FileType', 'DocumentOwner', 'DocumentName', 'DocumentVersion'] },
    { name: 'Workflow', keys: ['WorkflowType', 'EscalationCriteria', 'ApprovalDeadlineDays', 'NotifyApproversEachStep', 'AlertOnEscalationCompliance'] },
    { name: 'Approval Step', keys: ['StepRole', 'L1Approver', 'L2Approver', 'FinalApprover', 'ComplianceCheckApprover', 'ApprovalStepName', 'ActionBy', 'ActionStatus', 'ActionDate'] },
    { name: 'Guest Access', keys: ['GuestAccessID', 'GuestName', 'VendorName', 'EmailID', 'RequestorEmail', 'Note'] }
];

// Columns visible when the user opens the popup for the first time
// (or presses "Reset to default"). Change this list to change the starting view.
const DEFAULT_VISIBLE_KEYS: string[] = [
    'Created', 'AuditID', 'ProcessCode', 'EntryType', 'ProcessName', 'Department',
    'DocumentName', 'ActionBy', 'ActionStatus', 'ActionDate', 'Note'
];

const COLUMNS_STORAGE_KEY = 'DMSAuditLog_VisibleColumns_v1';

type VisibleMap = { [key: string]: boolean };

const keysToMap = (keys: string[]): VisibleMap => {
    const m: VisibleMap = {};
    keys.forEach((k) => { m[k] = true; });
    return m;
};

const countVisible = (m: VisibleMap): number => COLUMNS.filter((c) => m[c.key]).length;

// Saved choice from the browser, or the default set when nothing valid is saved.
const loadVisibleColumns = (): VisibleMap => {
    try {
        const raw = window.localStorage.getItem(COLUMNS_STORAGE_KEY);
        if (raw) {
            const arr = JSON.parse(raw);
            if (Array.isArray(arr)) {
                const valid = arr.filter((k: string) => COLUMNS.filter((c) => c.key === k).length > 0);
                if (valid.length > 0) return keysToMap(valid);
            }
        }
    } catch (e) {
        // storage not available - fall back to defaults
    }
    return keysToMap(DEFAULT_VISIBLE_KEYS);
};

interface IAuditRow {
    id: number;
    entryType: string;                    // only used for the "All Entry Types" filter
    processId: string;                    // only used for the "All Processes" filter
    created: number;                      // only used for the date range filter
    values: { [key: string]: string };    // display text for every column
    searchText: string;
}

const AUDIT_LOG_LIST_TITLE = 'DMS Audit Logs';
const MAX_ROWS = 1000;

const ENTRY_TYPE_COLORS: { [key: string]: { bg: string; fg: string } } = {
    'Document': { bg: '#dcfce7', fg: '#166534' },
    'Process': { bg: '#dbeafe', fg: '#1d4ed8' },
    'Assignment Group': { bg: '#cffafe', fg: '#0e7490' },
    'RACI': { bg: '#fee2e2', fg: '#b91c1c' },
    'Workflow': { bg: '#fef3c7', fg: '#92400e' },
    'Approval Step': { bg: '#ede9fe', fg: '#6d28d9' },
    'Guest Access': { bg: '#fce7f3', fg: '#be185d' }
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'];

const text = (v: any): string => (v === undefined || v === null ? '' : String(v));

const fmtDate = (v: any): string => {
    if (!v) return '';
    const d = new Date(v);
    return isNaN(d.getTime()) ? '' : d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });
};

const fmtDateTime = (v: any): string => {
    if (!v) return '';
    const d = new Date(v);
    if (isNaN(d.getTime())) return '';
    return `${d.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' })}, ${d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })}`;
};

// Calendar helpers — dates are passed around as local "YYYY-MM-DD" keys.
const pad = (n: number): string => (n < 10 ? `0${n}` : String(n));
const toKey = (d: Date): string => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const fromKey = (k: string): Date => {
    const p = k.split('-').map(Number);
    return new Date(p[0], p[1] - 1, p[2]);
};
const fmtShort = (k: string): string =>
    fromKey(k).toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' });

// ---------- Export helpers (CSV / Excel) ----------
const exportFileName = (ext: string): string => {
    const d = new Date();
    return `DMS_Audit_Logs_${toKey(d)}_${pad(d.getHours())}${pad(d.getMinutes())}.${ext}`;
};

const downloadBlob = (content: string, mime: string, fileName: string) => {
    const blob = new Blob([content], { type: mime });
    const nav: any = window.navigator;
    if (nav && nav.msSaveBlob) { nav.msSaveBlob(blob, fileName); return; }
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
};

const exportCsv = (cols: IColumn[], data: IAuditRow[]) => {
    const esc = (v: string) => `"${v.replace(/"/g, '""')}"`;
    const lines = [cols.map((c) => esc(c.key)).join(',')]
        .concat(data.map((r) => cols.map((c) => esc(r.values[c.key] || '')).join(',')));
    // \uFEFF = BOM, so Excel opens it correctly
    downloadBlob('\uFEFF' + lines.join('\r\n'), 'text/csv;charset=utf-8;', exportFileName('csv'));
};

const exportExcel = (cols: IColumn[], data: IAuditRow[]) => {
    const esc = (v: string) => v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const head = cols.map((c) => `<th style="background:#004d3d;color:#ffffff;font-weight:bold;">${esc(c.key)}</th>`).join('');
    const body = data.map((r) => `<tr>${cols.map((c) => `<td style="mso-number-format:'\\@';">${esc(r.values[c.key] || '')}</td>`).join('')}</tr>`).join('');
    const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8" /></head><body><table border="1"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></body></html>`;
    downloadBlob('\uFEFF' + html, 'application/vnd.ms-excel;charset=utf-8;', exportFileName('xls'));
};

// One row of the "DMS Audit Logs" list -> one row of the table.
const mapItem = (item: any): IAuditRow => {
    const values: { [key: string]: string } = {};
    COLUMNS.forEach((col) => {
        const raw = item[col.source || col.key];
        switch (col.type) {
            case 'date':
                values[col.key] = fmtDate(raw);
                break;
            case 'datetime':
                values[col.key] = fmtDateTime(raw);
                break;
            case 'bool':
                values[col.key] = raw === true ? 'Yes' : raw === false ? 'No' : '';
                break;
            case 'person':
                values[col.key] = raw && typeof raw === 'object' ? text(raw.Title) : '';
                break;
            default:
                values[col.key] = text(raw);
        }
    });

    return {
        id: item.Id,
        entryType: text(item.EntryType),
        processId: text(item.Title),
        created: item.Created ? new Date(item.Created).getTime() : 0,
        values,
        searchText: COLUMNS.map((c) => values[c.key]).join(' ').toLowerCase()
    };
};

// -----------------------------------------------------------------------------
// Custom dropdown — same look & behaviour as FormSelect in ProcessManagementForm.tsx
// (rounded white menu, green focus ring, light-green selected row with a ✓).
// -----------------------------------------------------------------------------
interface IAuditSelectProps {
    value: string;
    onChange: (val: string) => void;
    options: { label: string; value: string }[];
    width: string;
}

const AuditSelect: React.FC<IAuditSelectProps> = ({ value, onChange, options, width }) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const wrapperRef = useRef<HTMLDivElement>(null);
    const selected = options.filter((o) => o.value === value)[0];

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) setIsOpen(false);
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    return (
        <div style={{ position: 'relative', width }} ref={wrapperRef}>
            <div
                className="al-select-box"
                role="button"
                tabIndex={0}
                style={{ ...styles.selectBox, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none' }}
                onClick={() => setIsOpen((prev) => !prev)}
                onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setIsOpen((prev) => !prev); }
                    if (e.key === 'Escape') setIsOpen(false);
                }}
            >
                <span style={{ fontSize: '11px', color: '#1e293b', fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {selected ? selected.label : ''}
                </span>
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}>
                    <polyline points="6 9 12 15 18 9" />
                </svg>
            </div>
            {isOpen && (
                <div className="al-dropdown-menu al-scroll" style={styles.selectMenu}>
                    {options.map((opt) => {
                        const isSelected = opt.value === value;
                        return (
                            <div
                                key={opt.value}
                                className="al-select-option"
                                style={{ ...styles.selectOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent', color: isSelected ? '#004d3d' : '#334155', fontWeight: isSelected ? 600 : 400 }}
                                onClick={() => { onChange(opt.value); setIsOpen(false); }}
                            >
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

// -----------------------------------------------------------------------------
// Custom date range calendar (green theme).
// Click a start date, then an end date. Quick presets + month/year jump included.
// -----------------------------------------------------------------------------
interface IDateRangePickerProps {
    from: string;
    to: string;
    onChange: (from: string, to: string) => void;
}

const DateRangePicker: React.FC<IDateRangePickerProps> = ({ from, to, onChange }) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const [view, setView] = useState<'days' | 'months'>('days');
    const [viewYear, setViewYear] = useState<number>(new Date().getFullYear());
    const [viewMonth, setViewMonth] = useState<number>(new Date().getMonth());
    const [hoverKey, setHoverKey] = useState<string>('');
    const wrapperRef = useRef<HTMLDivElement>(null);

    const todayKey = toKey(new Date());

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) setIsOpen(false);
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    const toggleOpen = () => {
        if (!isOpen) {
            const base = from ? fromKey(from) : new Date();
            setViewYear(base.getFullYear());
            setViewMonth(base.getMonth());
            setView('days');
            setHoverKey('');
        }
        setIsOpen(!isOpen);
    };

    const shift = (delta: number) => {
        if (view === 'months') {
            setViewYear(viewYear + delta);
            return;
        }
        let m = viewMonth + delta;
        let y = viewYear;
        if (m < 0) { m = 11; y -= 1; }
        if (m > 11) { m = 0; y += 1; }
        setViewMonth(m);
        setViewYear(y);
    };

    const handlePick = (key: string) => {
        if (!from || (from && to)) {
            onChange(key, '');
            return;
        }
        if (key < from) onChange(key, from);
        else onChange(from, key);
        setHoverKey('');
        setIsOpen(false);
    };

    const applyPreset = (preset: 'today' | '7' | '30' | 'month') => {
        const now = new Date();
        let start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        if (preset === '7') start.setDate(start.getDate() - 6);
        if (preset === '30') start.setDate(start.getDate() - 29);
        if (preset === 'month') start = new Date(now.getFullYear(), now.getMonth(), 1);
        onChange(toKey(start), todayKey);
        setIsOpen(false);
    };

    // Range being drawn: while only a start date is picked, the hovered day previews the end.
    const pending = !!from && !to;
    let lo = from;
    let hi = to || (pending ? hoverKey : '');
    if (lo && hi && hi < lo) { const t = lo; lo = hi; hi = t; }
    const hasRange = !!lo && !!hi && lo !== hi;

    const firstDow = new Date(viewYear, viewMonth, 1).getDay();
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();
    const cells: (string | null)[] = [];
    for (let i = 0; i < firstDow; i++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(`${viewYear}-${pad(viewMonth + 1)}-${pad(d)}`);

    const label = from || to ? `${from ? fmtShort(from) : '…'} – ${to ? fmtShort(to) : '…'}` : 'Select Date Range';

    return (
        <div style={{ position: 'relative' }} ref={wrapperRef}>
            <button
                type="button"
                className="al-select-box"
                onClick={toggleOpen}
                style={{ ...styles.dateField, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none', color: from || to ? '#1e293b' : '#94a3b8' }}
            >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke={from || to ? '#004d3d' : '#64748b'} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
                    <line x1="16" y1="2" x2="16" y2="6" />
                    <line x1="8" y1="2" x2="8" y2="6" />
                    <line x1="3" y1="10" x2="21" y2="10" />
                </svg>
                <span style={{ flex: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</span>
            </button>

            {isOpen && (
                <div className="al-dropdown-menu" style={styles.datePopover}>
                    {/* Quick presets */}
                    <div style={styles.presetRow}>
                        <button type="button" className="al-chip-btn" style={styles.presetBtn} onClick={() => applyPreset('today')}>Today</button>
                        <button type="button" className="al-chip-btn" style={styles.presetBtn} onClick={() => applyPreset('7')}>Last 7 days</button>
                        <button type="button" className="al-chip-btn" style={styles.presetBtn} onClick={() => applyPreset('30')}>Last 30 days</button>
                        <button type="button" className="al-chip-btn" style={styles.presetBtn} onClick={() => applyPreset('month')}>This month</button>
                    </div>

                    {/* From / To readout */}
                    <div style={styles.rangeReadout}>
                        <div style={{ ...styles.readoutBox, borderColor: !from || (from && to) ? '#004d3d' : '#e2e8f0' }}>
                            <span style={styles.readoutLabel}>From</span>
                            <span style={{ ...styles.readoutValue, color: from ? '#0f172a' : '#94a3b8' }}>{from ? fmtShort(from) : 'Pick a date'}</span>
                        </div>
                        <span style={{ color: '#94a3b8', fontSize: '12px' }}>→</span>
                        <div style={{ ...styles.readoutBox, borderColor: pending ? '#004d3d' : '#e2e8f0' }}>
                            <span style={styles.readoutLabel}>To</span>
                            <span style={{ ...styles.readoutValue, color: to ? '#0f172a' : '#94a3b8' }}>{to ? fmtShort(to) : 'Pick a date'}</span>
                        </div>
                    </div>

                    {/* Month / year header */}
                    <div style={styles.calHeader}>
                        <button type="button" className="al-nav-btn" style={styles.navBtn} onClick={() => shift(-1)} title="Previous">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="15 18 9 12 15 6" /></svg>
                        </button>
                        <button type="button" className="al-chip-btn" style={styles.monthLabelBtn} onClick={() => setView(view === 'days' ? 'months' : 'days')}>
                            {view === 'days' ? `${MONTHS[viewMonth]} ${viewYear}` : viewYear}
                            <svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginLeft: '5px', transform: view === 'months' ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s ease' }}><polyline points="6 9 12 15 18 9" /></svg>
                        </button>
                        <button type="button" className="al-nav-btn" style={styles.navBtn} onClick={() => shift(1)} title="Next">
                            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6" /></svg>
                        </button>
                    </div>

                    {view === 'months' ? (
                        <div style={styles.monthGrid}>
                            {MONTHS.map((m, i) => {
                                const active = i === viewMonth;
                                return (
                                    <button
                                        key={m}
                                        type="button"
                                        className="al-month-btn"
                                        style={{ ...styles.monthBtn, backgroundColor: active ? '#004d3d' : 'transparent', color: active ? '#ffffff' : '#334155', fontWeight: active ? 700 : 500 }}
                                        onClick={() => { setViewMonth(i); setView('days'); }}
                                    >
                                        {m.slice(0, 3)}
                                    </button>
                                );
                            })}
                        </div>
                    ) : (
                        <>
                            <div style={styles.weekRow}>
                                {WEEKDAYS.map((w) => <span key={w} style={styles.weekDay}>{w}</span>)}
                            </div>
                            <div style={styles.dayGrid} onMouseLeave={() => setHoverKey('')}>
                                {cells.map((key, idx) => {
                                    if (!key) return <span key={`blank-${idx}`} />;
                                    const isStart = key === lo;
                                    const isEnd = !!hi && key === hi;
                                    const inRange = hasRange && key > lo && key < hi;
                                    let band = 'transparent';
                                    if (inRange) band = '#e6f4ea';
                                    else if (hasRange && isStart) band = 'linear-gradient(to right, transparent 50%, #e6f4ea 50%)';
                                    else if (hasRange && isEnd) band = 'linear-gradient(to right, #e6f4ea 50%, transparent 50%)';
                                    const cls = `al-day${isStart || isEnd ? ' sel' : ''}${key === todayKey ? ' today' : ''}${inRange ? ' in' : ''}`;
                                    return (
                                        <div key={key} style={{ ...styles.dayCell, background: band }} onMouseEnter={() => setHoverKey(key)}>
                                            <button type="button" className={cls} onClick={() => handlePick(key)}>
                                                {Number(key.slice(8))}
                                            </button>
                                        </div>
                                    );
                                })}
                            </div>
                        </>
                    )}

                    <div style={styles.calFooter}>
                        <button type="button" className="al-btn" style={styles.linkBtn} onClick={() => { onChange('', ''); setHoverKey(''); }}>Clear</button>
                        <button type="button" className="al-btn" style={styles.doneBtn} onClick={() => setIsOpen(false)}>Done</button>
                    </div>
                </div>
            )}
        </div>
    );
};

// -----------------------------------------------------------------------------
// Show / hide columns panel (SharePoint style, green theme).
// - Columns grouped by entry area, searchable
// - Tick = column visible in the table, untick = hidden
// - "All" / "None" per group, "Show all" and "Reset to default" at the bottom
// - At least one column always stays visible
// -----------------------------------------------------------------------------
interface IColumnPickerProps {
    visible: VisibleMap;
    onChange: (next: VisibleMap) => void;
}

const ColumnPicker: React.FC<IColumnPickerProps> = ({ visible, onChange }) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const [query, setQuery] = useState<string>('');
    const wrapperRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) setIsOpen(false);
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    const shown = countVisible(visible);
    const total = COLUMNS.length;
    const isCustom = shown !== total;

    const apply = (keys: string[], value: boolean) => {
        const next: VisibleMap = { ...visible };
        keys.forEach((k) => { next[k] = value; });
        if (countVisible(next) === 0) return; // never hide everything
        onChange(next);
    };

    const toggleOne = (key: string) => apply([key], !visible[key]);

    const q = query.trim().toLowerCase();
    const groups = COLUMN_GROUPS
        .map((g) => ({ name: g.name, keys: g.keys.filter((k) => q === '' || k.toLowerCase().indexOf(q) > -1) }))
        .filter((g) => g.keys.length > 0);

    return (
        <div style={{ position: 'relative' }} ref={wrapperRef}>
            <button
                type="button"
                className="al-select-box"
                onClick={() => setIsOpen(!isOpen)}
                onKeyDown={(e) => { if (e.key === 'Escape') setIsOpen(false); }}
                style={{ ...styles.columnsBtn, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none' }}
                title="Show or hide columns"
            >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <rect x="3" y="3" width="18" height="18" rx="2" ry="2" />
                    <line x1="9" y1="3" x2="9" y2="21" />
                    <line x1="15" y1="3" x2="15" y2="21" />
                </svg>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#1e293b' }}>Columns</span>
                <span style={{ ...styles.countBadge, backgroundColor: isCustom ? '#004d3d' : '#e6f4ea', color: isCustom ? '#ffffff' : '#004d3d' }}>
                    {shown}/{total}
                </span>
            </button>

            {isOpen && (
                <div className="al-dropdown-menu" style={styles.colPopover}>
                    <div style={styles.colHeader}>
                        <div>
                            <div style={styles.colTitle}>Show / hide columns</div>
                            <div style={styles.colSubtitle}>{shown} of {total} columns shown</div>
                        </div>
                    </div>

                    <div style={{ position: 'relative', marginBottom: '8px' }}>
                        <svg style={styles.searchIcon} width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <circle cx="11" cy="11" r="8" />
                            <line x1="21" y1="21" x2="16.65" y2="16.65" />
                        </svg>
                        <input
                            className="al-input"
                            type="text"
                            placeholder="Find a column..."
                            value={query}
                            onChange={(e) => setQuery(e.target.value)}
                            style={styles.colSearch}
                        />
                    </div>

                    <div className="al-scroll" style={styles.colList}>
                        {groups.length === 0 && <p style={{ ...styles.emptyState, padding: '20px 8px' }}>No column matches “{query}”.</p>}
                        {groups.map((g) => {
                            const onCount = g.keys.filter((k) => visible[k]).length;
                            return (
                                <div key={g.name} style={{ marginBottom: '6px' }}>
                                    <div style={styles.groupHeader}>
                                        <span style={styles.groupName}>{g.name} <span style={styles.groupCount}>{onCount}/{g.keys.length}</span></span>
                                        <span>
                                            <button type="button" className="al-link-btn" style={styles.groupBtn} onClick={() => apply(g.keys, true)}>All</button>
                                            <button type="button" className="al-link-btn" style={styles.groupBtn} onClick={() => apply(g.keys, false)}>None</button>
                                        </span>
                                    </div>
                                    {g.keys.map((k) => {
                                        const on = !!visible[k];
                                        return (
                                            <div
                                                key={k}
                                                className="al-check-row"
                                                role="checkbox"
                                                aria-checked={on}
                                                tabIndex={0}
                                                style={styles.checkRow}
                                                onClick={() => toggleOne(k)}
                                                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggleOne(k); } }}
                                            >
                                                <span style={{ ...styles.checkBox, backgroundColor: on ? '#004d3d' : '#ffffff', borderColor: on ? '#004d3d' : '#cbd5e1' }}>
                                                    {on && (
                                                        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#ffffff" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12" /></svg>
                                                    )}
                                                </span>
                                                <span style={{ fontSize: '11px', color: on ? '#0f172a' : '#64748b', fontWeight: on ? 600 : 400 }}>{k}</span>
                                            </div>
                                        );
                                    })}
                                </div>
                            );
                        })}
                    </div>

                    <div style={styles.calFooter}>
                        <span>
                            <button type="button" className="al-link-btn" style={styles.linkBtn} onClick={() => onChange(keysToMap(DEFAULT_VISIBLE_KEYS))}>Reset to default</button>
                            <button type="button" className="al-link-btn" style={styles.linkBtn} onClick={() => onChange(keysToMap(COLUMNS.map((c) => c.key)))}>Show all</button>
                        </span>
                        <button type="button" className="al-btn" style={styles.doneBtn} onClick={() => setIsOpen(false)}>Done</button>
                    </div>
                </div>
            )}
        </div>
    );
};

// -----------------------------------------------------------------------------
// Export button (CSV / Excel) - exports the filtered rows and visible columns.
// -----------------------------------------------------------------------------
interface IExportMenuProps {
    onCsv: () => void;
    onExcel: () => void;
    disabled: boolean;
}

const ExportMenu: React.FC<IExportMenuProps> = ({ onCsv, onExcel, disabled }) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const wrapperRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) setIsOpen(false);
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    return (
        <div style={{ position: 'relative' }} ref={wrapperRef}>
            <button
                type="button"
                className="al-select-box"
                disabled={disabled}
                onClick={() => setIsOpen(!isOpen)}
                onKeyDown={(e) => { if (e.key === 'Escape') setIsOpen(false); }}
                style={{ ...styles.columnsBtn, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none', opacity: disabled ? 0.5 : 1, cursor: disabled ? 'not-allowed' : 'pointer' }}
                title="Export the filtered logs"
            >
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
                    <polyline points="7 10 12 15 17 10" />
                    <line x1="12" y1="15" x2="12" y2="3" />
                </svg>
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#1e293b' }}>Export</span>
                <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease' }}>
                    <polyline points="6 9 12 15 18 9" />
                </svg>
            </button>

            {isOpen && (
                <div className="al-dropdown-menu" style={{ ...styles.selectMenu, left: 'auto', right: 0, width: '170px', minWidth: '170px', zIndex: 800 }}>
                    <div className="al-select-option" style={styles.selectOption} onClick={() => { onExcel(); setIsOpen(false); }}>
                        <span>Excel (.xls)</span>
                    </div>
                    <div className="al-select-option" style={styles.selectOption} onClick={() => { onCsv(); setIsOpen(false); }}>
                        <span>CSV (.csv)</span>
                    </div>
                </div>
            )}
        </div>
    );
};

const DMSAuditLog: React.FC<IDMSAuditLogProps> = ({ isOpen, onClose, context }) => {
    const [rows, setRows] = useState<IAuditRow[]>([]);
    const [loading, setLoading] = useState<boolean>(false);
    const [error, setError] = useState<string>('');

    const [search, setSearch] = useState<string>('');
    const [entryTypeFilter, setEntryTypeFilter] = useState<string>('All');
    const [processFilter, setProcessFilter] = useState<string>('All');
    const [dateFrom, setDateFrom] = useState<string>('');
    const [dateTo, setDateTo] = useState<string>('');

    // Which columns are visible (remembered in the browser).
    const [visibleCols, setVisibleCols] = useState<VisibleMap>(loadVisibleColumns);

    useEffect(() => {
        try {
            window.localStorage.setItem(
                COLUMNS_STORAGE_KEY,
                JSON.stringify(COLUMNS.filter((c) => visibleCols[c.key]).map((c) => c.key))
            );
        } catch (e) {
            // storage not available - choice just lasts for this session
        }
    }, [visibleCols]);

    const shownColumns = useMemo(() => COLUMNS.filter((c) => visibleCols[c.key]), [visibleCols]);

    // Fresh read from the SharePoint list every time the popup opens.
    useEffect(() => {
        if (!isOpen || !context) return;
        let cancelled = false;
        const load = async () => {
            setLoading(true);
            setError('');
            try {
                const sp = spfi().using(SPFx(context));
                const items: any[] = await sp.web.lists
                    .getByTitle(AUDIT_LOG_LIST_TITLE)
                    .items
                    // "*" = every normal column; the Person columns need their own expand to return a name.
                    .select('*', 'L1Approver/Title', 'L2Approver/Title', 'FinalApprover/Title')
                    .expand('L1Approver', 'L2Approver', 'FinalApprover')
                    .orderBy('Created', false)
                    .top(MAX_ROWS)();
                if (!cancelled) setRows(items.map(mapItem));
            } catch (err) {
                console.error('Error fetching DMS Audit Logs list:', err);
                if (!cancelled) setError('Could not load the "DMS Audit Logs" list. Check that the list exists and you have access to it.');
            } finally {
                if (!cancelled) setLoading(false);
            }
        };
        load();
        return () => { cancelled = true; };
    }, [isOpen, context]);

    // Start clean next time the popup is opened (the column choice is kept on purpose).
    useEffect(() => {
        if (!isOpen) {
            setSearch('');
            setEntryTypeFilter('All');
            setProcessFilter('All');
            setDateFrom('');
            setDateTo('');
        }
    }, [isOpen]);

    const entryTypeOptions = useMemo(() => {
        const types = Array.from(new Set(rows.map((r) => r.entryType).filter((v) => v !== ''))).sort();
        return [{ label: 'All Entry Types', value: 'All' }].concat(types.map((t) => ({ label: t, value: t })));
    }, [rows]);

    const processOptions = useMemo(() => {
        const codes = Array.from(new Set(rows.map((r) => r.processId).filter((v) => v !== ''))).sort();
        return [{ label: 'All Processes', value: 'All' }].concat(codes.map((c) => ({ label: c, value: c })));
    }, [rows]);

    const filtered = useMemo(() => {
        const q = search.trim().toLowerCase();
        const fromMs = dateFrom ? new Date(`${dateFrom}T00:00:00`).getTime() : 0;
        const toMs = dateTo ? new Date(`${dateTo}T23:59:59.999`).getTime() : 0;
        return rows.filter((r) => {
            if (entryTypeFilter !== 'All' && r.entryType !== entryTypeFilter) return false;
            if (processFilter !== 'All' && r.processId !== processFilter) return false;
            if (fromMs && r.created < fromMs) return false;
            if (toMs && r.created > toMs) return false;
            if (!q) return true;
            return r.searchText.indexOf(q) > -1;
        });
    }, [rows, search, entryTypeFilter, processFilter, dateFrom, dateTo]);

    if (!isOpen) return null;

    const cell = (value: string) => (value === '' ? '—' : value);

    return (
        <div className="al-overlay" style={styles.overlay} onClick={onClose}>
            <style>{`
        @keyframes alFadeIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes alScaleIn { from { opacity: 0; transform: translateY(10px) scale(0.97); } to { opacity: 1; transform: translateY(0) scale(1); } }
        @keyframes alSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .al-overlay { animation: alFadeIn 0.18s ease; }
        .al-modal { animation: alScaleIn 0.24s cubic-bezier(0.16, 1, 0.3, 1); }
        .al-dropdown-menu { animation: alSlideDown 0.15s ease; }
        .al-close-btn { transition: background-color 0.15s ease, transform 0.15s ease; }
        .al-close-btn:hover { background-color: #fee2e2; color: #dc2626; transform: rotate(90deg); }
        .al-input { transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .al-input:hover { border-color: #94a3b8; }
        .al-input:focus { border-color: #004d3d; box-shadow: 0 0 0 3px rgba(0, 77, 61, 0.10); }
        .al-select-box { transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .al-select-box:hover { border-color: #94a3b8; }
        .al-select-box:focus { outline: none; border-color: #004d3d; }
        .al-select-option:hover { background-color: #f1f5f9; }
        .al-row { transition: background-color 0.15s ease; }
        .al-row:hover { background-color: #f8fafc; }
        .al-btn { transition: transform 0.15s ease, filter 0.15s ease; }
        .al-btn:hover { transform: translateY(-1px); filter: brightness(0.98); }
        .al-btn:active { transform: translateY(0); }

        /* Columns panel */
        .al-check-row { transition: background-color 0.12s ease; outline: none; }
        .al-check-row:hover { background-color: #f1f5f9; }
        .al-check-row:focus-visible { box-shadow: inset 0 0 0 2px #004d3d; }
        .al-link-btn { transition: color 0.15s ease, background-color 0.15s ease; }
        .al-link-btn:hover { color: #004d3d; background-color: #e6f4ea; }

        /* Calendar */
        .al-day { width: 28px; height: 28px; border: none; border-radius: 8px; background: transparent; font-size: 11px; color: #334155; cursor: pointer; padding: 0; font-family: inherit; transition: background-color 0.12s ease, color 0.12s ease, transform 0.12s ease; }
        .al-day:hover { background-color: #e6f4ea; color: #004d3d; transform: scale(1.06); }
        .al-day.in { color: #004d3d; font-weight: 600; }
        .al-day.today { box-shadow: inset 0 0 0 1.5px #004d3d; color: #004d3d; font-weight: 700; }
        .al-day.sel, .al-day.sel:hover { background-color: #004d3d; color: #ffffff; font-weight: 700; box-shadow: none; }
        .al-nav-btn { transition: background-color 0.15s ease, color 0.15s ease; }
        .al-nav-btn:hover { background-color: #e6f4ea; color: #004d3d; }
        .al-chip-btn { transition: background-color 0.15s ease, color 0.15s ease, border-color 0.15s ease; }
        .al-chip-btn:hover { background-color: #e6f4ea; color: #004d3d; border-color: #b7dcc4; }
        .al-month-btn { transition: background-color 0.15s ease; }
        .al-month-btn:hover { background-color: #e6f4ea; }

        .al-scroll::-webkit-scrollbar { width: 8px; height: 8px; }
        .al-scroll::-webkit-scrollbar-thumb { background-color: #cbd5e1; border-radius: 8px; }
        .al-scroll::-webkit-scrollbar-thumb:hover { background-color: #94a3b8; }
        .al-scroll::-webkit-scrollbar-track { background-color: transparent; }
        .al-scroll::-webkit-scrollbar-corner { background-color: transparent; }
      `}</style>

            <div className="al-modal" style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
                <div style={styles.header}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <div style={styles.headerIcon}>
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#004d3d" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
                                <polyline points="9 12 11 14 15 10" />
                            </svg>
                        </div>
                        <h2 style={styles.title}>DMS Audit Logs</h2>
                    </div>
                    <button className="al-close-btn" style={styles.closeBtn} onClick={onClose} title="Close Modal">✕</button>
                </div>

                <div style={styles.filterBar}>
                    <div style={styles.searchWrap}>
                        <svg style={styles.searchIcon} width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                            <circle cx="11" cy="11" r="8" />
                            <line x1="21" y1="21" x2="16.65" y2="16.65" />
                        </svg>
                        <input
                            className="al-input"
                            type="text"
                            placeholder="Search across all columns..."
                            value={search}
                            onChange={(e) => setSearch(e.target.value)}
                            style={styles.searchInput}
                        />
                    </div>

                    <AuditSelect value={entryTypeFilter} onChange={setEntryTypeFilter} options={entryTypeOptions} width="160px" />
                    <AuditSelect value={processFilter} onChange={setProcessFilter} options={processOptions} width="160px" />
                    <DateRangePicker from={dateFrom} to={dateTo} onChange={(f, t) => { setDateFrom(f); setDateTo(t); }} />
                    <ColumnPicker visible={visibleCols} onChange={setVisibleCols} />
                    <ExportMenu
                        disabled={filtered.length === 0}
                        onCsv={() => exportCsv(shownColumns, filtered)}
                        onExcel={() => exportExcel(shownColumns, filtered)}
                    />
                </div>

                <div className="al-scroll" style={styles.tableWrap}>
                    {loading && rows.length === 0 ? (
                        <p style={styles.emptyState}>Loading audit logs…</p>
                    ) : error ? (
                        <p style={{ ...styles.emptyState, color: '#dc2626' }}>{error}</p>
                    ) : filtered.length === 0 ? (
                        <p style={styles.emptyState}>{rows.length === 0 ? 'No audit logs found yet.' : 'No audit logs match your filters.'}</p>
                    ) : (
                        <table style={styles.table}>
                            <thead>
                                <tr>
                                    {shownColumns.map((col) => (
                                        <th key={col.key} style={{ ...styles.th, minWidth: `${col.width}px` }}>
                                            {col.key}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {filtered.map((r) => (
                                    <tr key={r.id} className="al-row">
                                        {shownColumns.map((col) => {
                                            const isLong = col.type === 'long';
                                            return (
                                                <td
                                                    key={col.key}
                                                    style={{
                                                        ...styles.td,
                                                        minWidth: `${col.width}px`,
                                                        ...(col.key === 'AuditID' || col.key === 'ProcessCode' ? { fontWeight: 600, color: '#1e293b' } : {}),
                                                        ...(isLong ? { whiteSpace: 'normal', maxWidth: '320px' } : {})
                                                    }}
                                                >
                                                    {col.type === 'chip' && r.values[col.key] !== '' ? (
                                                        <span style={{ ...styles.chip, backgroundColor: (ENTRY_TYPE_COLORS[r.values[col.key]] || { bg: '#f1f5f9', fg: '#475569' }).bg, color: (ENTRY_TYPE_COLORS[r.values[col.key]] || { bg: '#f1f5f9', fg: '#475569' }).fg }}>
                                                            {r.values[col.key]}
                                                        </span>
                                                    ) : cell(r.values[col.key])}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    )}
                </div>

                <div style={styles.footer}>
                    <span style={styles.footerText}>
                        Showing {filtered.length} of {rows.length} audit logs · {shownColumns.length} of {COLUMNS.length} columns
                    </span>
                    <button type="button" className="al-btn" style={styles.closeFooterBtn} onClick={onClose}>Close</button>
                </div>
            </div>
        </div>
    );
};

const styles: { [key: string]: React.CSSProperties } = {
    overlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(15, 23, 42, 0.55)', backdropFilter: 'blur(3px)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 9999, padding: '32px 16px', boxSizing: 'border-box' },
    modalCard: { backgroundColor: '#ffffff', borderRadius: '16px', width: '100%', maxWidth: '1180px', height: 'min(88vh, 780px)', display: 'flex', flexDirection: 'column', boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)', overflow: 'hidden', border: '1px solid #e2e8f0', margin: 'auto', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '14px 20px', borderBottom: '1px solid #e2e8f0', backgroundColor: '#ffffff', flexShrink: 0 },
    headerIcon: { width: '30px', height: '30px', borderRadius: '8px', backgroundColor: '#e6f4ea', display: 'flex', alignItems: 'center', justifyContent: 'center' },
    title: { margin: 0, fontSize: '16px', fontWeight: 700, color: '#0f172a' },
    closeBtn: { background: 'none', border: 'none', fontSize: '16px', cursor: 'pointer', color: '#64748b', padding: '4px 8px', borderRadius: '6px' },
    filterBar: { display: 'flex', alignItems: 'center', gap: '10px', padding: '12px 20px', backgroundColor: '#ffffff', borderBottom: '1px solid #e2e8f0', flexWrap: 'wrap', flexShrink: 0 },
    searchWrap: { position: 'relative', flex: '1 1 260px', minWidth: '220px' },
    searchIcon: { position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none' },
    searchInput: { width: '100%', height: '32px', padding: '5px 10px 5px 30px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '11px', outline: 'none', backgroundColor: '#ffffff', color: '#1e293b', boxSizing: 'border-box' },

    // Dropdown (same look as FormSelect in ProcessManagementForm.tsx)
    selectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '6px', padding: '5px 10px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', height: '32px', boxSizing: 'border-box', cursor: 'pointer' },
    selectMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, minWidth: '100%', backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 8px 20px rgba(0, 0, 0, 0.10)', zIndex: 500, padding: '4px', maxHeight: '220px', overflowY: 'auto', boxSizing: 'border-box' },
    selectOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 8px', fontSize: '11px', borderRadius: '5px', cursor: 'pointer', margin: '1px 0' },

    // Date range field + calendar popover
    dateField: { height: '32px', width: '210px', display: 'flex', alignItems: 'center', gap: '8px', padding: '5px 10px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '11px', backgroundColor: '#ffffff', boxSizing: 'border-box', cursor: 'pointer', outline: 'none', textAlign: 'left', fontFamily: 'inherit' },
    datePopover: { position: 'absolute', top: 'calc(100% + 6px)', right: 0, width: '272px', backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '12px', boxShadow: '0 12px 28px rgba(15, 23, 42, 0.16)', padding: '12px', zIndex: 600, boxSizing: 'border-box' },
    presetRow: { display: 'flex', flexWrap: 'wrap', gap: '6px', marginBottom: '10px' },
    presetBtn: { border: '1px solid #e2e8f0', backgroundColor: '#ffffff', color: '#475569', fontSize: '10px', fontWeight: 600, padding: '4px 9px', borderRadius: '999px', cursor: 'pointer', fontFamily: 'inherit' },
    rangeReadout: { display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '10px' },
    readoutBox: { flex: 1, display: 'flex', flexDirection: 'column', gap: '1px', padding: '5px 8px', border: '1px solid #e2e8f0', borderRadius: '8px', backgroundColor: '#f8fafc', transition: 'border-color 0.15s ease' },
    readoutLabel: { fontSize: '9px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' },
    readoutValue: { fontSize: '11px', fontWeight: 600 },
    calHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '6px' },
    navBtn: { width: '26px', height: '26px', display: 'flex', alignItems: 'center', justifyContent: 'center', border: 'none', borderRadius: '8px', backgroundColor: 'transparent', color: '#475569', cursor: 'pointer', padding: 0 },
    monthLabelBtn: { display: 'flex', alignItems: 'center', border: '1px solid transparent', backgroundColor: 'transparent', color: '#0f172a', fontSize: '12px', fontWeight: 700, padding: '4px 10px', borderRadius: '8px', cursor: 'pointer', fontFamily: 'inherit' },
    weekRow: { display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', marginBottom: '2px' },
    weekDay: { textAlign: 'center', fontSize: '10px', fontWeight: 700, color: '#94a3b8', padding: '4px 0' },
    dayGrid: { display: 'grid', gridTemplateColumns: 'repeat(7, 1fr)', rowGap: '2px' },
    dayCell: { height: '30px', display: 'flex', alignItems: 'center', justifyContent: 'center' },
    monthGrid: { display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '6px', padding: '4px 0 6px 0' },
    monthBtn: { border: 'none', borderRadius: '8px', padding: '10px 0', fontSize: '11px', cursor: 'pointer', fontFamily: 'inherit' },
    calFooter: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '10px', paddingTop: '10px', borderTop: '1px solid #f1f5f9' },
    linkBtn: { background: 'none', border: 'none', color: '#64748b', fontSize: '11px', fontWeight: 600, cursor: 'pointer', padding: '4px 6px', borderRadius: '6px', fontFamily: 'inherit' },
    doneBtn: { backgroundColor: '#004d3d', color: '#ffffff', border: 'none', borderRadius: '6px', padding: '5px 14px', fontSize: '11px', fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' },

    // Columns button + panel
    columnsBtn: { height: '32px', display: 'flex', alignItems: 'center', gap: '7px', padding: '5px 10px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', boxSizing: 'border-box', cursor: 'pointer', outline: 'none', fontFamily: 'inherit' },
    countBadge: { fontSize: '10px', fontWeight: 700, padding: '1px 7px', borderRadius: '999px', transition: 'background-color 0.15s ease, color 0.15s ease' },
    colPopover: { position: 'absolute', top: 'calc(100% + 6px)', right: 0, width: '300px', backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '12px', boxShadow: '0 12px 28px rgba(15, 23, 42, 0.16)', padding: '12px', zIndex: 700, boxSizing: 'border-box' },
    colHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' },
    colTitle: { fontSize: '12px', fontWeight: 700, color: '#0f172a' },
    colSubtitle: { fontSize: '10px', color: '#64748b', marginTop: '1px' },
    colSearch: { width: '100%', height: '30px', padding: '5px 10px 5px 28px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '11px', outline: 'none', backgroundColor: '#ffffff', color: '#1e293b', boxSizing: 'border-box', fontFamily: 'inherit' },
    colList: { maxHeight: '300px', overflowY: 'auto', margin: '0 -4px', padding: '0 4px', overscrollBehavior: 'contain' },
    groupHeader: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 6px 3px 6px', position: 'sticky', top: 0, backgroundColor: '#ffffff', zIndex: 1 },
    groupName: { fontSize: '10px', fontWeight: 700, color: '#475569' },
    groupCount: { fontWeight: 600, color: '#94a3b8', marginLeft: '4px' },
    groupBtn: { background: 'none', border: 'none', color: '#64748b', fontSize: '10px', fontWeight: 600, cursor: 'pointer', padding: '2px 6px', borderRadius: '5px', fontFamily: 'inherit' },
    checkRow: { display: 'flex', alignItems: 'center', gap: '9px', padding: '5px 8px', borderRadius: '6px', cursor: 'pointer', userSelect: 'none' },
    checkBox: { width: '15px', height: '15px', borderRadius: '4px', border: '1.5px solid #cbd5e1', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, boxSizing: 'border-box', transition: 'background-color 0.12s ease, border-color 0.12s ease' },

    // Table (scrolls both ways; header pinned)
    tableWrap: { flex: 1, minHeight: 0, overflow: 'auto', backgroundColor: '#ffffff', overscrollBehavior: 'contain' },
    table: { width: 'max-content', minWidth: '100%', borderCollapse: 'collapse', fontSize: '11px' },
    th: { position: 'sticky', top: 0, zIndex: 2, textAlign: 'left', padding: '9px 12px', backgroundColor: '#f8fafc', color: '#475569', fontWeight: 700, fontSize: '10px', borderBottom: '1px solid #e2e8f0', whiteSpace: 'nowrap', boxSizing: 'border-box' },
    td: { padding: '10px 12px', borderBottom: '1px solid #f1f5f9', color: '#334155', fontSize: '11px', verticalAlign: 'middle', whiteSpace: 'nowrap', boxSizing: 'border-box' },
    chip: { display: 'inline-block', padding: '2px 9px', borderRadius: '999px', fontSize: '10px', fontWeight: 600, whiteSpace: 'nowrap' },
    emptyState: { textAlign: 'center', color: '#94a3b8', fontSize: '12px', padding: '40px 16px', margin: 0 },
    footer: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 20px', borderTop: '1px solid #e2e8f0', backgroundColor: '#ffffff', flexShrink: 0 },
    footerText: { fontSize: '11px', color: '#64748b' },
    closeFooterBtn: { backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '6px 16px', fontSize: '11px', fontWeight: 600, color: '#334155', cursor: 'pointer' }
};

export default DMSAuditLog;
