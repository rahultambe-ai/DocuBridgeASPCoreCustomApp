import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { SPFx } from '../../spfx-shim';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

/* ============================================================
   Document Type field — dropdown + "add new type" popup.
   No hardcoded option list: the dropdown's values are fetched live
   from the "Document Type List" SharePoint list (Title column), and
   adding a new type here writes a new item straight into that same
   list — https://mkx8.sharepoint.com/Lists/Document%20Type%20List/AllItems.aspx
   ============================================================ */

// Same list used for both reading the dropdown's options and writing new ones.
const DOCUMENT_TYPE_LIST_TITLE = 'Document Type List';

// NOTE: "Type List ID" is the column's DISPLAY name in the list view. Its internal
// SharePoint field name is assumed to be "TypeListID" (no spaces) below — if the
// column was created through "+ Add column" with the spaced label typed directly,
// SharePoint may instead have generated an internal name like "Type_x0020_List_x0020_ID".
// If new items don't land correctly, open the column's settings and confirm the
// internal name from its URL, then swap it in here.
const TYPE_LIST_ID_FIELD = 'TypeListID';

export interface IDocumentTypeFieldProps {
  /** Currently selected Document Type value. */
  value: string;
  /** Fired when the user picks a different option from the dropdown, or adds a new one. */
  onChange: (val: string) => void;
  /** SPFx web part context, needed to talk to SharePoint. */
  context: any;
}

const Icon = ({ name, color = 'currentColor', size = 14 }: { name: string; color?: string; size?: number }) => {
  const paths: { [key: string]: React.ReactNode } = {
    plus: <path d="M12 4v16m8-8H4" />,
    chevronDown: <path d="M6 9l6 6 6-6" />,
    check: <path d="M20 6L9 17l-5-5" />,
    close: <path d="M6 18L18 6M6 6l12 12" />
  };
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'inline-block', verticalAlign: 'middle', flexShrink: 0 }}>
      {paths[name] || null}
    </svg>
  );
};

// Same datetime-based id pattern used elsewhere in the app (e.g. Document ID on
// uploaded files): 31-Aug-2026 14:30:05 -> "20260831143005".
const generateTypeListId = (): string => {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
};

export const DocumentTypeField: React.FC<IDocumentTypeFieldProps> = ({ value, onChange, context }) => {
  const [options, setOptions] = useState<string[]>([]);
  const [isLoadingOptions, setIsLoadingOptions] = useState<boolean>(true);

  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const [isPopupOpen, setIsPopupOpen] = useState(false);
  const [newTypeInput, setNewTypeInput] = useState('');
  const [popupError, setPopupError] = useState('');
  const [isSavingType, setIsSavingType] = useState(false);

  const dropdownRef = useRef<HTMLDivElement>(null);
  const popupRef = useRef<HTMLDivElement>(null);
  const popupInputRef = useRef<HTMLInputElement>(null);

  // Live-fetches every Title from the Document Type List so the dropdown always
  // reflects what's actually in SharePoint — nothing hardcoded here.
  useEffect(() => {
    let cancelled = false;
    const fetchOptions = async () => {
      if (!context) return;
      setIsLoadingOptions(true);
      try {
        const sp = spfi().using(SPFx(context));
        const items: any[] = await sp.web.lists
          .getByTitle(DOCUMENT_TYPE_LIST_TITLE)
          .items.select('Title')
          .top(500)();
        const titles = items.map((i) => i.Title).filter((t): t is string => !!t);
        if (!cancelled) {
          setOptions(titles);
          if (titles.length > 0 && !titles.includes(value)) {
            onChange(titles[0]);
          }
        }
      } catch (err) {
        console.error('Error fetching Document Type List options:', err);
      } finally {
        if (!cancelled) setIsLoadingOptions(false);
      }
    };
    fetchOptions();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [context]);

  // Close the dropdown menu on outside click.
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // Close the "add type" popup on outside click, independent of the dropdown.
  useEffect(() => {
    const handleClickOutsidePopup = (event: MouseEvent) => {
      if (popupRef.current && !popupRef.current.contains(event.target as Node)) {
        setIsPopupOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutsidePopup);
    return () => document.removeEventListener('mousedown', handleClickOutsidePopup);
  }, []);

  useEffect(() => {
    if (isPopupOpen) {
      setTimeout(() => popupInputRef.current?.focus(), 0);
    }
  }, [isPopupOpen]);

  const handleTogglePopup = () => {
    setIsDropdownOpen(false);
    setPopupError('');
    setNewTypeInput('');
    setIsPopupOpen((prev) => !prev);
  };

  // Writes the new type straight into the Document Type List (Title = what the user
  // typed, the Type List ID column = a datetime-based id), then adds it to the
  // dropdown locally so it's selectable immediately without a re-fetch.
  const handleAddType = async () => {
    const trimmed = newTypeInput.trim();
    if (!trimmed) {
      setPopupError('Please enter a document type.');
      return;
    }
    if (options.some((opt) => opt.toLowerCase() === trimmed.toLowerCase())) {
      setPopupError('This document type already exists.');
      return;
    }

    setIsSavingType(true);
    setPopupError('');
    try {
      const sp = spfi().using(SPFx(context));
      await sp.web.lists.getByTitle(DOCUMENT_TYPE_LIST_TITLE).items.add({
        Title: trimmed,
        [TYPE_LIST_ID_FIELD]: generateTypeListId()
      });

      setOptions((prev) => [...prev, trimmed]);
      onChange(trimmed);
      setNewTypeInput('');
      setIsPopupOpen(false);
    } catch (err) {
      console.error('Error adding new Document Type to SharePoint:', err);
      setPopupError('Could not save this type. Please try again.');
    } finally {
      setIsSavingType(false);
    }
  };

  return (
    <div style={styles.wrapper}>
      <div style={styles.labelRow}>
        <label style={styles.fieldLabel}>Document Type *</label>
        <button
          type="button"
          onClick={handleTogglePopup}
          style={styles.addTypeBtn}
          title="Add new document type"
        >
          <Icon name="plus" color="#005a43" size={13} />
        </button>
      </div>

      <div style={styles.ddWrapper} ref={dropdownRef}>
        <div
          onClick={() => setIsDropdownOpen((prev) => !prev)}
          style={{
            ...styles.ddControl,
            borderColor: isDropdownOpen ? '#005a43' : '#cbd5e1',
            boxShadow: isDropdownOpen ? '0 0 0 2px rgba(0, 90, 67, 0.15)' : 'none'
          }}
        >
          <span style={{ ...styles.ddValueText, ...(!value && !isLoadingOptions ? styles.ddPlaceholderText : {}) }}>
            {isLoadingOptions && options.length === 0 ? 'Loading…' : (value || 'Select type')}
          </span>
          <span style={{ transform: isDropdownOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', display: 'flex' }}>
            <Icon name="chevronDown" color="#64748b" size={14} />
          </span>
        </div>

        {isDropdownOpen && (
          <div style={styles.ddMenu}>
            {options.length === 0 ? (
              <div style={styles.ddEmptyState}>
                {isLoadingOptions ? 'Loading types…' : 'No document types yet — use + to add one.'}
              </div>
            ) : (
              options.map((opt) => {
                const isSelected = opt === value;
                return (
                  <div
                    key={opt}
                    onClick={() => { onChange(opt); setIsDropdownOpen(false); }}
                    style={{
                      ...styles.ddOption,
                      backgroundColor: isSelected ? '#f0fdf4' : '#ffffff',
                      color: isSelected ? '#005a43' : '#334155',
                      fontWeight: isSelected ? 600 : 500
                    }}
                  >
                    <span>{opt}</span>
                    {isSelected && <Icon name="check" color="#005a43" size={14} />}
                  </div>
                );
              })
            )}
          </div>
        )}

        {isPopupOpen && (
          <div ref={popupRef} style={styles.popupCard}>
            <div style={styles.popupHeaderRow}>
              <span style={styles.popupTitle}>Add Document Type</span>
              <button type="button" onClick={() => setIsPopupOpen(false)} style={styles.popupCloseBtn}>
                <Icon name="close" color="#94a3b8" size={12} />
              </button>
            </div>
            <input
              ref={popupInputRef}
              type="text"
              value={newTypeInput}
              placeholder="e.g. Vendor Agreement"
              disabled={isSavingType}
              onChange={(e) => { setNewTypeInput(e.target.value); if (popupError) setPopupError(''); }}
              onKeyDown={(e) => { if (e.key === 'Enter') handleAddType(); }}
              style={styles.popupInput}
            />
            {popupError && <div style={styles.popupError}>{popupError}</div>}
            <div style={styles.popupFooterRow}>
              <button type="button" onClick={() => setIsPopupOpen(false)} style={styles.popupCancelBtn} disabled={isSavingType}>Cancel</button>
              <button type="button" onClick={handleAddType} style={{ ...styles.popupAddBtn, opacity: isSavingType ? 0.6 : 1 }} disabled={isSavingType}>
                {isSavingType ? 'Adding…' : 'Add'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  wrapper: { marginTop: '12px' },
  labelRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' },
  fieldLabel: { fontSize: '10px', fontWeight: 600, color: '#475569' },
  addTypeBtn: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '20px',
    height: '20px',
    borderRadius: '5px',
    border: '1px solid #cbd5e1',
    backgroundColor: '#f0fdf4',
    cursor: 'pointer',
    padding: 0,
    flexShrink: 0
  },

  ddWrapper: { position: 'relative', width: '100%' },
  ddControl: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 9px', backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#334155', cursor: 'pointer', transition: 'all 0.15s ease-in-out', boxSizing: 'border-box' },
  ddValueText: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  ddPlaceholderText: { color: '#94a3b8', fontWeight: 400 },
  ddMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1)', zIndex: 999, overflow: 'hidden', padding: '4px' },
  ddOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '7px 10px', borderRadius: '6px', fontSize: '11px', cursor: 'pointer', marginBottom: '2px' },
  ddEmptyState: { padding: '10px', fontSize: '10px', color: '#94a3b8', textAlign: 'center' },

  popupCard: {
    position: 'absolute',
    top: 'calc(100% + 4px)',
    right: 0,
    width: '220px',
    backgroundColor: '#ffffff',
    border: '1px solid #e2e8f0',
    borderRadius: '8px',
    boxShadow: '0 10px 25px -5px rgba(0,0,0,0.12), 0 8px 10px -6px rgba(0,0,0,0.08)',
    zIndex: 1000,
    padding: '10px'
  },
  popupHeaderRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' },
  popupTitle: { fontSize: '11px', fontWeight: 700, color: '#1e293b' },
  popupCloseBtn: { background: 'none', border: 'none', cursor: 'pointer', padding: '2px', display: 'flex' },
  popupInput: { width: '100%', padding: '7px 9px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '11px', color: '#1e293b', boxSizing: 'border-box', outline: 'none' },
  popupError: { fontSize: '9px', color: '#dc2626', marginTop: '5px' },
  popupFooterRow: { display: 'flex', justifyContent: 'flex-end', gap: '6px', marginTop: '10px' },
  popupCancelBtn: { padding: '5px 10px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', borderRadius: '5px', fontSize: '10px', fontWeight: 600, color: '#334155', cursor: 'pointer' },
  popupAddBtn: { padding: '5px 12px', border: 'none', backgroundColor: '#005a43', borderRadius: '5px', fontSize: '10px', fontWeight: 600, color: '#ffffff', cursor: 'pointer' }
};

export default DocumentTypeField;
