import * as React from 'react';
import { useState, useRef, useEffect } from 'react';

// ---------------------------------------------------------------------------
// Custom dropdown used in the Settings pages.
// The look and behaviour match the FormSelect dropdown from the Process
// Management form (same sizes, colours, check mark and open animation),
// so every dropdown in DMS feels the same.
// ---------------------------------------------------------------------------

export interface ISettingsSelectOption<T extends string | number> {
  label: string;
  value: T;
}

export interface ISettingsSelectProps<T extends string | number> {
  value: T;
  options: ISettingsSelectOption<T>[];
  onChange: (next: T) => void;
  placeholder?: string;
  width?: string;
  disabled?: boolean;
}

// Animation and hover rules. Class names are prefixed with "st-" so they never
// clash with the "fm-" classes used by the Process Management form.
const SELECT_CSS = `
  @keyframes stSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
  .st-dropdown-menu { animation: stSlideDown 0.15s ease; }
  .st-select-box:hover { border-color: #94a3b8; }
  .st-select-option:hover { background-color: #f1f5f9; }
`;

export function SettingsSelect<T extends string | number>({ value, options, onChange, placeholder = 'Select', width, disabled }: ISettingsSelectProps<T>) {
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const wrapperRef = useRef<HTMLDivElement>(null);
  const selected = options.filter(option => option.value === value)[0];

  // Close the menu when the user clicks anywhere outside of it.
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleKeyDown = (event: React.KeyboardEvent) => {
    if (disabled) return;
    if (event.key === 'Escape') {
      setIsOpen(false);
    } else if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      setIsOpen(prev => !prev);
    }
  };

  return (
    <div style={{ ...styles.field, width, position: 'relative' }} ref={wrapperRef}>
      <style>{SELECT_CSS}</style>

      <div
        className="st-select-box"
        role="button"
        tabIndex={disabled ? -1 : 0}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
        onKeyDown={handleKeyDown}
        style={{
          ...styles.selectBox,
          borderColor: isOpen ? '#004d3d' : '#cbd5e1',
          boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none',
          opacity: disabled ? 0.6 : 1,
          cursor: disabled ? 'not-allowed' : 'pointer'
        }}
        onClick={() => !disabled && setIsOpen(prev => !prev)}
      >
        <span style={{ fontSize: '11px', color: selected ? '#1e293b' : '#94a3b8', fontWeight: 500 }}>
          {selected ? selected.label : placeholder}
        </span>
        <svg
          width="10"
          height="10"
          viewBox="0 0 24 24"
          fill="none"
          stroke="#64748b"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s ease', flexShrink: 0 }}
        >
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </div>

      {isOpen && !disabled && (
        <div className="st-dropdown-menu" role="listbox" style={styles.selectMenu}>
          {options.map(option => {
            const isSelected = option.value === value;
            return (
              <div
                key={String(option.value)}
                className="st-select-option"
                role="option"
                aria-selected={isSelected}
                style={{
                  ...styles.selectOption,
                  backgroundColor: isSelected ? '#e6f4ea' : 'transparent',
                  color: isSelected ? '#004d3d' : '#334155',
                  fontWeight: isSelected ? 600 : 400
                }}
                onClick={() => {
                  onChange(option.value);
                  setIsOpen(false);
                }}
              >
                <span>{option.label}</span>
                {isSelected && <span style={{ color: '#004d3d', fontWeight: 700 }}>{'✓'}</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

const styles: { [key: string]: React.CSSProperties } = {
  field: { display: 'flex', flexDirection: 'column', gap: '3px' },
  selectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', height: '28px', boxSizing: 'border-box', transition: 'border-color 0.15s ease, box-shadow 0.15s ease', outline: 'none' },
  selectMenu: { position: 'absolute', top: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 8px 20px rgba(0, 0, 0, 0.10)', zIndex: 500, padding: '4px', maxHeight: '220px', overflowY: 'auto' },
  selectOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '6px 8px', fontSize: '11px', borderRadius: '5px', cursor: 'pointer', margin: '1px 0' }
};

export default SettingsSelect;
