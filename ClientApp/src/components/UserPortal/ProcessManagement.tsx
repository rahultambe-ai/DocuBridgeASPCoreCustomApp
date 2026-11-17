import * as React from 'react';
import { useState, useMemo, useRef, useEffect } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '../../spfx-shim';
import ProcessManagementForm, { IProcessFormData } from './ProcessManagementForm';
import ProcessWiseWorkflow from './ProcessWiseWorkflow';

export interface IProcessMasterItem {
  id: number;
  processCode: string;
  processName: string;
  department: string;
  kraCode: string;
  documentsCount: number;
  kpisCount: number;
  weightage: number;
  status: 'Active' | 'Pending Approval' | 'Draft' | 'Archived' | 'Inactive';
  category?: string;
  processOwner?: string;
  reviewFrequency?: string;
  effectiveDate?: string;
  nextReviewDate?: string;
  description?: string;
}

interface CustomSelectProps {
  label: string;
  options: { label: string; value: string }[];
  value: string;
  onChange: (val: string) => void;
}

const CustomSelect: React.FC<CustomSelectProps> = ({ label, options, value, onChange }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  const selectedOption = options.find((opt) => opt.value === value);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setIsOpen(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div style={styles.dropdownWrapper} ref={dropdownRef}>
      <label style={styles.dropdownLabel}>{label}</label>
      <div
        className="pm-select-box"
        style={{ ...styles.customSelectBox, borderColor: isOpen ? '#004d3d' : '#cbd5e1', boxShadow: isOpen ? '0 0 0 3px rgba(0, 77, 61, 0.10)' : 'none' }}
        onClick={() => setIsOpen(!isOpen)}
      >
        <span style={{ fontSize: '11px', color: '#334155', fontWeight: 500 }}>{selectedOption ? selectedOption.label : 'Select'}</span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s' }}>
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </div>

      {isOpen && (
        <div className="pm-dropdown-menu" style={styles.dropdownMenu}>
          {options.map((opt) => {
            const isSelected = opt.value === value;
            return (
              <div key={opt.value} className="pm-dropdown-option" style={{ ...styles.dropdownOption, backgroundColor: isSelected ? '#e6f4ea' : 'transparent', color: isSelected ? '#004d3d' : '#334155', fontWeight: isSelected ? 600 : 400 }} onClick={() => { onChange(opt.value); setIsOpen(false); }}>
                <span>{opt.label}</span>
                {isSelected && <span style={{ color: '#004d3d', fontWeight: 'bold' }}>✓</span>}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

interface ItemsPerPageSelectProps {
  value: number;
  onChange: (val: number) => void;
  options: number[];
}

const ItemsPerPageSelect: React.FC<ItemsPerPageSelectProps> = ({ value, onChange, options }) => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setIsOpen(false);
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  return (
    <div style={{ position: 'relative' }} ref={dropdownRef}>
      {isOpen && (
        <div className="pm-dropdown-menu" style={styles.itemsPerPageMenu}>
          {options.map((opt) => {
            const isSelected = opt === value;
            return (
              <div
                key={opt}
                className="pm-items-option"
                style={{
                  ...styles.itemsPerPageOption,
                  backgroundColor: isSelected ? '#004d3d' : '#ffffff',
                  color: isSelected ? '#ffffff' : '#1e293b',
                  fontWeight: isSelected ? 700 : 500,
                }}
                onClick={() => {
                  onChange(opt);
                  setIsOpen(false);
                }}
              >
                {opt}
              </div>
            );
          })}
        </div>
      )}

      <div className="pm-select-box" style={{ ...styles.itemsPerPageBox, borderColor: isOpen ? '#004d3d' : '#cbd5e1' }} onClick={() => setIsOpen(!isOpen)}>
        <span style={{ fontSize: '11px', color: '#1e293b', fontWeight: 600 }}>{value}</span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s' }}>
          <polyline points="6 9 12 15 18 9"></polyline>
        </svg>
      </div>
    </div>
  );
};

export interface IProcessManagementProps {
  graphContext?: any;
}

export const ProcessManagement: React.FC<IProcessManagementProps> = ({ graphContext }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [departmentFilter, setDepartmentFilter] = useState('All');
  const [statusFilter, setStatusFilter] = useState('All');
  const [kraFilter, setKraFilter] = useState('All');
  const [rowsPerPage, setRowsPerPage] = useState(10);
  const [currentPage, setCurrentPage] = useState(1);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState<'create' | 'edit' | 'view'>('create');
  const [selectedProcess, setSelectedProcess] = useState<IProcessMasterItem | null>(null);
  const [activeMenuId, setActiveMenuId] = useState<number | null>(null);
  const [menuPos, setMenuPos] = useState<{ top?: number; bottom?: number; right: number } | null>(null);
  const [workflowPopupProcessCode, setWorkflowPopupProcessCode] = useState<string | null>(null);

  const [processes, setProcesses] = useState<IProcessMasterItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);

  const PROCESS_MASTER_LIST_TITLE = 'Process Master';
  const DOCUMENT_LIBRARY_NAME = 'Shared Documents';

  const fetchFolderDocumentCount = async (siteUrl: string, folderName: string): Promise<number> => {
    if (!graphContext?.spHttpClient || !folderName) return 0;
    try {
      const webServerRelativeUrl: string = graphContext.pageContext.web.serverRelativeUrl || '';
      const libraryRelativeUrl = webServerRelativeUrl && webServerRelativeUrl !== '/'
        ? `${webServerRelativeUrl}/${DOCUMENT_LIBRARY_NAME}`
        : `/${DOCUMENT_LIBRARY_NAME}`;
      const folderServerRelativeUrl = `${libraryRelativeUrl}/${folderName}`;

      const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderServerRelativeUrl)}')/Files?$select=Id`;

      const response: SPHttpClientResponse = await graphContext.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) return 0;
      const data = await response.json();
      return Array.isArray(data.value) ? data.value.length : 0;
    } catch {
      return 0;
    }
  };

  // Fetch items from SharePoint "Process Master" list on load
  useEffect(() => {
    const fetchProcessMasterItems = async () => {
      if (!graphContext?.spHttpClient || !graphContext?.pageContext) return;
      setIsLoading(true);
      try {
        const siteUrl = graphContext.pageContext.web.absoluteUrl;
        const endpoint = `${siteUrl}/_api/web/lists/getByTitle('${PROCESS_MASTER_LIST_TITLE}')/items?$select=Id,Title,ProcessCode,Department,ProcessCategory,ProcessOwner,ReviewFrequency,KRACode,Weightage,EffectiveDate,NextReviewDate,Description`;
        
        const response: SPHttpClientResponse = await graphContext.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
            }
          }
        );

        if (response.ok) {
          const data = await response.json();
          const mappedItems: IProcessMasterItem[] = data.value.map((item: any) => ({
            id: item.Id,
            processCode: item.ProcessCode || `PR-${item.Id}`,
            processName: item.Title || 'Untitled Process',
            department: item.Department || 'General',
            kraCode: item.KRACode || 'KRA-001',
            documentsCount: 0,
            kpisCount: 0,
            weightage: Number(item.Weightage) || 0,
            status: 'Active',
            category: item.ProcessCategory,
            processOwner: item.ProcessOwner,
            reviewFrequency: item.ReviewFrequency,
            effectiveDate: item.EffectiveDate ? item.EffectiveDate.split('T')[0] : '',
            nextReviewDate: item.NextReviewDate ? item.NextReviewDate.split('T')[0] : '',
            description: item.Description,
          }));
          setProcesses(mappedItems);

          // Fetch real-time document counts from the SharePoint library folder for each process
          const counts = await Promise.all(
            mappedItems.map((item) => fetchFolderDocumentCount(siteUrl, item.processCode))
          );
          setProcesses((prev) =>
            prev.map((item) => {
              const idx = mappedItems.findIndex((m) => m.id === item.id);
              return idx >= 0 ? { ...item, documentsCount: counts[idx] } : item;
            })
          );
        }
      } catch (err) {
        console.error('Error fetching Process Master items:', err);
      } finally {
        setIsLoading(false);
      }
    };

    fetchProcessMasterItems();
  }, [graphContext]);


  // Close action dropdown on outside click
  useEffect(() => {
    const handleClickOutside = () => setActiveMenuId(null);
    document.addEventListener('click', handleClickOutside);
    return () => document.removeEventListener('click', handleClickOutside);
  }, []);

  const departments = useMemo(() => ['All', ...Array.from(new Set(processes.map((p) => p.department)))], [processes]);
  const kras = useMemo(() => ['All', ...Array.from(new Set(processes.map((p) => p.kraCode)))], [processes]);
  const statuses = useMemo(() => ['All', ...Array.from(new Set(processes.map((p) => p.status)))], [processes]);

  const filteredProcesses = useMemo(() => {
    return processes.filter((item) => {
      const matchesSearch = item.processName.toLowerCase().includes(searchTerm.toLowerCase()) || item.processCode.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesDept = departmentFilter === 'All' || item.department === departmentFilter;
      const matchesStatus = statusFilter === 'All' || item.status === statusFilter;
      const matchesKra = kraFilter === 'All' || item.kraCode === kraFilter;
      return matchesSearch && matchesDept && matchesStatus && matchesKra;
    });
  }, [processes, searchTerm, departmentFilter, statusFilter, kraFilter]);

  const totalPages = Math.ceil(filteredProcesses.length / rowsPerPage) || 1;
  const paginatedProcesses = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredProcesses.slice(start, start + rowsPerPage);
  }, [filteredProcesses, currentPage, rowsPerPage]);

  const handleSaveProcess = (data: IProcessFormData) => {
    if (modalMode === 'edit' && selectedProcess) {
      setProcesses((prev) =>
        prev.map((p) =>
          p.id === selectedProcess.id
            ? {
                ...p,
                processName: data.processName,
                processCode: data.processCode,
                department: data.department,
                kraCode: data.kraCode,
                weightage: data.weightage,
                status: data.status,
                category: data.category,
                processOwner: data.processOwner,
                reviewFrequency: data.reviewFrequency,
                effectiveDate: data.effectiveDate,
                nextReviewDate: data.nextReviewDate,
                description: data.description,
              }
            : p
        )
      );
    } else {
      const newItem: IProcessMasterItem = {
        id: processes.length ? Math.max(...processes.map((p) => p.id)) + 1 : 1,
        processCode: data.processCode || `PR-${String(processes.length + 1).padStart(3, '0')}`,
        processName: data.processName || 'Untitled Process',
        department: data.department,
        kraCode: data.kraCode,
        documentsCount: 0,
        kpisCount: 0,
        weightage: data.weightage,
        status: data.status,
        category: data.category,
        processOwner: data.processOwner,
        reviewFrequency: data.reviewFrequency,
        effectiveDate: data.effectiveDate,
        nextReviewDate: data.nextReviewDate,
        description: data.description,
      };
      setProcesses((prev) => [newItem, ...prev]);
    }
    setIsModalOpen(false);
    setSelectedProcess(null);
  };

  return (
    <div style={styles.container}>
      <style>{`
        @keyframes pmSlideDown { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .pm-dropdown-menu { animation: pmSlideDown 0.15s ease; }
        .pm-select-box:hover { border-color: #94a3b8; }
        .pm-dropdown-option:hover { background-color: #f1f5f9; }
        .pm-items-option:hover { filter: brightness(0.97); }
        .pm-create-btn { transition: transform 0.15s ease, filter 0.15s ease, box-shadow 0.15s ease; }
        .pm-create-btn:hover { transform: translateY(-1px); filter: brightness(1.08); box-shadow: 0 4px 10px rgba(0, 77, 61, 0.25); }
        .pm-create-btn:active { transform: translateY(0); }
        .pm-table-row { transition: background-color 0.15s ease; }
        .pm-table-row:hover { background-color: #f8fafc; }
        .pm-action-btn { transition: background-color 0.15s ease, transform 0.15s ease; border-radius: 6px; }
        .pm-action-btn:hover { background-color: #e2e8f0; }
        .pm-menu-item:hover { background-color: #f1f5f9; color: #004d3d; }
        .pm-page-btn { transition: all 0.15s ease; }
        .pm-page-btn:hover:not(:disabled) { border-color: #004d3d; }
        .pm-search-input { transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        .pm-search-input:focus { border-color: #004d3d; box-shadow: 0 0 0 3px rgba(0, 77, 61, 0.10); }
      `}</style>

      <div style={styles.header}>
        <div>
          <h1 style={styles.title}>Process Master</h1>
          <p style={styles.subtitle}>Manage all organizational processes</p>
        </div>
        <button
          className="pm-create-btn"
          style={styles.createBtn}
          onClick={() => {
            setModalMode('create');
            setSelectedProcess(null);
            setIsModalOpen(true);
          }}
        >
          <span style={{ fontSize: '13px', marginRight: '4px', fontWeight: 'bold' }}>+</span> Create Process
        </button>
      </div>

      <div style={styles.filterRow}>
        <div style={styles.searchContainer}>
          <svg style={styles.searchIcon} width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
          <input className="pm-search-input" type="text" placeholder="Search by title, number, type, tags..." value={searchTerm} onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }} style={styles.searchInput} />
        </div>

        <div style={styles.dropdownGroup}>
          <CustomSelect label="Department" value={departmentFilter} onChange={(val) => { setDepartmentFilter(val); setCurrentPage(1); }} options={departments.map((dept) => ({ label: dept === 'All' ? 'All Departments' : dept, value: dept }))} />
          <CustomSelect label="Status" value={statusFilter} onChange={(val) => { setStatusFilter(val); setCurrentPage(1); }} options={statuses.map((st) => ({ label: st === 'All' ? 'All Status' : st, value: st }))} />
          <CustomSelect label="KRA" value={kraFilter} onChange={(val) => { setKraFilter(val); setCurrentPage(1); }} options={kras.map((kra) => ({ label: kra === 'All' ? 'All KRAs' : kra, value: kra }))} />
        </div>
      </div>

      <div style={styles.tableCard}>
        <table style={styles.table}>
          <thead>
            <tr style={styles.thRow}>
              <th style={styles.th}>PROCESS CODE</th>
              <th style={styles.th}>PROCESS NAME</th>
              <th style={styles.th}>DEPARTMENT</th>
              <th style={styles.th}>KRA CODE</th>
              <th style={styles.th}>DOCUMENTS</th>
              <th style={styles.th}>KPIS</th>
              <th style={styles.th}>WEIGHTAGE (%)</th>
              <th style={styles.th}>STATUS</th>
              <th style={{ ...styles.th, textAlign: 'center' }}>ACTIONS</th>
            </tr>
          </thead>
          <tbody>
            {isLoading ? (
              <tr>
                <td style={{ ...styles.td, textAlign: 'center', padding: '24px', color: '#64748b' }} colSpan={9}>
                  Loading processes from SharePoint...
                </td>
              </tr>
            ) : paginatedProcesses.map((item) => (
              <tr key={item.id} className="pm-table-row" style={styles.tr}>
                <td style={styles.td}><span style={styles.docBadge}>{item.processCode}</span></td>
                <td style={{ ...styles.td, fontWeight: 600, color: '#0f172a' }}>{item.processName}</td>
                <td style={styles.td}>{item.department}</td>
                <td style={styles.td}>{item.kraCode}</td>
                <td style={styles.td}><span style={styles.greenTextCount}>{item.documentsCount}</span></td>
                <td style={styles.td}><span style={styles.greenTextCount}>{item.kpisCount}</span></td>
                <td style={styles.td}>{item.weightage}%</td>
                <td style={styles.td}>{getStatusPill(item.status)}</td>
                <td style={{ ...styles.td, textAlign: 'center', position: 'relative' }}>
                  <button
                    className="pm-action-btn"
                    style={styles.actionMenuBtn}
                    title="Actions"
                    onClick={(e) => {
                      e.stopPropagation();
                      if (activeMenuId === item.id) {
                        setActiveMenuId(null);
                      } else {
                        const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                        const estimatedMenuHeight = 130;
                        const spaceBelow = window.innerHeight - rect.bottom;
                        const openUp = spaceBelow < estimatedMenuHeight;
                        setMenuPos({
                          top: openUp ? undefined : rect.bottom + 4,
                          bottom: openUp ? window.innerHeight - rect.top + 4 : undefined,
                          right: window.innerWidth - rect.right
                        });
                        setActiveMenuId(item.id);
                      }
                    }}
                  >
                    <svg width="4" height="14" viewBox="0 0 4 14" fill="#64748b" xmlns="http://www.w3.org/2000/svg">
                      <circle cx="2" cy="2" r="1.5" />
                      <circle cx="2" cy="7" r="1.5" />
                      <circle cx="2" cy="12" r="1.5" />
                    </svg>
                  </button>

                  {activeMenuId === item.id && menuPos && (
                    <div
                      className="pm-dropdown-menu"
                      style={{
                        ...styles.actionsDropdown,
                        position: 'fixed',
                        left: 'auto',
                        right: menuPos.right,
                        top: menuPos.top,
                        bottom: menuPos.bottom
                      }}
                    >
                      <div
                        className="pm-menu-item"
                        style={styles.menuItem}
                        onClick={() => {
                          setSelectedProcess(item);
                          setModalMode('view');
                          setIsModalOpen(true);
                          setActiveMenuId(null);
                        }}
                      >
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>
                        <span>View</span>
                      </div>
                      <div
                        className="pm-menu-item"
                        style={styles.menuItem}
                        onClick={() => {
                          setSelectedProcess(item);
                          setModalMode('edit');
                          setIsModalOpen(true);
                          setActiveMenuId(null);
                        }}
                      >
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                        <span>Edit</span>
                      </div>
                      <div
                        className="pm-menu-item"
                        style={styles.menuItem}
                        onClick={() => {
                          setWorkflowPopupProcessCode(item.processCode);
                          setActiveMenuId(null);
                        }}
                      >
                        <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#475569" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                          <rect x="3" y="3" width="6" height="6" rx="1" /><rect x="15" y="15" width="6" height="6" rx="1" />
                          <rect x="3" y="15" width="6" height="6" rx="1" /><path d="M9 6h6a3 3 0 0 1 3 3v6" /><path d="M9 18h3" />
                        </svg>
                        <span>Workflow</span>
                      </div>
                    </div>
                  )}
                </td>
              </tr>
            ))}
            {!isLoading && paginatedProcesses.length === 0 && (
              <tr>
                <td style={{ ...styles.td, textAlign: 'center', padding: '24px', color: '#94a3b8' }} colSpan={9}>
                  No processes match your filters.
                </td>
              </tr>
            )}
          </tbody>
        </table>

        <div style={styles.paginationRow}>
          <span style={styles.pageInfo}>
            Showing {filteredProcesses.length > 0 ? (currentPage - 1) * rowsPerPage + 1 : 0} to {Math.min(currentPage * rowsPerPage, filteredProcesses.length)} of {filteredProcesses.length} items
          </span>

          <div style={styles.paginationControls}>
            <button
              className="pm-page-btn"
              style={styles.pageNavBtn}
              onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
              disabled={currentPage === 1}
            >
              <svg width="6" height="10" viewBox="0 0 6 10" fill="none" stroke={currentPage === 1 ? '#cbd5e1' : '#475569'} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="5 1 1 5 5 9"></polyline>
              </svg>
            </button>

            {Array.from({ length: totalPages }, (_, i) => i + 1).map((pageNum) => {
              const isActive = currentPage === pageNum;
              return (
                <button
                  key={pageNum}
                  className="pm-page-btn"
                  style={{
                    ...styles.pageBtn,
                    backgroundColor: isActive ? '#004d3d' : '#ffffff',
                    color: isActive ? '#ffffff' : '#334155',
                    borderColor: isActive ? '#004d3d' : '#e2e8f0',
                    fontWeight: isActive ? 700 : 500,
                  }}
                  onClick={() => setCurrentPage(pageNum)}
                >
                  {pageNum}
                </button>
              );
            })}

            <button
              className="pm-page-btn"
              style={styles.pageNavBtn}
              onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
              disabled={currentPage === totalPages}
            >
              <svg width="6" height="10" viewBox="0 0 6 10" fill="none" stroke={currentPage === totalPages ? '#cbd5e1' : '#475569'} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="1 1 5 5 1 9"></polyline>
              </svg>
            </button>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={styles.pageInfo}>Items per page:</span>
            <ItemsPerPageSelect
              value={rowsPerPage}
              options={[5, 10, 20, 50]}
              onChange={(num) => {
                setRowsPerPage(num);
                setCurrentPage(1);
              }}
            />
          </div>
        </div>
      </div>

      {isModalOpen && (
        <ProcessManagementForm
          key={selectedProcess ? selectedProcess.id : 'new'}
          isOpen={isModalOpen}
          onClose={() => setIsModalOpen(false)}
          onSubmit={handleSaveProcess}
          graphContext={graphContext}
          selectedId={selectedProcess ? selectedProcess.id : undefined}
          initialData={
            selectedProcess
              ? {
                  processName: selectedProcess.processName,
                  processCode: selectedProcess.processCode,
                  department: selectedProcess.department,
                  category: selectedProcess.category || 'SOP',
                  processOwner: selectedProcess.processOwner || '',
                  reviewFrequency: selectedProcess.reviewFrequency || 'Monthly',
                  kraCode: selectedProcess.kraCode,
                  weightage: selectedProcess.weightage,
                  effectiveDate: selectedProcess.effectiveDate || '',
                  nextReviewDate: selectedProcess.nextReviewDate || '',
                  description: selectedProcess.description || '',
                  status: selectedProcess.status as any,
                }
              : undefined
          }
          isViewMode={modalMode === 'view'}
        />
      )}

      {workflowPopupProcessCode && (
        <ProcessWiseWorkflow
          isOpen={!!workflowPopupProcessCode}
          onClose={() => setWorkflowPopupProcessCode(null)}
          context={graphContext}
          processCode={workflowPopupProcessCode}
        />
      )}
    </div>
  );
};

const getStatusPill = (status: string) => {
  const baseStyle: React.CSSProperties = { padding: '2px 8px', borderRadius: '10px', fontSize: '10px', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '3px' };
  switch (status) {
    case 'Active': return <span style={{ ...baseStyle, backgroundColor: '#e6f4ea', color: '#137333' }}>• Active</span>;
    case 'Pending Approval': return <span style={{ ...baseStyle, backgroundColor: '#fef3c7', color: '#b45309' }}>• Pending Approval</span>;
    case 'Draft': return <span style={{ ...baseStyle, backgroundColor: '#f1f5f9', color: '#475569' }}>• Draft</span>;
    case 'Archived': return <span style={{ ...baseStyle, backgroundColor: '#e2e8f0', color: '#475569' }}>• Archived</span>;
    default: return <span style={{ ...baseStyle, backgroundColor: '#f1f5f9', color: '#64748b' }}>• {status}</span>;
  }
};

const styles: { [key: string]: React.CSSProperties } = {
  container: { fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', backgroundColor: '#f8fafc', padding: '0 16px 16px 0', minHeight: '100vh', color: '#1e293b' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' },
  title: { margin: 0, fontSize: '18px', fontWeight: 700, color: '#0f172a' },
  subtitle: { margin: '2px 0 0 0', fontSize: '11px', color: '#64748b' },
  createBtn: { backgroundColor: '#004d3d', color: '#ffffff', border: 'none', borderRadius: '4px', padding: '6px 12px', fontSize: '11px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center' },
  filterRow: { display: 'flex', gap: '10px', marginBottom: '14px', alignItems: 'flex-end', flexWrap: 'wrap' },
  searchContainer: { position: 'relative', flex: '1 1 240px' },
  searchInput: { width: '100%', padding: '6px 10px 6px 30px', borderRadius: '6px', border: '1px solid #e2e8f0', fontSize: '11px', outline: 'none', backgroundColor: '#ffffff', boxSizing: 'border-box' },
  searchIcon: { position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)' },
  dropdownGroup: { display: 'flex', gap: '8px', alignItems: 'flex-end' },
  dropdownWrapper: { position: 'relative', display: 'flex', flexDirection: 'column', gap: '3px', minWidth: '130px' },
  dropdownLabel: { fontSize: '10px', fontWeight: 600, color: '#64748b' },
  customSelectBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 8px', borderRadius: '6px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', cursor: 'pointer', height: '30px', boxSizing: 'border-box', transition: 'border-color 0.15s ease, box-shadow 0.15s ease' },
  dropdownMenu: { position: 'absolute', top: 'calc(100% + 2px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '6px', boxShadow: '0 4px 12px rgba(0, 0, 0, 0.08)', zIndex: 100, padding: '4px', maxHeight: '180px', overflowY: 'auto' },
  dropdownOption: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 8px', fontSize: '11px', borderRadius: '4px', cursor: 'pointer', margin: '1px 0', transition: 'background-color 0.12s ease' },
  tableCard: { backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '6px', overflow: 'hidden' },
  table: { width: '100%', borderCollapse: 'collapse', textAlign: 'left' },
  thRow: { backgroundColor: '#ffffff', borderBottom: '1px solid #e2e8f0' },
  th: { padding: '8px 12px', fontSize: '10px', fontWeight: 700, color: '#64748b', letterSpacing: '0.04em', whiteSpace: 'nowrap' },
  tr: { borderBottom: '1px solid #f1f5f9' },
  td: { padding: '8px 12px', fontSize: '11px', color: '#334155', whiteSpace: 'nowrap' },
  docBadge: { backgroundColor: '#e6f4ea', color: '#004d3d', padding: '2px 6px', borderRadius: '4px', fontSize: '10px', fontWeight: 600 },
  greenTextCount: { color: '#004d3d', fontWeight: 600, fontSize: '11px' },
  actionMenuBtn: { 
    background: 'transparent', 
    border: 'none', 
    cursor: 'pointer', 
    padding: '2px 4px', 
    width: '11px', 
    height: '11px', 
    display: 'inline-flex', 
    alignItems: 'center', 
    justifyContent: 'center',
    borderRadius: '4px'
  },
  
  //actionMenuBtn: { background: 'transparent', border: 'none', cursor: 'pointer', padding: '4px 8px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center' },
  actionsDropdown: { position: 'absolute', right: '12px', top: 'calc(100% - 2px)', backgroundColor: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '6px', boxShadow: '0 4px 12px rgba(0, 0, 0, 0.1)', zIndex: 200, padding: '4px', minWidth: '90px', textAlign: 'left' },
  menuItem: { display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 8px', fontSize: '11px', borderRadius: '4px', cursor: 'pointer', color: '#334155' },
  paginationRow: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '12px 16px', backgroundColor: '#ffffff', borderTop: '1px solid #f1f5f9' },
  pageInfo: { fontSize: '11px', color: '#64748b' },
  paginationControls: { display: 'flex', gap: '6px', alignItems: 'center' },
  pageBtn: { width: '28px', height: '28px', borderRadius: '8px', border: '1px solid #e2e8f0', fontSize: '12px', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', transition: 'all 0.15s ease' },
  pageNavBtn: { width: '28px', height: '28px', borderRadius: '8px', border: '1px solid #e2e8f0', backgroundColor: '#ffffff', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center' },
  itemsPerPageBox: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '4px 8px', borderRadius: '8px', border: '1px solid #cbd5e1', backgroundColor: '#ffffff', cursor: 'pointer', minWidth: '48px', height: '28px', boxSizing: 'border-box', gap: '6px' },
  itemsPerPageMenu: { position: 'absolute', bottom: 'calc(100% + 4px)', left: 0, right: 0, backgroundColor: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '8px', boxShadow: '0 -4px 12px rgba(0,0,0,0.08)', zIndex: 100, overflow: 'hidden', padding: '2px 0' },
  itemsPerPageOption: { textAlign: 'center', padding: '5px 0', fontSize: '11px', cursor: 'pointer' },
};

export default ProcessManagement;