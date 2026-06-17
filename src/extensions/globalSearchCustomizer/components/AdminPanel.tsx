import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import {
  IAdminPanelProps,
  saveAdminConfig,
} from '../interface/IAdminPanelProps';
import {
  X,
  Plus,
  Trash2,
  Folder,
  Globe,
  Settings2,
  Info,
  Calendar,
} from 'lucide-react';
const SITE_PRESETS = ['Trivandi KSA', 'Trivandi Academy', 'Trivandi Canada', 'Trivandi UAE', 'Trivandi India', 'Trivandi Singapore'];
const FILE_TYPE_PRESETS = ['PDF', 'XLSX', 'PPTX', 'PNG', 'ZIP', 'CSV'];
const DATE_PRESETS = ['Past 24 Hours', 'Past 7 Days', 'Past 30 Days', 'Past Quarter', 'Past Year'];

export const AdminPanel: React.FC<IAdminPanelProps> = ({
  isOpen,
  onClose,
  config,
  onConfigChange,
  adminConfigService,
}) => {
  const [newSite, setNewSite] = useState('');
  const [newFileType, setNewFileType] = useState('');
  const [newDateFilter, setNewDateFilter] = useState('');
  const [errorText, setErrorText] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    type: 'site' | 'fileType' | 'dateFilter';
    value: string;
  } | null>(null);

  const confirmDelete = (type: 'site' | 'fileType' | 'dateFilter', value: string): void => {
    setDeleteConfirmation({ type, value });
  };

  const executeDelete = async (): Promise<void> => {
    if (!deleteConfirmation) return;
    const { type, value } = deleteConfirmation;
    setDeleteConfirmation(null);

    if (type === 'site') {
      await handleRemoveSite(value);
    } else if (type === 'fileType') {
      await handleRemoveFileType(value);
    } else if (type === 'dateFilter') {
      await handleRemoveDateFilter(value);
    }
  };

  const [siteSuggestions, setSiteSuggestions] = useState<{ title: string; slug: string }[]>([]);
  const [isLoadingSuggestions, setIsLoadingSuggestions] = useState(false);

  useEffect(() => {
    const val = newSite.trim();
    if (val.length < 2 || !adminConfigService) {
      setSiteSuggestions([]);
      return;
    }

    setIsLoadingSuggestions(true);
    const delayDebounce = setTimeout(async () => {
      try {
        const results = await adminConfigService.searchSharePointSites(val);
        setSiteSuggestions(results);
      } catch (e) {
        console.error('Failed to search site suggestions:', e);
      } finally {
        setIsLoadingSuggestions(false);
      }
    }, 350);

    return () => clearTimeout(delayDebounce);
  }, [newSite, adminConfigService]);

  if (!isOpen) return null;

  const updateConfig = (updated: typeof config): void => {
    saveAdminConfig(updated);
    onConfigChange(updated);
  };

  // ── Sites handlers ────────────────────────
  const handleAddSite = async (e: React.FormEvent): Promise<void> => {
    if (e) e.preventDefault();
    const val = newSite.trim();
    if (!val) return;
    
    if (config.sites.map(s => s.toLowerCase()).includes(val.toLowerCase())) {
      setErrorText('Site already exists');
      return;
    }
    
    setIsSaving(true);
    setErrorText('');
    try {
      const slug = val.replace(/\s+/g, '');
      if (adminConfigService) {
        await adminConfigService.addSite(val, slug);
      }
      updateConfig({ 
        ...config, 
        sites: [...config.sites, val] 
      });
      setNewSite('');
    } catch (err) {
      console.error('Add site failed:', err);
      // Fallback
      updateConfig({ 
        ...config, 
        sites: [...config.sites, val] 
      });
      setNewSite('');
    } finally {
      setIsSaving(false);
    }
  };

  const handleRemoveSite = async (site: string): Promise<void> => {
    if (site === 'All') return;
    setIsSaving(true);
    try {
      if (adminConfigService) {
        await adminConfigService.removeSite(site);
      }
      updateConfig({
        ...config,
        sites: config.sites.filter((s) => s !== site),
      });
    } catch (err) {
      console.error('Remove site failed:', err);
      updateConfig({
        ...config,
        sites: config.sites.filter((s) => s !== site),
      });
    } finally {
      setIsSaving(false);
    }
  };

  // ── File type handlers ────────────────────
  const handleAddFileType = async (e: React.FormEvent): Promise<void> => {
    if (e) e.preventDefault();
    const val = newFileType.trim();
    if (!val) return;

    if (config.fileTypes.map(f => f.toLowerCase()).includes(val.toLowerCase())) {
      setErrorText('File type already exists');
      return;
    }

    setIsSaving(true);
    setErrorText('');
    try {
      if (adminConfigService) {
        await adminConfigService.addFileType(val);
      }
      updateConfig({ 
        ...config, 
        fileTypes: [...config.fileTypes, val.toUpperCase()] 
      });
      setNewFileType('');
    } catch (err) {
      console.error('Add file type failed:', err);
      // Fallback
      updateConfig({ 
        ...config, 
        fileTypes: [...config.fileTypes, val.toUpperCase()] 
      });
      setNewFileType('');
    } finally {
      setIsSaving(false);
    }
  };

  const handleRemoveFileType = async (ft: string): Promise<void> => {
    if (ft === 'All') return;
    setIsSaving(true);
    try {
      if (adminConfigService) {
        await adminConfigService.removeFileType(ft);
      }
      updateConfig({
        ...config,
        fileTypes: config.fileTypes.filter((f) => f !== ft),
      });
    } catch (err) {
      console.error('Remove file type failed:', err);
      updateConfig({
        ...config,
        fileTypes: config.fileTypes.filter((f) => f !== ft),
      });
    } finally {
      setIsSaving(false);
    }
  };

  // ── Date filter handlers ──────────────────
  const handleAddDateFilter = async (e: React.FormEvent): Promise<void> => {
    if (e) e.preventDefault();
    const val = newDateFilter.trim();
    if (!val) return;

    if (config.dateFilters.map(d => d.toLowerCase()).includes(val.toLowerCase())) {
      setErrorText('Date filter already exists');
      return;
    }

    setIsSaving(true);
    setErrorText('');
    try {
      if (adminConfigService) {
        await adminConfigService.addDateFilter(val);
      }
      updateConfig({
        ...config,
        dateFilters: [...config.dateFilters, val],
      });
      setNewDateFilter('');
    } catch (err) {
      console.error('Add date filter failed:', err);
      // Fallback
      updateConfig({
        ...config,
        dateFilters: [...config.dateFilters, val],
      });
      setNewDateFilter('');
    } finally {
      setIsSaving(false);
    }
  };

  const handleRemoveDateFilter = async (df: string): Promise<void> => {
    setIsSaving(true);
    try {
      if (adminConfigService) {
        await adminConfigService.removeDateFilter(df);
      }
      updateConfig({
        ...config,
        dateFilters: config.dateFilters.filter((d) => d !== df),
      });
    } catch (err) {
      console.error('Remove date filter failed:', err);
      updateConfig({
        ...config,
        dateFilters: config.dateFilters.filter((d) => d !== df),
      });
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddSitePreset = async (siteName: string): Promise<void> => {
    if (config.sites.map(s => s.toLowerCase()).includes(siteName.toLowerCase())) return;
    setIsSaving(true);
    setErrorText('');
    try {
      const slug = siteName.replace(/\s+/g, '');
      if (adminConfigService) {
        await adminConfigService.addSite(siteName, slug);
      }
      updateConfig({ 
        ...config, 
        sites: [...config.sites, siteName] 
      });
    } catch (err) {
      console.error('Add site preset failed:', err);
      // Fallback
      updateConfig({ 
        ...config, 
        sites: [...config.sites, siteName] 
      });
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddFileTypePreset = async (ft: string): Promise<void> => {
    if (config.fileTypes.map(f => f.toLowerCase()).includes(ft.toLowerCase())) return;
    setIsSaving(true);
    setErrorText('');
    try {
      if (adminConfigService) {
        await adminConfigService.addFileType(ft);
      }
      updateConfig({ 
        ...config, 
        fileTypes: [...config.fileTypes, ft.toUpperCase()] 
      });
    } catch (err) {
      console.error('Add file type preset failed:', err);
      // Fallback
      updateConfig({ 
        ...config, 
        fileTypes: [...config.fileTypes, ft.toUpperCase()] 
      });
    } finally {
      setIsSaving(false);
    }
  };

  const handleAddDateFilterPreset = async (df: string): Promise<void> => {
    if (config.dateFilters.map(d => d.toLowerCase()).includes(df.toLowerCase())) return;
    setIsSaving(true);
    setErrorText('');
    try {
      if (adminConfigService) {
        await adminConfigService.addDateFilter(df);
      }
      updateConfig({
        ...config,
        dateFilters: [...config.dateFilters, df],
      });
    } catch (err) {
      console.error('Add date filter preset failed:', err);
      // Fallback
      updateConfig({
        ...config,
        dateFilters: [...config.dateFilters, df],
      });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className={styles.adminPanelOverlay}>
      {/* Backdrop */}
      <div className={styles.adminBackdrop} onClick={onClose} />

      {/* Side Drawer Panel */}
      <div className={styles.adminPanel}>
        {/* Header */}
        <div className={styles.adminPanelHeader}>
          <div className={styles.adminPanelHeaderLeft}>
            <div className={styles.adminPanelIconBox}>
              <Settings2 size={18} />
            </div>
            <div>
              <h2 className={styles.adminPanelTitle}>Admin Portal</h2>
              <p className={styles.adminPanelSubtitle}>Configure layouts & filters</p>
            </div>
          </div>
          {isSaving && (
            <span className={styles.adminSavingText}>
              Saving...
            </span>
          )}
          <button
            onClick={onClose}
            className={styles.adminPanelCloseBtn}
            title="Close Admin Portal"
          >
            <X size={18} strokeWidth={2.5} />
          </button>
        </div>

        {/* Error banner if any */}
        {errorText && (
          <div className={styles.adminErrorBanner}>
            <div className={styles.adminErrorDot} />
            {errorText}
          </div>
        )}

        {/* Content Body */}
        <div className={styles.adminPanelBody}>


          {/* 1. Configure Sidebar Sites */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <div className={styles.adminSectionHeader}>
              <Globe size={15} className={styles.adminSectionIcon} />
              <h3 className={styles.adminSectionTitle}>Configure Sidebar Sites</h3>
            </div>

            <form onSubmit={handleAddSite} className={styles.adminAddRow} style={{ position: 'relative' }}>
              <input
                type="text"
                placeholder="New site (e.g. communications)"
                value={newSite}
                onChange={(e) => {
                  setNewSite(e.target.value);
                  if (errorText) setErrorText('');
                }}
                onBlur={() => {
                  // Hide suggestions after brief delay to allow onClick/onMouseDown events to register first
                  setTimeout(() => setSiteSuggestions([]), 200);
                }}
                className={styles.adminAddInput}
              />
              <button type="submit" className={styles.adminAddBtn}>
                <Plus size={14} strokeWidth={2.5} />
                Add
              </button>

              {(() => {
                const filtered = siteSuggestions.filter(
                  s => !config.sites.some(cs => cs.toLowerCase() === s.title.toLowerCase())
                );
                if (filtered.length === 0) return null;
                return (
                  <div className={styles.adminSuggestionsDropdown}>
                    {filtered.map((s) => (
                      <div
                        key={s.slug}
                        className={styles.adminSuggestionItem}
                        onMouseDown={(e) => {
                          e.preventDefault(); // Prevents input focus loss
                          setNewSite(s.title);
                          setSiteSuggestions([]);
                        }}
                      >
                        <span className={styles.adminSuggestionTitle}>{s.title}</span>
                        <span className={styles.adminSuggestionSub}>/sites/{s.slug}</span>
                      </div>
                    ))}
                  </div>
                );
              })()}
            </form>

            <div className={styles.adminPresetsContainer}>
              <span className={styles.adminPresetsLabel}>Quick Presets (Click to Add)</span>
              {SITE_PRESETS.map((preset) => {
                const isAdded = config.sites.map(s => s.toLowerCase()).includes(preset.toLowerCase());
                return (
                  <span
                    key={preset}
                    onClick={() => !isAdded && handleAddSitePreset(preset)}
                    className={`${styles.adminPresetTag} ${styles.adminPresetTagSite} ${isAdded ? styles.adminPresetTagAdded : ''}`}
                  >
                    {preset} {!isAdded && '+'} {isAdded && '✓'}
                  </span>
                );
              })}
            </div>

            <div className={styles.adminConfigList}>
              <div className={styles.adminSystemDefault}>
                All (System Default)
              </div>
              {config.sites.map((site) => (
                <div key={site} className={styles.adminConfigRow}>
                  <span className={styles.adminConfigLabel}>{site}</span>
                  {site !== 'All' && (
                    <button
                      type="button"
                      onClick={() => confirmDelete('site', site)}
                      className={styles.adminDeleteBtn}
                      title={`Remove ${site}`}
                    >
                      <Trash2 size={13} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* 2. Configure File Types */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <div className={styles.adminSectionHeader}>
              <Folder size={15} className={styles.adminSectionIcon} />
              <h3 className={styles.adminSectionTitle}>Configure File Types</h3>
            </div>

            <form onSubmit={handleAddFileType} className={styles.adminAddRow}>
              <input
                type="text"
                placeholder="New file type (e.g. DOCX, Keynote)"
                value={newFileType}
                onChange={(e) => {
                  setNewFileType(e.target.value);
                  if (errorText) setErrorText('');
                }}
                className={styles.adminAddInput}
              />
              <button type="submit" className={styles.adminAddBtn}>
                <Plus size={14} strokeWidth={2.5} />
                Add
              </button>
            </form>

            <div className={styles.adminPresetsContainer}>
              <span className={styles.adminPresetsLabel}>Quick Presets (Click to Add)</span>
              {FILE_TYPE_PRESETS.map((preset) => {
                const isAdded = config.fileTypes.map(f => f.toLowerCase()).includes(preset.toLowerCase());
                return (
                  <span
                    key={preset}
                    onClick={() => !isAdded && handleAddFileTypePreset(preset)}
                    className={`${styles.adminPresetTag} ${styles.adminPresetTagFileType} ${isAdded ? styles.adminPresetTagAdded : ''}`}
                  >
                    {preset} {!isAdded && '+'} {isAdded && '✓'}
                  </span>
                );
              })}
            </div>

            <div className={styles.adminConfigList}>
              <div className={styles.adminSystemDefault}>
                All (System Default)
              </div>
              {config.fileTypes.map((ft) => (
                <div key={ft} className={styles.adminConfigRow}>
                  <span className={styles.adminConfigLabel}>{ft}</span>
                  {ft !== 'All' && (
                    <button
                      type="button"
                      onClick={() => confirmDelete('fileType', ft)}
                      className={styles.adminDeleteBtn}
                      title={`Remove ${ft}`}
                    >
                      <Trash2 size={13} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* 3. Configure Date Filters */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <div className={styles.adminSectionHeader}>
              <Calendar size={15} className={styles.adminSectionIcon} />
              <h3 className={styles.adminSectionTitle}>Configure Date Filters</h3>
            </div>

            <form onSubmit={handleAddDateFilter} className={styles.adminAddRow}>
              <input
                type="text"
                placeholder="New date group (e.g. Past 30 Days)"
                value={newDateFilter}
                onChange={(e) => {
                  setNewDateFilter(e.target.value);
                  if (errorText) setErrorText('');
                }}
                className={styles.adminAddInput}
              />
              <button type="submit" className={styles.adminAddBtn}>
                <Plus size={14} strokeWidth={2.5} />
                Add
              </button>
            </form>

            <div className={styles.adminPresetsContainer}>
              <span className={styles.adminPresetsLabel}>Quick Presets (Click to Add)</span>
              {DATE_PRESETS.map((preset) => {
                const isAdded = config.dateFilters.map(d => d.toLowerCase()).includes(preset.toLowerCase());
                return (
                  <span
                    key={preset}
                    onClick={() => !isAdded && handleAddDateFilterPreset(preset)}
                    className={`${styles.adminPresetTag} ${styles.adminPresetTagDate} ${isAdded ? styles.adminPresetTagAdded : ''}`}
                  >
                    {preset} {!isAdded && '+'} {isAdded && '✓'}
                  </span>
                );
              })}
            </div>

            <div className={styles.adminConfigList}>
              {config.dateFilters.map((df) => (
                <div key={df} className={styles.adminConfigRow}>
                  <span className={styles.adminConfigLabel}>{df}</span>
                  <button
                    type="button"
                    onClick={() => confirmDelete('dateFilter', df)}
                    className={styles.adminDeleteBtn}
                    title={`Remove ${df}`}
                  >
                    <Trash2 size={13} />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className={styles.adminPanelFooter}>
          <button
            type="button"
            className={styles.adminClosePanelBtn}
            onClick={onClose}
          >
            Close Panel
          </button>
        </div>
      </div>

      {/* Delete Confirmation Modal */}
      {deleteConfirmation && (
        <div className={styles.adminConfirmOverlay}>
          <div className={styles.adminConfirmBox}>
            <h3 className={styles.adminConfirmTitle}>Confirm Delete</h3>
            <p className={styles.adminConfirmMessage}>
              Are you sure you want to delete <strong>{deleteConfirmation.value}</strong>?
            </p>
            <div className={styles.adminConfirmActions}>
              <button
                type="button"
                className={styles.adminConfirmCancelBtn}
                onClick={() => setDeleteConfirmation(null)}
              >
                No
              </button>
              <button
                type="button"
                className={styles.adminConfirmConfirmBtn}
                onClick={executeDelete}
              >
                Yes
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
