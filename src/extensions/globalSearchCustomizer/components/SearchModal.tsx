import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { GraphSearchService } from '../services/GraphSearchService';
import { useSearch } from '../hooks/useSearch';
import { ISearchResult } from '../../../models/ISearchResult';
import { useDebounce } from 'use-debounce';


// Import modular subcomponents
import { Sidebar } from './Sidebar';
import { SearchTabs } from './SearchTabs';
import { PreviewPane } from './PreviewPane';
import { HistoryPane } from './HistoryPane';
import { Header } from './Header';
import { SearchResultsList } from './SearchResultsList';

import { ISearchModalProps } from '../interface/ISearchModalProps';

// File type design mapping helper
const getFileColor = (fileType: string) => {
  const ft = fileType ? fileType.toLowerCase() : '';
  if (ft === 'pdf') return { color: '#00acc1', bg: '#e0f7fa' };
  if (['xls', 'xlsx'].includes(ft)) return { color: '#00796b', bg: '#e0f2f1' };
  if (['ppt', 'pptx'].includes(ft)) return { color: '#e52592', bg: '#fce4ec' };
  if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ft)) return { color: '#9334e6', bg: '#f3e5f5' };
  if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) return { color: '#ea4335', bg: '#fce8e6' };
  if (ft === 'folder') return { color: '#f5b041', bg: '#fef5e7' };
  return { color: '#1a73e8', bg: '#e8f0fe' };
};

export default function SearchModal({ context, isOpen, onDismiss }: ISearchModalProps): React.ReactElement | null {
  if (!isOpen) return null;

  // --- Unified Column Widths & Resizing States ---
  const [sidebarWidth, setSidebarWidth] = useState(240);
  const [isResizingSidebar, setIsResizingSidebar] = useState(false);

  const [previewWidth, setPreviewWidth] = useState(380);
  const [isResizingPreview, setIsResizingPreview] = useState(false);

  // --- Search Query & Filter states ---
  const [searchQuery, setSearchQuery] = useState('');
  const [isSearchFocused, setIsSearchFocused] = useState(false);
  const [filters, setFilters] = useState({
    fileTypes: ['All'],
    selectedAuthors: [] as string[],
    date: ''
  });

  const [activeTopTab, setActiveTopTab] = useState('All');
  const [bottomTab, setBottomTab] = useState<'Search Results' | 'Recent Activities' | 'Starred Assets'>('Search Results');
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);

  // --- Inline state replacing zustand store ---
  const [recentlyViewedFiles, setRecentlyViewedFiles] = useState<ISearchResult[]>(() => {
    try {
      const saved = sessionStorage.getItem('trivandi_recent_files_data');
      return saved ? JSON.parse(saved) : [];
    } catch (e) { return []; }
  });
  const [starredFiles, setStarredFiles] = useState<ISearchResult[]>(() => {
    try {
      const saved = localStorage.getItem('trivandi_starred_files_data');
      return saved ? JSON.parse(saved) : [];
    } catch (e) { return []; }
  });
  const [starredIds, setStarredIds] = useState<Set<string>>(() => {
    try {
      const saved = localStorage.getItem('trivandi_starred_files_data');
      const arr: ISearchResult[] = saved ? JSON.parse(saved) : [];
      return new Set(arr.map(f => f.id));
    } catch (e) { return new Set(); }
  });

  const addRecentFile = (file: ISearchResult): void => {
    setRecentlyViewedFiles(prev => {
      const filtered = prev.filter(f => f.id !== file.id);
      const updated = [file, ...filtered];
      try { sessionStorage.setItem('trivandi_recent_files_data', JSON.stringify(updated)); } catch (e) { /* ignored */ }
      return updated;
    });
  };

  const toggleStar = (file: ISearchResult): void => {
    setStarredIds(prev => {
      const hasStar = prev.has(file.id);
      const next = new Set(prev);
      if (hasStar) { next.delete(file.id); } else { next.add(file.id); }
      return next;
    });
    setStarredFiles(prev => {
      const hasStar = prev.some(f => f.id === file.id);
      const next = hasStar ? prev.filter(f => f.id !== file.id) : [...prev, { ...file, isStarred: true }];
      try { localStorage.setItem('trivandi_starred_files_data', JSON.stringify(next)); } catch (e) { /* ignored */ }
      return next;
    });
    setRecentlyViewedFiles(prev =>
      prev.map(f => f.id === file.id ? { ...f, isStarred: !starredIds.has(file.id) } : f)
    );
  };

  // Top-level modal states for loose-coupling Copy link sharing dialog
  const [copyFileDialogFile, setCopyFileDialogFile] = useState<ISearchResult | null>(null);
  const [appCopyCopied, setAppCopyCopied] = useState(false);

  // Simulated Search History
  const [searchHistory, setSearchHistory] = useState<string[]>([
    'Quarterly Strategy plans',
    'Financial sheets FY24',
    'Brand design guide PDF',
    'HR employee handbook v3'
  ]);


  // --- Graph Service Setup ---
  const searchService = useMemo(() => {
    if (context && context.msGraphClientFactory) {
      return new GraphSearchService(context.msGraphClientFactory);
    }
    return null;
  }, [context]);

  // Handle drag mousemove event
  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (isResizingSidebar) {
        const newWidth = Math.max(180, Math.min(400, e.clientX));
        setSidebarWidth(newWidth);
      }
      if (isResizingPreview) {
        const newWidth = Math.max(280, Math.min(600, window.innerWidth - e.clientX));
        setPreviewWidth(newWidth);
      }
    };

    const handleMouseUp = () => {
      setIsResizingSidebar(false);
      setIsResizingPreview(false);
    };

    if (isResizingSidebar || isResizingPreview) {
      document.addEventListener('mousemove', handleMouseMove);
      document.addEventListener('mouseup', handleMouseUp);
    }

    return () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
  }, [isResizingSidebar, isResizingPreview]);

  // --- Hook Execution for Live Data ---
  const {
    results: liveResults,
    totalCount: liveTotalCount,
    loading: liveLoading,
    setQuery,
    setFileTypes,
    setActiveTopTab: hookSetActiveTopTab,
    setDate,
    from,
    setFrom
  } = useSearch({
    service: searchService,
    initialQuery: '',
    pageSize: 10
  });

  // Debounce search query changes using standard 'use-debounce' React library
  const [debouncedSearchQuery] = useDebounce(searchQuery, 450);

  // Track search query changes to synchronize hook with debounce
  useEffect(() => {
    setQuery(debouncedSearchQuery);
  }, [debouncedSearchQuery, setQuery]);

  // Track fileTypes filter changes to synchronize hook
  useEffect(() => {
    setFileTypes(filters.fileTypes);
  }, [filters.fileTypes, setFileTypes]);

  // Track activeTopTab changes to synchronize hook
  useEffect(() => {
    hookSetActiveTopTab(activeTopTab);
  }, [activeTopTab, hookSetActiveTopTab]);

  // Track date filter changes to synchronize hook
  useEffect(() => {
    setDate(filters.date);
  }, [filters.date, setDate]);


  // Convert raw Graph Search results into our formatted UI result cards
  const mappedLiveResults = useMemo(() => {
    return liveResults.map(res => {
      const isStarred = starredIds.has(res.id);
      return {
        id: res.id,
        title: res.title,
        author: res.author || 'SharePoint User',
        lastModified: res.lastModified || new Date().toISOString(),
        size: res.size || 1048576,
        summary: res.summary || 'No description preview available.',
        webUrl: res.webUrl,
        fileType: res.fileType || 'docx',
        siteName: res.siteName || 'SharePoint Site',
        siteUrl: res.siteUrl || '',
        isStarred: isStarred
      } as ISearchResult;
    });
  }, [liveResults, starredIds]);

  // Track selectedFileId change to push clicked files into recentlyViewedFiles via store action
  useEffect(() => {
    if (selectedFileId) {
      const fileObj = mappedLiveResults.find(f => f.id === selectedFileId);
      if (fileObj) {
        addRecentFile(fileObj);
      }
    }
  }, [selectedFileId, mappedLiveResults, addRecentFile]);

  // Listen to the custom copy-link event from any card's Action Menu to display the top-level overlay
  useEffect(() => {
    const handleCopyEvent = (e: CustomEvent) => {
      setCopyFileDialogFile(e.detail);
    };
    window.addEventListener('trivandi-copy-link', handleCopyEvent as EventListener);
    return () => {
      window.removeEventListener('trivandi-copy-link', handleCopyEvent as EventListener);
    };
  }, []);

  // Dynamic unique list of authors alphabetically sorted
  const authorsList = useMemo(() => {
    const collected = new Set<string>();
    // Default high-profile premium team members
    ['Akhilesh Gandhi', 'Paul Kerby', 'Carla Glossip', 'Ava Pevsner', 'Sam Lay'].forEach(a => collected.add(a));
    // Live loaded search authors
    mappedLiveResults.forEach(r => {
      if (r.author && r.author.trim() !== 'SharePoint User' && r.author.trim() !== 'SharePoint Portal') {
        collected.add(r.author.trim());
      }
    });
    return Array.from(collected).sort((a, b) => a.localeCompare(b));
  }, [mappedLiveResults]);

  // Robust live results filtering by Selected Authors
  const filteredLiveResults = useMemo(() => {
    let list = mappedLiveResults;
    
    // 1. Selected Authors Filter (Multi-select)
    if (filters.selectedAuthors.length > 0) {
      list = list.filter(item => filters.selectedAuthors.includes(item.author));
    }
    
    return list;
  }, [mappedLiveResults, filters.selectedAuthors]);

  // Set the final target results and states based on current Mode (Live vs Mock)
  const resultsToRender = useMemo(() => {
    if (bottomTab === 'Recent Activities') {
      return recentlyViewedFiles.map(f => ({
        ...f,
        isStarred: starredIds.has(f.id)
      }));
    }
    if (bottomTab === 'Starred Assets') {
      return starredFiles.map(f => ({
        ...f,
        isStarred: true
      }));
    }
    return filteredLiveResults;
  }, [bottomTab, recentlyViewedFiles, starredFiles, filteredLiveResults, starredIds]);

  const totalCountToRender: number = bottomTab === 'Search Results' 
    ? (filters.selectedAuthors.length > 0 ? filteredLiveResults.length : liveTotalCount)
    : resultsToRender.length;
  const isLoading: boolean = liveLoading;

  // Selected File Object derivation
  const selectedFile = useMemo(() => {
    if (!selectedFileId) return null;
    const file = resultsToRender.find(f => f.id === selectedFileId);
    if (!file) return null;
    const colors = getFileColor(file.fileType);
    
    // Format sizes cleanly
    const sizeInMB = file.size ? `${(file.size / (1024 * 1024)).toFixed(1)} MB` : '1.2 MB';
    
    return {
      id: file.id,
      title: file.title,
      author: file.author,
      date: file.lastModified ? new Date(file.lastModified).toLocaleDateString() : 'Today',
      size: sizeInMB,
      description: file.summary || 'No description extracted.',
      url: file.webUrl,
      type: file.fileType.toUpperCase(),
      starred: !!file.isStarred,
      color: colors.color,
      badgeColor: colors.color,
      summary: file.summary,
      siteName: file.siteName,
      siteUrl: file.siteUrl,
      webUrl: file.webUrl,
      lastModified: file.lastModified
    };
  }, [selectedFileId, resultsToRender]);

  // --- Drag Column trigger handlers ---
  const startResizingSidebar = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizingSidebar(true);
  };

  const startResizingPreview = (e: React.MouseEvent) => {
    e.preventDefault();
    setIsResizingPreview(true);
  };

  // --- File Starring Handler ---
  const handleToggleStar = (e: React.MouseEvent, id: string): void => {
    e.stopPropagation();
    const fileObj = mappedLiveResults.find(f => f.id === id) || recentlyViewedFiles.find(f => f.id === id) || starredFiles.find(f => f.id === id);
    if (fileObj) {
      toggleStar(fileObj);
    }
  };

  // --- Sidebar Type Filter checkbox handler ---
  const toggleFileType = (type: string): void => {
    setFilters(prev => {
      if (type === 'All') {
        return { ...prev, fileTypes: ['All'] };
      }
      let updated = prev.fileTypes.filter(t => t !== 'All');
      if (updated.includes(type)) {
        updated = updated.filter(t => t !== type);
      } else {
        updated.push(type);
      }
      if (updated.length === 0) {
        updated = ['All'];
      }
      return { ...prev, fileTypes: updated };
    });
  };

  // --- Remove single search history item ---
  const removeHistoryItem = (item: string): void => {
    setSearchHistory(prev => prev.filter(h => h !== item));
  };

  // --- Execute search instantly ---
  const handleSearch = (query: string): void => {
    setSearchQuery(query);
    setQuery(query); // Trigger instantly!
    setIsHistoryOpen(false);
  };

  // --- Clear all applied filters ---
  const handleClearAllFilters = (): void => {
    setFilters({
      fileTypes: ['All'],
      selectedAuthors: [],
      date: ''
    });
    setActiveTopTab('All');
    setFrom(0);
  };

  const [openMenuId, setOpenMenuId] = useState<string | null>(null);
  const _dummyIgnored = openMenuId; // Prevent lint warning for unused state

  return (
    <div className={styles.overlay}>
      <div className={styles.modal}>
        
        {/* ================= HEADER BAR ================= */}
        <Header 
          searchQuery={searchQuery}
          setSearchQuery={setSearchQuery}
          handleSearch={handleSearch}
          isSearchFocused={isSearchFocused}
          setIsSearchFocused={setIsSearchFocused}
          searchHistory={searchHistory}
          totalCountToRender={totalCountToRender}
          onDismiss={onDismiss}
        />

        {/* ================= BODY WRAPPER ================= */}
        <div className={styles.body}>
          
          <div className={styles.mainLayout}>
            
            {/* Sidebar Filters */}
            {bottomTab === 'Search Results' && (
              <Sidebar 
                sidebarWidth={sidebarWidth}
                fileTypes={filters.fileTypes}
                selectedAuthors={filters.selectedAuthors}
                date={filters.date}
                setFilters={setFilters}
                toggleFileType={toggleFileType}
                startResizingSidebar={startResizingSidebar}
                isResizingSidebar={isResizingSidebar}
                authorsList={authorsList}
              />
            )}

            {/* Main Center Panel (Tabs & Listing Pane) */}
            <div className={styles.centerContainer}>
              {/* Category Filter Tabs */}
              {bottomTab === 'Search Results' && (
                <SearchTabs 
                  activeTab={activeTopTab}
                  onTabChange={(tab) => {
                    setActiveTopTab(tab);
                    setFrom(0);
                  }}
                  onClearAll={handleClearAllFilters}
                />
              )}

              {/* Main Results Listing Pane */}
              <SearchResultsList 
                resultsToRender={resultsToRender}
                totalCountToRender={totalCountToRender}
                isLoading={isLoading}
                selectedFileId={selectedFileId}
                setSelectedFileId={setSelectedFileId}
                starredIds={starredIds}
                toggleStar={handleToggleStar}
                from={from}
                setFrom={setFrom}
                bottomTab={bottomTab}
                setBottomTab={setBottomTab}
                setIsHistoryOpen={setIsHistoryOpen}
                isHistoryOpen={isHistoryOpen}
                handleClearAllFilters={handleClearAllFilters}
                activeTopTab={activeTopTab}
                setActiveTopTab={setActiveTopTab}
                filters={filters}
                setFilters={setFilters}
                setOpenMenuId={setOpenMenuId}
                openMenuId={openMenuId}
              />
            </div>

            {/* Document preview analysis panel */}
            <PreviewPane 
              selectedFile={selectedFile}
              setSelectedFile={(file) => setSelectedFileId(file ? file.id : null)}
              previewWidth={previewWidth}
              searchQuery={searchQuery}
              isResizingPreview={isResizingPreview}
              startResizingPreview={startResizingPreview}
            />

            {/* Recent Searches history panel */}
            <HistoryPane 
              isHistoryOpen={isHistoryOpen}
              setIsHistoryOpen={setIsHistoryOpen}
              previewWidth={previewWidth}
              searchHistory={searchHistory}
              handleSearch={handleSearch}
              setSearchHistory={setSearchHistory}
              removeHistoryItem={removeHistoryItem}
              isResizingPreview={isResizingPreview}
              startResizingPreview={startResizingPreview}
            />

          </div>

        </div>
      </div>

      {/* ========================================================
           MODAL 2: Copy link SharePoint Share dialog (Top-level viewport)
         ======================================================== */}
      {copyFileDialogFile && (
        <div 
          className={styles.copyDialogOverlay}
          onClick={(e) => {
            e.stopPropagation();
            setCopyFileDialogFile(null);
          }}
        >
          <div 
            className={styles.copyDialogContent}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Close button top right */}
            <button 
              onClick={() => setCopyFileDialogFile(null)}
              className={styles.copyDialogClose}
            >
              ✕
            </button>
            
            <div className={styles.copyDialogHeader}>
              <div className={styles.copyDialogCheckCircle}>
                ✓
              </div>
              <h3 className={styles.copyDialogTitle}>
                Link created
              </h3>
            </div>

            <div className={styles.copyDialogInputRow}>
              <span className={styles.copyDialogUrlText}>
                {copyFileDialogFile.webUrl || `https://sharepoint.trivandi.com/Shared%20Documents/${copyFileDialogFile.title}`}
              </span>
              <button 
                onClick={async () => {
                  try {
                    const url = copyFileDialogFile.webUrl || `https://sharepoint.trivandi.com/Shared%20Documents/${copyFileDialogFile.title}`;
                    await navigator.clipboard.writeText(url);
                    setAppCopyCopied(true);
                    setTimeout(() => {
                      setAppCopyCopied(false);
                      setCopyFileDialogFile(null);
                    }, 1500);
                  } catch (err) {
                    console.error('Failed to copy', err);
                  }
                }}
                className={`${styles.copyDialogCopyButton} ${appCopyCopied ? styles.copyDialogCopyButtonCopied : ''}`}
              >
                {appCopyCopied ? 'Copied!' : 'Copy'}
              </button>
            </div>

            <div className={styles.copyDialogFooter}>
              <span>People in your organization with the link can view</span>
              <span className={styles.copyDialogSettingsLink}>
                ⚙ Settings
              </span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
