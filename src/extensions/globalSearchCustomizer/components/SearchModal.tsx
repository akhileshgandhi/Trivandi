import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { GraphSearchService } from '../services/GraphSearchService';
import { useSearch } from '../hooks/useSearch';
import { ISearchResult } from '../../../models/ISearchResult';

// Import modular subcomponents
import { Sidebar } from './Sidebar';
import { SearchTabs } from './SearchTabs';
import { PreviewPane } from './PreviewPane';
import { HistoryPane } from './HistoryPane';
import { Header } from './Header';
import { SearchResultsList } from './SearchResultsList';

export interface ISearchModalProps {
  context: any;
  isOpen: boolean;
  onDismiss: () => void;
}

// File type design mapping helper
const getFileColor = (fileType: string) => {
  const ft = fileType ? fileType.toLowerCase() : '';
  if (ft === 'pdf') return { color: '#00acc1', bg: '#e0f7fa' };
  if (['xls', 'xlsx'].includes(ft)) return { color: '#00796b', bg: '#e0f2f1' };
  if (['ppt', 'pptx'].includes(ft)) return { color: '#e52592', bg: '#fce4ec' };
  if (['png', 'jpg', 'jpeg'].includes(ft)) return { color: '#9334e6', bg: '#f3e5f5' };
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
    author: '',
    date: ''
  });

  const [activeTopTab, setActiveTopTab] = useState('All');
  const [bottomTab, setBottomTab] = useState<'Search Results' | 'Recent Activities' | 'Starred Assets'>('Search Results');
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const [starredIds, setStarredIds] = useState<Set<string>>(new Set(['m1', 'm4']));
  const [isAuthorDropdownOpen, setIsAuthorDropdownOpen] = useState(false);

  // --- Recently Viewed Files state ---
  const [recentlyViewedFiles, setRecentlyViewedFiles] = useState<ISearchResult[]>(() => {
    try {
      const saved = sessionStorage.getItem('trivandi_recent_files_data');
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

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

  // --- Search Results Mock Data (Fallback) ---
  const initialMockResults: ISearchResult[] = useMemo(() => [
    {
      id: 'm1',
      title: 'Marketing Strategy Q3 2024',
      author: 'Sarah Chen',
      lastModified: '2026-05-24T12:00:00Z',
      size: 1258291,
      summary: 'Comprehensive analysis of market trends and competitor performance in the APAC region.',
      webUrl: 'https://sharepoint.intel/marketing/strategies',
      fileType: 'docx',
      siteName: 'Marketing Hub',
      siteUrl: 'https://sharepoint.intel/marketing',
      isStarred: true
    },
    {
      id: 'm2',
      title: 'Annual Financial Report FY23',
      author: 'Robert Wilson',
      lastModified: '2026-05-21T09:15:00Z',
      size: 5033164,
      summary: 'Consolidated financial statements, auditor reports, and performance metrics for the fiscal year.',
      webUrl: 'https://sharepoint.intel/finance/reports',
      fileType: 'xlsx',
      siteName: 'Finance Portal',
      siteUrl: 'https://sharepoint.intel/finance',
      isStarred: false
    },
    {
      id: 'm3',
      title: 'Brand Guidelines V2.1',
      author: 'Elena Rodriguez',
      lastModified: '2026-05-16T16:45:00Z',
      size: 5033164,
      summary: 'Visual identity standards, logo usage, and typography rules for all corporate communications.',
      webUrl: 'https://sharepoint.intel/brand/assets',
      fileType: 'pdf',
      siteName: 'Brand Center',
      siteUrl: 'https://sharepoint.intel/brand',
      isStarred: false
    },
    {
      id: 'm4',
      title: 'Employee Onboarding Handbook',
      author: 'HR Department',
      lastModified: '2026-05-25T14:20:00Z',
      size: 1258291,
      summary: 'Essential information for new hires including benefits, culture, and operational procedures.',
      webUrl: 'https://sharepoint.intel/hr/portal',
      fileType: 'docx',
      siteName: 'HR Hub',
      siteUrl: 'https://sharepoint.intel/hr',
      isStarred: true
    },
    {
      id: 'm5',
      title: 'Product Roadmap 2024-2025',
      author: 'James T. Kirk',
      lastModified: '2026-05-22T10:00:00Z',
      size: 5033164,
      summary: 'Strategic product milestones, feature releases, and timeline planning for next-generation platform.',
      webUrl: 'https://sharepoint.intel/product/roadmap',
      fileType: 'pptx',
      siteName: 'Product Portal',
      siteUrl: 'https://sharepoint.intel/product',
      isStarred: false
    },
    {
      id: 'm6',
      title: 'IT Security Policy & Guidelines',
      author: 'Security Operations',
      lastModified: '2026-05-18T08:00:00Z',
      size: 2202009,
      summary: 'Information security standards, password requirements, and compliance guidelines for employees.',
      webUrl: 'https://sharepoint.intel/it/security',
      fileType: 'pdf',
      siteName: 'IT Operations',
      siteUrl: 'https://sharepoint.intel/it',
      isStarred: false
    },
    {
      id: 'm7',
      title: 'Trivandi Corporate Brand Video.mp4',
      author: 'Media Team',
      lastModified: '2026-05-24T10:00:00Z',
      size: 45097152,
      summary: 'Brand intro video, corporate milestones, and team introduction for marketing campaigns.',
      webUrl: 'https://sharepoint.intel/media/brandvideo',
      fileType: 'mp4',
      siteName: 'Media Portal',
      siteUrl: 'https://sharepoint.intel/media',
      isStarred: false
    },
    {
      id: 'm8',
      title: 'Global Townhall Meeting May 2026.mov',
      author: 'Internal Communications',
      lastModified: '2026-05-26T09:00:00Z',
      size: 209715200,
      summary: 'Full recording of the May 2026 company hub Townhall meeting with CEO strategy review.',
      webUrl: 'https://sharepoint.intel/media/townhall',
      fileType: 'mov',
      siteName: 'Media Portal',
      siteUrl: 'https://sharepoint.intel/media',
      isStarred: false
    }
  ], []);

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
    error: liveError,
    setQuery,
    setFileTypes,
    from,
    setFrom
  } = useSearch({
    service: searchService,
    initialQuery: '',
    pageSize: 20
  });

  // Track search query changes to synchronize hook
  useEffect(() => {
    setQuery(searchQuery);
  }, [searchQuery, setQuery]);

  // Track fileTypes filter changes to synchronize hook
  useEffect(() => {
    setFileTypes(filters.fileTypes);
  }, [filters.fileTypes, setFileTypes]);

  // Determine if we should use Live Mode or Fallback Mock Mode
  const isLiveMode = !!searchService && !liveError;

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

  // Track selectedFileId change to push clicked files into recentlyViewedFiles
  useEffect(() => {
    if (selectedFileId) {
      const allPossibleFiles = isLiveMode ? mappedLiveResults : initialMockResults;
      const fileObj = allPossibleFiles.find(f => f.id === selectedFileId);
      
      if (fileObj) {
        setRecentlyViewedFiles(prev => {
          const filtered = prev.filter(f => f.id !== selectedFileId);
          const updated = [{ ...fileObj, isStarred: starredIds.has(fileObj.id) }, ...filtered];
          try {
            sessionStorage.setItem('trivandi_recent_files_data', JSON.stringify(updated));
          } catch (e) {
            // ignored
          }
          return updated;
        });
      }
    }
  }, [selectedFileId, isLiveMode, mappedLiveResults, initialMockResults, starredIds]);

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

  const filteredLiveResults = useMemo(() => {
    let list = mappedLiveResults;
    if (activeTopTab === 'Files') {
      list = list.filter(item => ['doc', 'docx', 'xls', 'xlsx', 'pdf', 'ppt', 'pptx'].includes(item.fileType.toLowerCase()));
    } else if (activeTopTab === 'Images') {
      list = list.filter(item => ['png', 'jpg', 'jpeg', 'gif', 'svg', 'tiff'].includes(item.fileType.toLowerCase()));
    } else if (activeTopTab === 'Videos') {
      list = list.filter(item => ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(item.fileType.toLowerCase()));
    }
    return list;
  }, [mappedLiveResults, activeTopTab]);

  // --- Filtering Mock Data (Offline Fallback Logic) ---
  const filteredMockResults = useMemo(() => {
    let list = initialMockResults.map(item => ({
      ...item,
      isStarred: starredIds.has(item.id)
    }));

    if (bottomTab === 'Starred Assets') {
      list = list.filter(item => item.isStarred);
    }

    if (activeTopTab === 'Files') {
      list = list.filter(item => ['doc', 'docx', 'xls', 'xlsx', 'pdf', 'ppt', 'pptx'].includes(item.fileType.toLowerCase()));
    } else if (activeTopTab === 'Images') {
      list = list.filter(item => ['png', 'jpg', 'jpeg', 'gif', 'svg', 'tiff'].includes(item.fileType.toLowerCase())); 
    } else if (activeTopTab === 'Videos') {
      list = list.filter(item => ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(item.fileType.toLowerCase()));
    }

    if (searchQuery.trim() !== '') {
      const q = searchQuery.toLowerCase();
      list = list.filter(item => 
        item.title.toLowerCase().includes(q) || 
        (item.summary && item.summary.toLowerCase().includes(q)) ||
        item.author.toLowerCase().includes(q)
      );
    }

    if (filters.fileTypes.length > 0 && !filters.fileTypes.includes('All')) {
      list = list.filter(item => {
        const typeMatch = filters.fileTypes.map(t => t.toLowerCase());
        return typeMatch.some(tm => item.fileType.toLowerCase().includes(tm));
      });
    }

    if (filters.author.trim() !== '') {
      const auth = filters.author.toLowerCase();
      list = list.filter(item => item.author.toLowerCase().includes(auth));
    }

    return list;
  }, [searchQuery, filters, activeTopTab, bottomTab, starredIds, initialMockResults]);

  // Set the final target results and states based on current Mode (Live vs Mock)
  const resultsToRender = useMemo(() => {
    if (bottomTab === 'Recent Activities') {
      return recentlyViewedFiles.map(f => ({
        ...f,
        isStarred: starredIds.has(f.id)
      }));
    }
    return isLiveMode ? filteredLiveResults : filteredMockResults;
  }, [bottomTab, recentlyViewedFiles, isLiveMode, filteredLiveResults, filteredMockResults, starredIds]);

  const totalCountToRender: number = resultsToRender.length;
  const isLoading: boolean = isLiveMode ? liveLoading : false;

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
  const toggleStar = (e: React.MouseEvent, id: string): void => {
    e.stopPropagation();
    setStarredIds(prev => {
      const updated = new Set(prev);
      if (updated.has(id)) {
        updated.delete(id);
      } else {
        updated.add(id);
      }
      return updated;
    });
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

  // --- Execute search ---
  const handleSearch = (query: string): void => {
    setSearchQuery(query);
    setIsHistoryOpen(false);
  };

  // --- Clear all applied filters ---
  const handleClearAllFilters = (): void => {
    setFilters({
      fileTypes: ['All'],
      author: '',
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
            <Sidebar 
              sidebarWidth={sidebarWidth}
              fileTypes={filters.fileTypes}
              author={filters.author}
              date={filters.date}
              setFilters={setFilters}
              toggleFileType={toggleFileType}
              isAuthorDropdownOpen={isAuthorDropdownOpen}
              setIsAuthorDropdownOpen={setIsAuthorDropdownOpen}
              startResizingSidebar={startResizingSidebar}
              isResizingSidebar={isResizingSidebar}
            />

            {/* Main Center Panel (Tabs & Listing Pane) */}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              {/* Category Filter Tabs */}
              <SearchTabs 
                activeTab={activeTopTab}
                onTabChange={(tab) => {
                  setActiveTopTab(tab);
                  setFrom(0);
                }}
                onClearAll={handleClearAllFilters}
              />

              {/* Main Results Listing Pane */}
              <SearchResultsList 
                resultsToRender={resultsToRender}
                totalCountToRender={totalCountToRender}
                isLoading={isLoading}
                selectedFileId={selectedFileId}
                setSelectedFileId={setSelectedFileId}
                starredIds={starredIds}
                toggleStar={toggleStar}
                from={from}
                setFrom={setFrom}
                isLiveMode={isLiveMode}
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
          style={{ 
            position: 'fixed', 
            inset: 0, 
            backgroundColor: 'rgba(15, 23, 42, 0.4)', 
            backdropFilter: 'blur(4px)',
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center', 
            zIndex: 999999 
          }}
          onClick={(e) => {
            e.stopPropagation();
            setCopyFileDialogFile(null);
          }}
        >
          <div 
            style={{ 
              width: '450px', 
              background: '#ffffff', 
              borderRadius: '16px', 
              boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
              padding: '24px',
              boxSizing: 'border-box',
              position: 'relative'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Close button top right */}
            <button 
              onClick={() => setCopyFileDialogFile(null)}
              style={{ 
                position: 'absolute', 
                top: '16px', 
                right: '16px', 
                background: 'none', 
                border: 'none', 
                color: '#94a3b8', 
                cursor: 'pointer',
                fontSize: '18px'
              }}
            >
              ✕
            </button>
            
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', width: '20px', height: '20px', borderRadius: '50%', backgroundColor: '#4caf50', color: '#ffffff', fontSize: '12px', fontWeight: 'bold' }}>
                ✓
              </div>
              <h3 style={{ fontSize: '15px', fontWeight: 700, color: '#202124', margin: 0, fontFamily: 'Segoe UI, sans-serif' }}>
                Link created
              </h3>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', border: '1px solid #dadce0', borderRadius: '8px', padding: '6px 12px', background: '#f8f9fa', marginBottom: '12px' }}>
              <span style={{ fontSize: '13px', color: '#3c4043', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1, fontFamily: 'Segoe UI, sans-serif' }}>
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
                style={{
                  padding: '6px 16px',
                  borderRadius: '20px',
                  border: '1px solid #dadce0',
                  background: appCopyCopied ? '#e8f5e9' : '#ffffff',
                  color: appCopyCopied ? '#2e7d32' : '#1a73e8',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  transition: 'all 0.15s ease',
                  fontFamily: 'Segoe UI, sans-serif'
                }}
              >
                {appCopyCopied ? 'Copied!' : 'Copy'}
              </button>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: '12px', color: '#5f6368', fontFamily: 'Segoe UI, sans-serif' }}>
              <span>People in your organization with the link can view</span>
              <span style={{ display: 'flex', alignItems: 'center', gap: '4px', cursor: 'pointer', color: '#1a73e8', fontWeight: 600 }}>
                ⚙ Settings
              </span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
