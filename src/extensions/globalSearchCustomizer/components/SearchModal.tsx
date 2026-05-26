import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { 
  Search, 
  X, 
  Clock, 
  Star, 
  User, 
  Calendar, 
  MoreVertical, 
  ChevronLeft, 
  ChevronRight, 
  FileText, 
  Download, 
  ExternalLink,
  SlidersHorizontal,
  Wifi,
  WifiOff,
  History as HistoryIcon,
  Image as ImageIcon,
  FileSpreadsheet
} from 'lucide-react';
import { GraphSearchService } from '../services/GraphSearchService';
import { useSearch } from '../hooks/useSearch';
import { ISearchResult } from '../../../models/ISearchResult';

// Import subcomponents
import { FiltersPanel } from './FiltersPanel';
import { DocumentAnalysis } from './DocumentAnalysis';
import { SearchHistory } from './SearchHistory';
import { ResultCard, SearchTabs } from './ResultCard';
import { FileActionMenu } from './FileActionMenu';
import { PaginationComponent } from '../Common/PaginationComponent';
import { SkeletonLoader } from '../Common/SkeletonLoader';

import { ISearchModalProps } from '../interface/ISearchModalProps';

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
    pageSize: 10
  });

  // Track search query changes to synchronize hook with debounce
  useEffect(() => {
    const timeoutId = setTimeout(() => {
      setQuery(searchQuery);
    }, 500);
    return () => clearTimeout(timeoutId);
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
      list = list.filter(item => ['docx', 'xlsx', 'pdf', 'pptx'].includes(item.fileType.toLowerCase()));
    } else if (activeTopTab === 'Images') {
      list = list.filter(item => ['png', 'jpg', 'jpeg', 'pdf'].includes(item.fileType.toLowerCase())); 
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
  const resultsToRender: ISearchResult[] = isLiveMode ? mappedLiveResults : filteredMockResults;
  const totalCountToRender: number = isLiveMode ? liveTotalCount : filteredMockResults.length;
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

  const [openMenuId, setOpenMenuId] = useState<string | null>(null);

  return (
    <div className={styles.overlay}>
      <div className={styles.modal}>
        
        {/* ================= HEADER BAR ================= */}
        <header className={styles.header}>
          <div className={styles.logoWrapper}>
            <svg width="120" height="32" viewBox="0 0 120 32" fill="none" xmlns="http://www.w3.org/2000/svg" className={styles.logoSvg}>
              <text x="5" y="24" className={styles.logoText}>
                trivandi
              </text>
              <circle cx="43.5" cy="8" r="2.5" fill="#F22797" />
              <circle cx="114.5" cy="8" r="2.5" fill="#F22797" />
              <path d="M6 10.5C9 8.5 14 8.5 17 10.5" stroke="#F22797" strokeWidth="2.5" strokeLinecap="round" />
            </svg>
          </div>

          <div className={styles.searchBarContainer}>
            <div className={styles.searchBarWrapper}>
              <span className={styles.searchIcon}>
                <Search size={18} strokeWidth={2.5} />
              </span>
              <input 
                type="text"
                placeholder="Search documents, pages, files..."
                value={searchQuery}
                onFocus={() => setIsSearchFocused(true)}
                onBlur={() => setTimeout(() => setIsSearchFocused(false), 200)}
                onChange={(e) => setSearchQuery(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleSearch(searchQuery)}
                className={styles.searchInput}
              />
              
              {isSearchFocused && searchHistory.length > 0 && (
                <div className={styles.searchSuggestionBox}>
                  <div className={styles.suggestionHeader}>
                    <span>Recent Searches</span>
                  </div>
                  <div>
                    {searchHistory.map((historyItem, idx) => (
                      <button
                        key={idx}
                        onClick={() => handleSearch(historyItem)}
                        className={styles.suggestionItem}
                      >
                        <Search size={14} className={styles.suggestionIcon} />
                        {historyItem}
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Connection status badge */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '4px 12px', borderRadius: '20px', background: isLiveMode ? '#ECFDF5' : '#FFF7ED', border: `1px solid ${isLiveMode ? '#10B981' : '#F97316'}`, marginLeft: '12px' }}>
            {isLiveMode ? (
              <>
                <Wifi size={14} style={{ color: '#10B981' }} />
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#065F46' }}>Live Mode</span>
              </>
            ) : (
              <>
                <WifiOff size={14} style={{ color: '#F97316' }} />
                <span style={{ fontSize: '11px', fontWeight: 600, color: '#9A3412' }}>Offline Mode</span>
              </>
            )}
          </div>

          <div className={styles.headerMeta}>
            <div className={styles.metaTextContainer}>
              <div className={styles.metaResultsCount}>{totalCountToRender.toLocaleString()} results</div>
              <div className={styles.metaSubTitle}>SHOWING TOP 10,000</div>
            </div>
            
            <div className={styles.metaDivider} />
            
            <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close search">
              <X size={20} />
            </button>
          </div>
        </header>

        {/* ================= BODY WRAPPER ================= */}
        <div className={styles.body}>
          
          <div className={styles.mainLayout}>
            
            {/* Sidebar Filters */}
            <FiltersPanel 
              sidebarWidth={sidebarWidth}
              fileTypes={filters.fileTypes}
              author={filters.author}
              date={filters.date}
              setFilters={setFilters}
              toggleFileType={toggleFileType}
              isAuthorDropdownOpen={false}
              setIsAuthorDropdownOpen={() => {}}
              startResizingSidebar={startResizingSidebar}
              isResizingSidebar={isResizingSidebar}
            />

            {/* Main Listing Panel */}
            <main className={styles.mainContainer}>
              
              {/* Search Tabs */}
              <SearchTabs 
                activeTab={activeTopTab}
                onTabChange={(tab) => {
                  setActiveTopTab(tab);
                  setFrom(0);
                }}
              />

              {/* Statistics & Filter Flags Bar */}
              <div className={styles.statsBar}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <span className={styles.statsLabel}>
                    Showing {resultsToRender.length} of {totalCountToRender.toLocaleString()} results
                  </span>
                  
                  <div className={styles.filterBadgeWrapper}>
                    {activeTopTab !== 'All' && (
                      <div className={styles.filterBadge}>
                        Tab: {activeTopTab}
                        <button onClick={() => setActiveTopTab('All')} className={styles.filterBadgeClose}><X size={10} /></button>
                      </div>
                    )}
                    {filters.fileTypes.length > 0 && !filters.fileTypes.includes('All') && (
                      <div className={styles.filterBadge}>
                        Types: {filters.fileTypes.join(', ')}
                        <button onClick={() => setFilters({ ...filters, fileTypes: ['All'] })} className={styles.filterBadgeClose}><X size={10} /></button>
                      </div>
                    )}
                    {filters.date && (
                      <div className={styles.filterBadge}>
                        Date: {filters.date}
                        <button onClick={() => setFilters({ ...filters, date: '' })} className={styles.filterBadgeClose}><X size={10} /></button>
                      </div>
                    )}
                    {filters.author && (
                      <div className={styles.filterBadge}>
                        Author: {filters.author}
                        <button onClick={() => setFilters({ ...filters, author: '' })} className={styles.filterBadgeClose}><X size={10} /></button>
                      </div>
                    )}
                  </div>
                </div>

                <button 
                  onClick={() => {
                    setIsHistoryOpen(!isHistoryOpen);
                    setSelectedFileId(null);
                  }}
                  className={styles.historyNavButton}
                >
                  <HistoryIcon size={14} strokeWidth={2.5} />
                  <span style={{ fontWeight: 900 }}>History</span>
                </button>
              </div>

              {/* Results cards panel */}
              <div className={styles.resultsListWrapper}>
                {isLoading ? (
                  <SkeletonLoader count={4} />
                ) : resultsToRender.length === 0 ? (
                  <div className={styles.emptyState}>
                    <SlidersHorizontal size={40} />
                    <h3>No results found</h3>
                    <p>Try broadening your terms or resetting filters.</p>
                  </div>
                ) : (
                  resultsToRender.map((result, idx) => {
                    const isSelected = selectedFileId === result.id;
                    const colors = getFileColor(result.fileType);
                    const isStarred = starredIds.has(result.id);
                    
                    const formatBytes = (bytes: number): string => {
                      if (!bytes) return '4.2 MB';
                      const k = 1024;
                      const sizes = ['Bytes', 'KB', 'MB', 'GB'];
                      const i = Math.floor(Math.log(bytes) / Math.log(k));
                      return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
                    };
                    const sizeLabel = result.size ? formatBytes(result.size) : '4.2 MB';
                    const dateLabel = result.lastModified ? new Date(result.lastModified).toLocaleDateString() : 'Today';

                    return (
                      <div
                        key={result.id}
                        className={`${styles.resultCard} ${isSelected ? styles.resultCardSelected : ''}`}
                        onClick={() => setSelectedFileId(isSelected ? null : result.id)}
                        style={{ borderLeftColor: colors.color } as React.CSSProperties}
                      >
                        {/* File action dot menu */}
                        <div className={styles.actionMenuAnchor}>
                          <FileActionMenu 
                            file={result} 
                            onOpenChange={(open) => setOpenMenuId(open ? result.id : null)} 
                          />
                        </div>

                        <div className={styles.cardLeftBlock}>
                          <div 
                            className={styles.cardIconBox}
                            style={{ backgroundColor: colors.color + '15', color: colors.color }}
                          >
                            {result.fileType?.toLowerCase() === 'png' ? <ImageIcon size={22} strokeWidth={2} /> : 
                             ['xls', 'xlsx'].includes(result.fileType?.toLowerCase()) ? <FileSpreadsheet size={22} strokeWidth={2} /> :
                             ['pdf'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
                             ['doc', 'docx'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
                             <FileText size={22} strokeWidth={2} />}
                          </div>
                          
                          <button 
                            onClick={(e) => toggleStar(e, result.id)}
                            className={styles.starIconButton}
                            style={{ color: isStarred ? '#FBBF24' : undefined }}
                          >
                            <Star size={16} fill={isStarred ? '#FBBF24' : 'transparent'} />
                          </button>
                        </div>
                        
                        <div className={styles.cardBody}>
                          <div className={styles.cardHeaderRow}>
                            <h4 
                              className={styles.cardTitle}
                              style={{ color: isSelected ? colors.color : undefined }}
                            >
                              {result.title}
                            </h4>
                            <div className={styles.scoreContainer}>
                              <div className={styles.scoreRow}>
                                <span className={`${styles.scoreBadge} ${styles.scoreBM25}`}>BM25: 4.80</span>
                                <span className={`${styles.scoreBadge} ${styles.scoreTFIDF}`}>TF-IDF: 2.12</span>
                              </div>
                            </div>
                          </div>

                          <div className={styles.cardMetadataRow}>
                            <span className={styles.metaBoldLabel}>BY: <span className={styles.metaValueDark}>{result.author}</span></span>
                            <span className={styles.metaDot} />
                            <span>{dateLabel}</span>
                            <span className={styles.metaDot} />
                            <span>{sizeLabel}</span>
                          </div>

                          <p className={styles.cardDescription}>{result.summary}</p>
                          
                          <div className={styles.cardProjectHubRow}>
                            <span>PROJECT HUB: </span>
                            <span className={styles.projectHubValue} style={{ color: colors.color }}>{result.siteName}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>

              {/* Pagination controls footer */}
              {!isLoading && resultsToRender.length > 0 && (
                <footer className={styles.pagination}>
                  <span className={styles.paginationInfo}>
                    PAGE {Math.floor(from / 10) + 1} <span className={styles.infoDivider}>|</span> SHOWING <span className={styles.infoHighlight}>{from + 1}-{from + resultsToRender.length}</span> OF <span className={styles.infoHighlight}>{totalCountToRender}</span>
                  </span>

                  <PaginationComponent
                    totalCount={totalCountToRender}
                    pageSize={10}
                    from={from}
                    onPageChange={(newFrom) => {
                      if (isLiveMode) setFrom(newFrom);
                    }}
                  />

                  <div className={styles.jumpToWrapper}>
                    <span className={styles.jumpToLabel}>Jump to</span>
                    <input 
                      type="text" 
                      placeholder={(Math.floor(from / 10) + 1).toString()}
                      className={styles.jumpToInput}
                      disabled
                    />
                  </div>
                </footer>
              )}

              {/* Bottom tabs for sub-navigation */}
              <nav className={styles.bottomTabNav}>
                <div className={styles.bottomTabWrapper}>
                  {(['Search Results', 'Recent Activities', 'Starred Assets'] as const).map((tab) => {
                    const isActive = bottomTab === tab;
                    return (
                      <button 
                        key={tab}
                        onClick={() => {
                          setBottomTab(tab);
                          setIsHistoryOpen(tab === 'Recent Activities');
                        }}
                        className={`${styles.bottomTabItem} ${isActive ? styles.bottomTabActive : ''}`}
                      >
                        {tab}
                        {isActive && (
                          <div className={styles.bottomActiveLine} />
                        )}
                      </button>
                    );
                  })}
                </div>
              </nav>

            </main>

            {/* Document preview analysis panel */}
            <DocumentAnalysis 
              selectedFile={selectedFile}
              setSelectedFile={(file) => setSelectedFileId(file ? file.id : null)}
              previewWidth={previewWidth}
              searchQuery={searchQuery}
              isResizingPreview={isResizingPreview}
              startResizingPreview={startResizingPreview}
            />

            {/* Recent Searches history panel */}
            <SearchHistory 
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
    </div>
  );
}
