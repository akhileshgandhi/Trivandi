import * as React from 'react';
import { SlidersHorizontal, History as HistoryIcon, Star, ChevronLeft, ChevronRight, X, ImageIcon, FileSpreadsheet, FileText, Video } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';
import { FileActionMenu } from './FileActionMenu';
import { SkeletonLoader } from '../Common/SkeletonLoader';
import { PaginationComponent } from '../Common/PaginationComponent';

import { ISearchResultsListProps } from '../interface/ISearchResultsListProps';

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

export const SearchResultsList: React.FC<ISearchResultsListProps> = ({
  resultsToRender,
  totalCountToRender,
  isLoading,
  selectedFileId,
  setSelectedFileId,
  starredIds,
  toggleStar,
  from,
  setFrom,
  bottomTab,
  setBottomTab,
  setIsHistoryOpen,
  isHistoryOpen,
  handleClearAllFilters,
  activeTopTab,
  setActiveTopTab,
  filters,
  setFilters,
  setOpenMenuId,
  openMenuId
}) => {
  const [jumpPage, setJumpPage] = React.useState('');

  return (
    <main className={styles.mainContainer}>
      {/* Statistics & Filter Flags Bar */}
      <div className={styles.statsBar}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flex: 1 }}>
          <span className={styles.statsLabel}>
            Showing {resultsToRender.length} of {totalCountToRender.toLocaleString()} results
          </span>
          
          <div className={styles.filterBadgeWrapper}>
            {activeTopTab !== 'All' && (
              <div className={styles.filterBadge}>
                Tab: {activeTopTab}
                <button onClick={() => setActiveTopTab('All')} className={styles.filterBadgeClose}><X size={12} /></button>
              </div>
            )}
            {filters.fileTypes.length > 0 && !filters.fileTypes.includes('All') && (
              filters.fileTypes.map(type => (
                <div key={type} className={styles.filterBadge}>
                  Type: {type}
                  <button 
                    onClick={() => {
                      const updated = filters.fileTypes.filter(t => t !== type);
                      setFilters({
                        ...filters,
                        fileTypes: updated.length === 0 ? ['All'] : updated
                      });
                    }} 
                    className={styles.filterBadgeClose}
                  >
                    <X size={12} />
                  </button>
                </div>
              ))
            )}
            {filters.date && (
              <div className={styles.filterBadge}>
                Date: {filters.date}
                <button onClick={() => setFilters({ ...filters, date: '' })} className={styles.filterBadgeClose}><X size={12} /></button>
              </div>
            )}
            {filters.author && (
              <div className={styles.filterBadge}>
                Author: {filters.author}
                <button onClick={() => setFilters({ ...filters, author: '' })} className={styles.filterBadgeClose}><X size={12} /></button>
              </div>
            )}
          </div>
        </div>

        <div className={styles.statsDivider} />

        {/* History Action Button */}
        <button 
          onClick={() => {
            setIsHistoryOpen(!isHistoryOpen);
            setSelectedFileId(null);
          }}
          className={styles.historyNavButton}
        >
          <HistoryIcon size={12} strokeWidth={2.5} />
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

            const isMenuOpen = openMenuId === result.id;

            return (
              <div
                key={result.id}
                className={`${styles.resultCard} ${isSelected ? styles.resultCardSelected : ''}`}
                onClick={() => setSelectedFileId(isSelected ? null : result.id)}
                style={{ 
                  borderLeftColor: colors.color,
                  zIndex: isMenuOpen ? 50 : undefined
                } as React.CSSProperties}
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
                    {['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(result.fileType?.toLowerCase()) ? <ImageIcon size={22} strokeWidth={2} /> : 
                     ['xls', 'xlsx'].includes(result.fileType?.toLowerCase()) ? <FileSpreadsheet size={22} strokeWidth={2} /> :
                     ['pdf'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
                     ['doc', 'docx'].includes(result.fileType?.toLowerCase()) ? <FileText size={22} strokeWidth={2} /> :
                     ['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(result.fileType?.toLowerCase()) ? <Video size={22} strokeWidth={2} /> :
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
              setFrom(newFrom);
            }}
          />

          <div className={styles.jumpToWrapper}>
            <span className={styles.jumpToLabel}>Jump to</span>
            <input 
              type="text" 
              placeholder={(Math.floor(from / 10) + 1).toString()}
              value={jumpPage}
              onChange={(e) => setJumpPage(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  const pageNum = parseInt(jumpPage, 10);
                  const pageCount = Math.ceil(totalCountToRender / 10);
                  if (!isNaN(pageNum) && pageNum >= 1 && pageNum <= pageCount) {
                    setFrom((pageNum - 1) * 10);
                    setJumpPage('');
                  }
                }
              }}
              className={styles.jumpToInput}
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
                  setIsHistoryOpen(false);
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
  );
};
