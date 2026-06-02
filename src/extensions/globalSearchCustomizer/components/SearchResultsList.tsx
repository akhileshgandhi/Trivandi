import * as React from 'react';
import { SlidersHorizontal, History as HistoryIcon, Star, ChevronLeft, ChevronRight, X, ImageIcon, FileSpreadsheet, FileText, Video } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchResult } from '../../../models/ISearchResult';
import { FileActionMenu } from './FileActionMenu';
import { renderFormattedSummary, getSharePointThumbnailUrl } from '../../../utils/SearchHelpers';
import { SkeletonLoader } from '../Common/SkeletonLoader';
import { PaginationComponent } from '../Common/PaginationComponent';

import { ISearchResultsListProps } from '../interface/ISearchResultsListProps';

interface ISearchResultThumbnailProps {
  result: ISearchResult;
}

const SearchResultThumbnail: React.FC<ISearchResultThumbnailProps> = ({ result }) => {
  const [imageError, setImageError] = React.useState(false);
  const isImage = ['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(result.fileType?.toLowerCase() || '');

  if (isImage && !imageError) {
    const thumbUrl = getSharePointThumbnailUrl(result.webUrl, result.title);
    if (thumbUrl) {
      return (
        <img
          src={thumbUrl}
          alt={result.title}
          onError={() => setImageError(true)}
          className={styles.cardThumbnailImage}
        />
      );
    }
  }

  // Fallback to standard icons
  const ft = result.fileType ? result.fileType.toLowerCase() : '';
  if (isImage) {
    return <ImageIcon size={22} strokeWidth={2} />;
  }
  if (['xls', 'xlsx'].includes(ft)) {
    return <FileSpreadsheet size={22} strokeWidth={2} />;
  }
  if (['pdf'].includes(ft)) {
    return <FileText size={22} strokeWidth={2} />;
  }
  if (['doc', 'docx'].includes(ft)) {
    return <FileText size={22} strokeWidth={2} />;
  }
  if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) {
    return <Video size={22} strokeWidth={2} />;
  }
  return <FileText size={22} strokeWidth={2} />;
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
  openMenuId,
  searchHistory,
  sortBy,
  setSortBy,
  promotedResults
}) => {

  return (
    <main className={styles.mainContainer}>
      {/* Statistics & Filter Flags Bar */}
      <div className={styles.statsBar}>
        <div className={styles.statsLeft}>
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
            {filters.selectedAuthors && filters.selectedAuthors.length > 0 && (
              filters.selectedAuthors.map(authorName => (
                <div key={authorName} className={styles.filterBadge}>
                  Author: {authorName}
                  <button 
                    onClick={() => {
                      setFilters({
                        ...filters,
                        selectedAuthors: filters.selectedAuthors.filter(a => a !== authorName)
                      });
                    }} 
                    className={styles.filterBadgeClose}
                  >
                    <X size={12} />
                  </button>
                </div>
              ))
            )}
          </div>
        </div>

        <div className={styles.statsDivider} />

        {/* Dynamic Sort Ranking Option */}
        {bottomTab === 'Search Results' && (
          <div className={styles.sortContainer}>
            <span className={styles.sortLabel}>SORT BY</span>
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value)}
              className={styles.sortSelect}
              title="Change search ranking criteria"
            >
              <option value="relevance">Relevance Tuning</option>
              <option value="dateDesc">Newest Modified</option>
              <option value="dateAsc">Oldest Modified</option>
              <option value="sizeDesc">Largest Assets</option>
            </select>
          </div>
        )}

        <div className={styles.statsDivider} />

        <button 
          onClick={() => {
            setIsHistoryOpen(!isHistoryOpen);
            setSelectedFileId(null);
          }}
          className={styles.historyNavButton}
        >
          <HistoryIcon size={12} strokeWidth={2.5} />
          <span>History ({searchHistory?.length || 0})</span>
        </button>
      </div>

      {/* Results cards panel */}
      <div className={styles.resultsListWrapper}>
        {/* Render High-Fidelity Promoted Resources / Bookmarks */}
        {!isLoading && bottomTab === 'Search Results' && promotedResults && promotedResults.length > 0 && (
          <div className={styles.promotedSection}>
            {promotedResults.map((promo) => (
              <a
                key={promo.id}
                href={promo.url}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.promotedCard}
              >
                <div className={styles.promotedCrownBox}>
                  👑
                </div>
                <div className={styles.promotedCardBody}>
                  <div className={styles.promotedTitleRow}>
                    <h4 className={styles.promotedCardTitle}>{promo.title}</h4>
                    <span className={styles.promotedCategoryBadge}>{promo.category}</span>
                  </div>
                  <p className={styles.promotedCardDescription}>{promo.description}</p>
                  <span className={styles.promotedUrlLabel}>Explore Official Portal ↗</span>
                </div>
              </a>
            ))}
          </div>
        )}

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
            const isStarred = starredIds.has(result.id);
            
            const formatBytes = (bytes: number): string => {
              if (!bytes) return '---';
              const k = 1024;
              const sizes = ['Bytes', 'KB', 'MB', 'GB'];
              const i = Math.floor(Math.log(bytes) / Math.log(k));
              return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
            };
            const sizeLabel = result.fileType === 'folder'
              ? 'Folder'
              : (result.size ? formatBytes(result.size) : '---');
            const dateLabel = result.lastModified ? new Date(result.lastModified).toLocaleDateString() : 'Today';

            const isMenuOpen = openMenuId === result.id;

            const ft = result.fileType ? result.fileType.toLowerCase() : '';
            let typeClass = styles.type_default;
            let accentColor = '#1a73e8';
            if (ft === 'pdf') {
              typeClass = styles.type_pdf;
              accentColor = '#00acc1';
            } else if (['xls', 'xlsx'].includes(ft)) {
              typeClass = styles.type_xlsx;
              accentColor = '#00796b';
            } else if (['ppt', 'pptx'].includes(ft)) {
              typeClass = styles.type_pptx;
              accentColor = '#e52592';
            } else if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ft)) {
              typeClass = styles.type_image;
              accentColor = '#9334e6';
            } else if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) {
              typeClass = styles.type_video;
              accentColor = '#ea4335';
            } else if (ft === 'folder') {
              typeClass = styles.type_folder;
              accentColor = '#f5b041';
            }

            return (
              <div
                key={result.id}
                className={`${styles.resultCard} ${typeClass} ${isSelected ? styles.resultCardSelected : ''} ${isMenuOpen ? styles.resultCardMenuOpen : ''}`}
                onClick={() => setSelectedFileId(isSelected ? null : result.id)}
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
                    className={`${styles.cardIconBox} ${['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(result.fileType?.toLowerCase()) ? styles.cardIconBoxImage : ''}`}
                  >
                    <SearchResultThumbnail result={result} />
                  </div>
                  
                  <button 
                    onClick={(e) => toggleStar(e, result.id)}
                    className={`${styles.starIconButton} ${isStarred ? styles.starIconButtonStarred : ''}`}
                  >
                    <Star size={16} fill={isStarred ? '#FBBF24' : 'transparent'} />
                  </button>
                </div>
                
                <div className={styles.cardBody}>
                  <div className={styles.cardHeaderRow}>
                    <h4 className={`${styles.cardTitle} ${isSelected ? styles.cardTitleSelected : ''}`}>
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

                  <p className={styles.cardDescription}>{renderFormattedSummary(result.summary, '', accentColor)}</p>
                  
                  <div className={styles.cardProjectHubRow}>
                    <span>PROJECT HUB: </span>
                    <span className={styles.projectHubValue}>{result.siteName}</span>
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
