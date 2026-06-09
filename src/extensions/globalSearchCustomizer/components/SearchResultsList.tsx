import * as React from 'react';
import { SlidersHorizontal, History as HistoryIcon, Star, ChevronLeft, ChevronRight, X, ChevronDown } from 'lucide-react';
import { BsFiletypeXls, BsFiletypeXlsx, BsFiletypePpt, BsFiletypePptx, BsFiletypeDoc, BsFiletypeDocx, BsFileEarmarkText, BsFilePlay } from 'react-icons/bs';
import { IoMdImages } from 'react-icons/io';
import { FaFolderOpen, FaRegFilePdf } from 'react-icons/fa6';

const IconXls = BsFiletypeXls as any;
const IconXlsx = BsFiletypeXlsx as any;
const IconPpt = BsFiletypePpt as any;
const IconPptx = BsFiletypePptx as any;
const IconPdf = FaRegFilePdf as any;
const IconDoc = BsFiletypeDoc as any;
const IconDocx = BsFiletypeDocx as any;
const IconFolder = FaFolderOpen as any;
const IconPlay = BsFilePlay as any;
const IconText = BsFileEarmarkText as any;
const IconImages = IoMdImages as any;
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
  const isMedia = ['png', 'jpg', 'jpeg', 'gif', 'svg', 'mp4', 'mov', 'avi'].includes(result.fileType?.toLowerCase() || '');

  if (isMedia && !imageError) {
    const thumbUrl = result.thumbnailUrl || '';
    if (thumbUrl) {
      return (
        <img
          src={thumbUrl}
          loading="lazy"
          alt={result.title}
          onError={() => setImageError(true)}
          className={styles.cardThumbnailImage}
        />
      );
    }
  }

  // Fallback to standard icons
  const ft = result.fileType ? result.fileType.toLowerCase() : '';
  if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ft)) {
    return <IconImages size={22} />;
  }
  if (ft === 'xls') {
    return <IconXls size={22} />;
  }
  if (ft === 'xlsx') {
    return <IconXlsx size={22} />;
  }
  if (ft === 'ppt') {
    return <IconPpt size={22} />;
  }
  if (ft === 'pptx') {
    return <IconPptx size={22} />;
  }
  if (ft === 'pdf') {
    return <IconPdf size={22} />;
  }
  if (ft === 'doc') {
    return <IconDoc size={22} />;
  }
  if (ft === 'docx') {
    return <IconDocx size={22} />;
  }
  if (ft === 'folder') {
    return <IconFolder size={22} />;
  }
  if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ft)) {
    return <IconPlay size={22} />;
  }
  return <IconText size={22} />;
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
  setSortBy
}) => {
  const [isSortOpen, setIsSortOpen] = React.useState(false);
  const sortRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (sortRef.current && !sortRef.current.contains(event.target as Node)) {
        setIsSortOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const sortOptions = [
    { value: 'relevance', label: 'Relevance Tuning' },
    { value: 'dateDesc', label: 'Newest Modified' },
    { value: 'dateAsc', label: 'Oldest Modified' },
    { value: 'sizeDesc', label: 'Largest Assets' }
  ];

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
            {filters.selectedSites && filters.selectedSites.length > 0 && (
              filters.selectedSites.map(siteId => (
                <div key={siteId} className={styles.filterBadge}>
                  Site: {siteId}
                  <button 
                    onClick={() => {
                      setFilters({
                        ...filters,
                        selectedSites: filters.selectedSites.filter(s => s !== siteId)
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
            <div className={styles.customSortWrapper} ref={sortRef}>
              <button 
                className={styles.customSortButton} 
                onClick={() => setIsSortOpen(!isSortOpen)}
                title="Change search ranking criteria"
              >
                {sortOptions.find(opt => opt.value === sortBy)?.label || 'Relevance Tuning'}
                <ChevronDown size={14} />
              </button>
              {isSortOpen && (
                <div className={styles.customSortDropdown}>
                  {sortOptions.map(opt => (
                    <button
                      key={opt.value}
                      className={`${styles.customSortOption} ${sortBy === opt.value ? styles.activeSortOption : ''}`}
                      onClick={() => {
                        setSortBy(opt.value);
                        setIsSortOpen(false);
                      }}
                    >
                      {opt.label}
                    </button>
                  ))}
                </div>
              )}
            </div>
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
        {/* Promoted Results feature has been removed as per user request */}

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
              accentColor = '#d93025';
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
      {!isLoading && resultsToRender.length > 0 && bottomTab === 'Search Results' && (
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
