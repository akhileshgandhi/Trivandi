import * as React from 'react';
import { History as HistoryIcon, X, Search } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { ISearchHistoryProps } from '../interface/ISearchHistoryProps';

export const HistoryPane: React.FC<ISearchHistoryProps> = ({
  isHistoryOpen,
  setIsHistoryOpen,
  previewWidth,
  searchHistory,
  handleSearch,
  setSearchHistory,
  removeHistoryItem,
  isResizingPreview,
  startResizingPreview,
  topQueries,
  topClickedDocs,
  onClearAnalytics
}) => {
  if (!isHistoryOpen) return null;

  return (
    <>
      <div 
        onMouseDown={startResizingPreview}
        className={`${styles.resizeHandle} ${isResizingPreview ? styles.resizeHandleActive : ''}`}
      />

      <aside
        className={styles.historyPane}
        style={{ width: previewWidth }}
      >
        {/* Header Block */}
        <div className={styles.historyHeader}>
          <div className={styles.historyTitleBox}>
            <HistoryIcon size={18} className={styles.historyHeaderIcon} strokeWidth={2} />
            <span className={styles.historyTitleLabel}>Search History ({searchHistory.length})</span>
          </div>
          <button 
            onClick={() => setIsHistoryOpen(false)}
            className={styles.historyCloseBtn}
          >
            <X size={20} />
          </button>
        </div>

        {/* Scrollable Query items and Analytics panels */}
        <div className={styles.historyScrollContent}>
          
          {/* Section 1: Recent Searches */}
          <div className={styles.historySectionHeader}>
            <span>RECENT SEARCH HISTORY</span>
          </div>
          <div className={styles.historyListGroup}>
            {searchHistory.length > 0 ? (
              searchHistory.slice(0, 8).map((item, idx) => (
                <div 
                  key={idx}
                  className={styles.historyItemCard}
                >
                  <button
                    onClick={() => {
                      handleSearch(item);
                      setIsHistoryOpen(false);
                    }}
                    className={styles.historyItemTrigger}
                  >
                    <div className={styles.historyItemIconBox}>
                      <Search size={14} className={styles.suggestionIcon} />
                    </div>
                    <div className={styles.historyItemInfo}>
                      <p className={styles.historyItemQueryText}>{item}</p>
                    </div>
                  </button>
                  <button 
                    onClick={(e) => {
                      e.stopPropagation();
                      removeHistoryItem(item);
                    }}
                    className={styles.historyItemRemoveBtn}
                    title="Remove history item"
                  >
                    <X size={14} />
                  </button>
                </div>
              ))
            ) : (
              <div className={styles.historyEmptyState}>
                <div className={styles.historyEmptyIconFrame}>
                  <HistoryIcon size={24} className={styles.historyEmptyIcon} />
                </div>
                <p className={styles.historyEmptyTitle}>No Recent Searches</p>
                <p className={styles.historyEmptyDesc}>
                  Your search terms will be archived here for instant retrieval.
                </p>
              </div>
            )}
          </div>

          {/* Section 2: Top Search Analytics (Top Queries) */}
          {topQueries && topQueries.length > 0 && (
            <div className={styles.analyticsWrapper}>
              <div className={styles.historySectionHeader}>
                <span>YOUR TOP SEARCH QUERIES</span>
              </div>
              <div className={styles.analyticsQueriesList}>
                {topQueries.map((item, idx) => (
                  <button 
                    key={idx} 
                    onClick={() => {
                      handleSearch(item.query);
                      setIsHistoryOpen(false);
                    }}
                    className={styles.analyticsQueryRow}
                    title={`Search for ${item.query} again`}
                  >
                    <span className={styles.analyticsQueryText}>{item.query}</span>
                    <span className={styles.analyticsCountBadge}>{item.count} {item.count === 1 ? 'search' : 'searches'}</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Section 3: Click Analytics (Frequently Visited Documents) */}
          {topClickedDocs && topClickedDocs.length > 0 && (
            <div className={styles.analyticsWrapper}>
              <div className={styles.historySectionHeader}>
                <span>FREQUENTLY VISITED DOCUMENTS</span>
              </div>
              <div className={styles.analyticsDocsList}>
                {topClickedDocs.map((doc, idx) => (
                  <a 
                    key={idx} 
                    href={doc.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className={styles.analyticsDocRow}
                    title={`Open ${doc.title} in a new tab`}
                  >
                    <div className={styles.analyticsDocIcon}>📄</div>
                    <div className={styles.analyticsDocInfo}>
                      <p className={styles.analyticsDocTitle}>{doc.title}</p>
                      <p className={styles.analyticsDocClicks}>{doc.clickCount} {doc.clickCount === 1 ? 'view' : 'views'}</p>
                    </div>
                  </a>
                ))}
              </div>
            </div>
          )}

        </div>

        {/* Clear Button Footer */}
        <div className={styles.historyFooter}>
          <button 
            onClick={() => {
              setSearchHistory([]);
              if (onClearAnalytics) onClearAnalytics();
            }}
            className={styles.clearHistoryButton}
          >
            Reset Analytics & History
          </button>
        </div>
      </aside>
    </>
  );
};
