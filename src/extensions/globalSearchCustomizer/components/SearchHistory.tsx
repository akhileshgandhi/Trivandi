import * as React from 'react';
import { History as HistoryIcon, X, Search } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { ISearchHistoryProps } from '../interface/ISearchHistoryProps';

export const SearchHistory: React.FC<ISearchHistoryProps> = ({
  isHistoryOpen,
  setIsHistoryOpen,
  previewWidth,
  searchHistory,
  handleSearch,
  setSearchHistory,
  removeHistoryItem,
  isResizingPreview,
  startResizingPreview
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
            <HistoryIcon size={18} style={{ color: '#475569' }} strokeWidth={2} />
            <span className={styles.historyTitleLabel}>Search History</span>
          </div>
          <button 
            onClick={() => setIsHistoryOpen(false)}
            className={styles.historyCloseBtn}
          >
            <X size={20} />
          </button>
        </div>

        {/* Scrollable Query items */}
        <div className={styles.historyScrollContent}>
          <div className={styles.historyListGroup}>
            {searchHistory.length > 0 ? (
              searchHistory.map((item, idx) => (
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
                      <Search size={16} className={styles.suggestionIcon} />
                    </div>
                    <div className={styles.historyItemInfo}>
                      <p className={styles.historyItemQueryText}>{item}</p>
                      <p className={styles.historyItemTimeLabel}>Indexed 2m ago</p>
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
                  <HistoryIcon size={24} style={{ color: '#cbd5e1' }} />
                </div>
                <p className={styles.historyEmptyTitle}>Terminal Void</p>
                <p className={styles.historyEmptyDesc}>
                  Your recent academic queries will be archived here for instant retrieval.
                </p>
              </div>
            )}
          </div>
        </div>

        {/* Clear Button Footer */}
        <div className={styles.historyFooter}>
          <button 
            onClick={() => setSearchHistory([])}
            className={styles.clearHistoryButton}
          >
            Clear History
          </button>
        </div>
      </aside>
    </>
  );
};
