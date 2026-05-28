import * as React from 'react';
import { Search, X } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { IHeaderProps } from '../interface/IHeaderProps';

export const Header: React.FC<IHeaderProps> = ({
  searchQuery,
  setSearchQuery,
  handleSearch,
  isSearchFocused,
  setIsSearchFocused,
  searchHistory,
  totalCountToRender,
  onDismiss
}) => {
  return (
    <header className={styles.header}>
      {/* Logo */}
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

      {/* Centered Search Bar */}
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
          
          {searchQuery && (
            <button 
              onClick={() => {
                setSearchQuery('');
                handleSearch('');
              }}
              className={styles.searchClearBtn}
              title="Clear search"
              aria-label="Clear search"
            >
              <X size={14} strokeWidth={2.5} />
            </button>
          )}
          
          {isSearchFocused && searchHistory.length > 0 && (
            <div className={styles.searchSuggestionBox}>
              <div className={styles.suggestionHeader}>
                <span>Recent Searches</span>
              </div>
              <div>
                {searchHistory?.slice(0, 5).map((historyItem, idx) => (
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

      {/* Header Meta Actions (Total counts and close button) */}
      <div className={styles.headerMeta}>
        <div className={styles.metaTextContainer}>
          <div className={styles.metaResultsCount}>{totalCountToRender.toLocaleString()} results</div>
        </div>
        
        <div className={styles.metaDivider} />
        
        <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close search">
          <X size={20} />
        </button>
      </div>
    </header>
  );
};
