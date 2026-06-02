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
        <img 
          src={require('../Assets/moreyeahLogo.png')} 
          alt="Moreyeahs Logo" 
          className={styles.logoImage} 
        />
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
