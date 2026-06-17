import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { Search, X } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { IHeaderProps } from '../interface/IHeaderProps';
import { useAutocomplete } from '../hooks/useAutocomplete';
import { AutocompleteDropdown } from './AutocompleteDropdown';

export const Header: React.FC<IHeaderProps> = ({
  searchQuery,
  setSearchQuery,
  handleSearch,
  isSearchFocused,
  setIsSearchFocused,
  searchHistory,
  totalCountToRender,
  onDismiss,
  service,
  isAdmin,
  storeIsOwner,
  isAdminPanelOpen,
  setIsAdminPanelOpen,
  selectedSites = [],
}) => {
  const { suggestions, clearSuggestions } = useAutocomplete(searchQuery, service, selectedSites);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [showDropdown, setShowDropdown] = useState(false);
  const wrapperRef = useRef<HTMLDivElement>(null);

  // Click outside to clear suggestions
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
        clearSuggestions();
        setActiveIndex(-1);
        setShowDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [clearSuggestions]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveIndex(prev => Math.min(prev + 1, suggestions.length - 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIndex(prev => Math.max(prev - 1, -1));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      setShowDropdown(false);
      if (activeIndex >= 0 && suggestions[activeIndex]) {
        handleSelect(suggestions[activeIndex].label);
      } else {
        handleSearch(searchQuery);
        clearSuggestions();
      }
    } else if (e.key === 'Escape') {
      clearSuggestions();
      setActiveIndex(-1);
      setShowDropdown(false);
    }
  };

  const handleSelect = (label: string) => {
    setSearchQuery(label);
    clearSuggestions();
    setActiveIndex(-1);
    setShowDropdown(false);
    handleSearch(label);
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearchQuery(e.target.value);
    setActiveIndex(-1);
    setShowDropdown(true);
  };

  return (
    <header className={styles.header}>
      {/* Logo */}
      <div className={styles.logoWrapper}>
        <img 
          src={require('../Assets/TrivandiLogo.png')} 
          alt="Trivandi Logo" 
          className={styles.logoImage} 
        />
      </div>

      {/* Centered Search Bar */}
      <div className={styles.searchBarContainer}>
        <div className={styles.searchBarWrapper} ref={wrapperRef} style={{ position: 'relative' }}>
          <span className={styles.searchIcon}>
            <Search size={18} strokeWidth={2.5} />
          </span>
          <input 
            type="text"
            placeholder="Search documents, pages, files..."
            value={searchQuery}
            onFocus={() => {
              setIsSearchFocused(true);
              if (searchQuery.trim().length > 0) {
                setShowDropdown(true);
              }
            }}
            onBlur={() => setTimeout(() => setIsSearchFocused(false), 200)}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            className={styles.searchInput}
          />
          
          {searchQuery && (
            <button 
              onClick={() => {
                setSearchQuery('');
                handleSearch('');
                clearSuggestions();
                setActiveIndex(-1);
                setShowDropdown(false);
              }}
              className={styles.searchClearBtn}
              title="Clear search"
              aria-label="Clear search"
            >
              <X size={14} strokeWidth={2.5} />
            </button>
          )}

          {showDropdown && (
            <AutocompleteDropdown
              suggestions={suggestions}
              onSelect={handleSelect}
              activeIndex={activeIndex}
            />
          )}
          
          {isSearchFocused && searchHistory.length > 0 && (!suggestions || suggestions.length === 0) && searchQuery.length < 2 && (
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
        
        {(isAdmin || storeIsOwner) && (
          <>
            <div className={styles.metaDivider} />
            <button 
              className={`${styles.adminHeaderGearBtn} ${isAdminPanelOpen ? styles.adminHeaderGearBtnActive : ''}`} 
              onClick={() => setIsAdminPanelOpen && setIsAdminPanelOpen(!isAdminPanelOpen)} 
              title="Admin Portal"
            >
              <svg width="16" height="16" viewBox="0 0 24 24"
                fill="none" stroke="currentColor" strokeWidth="2"
                strokeLinecap="round" strokeLinejoin="round">
                <circle cx="12" cy="12" r="3"/>
                <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 
                  2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82
                  -.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09
                  A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l
                  -.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 
                  0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h
                  .09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82
                  l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 
                  0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09
                  a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l
                  .06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 
                  19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09
                  a1.65 1.65 0 0 0-1.51 1z"/>
              </svg>
            </button>
          </>
        )}

        <div className={styles.metaDivider} />
        
        <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close search">
          <X size={20} />
        </button>
      </div>
    </header>
  );
};
