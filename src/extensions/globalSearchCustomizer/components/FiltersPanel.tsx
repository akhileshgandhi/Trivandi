import * as React from 'react';
import { User, Calendar } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';

import { IFiltersPanelProps } from '../interface/IFiltersPanelProps';

const FILE_TYPE_OPTIONS = ['All', 'PDF', 'DOC', 'XLS', 'PPT'];
const DATE_OPTIONS = ['Today', 'Yesterday', 'This Week', 'This Month', 'This Year'];

export const FiltersPanel: React.FC<IFiltersPanelProps> = ({
  sidebarWidth,
  fileTypes,
  author,
  date,
  setFilters,
  toggleFileType,
  isAuthorDropdownOpen,
  setIsAuthorDropdownOpen,
  startResizingSidebar,
  isResizingSidebar,
  authorsList = ['Sarah Chen', 'Robert Wilson', 'Elena Rodriguez', 'James T. Kirk', 'Security Operations', 'HR Department']
}) => {
  return (
    <>
      <aside 
        className={styles.sidebar}
        style={{ width: sidebarWidth }}
      >
        <div className={styles.sidebarInner}>
          {/* Filter by Type */}
          <div>
            <h3 className={styles.sidebarTitle}>Filter by Type</h3>
            
            <div className={styles.sidebarOptionList}>
              {FILE_TYPE_OPTIONS.map(type => {
                const isActive = fileTypes.includes(type);
                return (
                  <label 
                    key={type} 
                    className={`${styles.checkboxLabel} ${isActive ? styles.checkboxActive : ''}`}
                  >
                    <input 
                      type="checkbox" 
                      checked={isActive}
                      onChange={() => toggleFileType(type)}
                      className={styles.checkboxInput}
                    />
                    <span>{type}</span>
                    {type === 'All' && <span className={styles.checkboxBadge}>1.3M</span>}
                  </label>
                );
              })}
            </div>

            {/* Filter by Author */}
            <h3 className={styles.sidebarSectionHeader}>
              Author <span className={styles.countLabel}>({authorsList.length})</span>
            </h3>
            <div className={styles.authorInputWrapper}>
              <div style={{ position: 'relative' }}>
                <User className={styles.authorIcon} size={16} />
                <input 
                  type="text" 
                  placeholder="Find an author..."
                  value={author}
                  onFocus={() => setIsAuthorDropdownOpen(true)}
                  onBlur={() => setTimeout(() => setIsAuthorDropdownOpen(false), 200)}
                  onChange={(e) => setFilters({ fileTypes, author: e.target.value, date })}
                  className={styles.authorInput}
                />
              </div>
              
              {isAuthorDropdownOpen && (
                <div className={styles.authorDropdown}>
                  {authorsList
                    .filter(name => name.toLowerCase().includes(author.toLowerCase()))
                    .map(name => (
                      <button
                        key={name}
                        onMouseDown={(e) => e.preventDefault()} // Prevent input blur before click
                        onClick={() => setFilters({ fileTypes, author: name, date })}
                        className={styles.authorDropdownItem}
                      >
                        <div className={styles.authorInitialsAvatar}>
                          {name.charAt(0)}
                        </div>
                        {name}
                      </button>
                    ))
                  }
                  {authorsList.filter(name => name.toLowerCase().includes(author.toLowerCase())).length === 0 && (
                    <div style={{ padding: '16px', fontSize: '14px', color: '#94a3b8', fontStyle: 'italic', textAlign: 'center' }}>
                      No authors found
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Filter by Date */}
            <h3 className={styles.sidebarSectionHeader}>Modified</h3>
            <div className={styles.sidebarOptionList}>
              {DATE_OPTIONS.map(opt => {
                const isActive = date === opt;
                return (
                  <button 
                    key={opt}
                    onClick={() => setFilters({ fileTypes, author, date: isActive ? '' : opt })}
                    className={`${styles.dateFilterButton} ${isActive ? styles.dateFilterActive : ''}`}
                  >
                    {opt}
                  </button>
                );
              })}

              <div className={styles.datePickerLabel}>Or pick a specific date</div>
              <div className={styles.datePickerBox}>
                <div className={styles.datePickerIcon}>
                  <Calendar size={18} strokeWidth={2.5} />
                </div>
                <input 
                  type="date"
                  value={date && !DATE_OPTIONS.includes(date) ? date : ''}
                  onChange={(e) => setFilters({ fileTypes, author, date: e.target.value })}
                  className={styles.dateInput}
                />
              </div>
            </div>
          </div>

          {/* Action Row */}
          <div className={styles.sidebarActionWrapper}>
            <button 
              className={styles.applyButton}
              onClick={() => {}}
            >
              Apply Filter
            </button>
            <button 
              onClick={() => setFilters({ fileTypes: ['All'], author: '', date: '' })}
              className={styles.resetButton}
            >
              Reset Filters
            </button>
          </div>
        </div>
      </aside>

      <div 
        onMouseDown={startResizingSidebar}
        className={`${styles.resizeHandle} ${isResizingSidebar ? styles.resizeHandleActive : ''}`}
      />
    </>
  );
};
