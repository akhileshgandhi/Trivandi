import * as React from 'react';
import { User, Calendar, GripVertical } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { IFiltersPanelProps } from '../interface/IFiltersPanelProps';

const FILE_TYPE_OPTIONS = ['All', 'PDF', 'DOC', 'XLS', 'PPT'];
const DATE_OPTIONS = ['Today', 'Yesterday', 'This Week', 'This Month', 'This Year'];

export const Sidebar: React.FC<IFiltersPanelProps> = ({
  sidebarWidth,
  fileTypes,
  selectedAuthors,
  date,
  setFilters,
  toggleFileType,
  startResizingSidebar,
  isResizingSidebar,
  authorsList,
  searchService
}) => {
  const [authorSearchQuery, setAuthorSearchQuery] = React.useState('');
  const [isDropdownOpen, setIsDropdownOpen] = React.useState(false);
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  const [apiAuthors, setApiAuthors] = React.useState<string[]>([]);
  const [isLoadingAuthors, setIsLoadingAuthors] = React.useState(false);

  // Fetch authors dynamically from Graph API with debounce
  React.useEffect(() => {
    if (!searchService) {
      setApiAuthors([]);
      return;
    }

    setIsLoadingAuthors(true);
    const delayDebounceFn = setTimeout(async () => {
      try {
        const fetched = await searchService.getAuthors(authorSearchQuery);
        setApiAuthors(fetched);
      } catch (e) {
        console.error('Error fetching authors from API:', e);
      } finally {
        setIsLoadingAuthors(false);
      }
    }, 350);

    return () => clearTimeout(delayDebounceFn);
  }, [authorSearchQuery, searchService]);

  // Combine selected authors, API authors, and fallback authors
  const filteredAuthors = React.useMemo(() => {
    const set = new Set<string>();
    
    // Always include currently selected authors so they don't disappear from the UI list
    selectedAuthors.forEach(a => set.add(a));
    
    // Add API fetched authors
    apiAuthors.forEach(a => set.add(a));
    
    // Fallback to offline authors if API isn't present or returned nothing
    if (apiAuthors.length === 0) {
      authorsList.forEach(a => {
        if (!authorSearchQuery || a.toLowerCase().includes(authorSearchQuery.toLowerCase())) {
          set.add(a);
        }
      });
    }

    const uniqueNames = Array.from(set);
    const alphabetic = uniqueNames.filter(name => /^[a-zA-Z]/.test(name.trim().charAt(0)));
    const nonAlphabetic = uniqueNames.filter(name => !/^[a-zA-Z]/.test(name.trim().charAt(0)));
    alphabetic.sort((a, b) => a.localeCompare(b));
    nonAlphabetic.sort((a, b) => a.localeCompare(b));
    return [...alphabetic, ...nonAlphabetic];
  }, [apiAuthors, selectedAuthors, authorsList, authorSearchQuery]);

  const toggleAuthor = (name: string) => {
    const updated = selectedAuthors.includes(name)
      ? selectedAuthors.filter(a => a !== name)
      : [...selectedAuthors, name];
    setFilters({ fileTypes, selectedAuthors: updated, date });
  };

  // Close dropdown when clicking outside
  React.useEffect(() => {
    const handleOutsideClick = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setIsDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleOutsideClick);
    return () => {
      document.removeEventListener('mousedown', handleOutsideClick);
    };
  }, []);

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
                  </label>
                );
              })}
            </div>

            {/* Filter by Author Dropdown */}
            <h3 className={`${styles.sidebarSectionHeader} ${styles.authorDropdownHeader}`}>
              Author <span className={styles.countLabel}>({selectedAuthors.length} selected)</span>
            </h3>
            <div ref={dropdownRef} className={`${styles.authorInputWrapper} ${styles.authorDropdownWrapper}`}>
              <div 
                className={styles.authorDropdownTrigger}
                onClick={() => setIsDropdownOpen(true)}
              >
                <User className={styles.authorIcon} size={16} />
                <input 
                  type="text" 
                  placeholder="Find an author..."
                  value={authorSearchQuery}
                  onFocus={() => setIsDropdownOpen(true)}
                  onChange={(e) => {
                    setAuthorSearchQuery(e.target.value);
                    setIsDropdownOpen(true);
                  }}
                  className={styles.authorInput}
                />
              </div>
              
              {isDropdownOpen && (
                <div className={styles.authorDropdownMenu}>
                   <div className={styles.authorDropdownScroll}>
                    {isLoadingAuthors && (
                      <div className={styles.authorDropdownEmpty}>
                        Searching authors...
                      </div>
                    )}
                    {!isLoadingAuthors && filteredAuthors.map(name => {
                      const isChecked = selectedAuthors.includes(name);
                      return (
                        <div 
                          key={name}
                          onClick={(e) => {
                            e.stopPropagation();
                            toggleAuthor(name);
                          }}
                          className={`${styles.authorItemLabel} ${isChecked ? styles.authorItemLabelChecked : ''}`}
                        >
                          <div className={`${styles.authorItemAvatar} ${isChecked ? styles.authorItemAvatarChecked : ''}`}>
                            {name.charAt(0).toUpperCase()}
                          </div>
                          <span className={`${styles.authorItemName} ${isChecked ? styles.authorItemNameChecked : ''}`}>
                            {name}
                          </span>
                        </div>
                      );
                    })}
                    {!isLoadingAuthors && filteredAuthors.length === 0 && (
                      <div className={styles.authorDropdownEmpty}>
                        No authors found
                      </div>
                    )}
                  </div>
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
                    onClick={() => setFilters({ fileTypes, selectedAuthors, date: isActive ? '' : opt })}
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
                  onChange={(e) => setFilters({ fileTypes, selectedAuthors, date: e.target.value })}
                  className={styles.dateInput}
                />
              </div>
            </div>
          </div>

          {/* Action Row */}
          <div className={styles.sidebarActionWrapper}>
            <button 
              onClick={() => setFilters({ fileTypes, selectedAuthors, date })}
              className={styles.applyButton}
            >
              Apply Filters
            </button>
            <div 
              style={{ textAlign: 'center', fontSize: '10px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginTop: '6px', cursor: 'pointer' }}
              onClick={() => setFilters({ fileTypes: ['All'], selectedAuthors: [], date: '' })}
            >
              Clear All Filters
            </div>
          </div>
        </div>
      </aside>

      <div 
        onMouseDown={startResizingSidebar}
        className={`${styles.resizeHandle} ${isResizingSidebar ? styles.resizeHandleActive : ''}`}
      >
        <div className={styles.resizeGrip}>
          <GripVertical size={12} strokeWidth={2.5} />
        </div>
      </div>
    </>
  );
};
