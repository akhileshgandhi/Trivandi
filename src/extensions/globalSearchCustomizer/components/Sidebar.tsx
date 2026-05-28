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
  authorsList
}) => {
  const [authorSearchQuery, setAuthorSearchQuery] = React.useState('');
  const [isDropdownOpen, setIsDropdownOpen] = React.useState(false);
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  // Filter authors alphabetically matching the search query
  const filteredAuthors = React.useMemo(() => {
    return authorsList
      .filter(name => name.toLowerCase().includes(authorSearchQuery.toLowerCase()))
      .sort((a, b) => a.localeCompare(b));
  }, [authorsList, authorSearchQuery]);

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
                    {type === 'All' && <span className={styles.checkboxBadge}>1.3M</span>}
                  </label>
                );
              })}
            </div>

            {/* Filter by Author Dropdown */}
            <h3 className={styles.sidebarSectionHeader} style={{ marginTop: '24px' }}>
              Author <span className={styles.countLabel}>({selectedAuthors.length} selected)</span>
            </h3>
            <div ref={dropdownRef} className={styles.authorInputWrapper} style={{ position: 'relative', marginBottom: '16px' }}>
              <div 
                style={{ position: 'relative', cursor: 'pointer' }}
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
                <div 
                  style={{ 
                    position: 'absolute',
                    top: '100%',
                    left: 0,
                    right: 0,
                    zIndex: 100,
                    marginTop: '6px',
                    maxHeight: '230px', 
                    display: 'flex', 
                    flexDirection: 'column', 
                    padding: '10px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '12px',
                    background: '#ffffff',
                    boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.1)',
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '2px 4px 8px 4px', borderBottom: '1px solid #f1f5f9', marginBottom: '8px' }}>
                    <span style={{ fontSize: '11px', color: '#94a3b8', fontWeight: 600 }}>SELECT AUTHORS</span>
                    <button 
                      onClick={(e) => {
                        e.stopPropagation();
                        setIsDropdownOpen(false);
                      }}
                      style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: '11px', cursor: 'pointer', fontWeight: 700 }}
                    >
                      Done
                    </button>
                  </div>
                  
                  <div 
                    style={{ 
                      display: 'flex', 
                      flexDirection: 'column', 
                      gap: '5px', 
                      overflowY: 'auto', 
                      maxHeight: '170px',
                      paddingRight: '2px'
                    }}
                  >
                    {filteredAuthors.map(name => {
                      const isChecked = selectedAuthors.includes(name);
                      return (
                        <label 
                          key={name}
                          onClick={(e) => e.stopPropagation()} // Prevent closing dropdown on item click
                          style={{ 
                            display: 'flex', 
                            alignItems: 'center', 
                            gap: '8px', 
                            cursor: 'pointer', 
                            padding: '6px 8px', 
                            borderRadius: '6px', 
                            backgroundColor: isChecked ? '#eff6ff' : 'transparent',
                            transition: 'all 0.15s ease',
                            userSelect: 'none'
                          }}
                        >
                          <input 
                            type="checkbox" 
                            checked={isChecked}
                            onChange={() => toggleAuthor(name)}
                            style={{
                              width: '14px',
                              height: '14px',
                              accentColor: '#2563eb',
                              cursor: 'pointer'
                            }}
                          />
                          <div 
                            style={{ 
                              width: '24px', 
                              height: '24px', 
                              borderRadius: '50%', 
                              backgroundColor: isChecked ? '#2563eb' : '#64748b', 
                              color: '#ffffff', 
                              display: 'flex', 
                              alignItems: 'center', 
                              justifyContent: 'center', 
                              fontSize: '11px', 
                              fontWeight: 700 
                            }}
                          >
                            {name.charAt(0).toUpperCase()}
                          </div>
                          <span 
                            style={{ 
                              fontSize: '13px', 
                              color: isChecked ? '#1e40af' : '#334155', 
                              fontWeight: isChecked ? 600 : 400,
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              whiteSpace: 'nowrap'
                            }}
                          >
                            {name}
                          </span>
                        </label>
                      );
                    })}
                    {filteredAuthors.length === 0 && (
                      <div style={{ padding: '16px', fontSize: '12px', color: '#94a3b8', fontStyle: 'italic', textAlign: 'center' }}>
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
              className={styles.applyButton}
              onClick={() => {}}
            >
              Apply Filter
            </button>
            <button 
              onClick={() => setFilters({ fileTypes: ['All'], selectedAuthors: [], date: '' })}
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
      >
        <div className={styles.resizeGrip}>
          <GripVertical size={12} strokeWidth={2.5} />
        </div>
      </div>
    </>
  );
};
