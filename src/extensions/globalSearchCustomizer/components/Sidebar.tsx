import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { User, Calendar, GripVertical, ChevronLeft, ChevronRight } from 'lucide-react';
import styles from '../../../styles/PremiumSearch.module.scss';
import { IFiltersPanelProps } from '../interface/IFiltersPanelProps';
import { loadAdminConfig } from '../interface/IAdminPanelProps';

const getSiteId = (label: string): string => {
  const mapping: { [key: string]: string } = {
    'Intranet': 'TrivandiHub',
    'People': 'PeopleHub',
    'Company': 'CompanyHub',
    'Marketing': 'BrandingMarketing',
    'Projects': 'Projects',
    'Trivandi London': 'TrivandiLondon',
    'TDMCC': 'TDMCC',
    'Trivandi USA': 'TrivandiUSA',
    'Trivandi Australia': 'TrivandiAustralia',
    'Trivandi KSA': 'TrivandiKSA',
  };
  return mapping[label] || label.replace(/\s+/g, '');
};

// const SITES = [
//   { label: 'Operations Hub', siteId: 'OperationsHub' },
//   { label: 'Freudiger', siteId: 'Freudiger' },
//   { label: 'MoreYeahs Departments DMS', siteId: 'moreYeahsdepartmentsDMS' },
//   { label: 'Pembe Portal', siteId: 'PembePortal' }
// ];

export const Sidebar: React.FC<IFiltersPanelProps> = ({
  sidebarWidth,
  fileTypes,
  selectedAuthors,
  selectedSites,
  selectedProjects,
  date,
  setFilters,
  toggleFileType,
  startResizingSidebar,
  isResizingSidebar,
  authorsList,
  projectsList,
  searchService,
  adminConfig: propAdminConfig
}) => {
  const adminConfig = propAdminConfig || loadAdminConfig();
  const FILE_TYPE_OPTIONS = React.useMemo(() => ['All', ...adminConfig.fileTypes], [adminConfig.fileTypes]);
  const DATE_OPTIONS = adminConfig.dateFilters;
  const SITES = React.useMemo(() => adminConfig.sites.map(s => ({
    label: s,
    siteId: getSiteId(s)
  })), [adminConfig.sites]);

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
    setFilters({ fileTypes, selectedAuthors: updated, selectedSites, selectedProjects, date });
  };

  const [startDate, endDate] = React.useMemo(() => {
    if (date && !DATE_OPTIONS.includes(date) && date.includes('_')) {
      const [s, e] = date.split('_');
      const startObj = s ? new Date(s + 'T00:00:00') : null;
      const endObj = e ? new Date(e + 'T00:00:00') : null;
      return [
        startObj && !isNaN(startObj.getTime()) ? startObj : null,
        endObj && !isNaN(endObj.getTime()) ? endObj : null
      ];
    }
    return [null, null];
  }, [date]);

  const handleDateRangeChange = (update: [Date | null, Date | null]) => {
    const [start, end] = update;
    const formatDate = (d: Date | null) => {
      if (!d) return '';
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };
    const startStr = formatDate(start);
    const endStr = formatDate(end);
    setFilters({
      fileTypes,
      selectedAuthors,
      selectedSites,
      selectedProjects,
      date: startStr || endStr ? `${startStr}_${endStr}` : ''
    });
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
            
            <div className={styles.fileTypeGrid}>
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

            <h3 className={styles.sidebarTitle} style={{ marginTop: '16px' }}>Sites</h3>
            <div className={styles.sidebarOptionList}>
              {SITES.map(site => {
                const isActive = (selectedSites || []).includes(site.siteId);
                return (
                  <label 
                    key={site.siteId} 
                    className={`${styles.checkboxLabel} ${isActive ? styles.checkboxActive : ''}`}
                  >
                    <input 
                      type="checkbox" 
                      checked={isActive}
                      onChange={() => {
                        const current = selectedSites || [];
                        const updated = current.includes(site.siteId)
                          ? current.filter(s => s !== site.siteId)
                          : [...current, site.siteId];
                        setFilters({ fileTypes, selectedAuthors, selectedSites: updated, selectedProjects, date });
                      }}
                      className={styles.checkboxInput}
                    />
                    <span>{site.label}</span>
                  </label>
                );
              })}
            </div>

            {/* Project Filter */}
            {projectsList.length > 0 && (
              <>
                <h3 className={`${styles.sidebarSectionHeader}`} style={{ marginTop: '16px' }}>
                  Projects <span className={styles.countLabel}>({selectedProjects.length} selected)</span>
                </h3>
                <div className={styles.sidebarOptionList}>
                  {projectsList.map(project => {
                    const isActive = selectedProjects.includes(project);
                    return (
                      <label 
                        key={project} 
                        className={`${styles.checkboxLabel} ${isActive ? styles.checkboxActive : ''}`}
                      >
                        <input 
                          type="checkbox" 
                          checked={isActive}
                          onChange={() => {
                            const updated = selectedProjects.includes(project)
                              ? selectedProjects.filter(p => p !== project)
                              : [...selectedProjects, project];
                            setFilters({ fileTypes, selectedAuthors, selectedSites, selectedProjects: updated, date });
                          }}
                          className={styles.checkboxInput}
                        />
                        <span>{project}</span>
                      </label>
                    );
                  })}
                </div>
              </>
            )}

            {/* Filter by Author Dropdown */}
            <h3 className={`${styles.sidebarSectionHeader} ${styles.authorDropdownHeader}`}>
              Author <span className={styles.countLabel}>({selectedAuthors.length} selected)</span>
            </h3>
            <div ref={dropdownRef} className={`${styles.authorInputWrapper} ${styles.authorDropdownWrapper}`}>
              <div 
                className={styles.authorDropdownTrigger}
                onClick={() => setIsDropdownOpen(true)}
              >
                <div className={styles.authorIcon}>
                  <User size={18} strokeWidth={2.5} />
                </div>
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
              <CustomDateRangePicker
                startDate={startDate}
                endDate={endDate}
                onChange={handleDateRangeChange}
              />

              <div className={styles.datePickerLabel} style={{ marginTop: '12px' }}>Or choose a preset</div>
              <div className={styles.dateFilterGrid}>
                {DATE_OPTIONS.map(opt => {
                  const isActive = date === opt;
                  return (
                    <button 
                      key={opt}
                      onClick={() => setFilters({ fileTypes, selectedAuthors, selectedSites, selectedProjects, date: isActive ? '' : opt })}
                      className={`${styles.dateFilterButton} ${isActive ? styles.dateFilterActive : ''}`}
                    >
                      {opt}
                    </button>
                  );
                })}
              </div>
            </div>
          </div>

          {/* Action Row */}
          <div className={styles.sidebarActionWrapper}>
            <button 
              onClick={() => setFilters({ fileTypes, selectedAuthors, selectedSites, selectedProjects, date })}
              className={styles.applyButton}
            >
              Apply Filters
            </button>
            <div 
              style={{ textAlign: 'center', fontSize: '10px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase', letterSpacing: '0.05em', marginTop: '6px', cursor: 'pointer' }}
              onClick={() => setFilters({ fileTypes: ['All'], selectedAuthors: [], selectedSites: [], selectedProjects: [], date: '' })}
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

const CustomDateRangePicker: React.FC<{
  startDate: Date | null;
  endDate: Date | null;
  onChange: (update: [Date | null, Date | null]) => void;
}> = ({ startDate, endDate, onChange }) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const [currentMonth, setCurrentMonth] = React.useState(new Date());
  const containerRef = React.useRef<HTMLDivElement>(null);
  const boxRef = React.useRef<HTMLDivElement>(null);
  const [coords, setCoords] = React.useState({ top: 0, left: 0 });

  React.useEffect(() => {
    const handleOutside = (e: MouseEvent) => {
      // If clicking inside the calendar popup portal, don't close
      const popupEl = document.getElementById('trivandi-calendar-portal-popup');
      if (
        containerRef.current && containerRef.current.contains(e.target as Node) ||
        popupEl && popupEl.contains(e.target as Node)
      ) {
        return;
      }
      setIsOpen(false);
    };
    document.addEventListener('mousedown', handleOutside);
    return () => document.removeEventListener('mousedown', handleOutside);
  }, []);

  React.useEffect(() => {
    const updatePosition = () => {
      if (boxRef.current) {
        const rect = boxRef.current.getBoundingClientRect();
        setCoords({
          top: rect.bottom + window.pageYOffset,
          left: rect.left + window.pageXOffset
        });
      }
    };

    if (isOpen) {
      updatePosition();
      // Listen on capture phase to capture scroll inside any scrollable container
      window.addEventListener('scroll', updatePosition, true);
      window.addEventListener('resize', updatePosition);
    }

    return () => {
      window.removeEventListener('scroll', updatePosition, true);
      window.removeEventListener('resize', updatePosition);
    };
  }, [isOpen]);

  const displayValue = React.useMemo(() => {
    if (!startDate) return '';
    const format = (d: Date) => {
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      const y = d.getFullYear();
      return `${m}/${day}/${y}`;
    };
    if (!endDate) return `${format(startDate)} – MM/DD/YYYY`;
    return `${format(startDate)} – ${format(endDate)}`;
  }, [startDate, endDate]);

  const daysInMonth = (year: number, month: number) => new Date(year, month + 1, 0).getDate();
  const startDayOfWeek = (year: number, month: number) => new Date(year, month, 1).getDay();

  const handleDayClick = (clickedDate: Date) => {
    if (!startDate || (startDate && endDate)) {
      onChange([clickedDate, null]);
    } else if (startDate && !endDate) {
      if (clickedDate < startDate) {
        onChange([clickedDate, null]);
      } else {
        onChange([startDate, clickedDate]);
        setIsOpen(false);
      }
    }
  };

  const nextMonth = () => {
    setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1));
  };
  const prevMonth = () => {
    setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1));
  };

  const renderMonthCalendar = (monthDate: Date, showPrevBtn: boolean, showNextBtn: boolean) => {
    const year = monthDate.getFullYear();
    const month = monthDate.getMonth();
    const totalDays = daysInMonth(year, month);
    const startDay = startDayOfWeek(year, month);
    const cells = [];

    for (let i = 0; i < startDay; i++) {
      cells.push(<div key={`empty-${i}`} className={styles.calendarEmptyCell} />);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    for (let day = 1; day <= totalDays; day++) {
      const dateObj = new Date(year, month, day);
      const isFuture = dateObj > today;

      const isStart = startDate && dateObj.toDateString() === startDate.toDateString();
      const isEnd = endDate && dateObj.toDateString() === endDate.toDateString();
      const isInRange = startDate && endDate && dateObj > startDate && dateObj < endDate;

      let cellClass = styles.calendarDayCell;
      if (isFuture) {
        cellClass += ` ${styles.calendarDayDisabled}`;
      } else if (isStart || isEnd) {
        cellClass += ` ${styles.calendarDaySelected}`;
      } else if (isInRange) {
        cellClass += ` ${styles.calendarDayInRange}`;
      }

      cells.push(
        <button
          key={`day-${day}`}
          type="button"
          disabled={isFuture}
          onClick={() => !isFuture && handleDayClick(dateObj)}
          className={cellClass}
        >
          {day}
        </button>
      );
    }

    return (
      <div className={styles.calendarSingleMonth}>
        <div className={styles.calendarHeader}>
          {showPrevBtn ? (
            <button type="button" onClick={prevMonth} className={styles.calendarNavBtn}>
              <ChevronLeft size={16} />
            </button>
          ) : (
            <div style={{ width: 24 }} />
          )}
          <span className={styles.calendarMonthTitle}>
            {monthDate.toLocaleString('default', { month: 'long', year: 'numeric' })}
          </span>
          {showNextBtn ? (
            <button type="button" onClick={nextMonth} className={styles.calendarNavBtn}>
              <ChevronRight size={16} />
            </button>
          ) : (
            <div style={{ width: 24 }} />
          )}
        </div>

        <div className={styles.calendarWeekdays}>
          {['S', 'M', 'T', 'W', 'T', 'F', 'S'].map(w => (
            <span key={w} className={styles.calendarWeekdayLabel}>{w}</span>
          ))}
        </div>

        <div className={styles.calendarGrid}>
          {cells}
        </div>
      </div>
    );
  };

  const rightMonth = currentMonth;
  const leftMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1);

  const todayDate = new Date();
  const isRightMonthCurrentOrFuture = rightMonth.getFullYear() > todayDate.getFullYear() || 
    (rightMonth.getFullYear() === todayDate.getFullYear() && rightMonth.getMonth() >= todayDate.getMonth());

  return (
    <div className={styles.customDatePickerContainer} ref={containerRef}>
      <div className={styles.dateRangePickerBox} ref={boxRef} onClick={() => setIsOpen(!isOpen)}>
        <input
          type="text"
          readOnly
          placeholder="Select date range..."
          value={displayValue}
          className={styles.dateRangeInput}
        />
        <Calendar size={18} strokeWidth={2.5} className={styles.dateRangeCalendarIcon} />
      </div>

      {isOpen && ReactDOM.createPortal(
        <div 
          id="trivandi-calendar-portal-popup"
          className={styles.calendarPopup}
          style={{
            position: 'absolute',
            top: coords.top + 8,
            left: coords.left,
            zIndex: 99999
          }}
        >
          <div className={styles.calendarMonthsContainer}>
            {renderMonthCalendar(leftMonth, true, false)}
            {renderMonthCalendar(rightMonth, false, !isRightMonthCurrentOrFuture)}
          </div>
          
          {(startDate || endDate) && (
            <div className={styles.calendarFooter}>
              <button 
                type="button" 
                onClick={() => { onChange([null, null]); setIsOpen(false); }}
                className={styles.calendarClearBtn}
              >
                Clear Range
              </button>
            </div>
          )}
        </div>,
        document.body
      )}
    </div>
  );
};

