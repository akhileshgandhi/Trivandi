/* eslint-disable @typescript-eslint/explicit-function-return-type */

/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import * as ReactDOM from "react-dom";
import styles from "./DataTable.module.scss";
import filterIcon from "../../assets/sort.png";
import searchIcon from "../../assets/search-normal.png";
import { ChevronLeft, ChevronRight, ArrowUp, ArrowDown, ArrowUpDown } from "lucide-react";
import Pagination from "../Pagination/Pagination";

/* ===================== TYPES ===================== */

export interface TableColumn {
  key: string;
  label: string;
  filterable?: boolean;
  filterType?: "search";
  sortable?: boolean;
  render?: (value: any, row: any) => React.ReactNode;
  minWidth?: number | string;
}

interface DataTableProps {
  columns: TableColumn[];
  rows: any[];
  // Server side pagination support
  serverSide?: boolean;
  totalItems?: number;
  currentPage?: number;
  pageSize?: number;
  onPageChange?: (page: number) => void;
  // Sorting support
  sortColumn?: string;
  sortAscending?: boolean;
  onSort?: (columnKey: string, ascending: boolean) => void;
  // Provide unique values for dropdowns from parent when server-side
  uniqueValues?: Record<string, string[]>;
  // Controlled filter handlers (server-side)
  onFilterChange?: (filters: Record<string, string>) => void;
  onDateFilterChange?: (dateFilters: Record<string, { from?: string; to?: string }>) => void;
  // Reset filters callback
  onResetFilters?: () => void;
  // Show/hide filter controls
  showFilterControls?: boolean;
  // Active filters from parent (for server-side)
  activeFilters?: Record<string, string>;
  // Row click handler
  onRowClick?: (row: any) => void;
}

/* ===================== COMPONENT ===================== */

const DataTable: React.FC<DataTableProps> = ({
  columns,
  rows,
  serverSide = false,
  totalItems,
  currentPage: propsCurrentPage,
  pageSize: propsPageSize,
  onPageChange,
  uniqueValues,
  onFilterChange,
  onDateFilterChange,
  onResetFilters,
  showFilterControls = false,
  activeFilters: propsActiveFilters,
  onRowClick,
  sortColumn,
  sortAscending,
  onSort,
}) => {
  const [activeFilter, setActiveFilter] = React.useState<string | null>(null);
  const [searchText, setSearchText] = React.useState("");
  const [filters, setFilters] = React.useState<Record<string, string>>(propsActiveFilters || {});
  const [anchorRect, setAnchorRect] = React.useState<DOMRect | null>(null);
  const MAX_FILTERS = 5;

  // Sync local filters with parent filters when server-side
  React.useEffect(() => {
    if (serverSide && propsActiveFilters) {
      setFilters(propsActiveFilters);
    }
  }, [propsActiveFilters, serverSide]);

  const tableScrollRef = React.useRef<HTMLDivElement | null>(null);

  const DROPDOWN_WIDTH = 260;
  const dropdownRef = React.useRef<HTMLDivElement | null>(null);

  /* ================= PAGINATION ================= */

  const ITEMS_PER_PAGE = 12;
  const pageSize = propsPageSize ?? ITEMS_PER_PAGE;

  const [localPage, setLocalPage] = React.useState(1);
  const currentPage = serverSide ? propsCurrentPage ?? 1 : localPage;

  const [dateFilters, setDateFilters] = React.useState<Record<string, { from?: string; to?: string }>>({});
  const isDateColumn = (key: string | null) => !!key && key.toLowerCase().includes("date");

  /* ================= CHECK IF ANY FILTERS ARE ACTIVE ================= */
  const hasActiveFilters = React.useMemo(() => {
    return Object.keys(filters).length > 0 || Object.keys(dateFilters).length > 0;
  }, [filters, dateFilters]);

  const totalActiveFilters = Object.keys(filters).length + Object.keys(dateFilters).length;
  const canAddMoreFilters = totalActiveFilters < MAX_FILTERS;

  /* ================= FILTER LOGIC - MULTI-COLUMN AND LOGIC ================= */

  const filteredRows = React.useMemo(() => {
    return rows.filter((row) => {
      // ✅ ALL text filters must match (AND logic)
      const textMatch = Object.entries(filters).every(([key, value]) =>
        String(row[key] ?? "").toLowerCase().includes(value.toLowerCase())
      );

      // ✅ ALL date filters must match (AND logic)
      const dateMatch = Object.entries(dateFilters).every(([key, range]) => {
        const cellValue = row[key];
        if (!cellValue) return true;

        const cellDate = new Date(cellValue);
        if (range.from && cellDate < new Date(range.from)) return false;
        if (range.to && cellDate > new Date(range.to)) return false;

        return true;
      });

      return textMatch && dateMatch;
    });
  }, [rows, filters, dateFilters]);

  /* reset page when filters change (client-side only) */
  React.useEffect(() => {
    if (!serverSide) setLocalPage(1);
  }, [filters, dateFilters, serverSide]);

  /* pagination */
  const totalPages = serverSide
    ? Math.max(1, Math.ceil((totalItems ?? filteredRows.length) / pageSize))
    : Math.max(1, Math.ceil(filteredRows.length / pageSize));

  const paginatedRows = serverSide
    ? rows
    : filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  /* unique values */
  const getUniqueValues = (key: string | null) =>
    // prefer parent-provided unique values when server-side
    (serverSide && uniqueValues && key ? uniqueValues[key] : undefined) ||
    Array.from(new Set(rows.map((r) => r[key]))).filter(Boolean);

  /* filter icon click */
  const onFilterClick = (e: React.MouseEvent, key: string) => {
    e.stopPropagation();
    const hasFilterForColumn = filters[key] || dateFilters[key];
    if (!hasFilterForColumn && !canAddMoreFilters) {
      alert(`Maximum ${MAX_FILTERS} filters allowed. Please remove a filter first.`);
      return;
    }
    setAnchorRect(e.currentTarget.getBoundingClientRect());
    setActiveFilter((prev) => (prev === key ? null : key));
    setSearchText("");
  };

  /* reset all filters */
  const handleResetFilters = () => {
    setFilters({});
    setDateFilters({});
    if (serverSide) {
      onFilterChange?.({});
      onDateFilterChange?.({});
      onResetFilters?.();
    }
  };

  /* close on outside click */
  React.useEffect(() => {
    const close = () => setActiveFilter(null);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, []);

  /* close on scroll */
  React.useEffect(() => {
    if (!activeFilter) return;

    const closeOnScroll = (e: Event) => {
      if (dropdownRef.current?.contains(e.target as Node)) return;
      setActiveFilter(null);
    };

    window.addEventListener("scroll", closeOnScroll, true);
    return () => window.removeEventListener("scroll", closeOnScroll, true);
  }, [activeFilter]);

  /* render */
  return (
    <>
      {/* Active Filters Display - Always show when filters are active */}
      {hasActiveFilters && (
        <div className={styles.filterControlsSection}>
          <div className={styles.activeFiltersSection}>
            {Object.entries(filters).map(([key, value]) => {
              const column = columns.find(c => c.key === key);
              return (
                <div key={key} className={styles.filterChip}>
                  <span className={styles.chipText}>
                    <strong>{column?.label || key}:</strong> {value}
                  </span>
                  <button
                    className={styles.chipRemove}
                    onClick={() => {
                      const updated = { ...filters };
                      delete updated[key];
                      setFilters(updated);
                      if (serverSide) {
                        onFilterChange?.(updated);
                      }
                    }}
                    title="Remove filter"
                  >
                    ×
                  </button>
                </div>
              );
            })}
            
            {Object.entries(dateFilters).map(([key, range]) => {
              const column = columns.find(c => c.key === key);
              const dateRangeText = range.from && range.to
                ? `${range.from} to ${range.to}`
                : range.from
                ? `From ${range.from}`
                : `To ${range.to}`;
              
              return (
                <div key={key} className={styles.filterChip}>
                  <span className={styles.chipText}>
                    <strong>{column?.label || key}:</strong> {dateRangeText}
                  </span>
                  <button
                    className={styles.chipRemove}
                    onClick={() => {
                      const updated = { ...dateFilters };
                      delete updated[key];
                      setDateFilters(updated);
                      if (serverSide) {
                        onDateFilterChange?.(updated);
                      }
                    }}
                    title="Remove filter"
                  >
                    ×
                  </button>
                </div>
              );
            })}

            <button 
              className={styles.resetFilterBtn} 
              onClick={handleResetFilters}
              type="button"
            >
              <span className={styles.resetIcon}>&#8635;</span>
              Reset
            </button>

            <button 
              className={styles.clearAllFiltersBtn} 
              onClick={handleResetFilters}
              type="button"
            >
              Clear All
            </button>
          </div>
        </div>
      )}

      <div className={styles.tableWrapper}>
        <div className={styles.tableScroll} ref={tableScrollRef}>
          <table className={styles.table}>
            <thead>
              <tr>
                {columns.map((col) => (
                  <th key={col.key} style={col.minWidth ? { minWidth: col.minWidth, maxWidth: col.minWidth } : undefined}>
                    <div className={styles.headerCell}>
                      <span>{col.label}</span>

                      <div className={styles.headerActions}>
                        {col.sortable && (
                          <button
                            type="button"
                            className={styles.sortIconBtn}
                            onClick={() => onSort?.(col.key, sortColumn === col.key ? !sortAscending : true)}
                            title={`Sort by ${col.label}`}
                          >
                            {sortColumn === col.key ? (
                              sortAscending ? <ArrowUp size={14} /> : <ArrowDown size={14} />
                            ) : (
                              <ArrowUpDown size={14} />
                            )}
                          </button>
                        )}
                        {col.filterable && (
                          <button
                            type="button"
                            className={`${styles.filterIconBtn} ${
                              filters[col.key] || dateFilters[col.key] ? styles.filterActive : ''
                            }`}
                            onClick={(e) => onFilterClick(e, col.key)}
                            disabled={!canAddMoreFilters && !(filters[col.key] || dateFilters[col.key])}
                          >
                            <img 
                              src={filterIcon} 
                              className={styles.filterIcon}
                            />
                          </button>
                        )}
                      </div>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>

            <tbody>
              {paginatedRows.length > 0 ? (
                paginatedRows.map((row, i) => (
                  <tr 
                    key={i}
                    onClick={() => onRowClick?.(row)}
                    className={onRowClick ? styles.clickableRow : ''}
                  >
                    {columns.map((col) => (
                      <td key={col.key} style={col.minWidth ? { minWidth: col.minWidth, maxWidth: col.minWidth } : undefined}>
                        {col.render ? col.render(row[col.key], row) : (row[col.key] ?? '-')}
                      </td>
                    ))}
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={columns.length} className={styles.noData}>
                    No results found
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* pagination UI */}
      <Pagination
        currentPage={currentPage}
        totalPages={totalPages}
        pageSize={pageSize}
        totalItems={totalItems ?? filteredRows.length}
        onPageChange={(p) => {
          if (serverSide) onPageChange?.(p);
          else setLocalPage(p);
        }}
        onPageSizeChange={serverSide ? undefined : undefined} // DataTable doesn't support local pageSize change yet in props
      />

      {/* filter dropdown */}
      {activeFilter &&
        anchorRect &&
        (() => {
          const tableRect = tableScrollRef.current?.getBoundingClientRect();

          let left = anchorRect.left + window.scrollX;
          const top = anchorRect.bottom + window.scrollY + 8;

          if (tableRect) {
            const tableLeft = tableRect.left + window.scrollX;
            const tableRight = tableRect.right + window.scrollX;

            if (left + DROPDOWN_WIDTH > tableRight) {
              left = tableRight - DROPDOWN_WIDTH - 8;
            }
            if (left < tableLeft + 8) {
              left = tableLeft + 8;
            }
          }

          return ReactDOM.createPortal(
            <div ref={dropdownRef} className={styles.filterDropdown} style={{ top, left, position: "absolute" }} onClick={(e) => e.stopPropagation()}>
              <div className={styles.searchBox}>
                <input type="text" placeholder="Search " value={searchText} onChange={(e) => setSearchText(e.target.value)} />
                <img src={searchIcon} className={styles.searchIcon} />
              </div>

              <div className={styles.dropdownList}>
                {isDateColumn(activeFilter) ? (
                  <>
                    <label className={styles.dateLabel}>From</label>
                    <input 
                      type="date" 
                      className={styles.dateInput} 
                      value={dateFilters[activeFilter]?.from || ""} 
                      onChange={(e) => {
                        const updated = {
                          ...dateFilters,
                          [activeFilter]: {
                            ...dateFilters[activeFilter],
                            from: e.target.value
                          }
                        };
                        setDateFilters(updated);
                      }}
                    />

                    <label className={styles.dateLabel}>To</label>
                    <input 
                      type="date" 
                      className={styles.dateInput} 
                      value={dateFilters[activeFilter]?.to || ""} 
                      onChange={(e) => {
                        const updated = {
                          ...dateFilters,
                          [activeFilter]: {
                            ...dateFilters[activeFilter],
                            to: e.target.value
                          }
                        };
                        setDateFilters(updated);
                      }}
                    />

                    <div className={styles.dateActions}>
                      <button
                        className={styles.applyBtn}
                        onClick={() => {
                          if (serverSide) {
                            onDateFilterChange?.(dateFilters);
                          }
                          setActiveFilter(null);
                        }}
                      >
                        Apply
                      </button>

                      <button
                        className={styles.clearBtn}
                        onClick={() => {
                          const updated = { ...dateFilters };
                          if (activeFilter) delete updated[activeFilter];
                          setDateFilters(updated);
                          if (serverSide) {
                            onDateFilterChange?.(updated);
                          }
                          setActiveFilter(null);
                        }}
                      >
                        Clear
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    <div 
                      className={styles.dropdownItem} 
                      onClick={() => {
                        // Remove filter for this column (keep other columns' filters)
                        const updated = { ...filters };
                        if (activeFilter) delete updated[activeFilter];
                        setFilters(updated);
                        
                        if (serverSide) {
                          onFilterChange?.(updated);
                        }
                        setActiveFilter(null);
                      }}
                    >
                      All
                    </div>

                    {getUniqueValues(activeFilter)
                      .filter((v) => String(v).toLowerCase().includes(searchText.toLowerCase()))
                      .map((v) => (
                        <div 
                          key={v} 
                          className={styles.dropdownItem} 
                          onClick={() => {
                            // ✅ KEY FIX: Add to existing filters, don't replace
                            const updated = {
                              ...filters,  // Keep existing filters
                              [activeFilter as string]: String(v)  // Add/update this column's filter
                            };
                            setFilters(updated);
                            
                            
                            if (serverSide) {
                              onFilterChange?.(updated);
                            }
                            setActiveFilter(null);
                          }}
                        >
                          {v}
                        </div>
                      ))}
                  </>
                )}
              </div>
            </div>,
            document.body
          );
        })()}
    </>
  );
};

export default DataTable;