import * as React from "react";
import styles from "./TrackerTable.module.scss";
import DataTable, { TableColumn } from "../DataTable/DataTable";
import filterIcon from "../../assets/sort.png";

interface TrackerTableProps {
  title: string;
  tabs: string[];
  activeTab: string;
  onTabChange: (tab: string) => void;
  columns: TableColumn[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  rows: any[];
  showDivider?: boolean;
  showFilterIcon?: boolean;
  onFilterClick?: () => void;
  isFilterActive?: boolean;
  onResetFilters?: () => void;
  activeFilters?: Record<string, string>;
  children?: React.ReactNode;
  onSettingsClick?: () => void;
  onCreateClick?: () => void;
  // server-side passthrough
  serverSide?: boolean;
  totalItems?: number;
  currentPage?: number;
  pageSize?: number;
  onPageChange?: (page: number) => void;
  onPageSizeChange?: (pageSize: number) => void;
  uniqueValues?: Record<string, string[]>;
  onFilterChange?: (filters: Record<string, string>) => void;
  onDateFilterChange?: (dateFilters: Record<string, { from?: string; to?: string }>) => void;
  showFilterControls?: boolean;
  // Row click handler
  onRowClick?: (row: any) => void;
  // Tab counts
  tabCounts?: Record<string, number>;
  // Sorting
  sortColumn?: string;
  sortAscending?: boolean;
  onSort?: (columnKey: string, ascending: boolean) => void;
  // Search
  searchTerm?: string;
  onSearch?: (value: string) => void;
  searchPlaceholder?: string;
}

const TrackerTable: React.FC<TrackerTableProps> = ({
  title,
  tabs,
  activeTab,
  onTabChange,
  columns,
  rows,
  children,
  showDivider = false,
  showFilterIcon = true,
  isFilterActive,
  onFilterClick,
  onResetFilters,
  activeFilters,
  onSettingsClick,
  onCreateClick,
  serverSide,
  totalItems,
  currentPage,
  pageSize,
  onPageChange,
  onPageSizeChange,
  uniqueValues,
  onFilterChange,
  onDateFilterChange,
  showFilterControls = false,
  onRowClick,
  tabCounts,
  sortColumn,
  sortAscending,
  onSort,
  searchTerm,
  onSearch,
  searchPlaceholder = "Search Projects...",
}) => {
  return (
    <div className={styles.trackerWrapper}>
      {/* TITLE */}
      <div className={styles.title}>{title}</div>

      <div className={styles.outerContainer}>
        {/* CONTROLS ROW */}
        <div
          className={`${styles.controls} ${showDivider ? styles.withDivider : ""
            }`}
        >
          {/* TABS */}
          <div className={styles.tabs}>
            {tabs.map((tab) => (
              <button type="button"
                key={tab}
                className={`${styles.tab} ${activeTab === tab ? styles.activeTab : ""
                  }`}
                onClick={() => onTabChange(tab)}
              >
                {tab}
                {tabCounts && tabCounts[tab] !== undefined && (
                  <span className={styles.tabCount}>({tabCounts[tab]})</span>
                )}
              </button>
            ))}
          </div>

          <div className={styles.searchBox}>
            <input
              type="text"
              placeholder={searchPlaceholder}
              value={searchTerm || ""}
              onChange={(e) => onSearch?.(e.target.value)}
              className={styles.searchInput}
            />
            <svg
              className={styles.searchIcon}
              width="18"
              height="18"
              viewBox="0 0 24 24"
              fill="none"
              xmlns="http://www.w3.org/2000/svg"
            >
              <path
                d="M11 19C15.4183 19 19 15.4183 19 11C19 6.58172 15.4183 3 11 3C6.58172 3 3 6.58172 3 11C3 15.4183 6.58172 19 11 19Z"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
              <path
                d="M21 21L16.65 16.65"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
            {searchTerm && (
              <button
                type="button"
                className={styles.clearSearchBtn}
                onClick={() => onSearch?.("")}
                title="Clear search"
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M18 6L6 18M6 6l12 12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
              </button>
            )}
          </div>

          <div className={styles.reset}>
            {/* RESET FILTERS BUTTON */}
            {/* {onResetFilters && activeFilters && Object.keys(activeFilters).length > 0 && (
              <button type="button"
                className={styles.resetBtn}
                onClick={onResetFilters}
                aria-label="Reset filters"
              >
                Reset
              </button>
            )} */}

            {/* FILTER TOGGLE BUTTON */}
            {/* {onFilterClick && (
              <button type="button"
                className={`${styles.filterBtn} ${isFilterActive ? styles.filterBtnActive : ""
                  }`}
                onClick={onFilterClick}
                aria-label="Toggle filters"
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M3 6h18M7 12h10M11 18h2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                </svg>
              </button>
            )} */}

            {/* CREATE PROJECT BUTTON */}
            {onCreateClick && (
              <button type="button"
                className={styles.createBtn}
                onClick={onCreateClick}
                aria-label="Add Non-CMAP Project"
              >
                Add Non-CMAP Project
              </button>
            )}

            {/* SETTINGS ICON */}
            {showFilterIcon && onSettingsClick && activeTab !== "Non-CMAP" && (
              <button type="button"
                className={styles.filterBtn}
                onClick={onSettingsClick}
                aria-label="Customize columns"
              >
                <img src={filterIcon} alt="Customize Columns" />
              </button>
            )}
          </div>
        </div>

        {/* FILTER ROW (CHILDREN) */}
        {children && <div className={styles.filtersRow}>{children}</div>}

        {/* TABLE */}
        <DataTable
          columns={columns}
          rows={rows}
          serverSide={serverSide}
          totalItems={totalItems}
          currentPage={currentPage}
          pageSize={pageSize}
          onPageChange={onPageChange}
          uniqueValues={uniqueValues}
          onFilterChange={onFilterChange}
          onDateFilterChange={onDateFilterChange}
          showFilterControls={showFilterControls}
          activeFilters={activeFilters}
          onRowClick={onRowClick}
          sortColumn={sortColumn}
          sortAscending={sortAscending}
          onSort={onSort}
        />
      </div>
    </div>
  );
};

export default TrackerTable;