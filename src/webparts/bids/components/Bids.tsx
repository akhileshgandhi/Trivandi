/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { useState, useEffect } from "react";
import type { IBidsProps } from "./IBidsProps";

import Hero from "../../../shared/component/Hero/Hero";
import heroBg from "../../../shared/assets/hero_new.jpg";

import TrackerTable from "../../../shared/component/TrackerTable/TrackerTable";
import { TableColumn } from "../../../shared/component/DataTable/DataTable";
import GlobalLoader from "../../../shared/component/GlobalLoader";
import {
  getProjectsPage,
  getProjectColumns,
} from "../../../shared/services/projectService";

import styles from "./Bids.module.scss";
import "../../../shared/globalcss/globalcss.scss";

/* ===================== DEFAULT UI COLUMNS ===================== */
const DEFAULT_COLUMNS: TableColumn[] = [
  {
    key: "Title",
    label: "Project Title",
    minWidth: 350,
    filterable: true,
    render: (value: any, row: any) => (
      <span className={styles.pipelineTitle}>
        {row.Code ? `${row.Code} - ` : ""}{value ?? "-"}
      </span>
    ),
    sortable: true,
  },
  
  { key: "Company", label: "Company", minWidth: 200, filterable: true, sortable: true },
  { key: "Owner", label: "Owner", filterable: true, sortable: true },
  { key: "BusinessUnit", label: "Business Unit", filterable: true, sortable: true },
  { key: "Sector", label: "Sector", filterable: true, sortable: true },
  { key: "ContractedEntity", label: "Contracted Entity", filterable: true, sortable: true },
  { key: "Country", label: "Country", filterable: true, sortable: true },
  { key: "EndDate", label: "End Date", filterable: true, sortable: true },
];

const Bids: React.FC<IBidsProps> = (props) => {
  const sanitizeColumnKeys = (keys: string[]): string[] => {
    const uniqueKeys = Array.from(new Set(keys.filter(Boolean)));
    return uniqueKeys.filter((k) => k !== "ProjectID");
  };

  const [activeTab, setActiveTab] = useState("Bid Documents");
  const [isLoading, setIsLoading] = useState(true);
  const [rows, setRows] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [totalItems, setTotalItems] = useState(0);
  const PAGE_SIZE = 12;
  const [serverFilters, setServerFilters] = useState<Record<string, string>>({});
  const [uniqueValues, setUniqueValues] = useState<Record<string, string[]>>({});
  const [tabCounts, setTabCounts] = useState<Record<string, number>>({});
  const [sortColumn, setSortColumn] = useState<string>("Title");
  const [sortAscending, setSortAscending] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>("");

  const COLUMN_STORAGE_KEY = `bids_column_order_${props.userDisplayName}`;
  const COLUMN_VERSION_KEY = `bids_column_version_${props.userDisplayName}`;
  const CURRENT_COLUMN_VERSION = '7'; // Incremented after hiding ProjectID column

  const [allColumns, setAllColumns] = useState<TableColumn[]>(DEFAULT_COLUMNS);

  const [orderedColumnKeys, setOrderedColumnKeys] = useState<string[]>(() => {
    const savedVersion = localStorage.getItem(COLUMN_VERSION_KEY);
    const saved = localStorage.getItem(COLUMN_STORAGE_KEY);

    // If version mismatch or no version, use defaults
    if (savedVersion !== CURRENT_COLUMN_VERSION || !saved) {
      localStorage.setItem(COLUMN_VERSION_KEY, CURRENT_COLUMN_VERSION);
      return sanitizeColumnKeys(DEFAULT_COLUMNS.map((c) => c.key));
    }

    try {
      const parsed = JSON.parse(saved) as string[];
      return sanitizeColumnKeys(parsed);
    } catch {
      return sanitizeColumnKeys(DEFAULT_COLUMNS.map((c) => c.key));
    }
  });

  /* ✅ TEMP STATE FOR POPUP */
  const [tempColumnKeys, setTempColumnKeys] =
    useState<string[]>(orderedColumnKeys);

  const [showColumnPopup, setShowColumnPopup] = useState(false);
  const [columnSearchQuery, setColumnSearchQuery] = useState("");

  const dragFromIndex = React.useRef<number | null>(null);
  const rowRefs = React.useRef<HTMLDivElement[]>([]);

  // Map custom column keys to SharePoint column names
  const columnKeyToSharePointName: Record<string, string> = {
    Title: "Title",
    Company: "Company",
    Budget: "TotalProjectValue",
    Owner: "Owner",
    Office: "Office",
    ContractedEntity: "ContractedEntity",
    Status: "Status",
  };

  const getSortValue = (item: any, columnKey: string): string | number => {
    // Project Title should sort by Code Name first, then Title.
    if (columnKey === "Title") {
      const code = String(item?.Code ?? "").trim().toLowerCase();
      const title = String(item?.Title ?? "").trim().toLowerCase();
      return `${code} ${title}`;
    }

    const value = item?.[columnKey];
    if (typeof value === "number") return value;
    if (value instanceof Date) return value.getTime();
    return String(value ?? "").toLowerCase();
  };

  /* ===================== LOAD BIDS ===================== */
  const loadBids = async (
    page = 1,
    filters: Record<string, string> = {},
    sTerm?: string,
    silent: boolean = false,
    sCol?: string,
    sAsc?: boolean
  ) => {
    if (!silent) setIsLoading(true);
    try {
      const termToUse = sTerm ?? searchTerm;
      console.log(`📋 loadBids - Tab: ${activeTab}, Page: ${page}`);

      let allItems: any[] = [];

      if (activeTab === 'Bid Documents') {
        // Bid Documents tab - Only Potential (removed Lead status)
        const statuses = ['Potential'];
        for (const statuss of statuses) {
          try {
            const { items } = await getProjectsPage(statuss, 1, 5000, filters);
            console.log(`[Bid Documents] Fetched items:`, items?.length || 0);
            allItems.push(...items);
          } catch (err) {
            console.error(`Error fetching Bid Documents:`, err);
          }
        }
      } else {
        // Lost tab - Fetch ALL statuses and filter client-side by Status field
        // const allStatuses = ['Lead', 'Potential', 'Project', 'Closed'];
        // const lostStatusValues = ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];
        const lostStatusValues = ['Dead'];
        console.log(`[Lost] Using statuses:`, lostStatusValues);

        for (const status of lostStatusValues) {
          try {
            const { items } = await getProjectsPage(status, 1, 5000, filters);
            console.log(`[Lost] Fetched ${status} items:`, items?.length || 0);
            allItems.push(...items);
          } catch (err) {
            console.error(`Error fetching Lost/${status}:`, err);
          }
        }

        // Apply filters after fetching
        if (Object.keys(filters).length > 0) {
          allItems = allItems.filter(item => {
            return Object.entries(filters).every(([key, value]) => {
              const itemValue = String(item[key] ?? '').toLowerCase();
              return itemValue.includes(value.toLowerCase());
            });
          });
        }
      }

      // Apply global search
      if (termToUse && termToUse.trim()) {
        const lowerTerm = termToUse.toLowerCase();
        allItems = allItems.filter(item => {
          return (
            String(item.Title ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Company ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Code ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Owner ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Sector ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Country ?? "").toLowerCase().includes(lowerTerm)
          );
        });
      }

      // Remove duplicates
      const uniqueItems = allItems.filter((item, index, self) =>
        index === self.findIndex(i => i.ID === item.ID)
      );
      
      console.log(`📊 Total unique items after all filters:`, uniqueItems.length);

      // Apply sort before pagination so paging stays consistent
      const sortKey = sCol ?? sortColumn;
      const sortAsc = sAsc ?? sortAscending;
      const sortedItems = [...uniqueItems].sort((a, b) => {
        const valA = getSortValue(a, sortKey);
        const valB = getSortValue(b, sortKey);

        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
      });

      // Apply pagination
      const startIndex = (page - 1) * PAGE_SIZE;
      const endIndex = startIndex + PAGE_SIZE;
      const paginatedItems = sortedItems
        .slice(startIndex, endIndex)
        .map((item) => ({
          ...item,
          ProjectID: item.ProjectID ?? item.ProjectId ?? "",
        }));

      setRows(paginatedItems);
      setTotalItems(uniqueItems.length);
      setCurrentPage(page);
    } catch (err) {
      console.error('❌ loadBids error:', err);
      setRows([]);
      setTotalItems(0);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSort = (columnKey: string, ascending: boolean) => {
    setSortColumn(columnKey);
    setSortAscending(ascending);
    setCurrentPage(1);
    void loadBids(1, serverFilters, searchTerm, true, columnKey, ascending);
  };

  const handleSearch = (value: string) => {
    setSearchTerm(value);
    setCurrentPage(1);
    void loadBids(1, serverFilters, value, true);
  };

  const loadUniqueValues = async (tab?: string) => {
    const currentTab = tab || activeTab;
    const statuses = currentTab === 'Bid Documents'
      ? ['Potential']
      : ['Dead'];
      // : ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];

    try {
      // Fetch ALL items for this tab using the exact same filters as the table
      let allItems: any[] = [];
      for (const status of statuses) {
        try {
          const { items } = await getProjectsPage(status, 1, 5000, {});
          allItems = allItems.concat(items);
        } catch (err) {

        }
      }

      // Remove duplicates
      const uniqueItems = allItems.filter((item, index, self) =>
        index === self.findIndex(i => i.ID === item.ID)
      );

      const filterableKeys = allColumns.filter((c) => c.filterable).map((c) => c.key);
      const map: Record<string, string[]> = {};

      filterableKeys.forEach((k) => {
        if (k === 'Status') {
          map[k] = statuses;
        } else if (k === "Title") {
          // For Title, combine Code + Title so the dropdown shows e.g. "25048b - Extension..."
          const values = Array.from(
            new Set(
              uniqueItems
                .map((item: any) => {
                  const title = String(item["Title"] ?? "").trim();
                  if (!title) return "";
                  const code = String(item["Code"] ?? "").trim();
                  return code ? `${code} - ${title}` : title;
                })
                .filter((v: string) => v !== "")
            )
          ).sort() as string[];
          map[k] = values;
        } else {
          const values = Array.from(
            new Set(uniqueItems.map((item: any) => String(item[k] ?? "")).filter((v: string) => v !== ""))
          ).sort() as string[];
          map[k] = values;
        }
      });

      setUniqueValues(map);
    } catch (err) {

    }
  };

  const loadTabCounts = async () => {
    const counts: Record<string, number> = {};
    const tabs = ["Bid Documents", "Lost"];

    console.log('🔄 loadTabCounts started');

    try {
      await Promise.all(
        tabs.map(async (tab) => {
          try {
            let allItems: any[] = [];
            let statuses: string[] = [];

            if (tab === 'Bid Documents') {
              statuses = ['Potential'];
            } else {
              statuses = ['Dead'];
              // statuses = ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];
            }

            console.log(`Tab: ${tab}, Statuses:`, statuses);

            for (const status of statuses) {
              try {
                const { items } = await getProjectsPage(status, 1, 5000, {});
                console.log(`Status ${status} items count:`, items?.length || 0);
                allItems.push(...items);
              } catch (err) {
                console.error(`Error fetching ${status}:`, err);
              }
            }

            // Remove duplicates
            const uniqueItems = allItems.filter((item, index, self) =>
              index === self.findIndex(i => i.ID === item.ID)
            );

            counts[tab] = uniqueItems.length;
            console.log(`Tab ${tab} final count:`, counts[tab]);
          } catch (err) {
            console.error(`Error processing tab ${tab}:`, err);
            counts[tab] = 0;
          }
        })
      );
      console.log('✅ Final Tab Counts:', counts);
      setTabCounts(counts);
    } catch (err) {
      console.error('❌ loadTabCounts error:', err);
    }
  };

  const handleResetFilters = () => {
    setServerFilters({});
    setSearchTerm("");
    setCurrentPage(1);
    void loadBids(1, {}, "");
  };

  const handleTabChange = (tab: string) => {
    setActiveTab(tab);
    setServerFilters({});
    setCurrentPage(1);
    setUniqueValues({}); // clear stale dropdown values immediately
    void loadUniqueValues(tab);
  };

  /* ===================== LOAD SHAREPOINT COLUMNS ===================== */
  useEffect(() => {
    const loadColumns = async () => {
      // fetch list fields (lightweight) rather than items to avoid heavy queries
      const fields = await getProjectColumns();
      if (!fields || !fields.length) return;

      const spColumns: TableColumn[] = fields.map((f: any) => ({
        key: f.key,
        label: f.label,
        filterable: true,
      }));

      const merged = [
        ...DEFAULT_COLUMNS,
        ...spColumns.filter((c) => !DEFAULT_COLUMNS.some((d) => d.key === c.key)),
      ];

      setAllColumns(merged);
    };

    void loadColumns();
  }, []);

  useEffect(() => {
    void loadBids();
  }, [activeTab, allColumns]);

  useEffect(() => {
    if (allColumns.length > 0) {
      void loadUniqueValues();
    }
  }, [allColumns, activeTab]);

  useEffect(() => {
    void loadTabCounts();
  }, [allColumns]);

  useEffect(() => {
    localStorage.setItem(COLUMN_STORAGE_KEY, JSON.stringify(orderedColumnKeys));
    localStorage.setItem(COLUMN_VERSION_KEY, CURRENT_COLUMN_VERSION);
  }, [orderedColumnKeys]);

  const visibleColumns = sanitizeColumnKeys(orderedColumnKeys)
    .map((k) => allColumns.find((c) => c.key === k))
    .filter(Boolean) as TableColumn[];

  return (
    <>
      {/* <Hero
        title={`Hey ${props.userDisplayName},`}
        subtitle="Step into your project control hub!"
        bg={heroBg}
      /> */}

      {isLoading ? (
        <GlobalLoader variant="content" />
      ) : (
        <div style={{ marginTop: 16 }}>
          <TrackerTable
            title="Pipeline"
            tabs={["Bid Documents", "Lost"]}
            activeTab={activeTab}
            onTabChange={handleTabChange}
            columns={visibleColumns}
            rows={rows}
            serverSide
            totalItems={totalItems}
            currentPage={currentPage}
            pageSize={PAGE_SIZE}
            onPageChange={(p) => void loadBids(p, serverFilters)}
            uniqueValues={uniqueValues}
            onFilterChange={(filters) => {
              // Filters use SharePoint column names directly
              setServerFilters(filters);

              // Clean Title filter if it contains the Code separator
              const cleanedFilters = { ...filters };
              if (cleanedFilters.Title && cleanedFilters.Title.includes(" - ")) {
                cleanedFilters.Title = cleanedFilters.Title.split(" - ").slice(1).join(" - ");
              }

              void loadBids(1, cleanedFilters, searchTerm);
            }}
            onResetFilters={handleResetFilters}
            activeFilters={serverFilters}
            onRowClick={(item) => {
              window.location.href = `/sites/Projects/SitePages/PipelineDetails.aspx?projectId=${item.ID}`;
            }}
            showFilterIcon
            onSettingsClick={() => {
              setTempColumnKeys(sanitizeColumnKeys(orderedColumnKeys));
              setShowColumnPopup(true);
            }}
            tabCounts={tabCounts}
            sortColumn={sortColumn}
            sortAscending={sortAscending}
            onSort={handleSort}
            searchTerm={searchTerm}
            onSearch={handleSearch}
            searchPlaceholder="Search by Title, Company or Code..."
          />
        </div>
      )}

      {/* ================= COLUMN CUSTOMISER ================= */}
      {showColumnPopup && (
        <div className={styles.columnPopupOverlay}>
          <div className={styles.columnPopup}>
            <div className={styles.columnPopupHeader}>
              <div>
                <h3>Customise Columns</h3>
                <p>Reorder or toggle columns</p>
              </div>
              <span
                onClick={() => {
                  setTempColumnKeys(sanitizeColumnKeys(orderedColumnKeys));
                  setColumnSearchQuery("");
                  setShowColumnPopup(false);
                }}
              >
                ✕
              </span>
            </div>

            {/* Search Input */}
            <div className={styles.columnSearch}>
              <input
                type="text"
                placeholder="Search columns..."
                value={columnSearchQuery}
                onChange={(e) => setColumnSearchQuery(e.target.value)}
              />
            </div>

            <div className={styles.columnList}>
              <h4>Visible Columns</h4>

              {tempColumnKeys
                .filter((key) => {
                  const col = allColumns.find((c) => c.key === key);
                  return col?.label.toLowerCase().includes(columnSearchQuery.toLowerCase());
                })
                .map((key, index) => {
                  const col = allColumns.find((c) => c.key === key);
                  if (!col) return null;

                  return (
                    <div
                      key={key}
                      ref={(el) => el && (rowRefs.current[index] = el)}
                      className={styles.columnRow}
                      draggable
                      onDragStart={(e) => {
                        dragFromIndex.current = index;
                        e.dataTransfer.effectAllowed = "move";
                        e.currentTarget.style.opacity = "0.5";
                      }}
                      onDragEnd={(e) => {
                        e.currentTarget.style.opacity = "1";
                        dragFromIndex.current = null;
                      }}
                      onDragOver={(e) => {
                        e.preventDefault();
                        e.dataTransfer.dropEffect = "move";
                      }}
                      onDragEnter={(e) => {
                        e.preventDefault();
                      }}
                      onDrop={(e) => {
                        e.preventDefault();
                        const from = dragFromIndex.current;
                        if (from === null || from === index) return;

                        setTempColumnKeys((prev) => {
                          const copy = [...prev];
                          const [moved] = copy.splice(from, 1);
                          copy.splice(index, 0, moved);
                          return copy;
                        });

                        dragFromIndex.current = null;
                      }}
                    >
                      <input
                        type="checkbox"
                        checked
                        onChange={() =>
                          setTempColumnKeys((p) => p.filter((k) => k !== key))
                        }
                      />
                      <span className={styles.columnLabel}>{col.label}</span>
                      <span className={styles.dragHandle}>⋮⋮</span>
                    </div>
                  );
                })}

              <h4>Available Columns</h4>

              {allColumns
                .filter((c) => !tempColumnKeys.includes(c.key) && c.label.toLowerCase().includes(columnSearchQuery.toLowerCase()))
                .map((col) => (
                  <div key={col.key} className={styles.availableColumnRow}>
                    <input
                      type="checkbox"
                      onChange={() => setTempColumnKeys((p) => [...p, col.key])}
                    />
                    <span>{col.label}</span>
                  </div>
                ))}
            </div>

            {/* ✅ APPLY / RESET */}
            <div className={styles.popupActions}>
              <button
                onClick={() => {
                  setOrderedColumnKeys(sanitizeColumnKeys(tempColumnKeys));
                  setShowColumnPopup(false);
                }}
              >
                Apply
              </button>

              <button
                onClick={() => {
                  const defaults = DEFAULT_COLUMNS.map((c) => c.key);
                  setTempColumnKeys(defaults);
                  setOrderedColumnKeys(defaults);
                  localStorage.setItem(COLUMN_VERSION_KEY, CURRENT_COLUMN_VERSION);
                  localStorage.removeItem(COLUMN_STORAGE_KEY);
                  setShowColumnPopup(false);
                }}
              >
                Reset
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default Bids;
