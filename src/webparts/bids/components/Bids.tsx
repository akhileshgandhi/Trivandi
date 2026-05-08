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
    filterable: true,
    render: (value: any, row: any) => (
      <span className={styles.pipelineTitle}>
        {row.Code ? `${row.Code} - ` : ""}{value ?? "-"}
      </span>
    ),
    sortable: true,
  },
  { key: "Company", label: "Company", filterable: true, sortable: true },
  { key: "Owner", label: "Owner", filterable: true, sortable: true },
  { key: "BusinessUnit", label: "Business Unit", filterable: true, sortable: true },
  { key: "Sector", label: "Sector", filterable: true, sortable: true },
  { key: "ContractedEntity", label: "Contracted Entity", filterable: true, sortable: true },
  { key: "Country", label: "Country", filterable: true, sortable: true },
  { key: "EndDate", label: "End Date", filterable: true, sortable: true },
];

const Bids: React.FC<IBidsProps> = (props) => {
  const [activeTab, setActiveTab] = useState("Bid Documents");
  const [isLoading, setIsLoading] = useState(true);
  const [rows, setRows] = useState<any[]>([]);
  const [currentPage, setCurrentPage] = useState(1);
  const [totalItems, setTotalItems] = useState(0);
  const PAGE_SIZE = 12;
  const [serverFilters, setServerFilters] = useState<Record<string, string>>({});
  const [uniqueValues, setUniqueValues] = useState<Record<string, string[]>>({});
  const [tabCounts, setTabCounts] = useState<Record<string, number>>({});
  const [sortColumn, setSortColumn] = useState<string>("ID");
  const [sortAscending, setSortAscending] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>("");

  const COLUMN_STORAGE_KEY = `bids_column_order_${props.userDisplayName}`;
  const COLUMN_VERSION_KEY = `bids_column_version_${props.userDisplayName}`;
  const CURRENT_COLUMN_VERSION = '5'; // Increment when changing default columns

  const [allColumns, setAllColumns] = useState<TableColumn[]>(DEFAULT_COLUMNS);

  const [orderedColumnKeys, setOrderedColumnKeys] = useState<string[]>(() => {
    const savedVersion = localStorage.getItem(COLUMN_VERSION_KEY);
    const saved = localStorage.getItem(COLUMN_STORAGE_KEY);

    // If version mismatch or no version, use defaults
    if (savedVersion !== CURRENT_COLUMN_VERSION || !saved) {
      localStorage.setItem(COLUMN_VERSION_KEY, CURRENT_COLUMN_VERSION);
      return DEFAULT_COLUMNS.map((c) => c.key);
    }

    return JSON.parse(saved);
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

  /* ===================== LOAD BIDS ===================== */
  const loadBids = async (page = 1, filters: Record<string, string> = {}, sTerm?: string) => {
    setIsLoading(true);
    try {
      const termToUse = sTerm ?? searchTerm;
      console.log(`🔍 Loading ${activeTab} tab`);

      let allItems: any[] = [];

      if (activeTab === 'Bid Documents') {
        // Bid Documents tab - Only Potential (removed Lead status)
        const statuses = ['Potential'];
        for (const statuss of statuses) {
          try {
            const { items } = await getProjectsPage(statuss, 1, 5000, filters);
            allItems.push(...items);
          } catch (err) {
            console.error(`Error fetching ${statuss}:`, err);
          }
        }
      } else {
        // Lost tab - Fetch ALL statuses and filter client-side by Status field
        // const allStatuses = ['Lead', 'Potential', 'Project', 'Closed'];
        const lostStatusValues = ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];

        console.log(`🔍 Lost tab: Fetching all items and filtering by Status field...`);

        for (const status of lostStatusValues) {
          try {
            // const { items } = await getProjectsPage(status, 1, 5000, {});
            const { items } = await getProjectsPage(status, 1, 5000, filters);
            allItems.push(...items);
          } catch (err) {
            console.error(`Error checking ${status}:`, err);
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
            String(item.Owner ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Sector ?? "").toLowerCase().includes(lowerTerm) ||
            String(item.Country ?? "").toLowerCase().includes(lowerTerm)
          );
        });
      }

      console.log(`📊 Total items for ${activeTab}: ${allItems.length}`);

      // Remove duplicates
      const uniqueItems = allItems.filter((item, index, self) =>
        index === self.findIndex(i => i.ID === item.ID)
      );

      // Apply pagination
      const startIndex = (page - 1) * PAGE_SIZE;
      const endIndex = startIndex + PAGE_SIZE;
      const paginatedItems = uniqueItems.slice(startIndex, endIndex);

      setRows(paginatedItems);
      setTotalItems(uniqueItems.length);
      setCurrentPage(page);
    } catch (err) {
      console.error('Fatal error:', err);
      setRows([]);
      setTotalItems(0);
    } finally {
      setIsLoading(false);
    }
  };

  const handleSort = (columnKey: string, ascending: boolean) => {
    setSortColumn(columnKey);
    setSortAscending(ascending);

    // Client-side sort for Bids since it fetches 5000 items
    const sortedRows = [...rows].sort((a, b) => {
      const valA = a[columnKey] ?? "";
      const valB = b[columnKey] ?? "";

      if (valA < valB) return ascending ? -1 : 1;
      if (valA > valB) return ascending ? 1 : -1;
      return 0;
    });

    setRows(sortedRows);
  };

  const handleSearch = (value: string) => {
    setSearchTerm(value);
    setCurrentPage(1);
    void loadBids(1, serverFilters, value);
  };

  const loadUniqueValues = async (tab?: string) => {
    const currentTab = tab || activeTab;
    const statuses = currentTab === 'Bid Documents'
      ? ['Potential']
      : ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];

    try {
      // Fetch ALL items for this tab using the exact same filters as the table
      let allItems: any[] = [];
      for (const status of statuses) {
        try {
          const { items } = await getProjectsPage(status, 1, 5000, {});
          allItems = allItems.concat(items);
        } catch (err) {
          console.error(`Error fetching ${status}:`, err);
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
        } else {
          const values = Array.from(
            new Set(uniqueItems.map((item: any) => String(item[k] ?? "")).filter((v: string) => v !== ""))
          ).sort() as string[];
          map[k] = values;
        }
      });

      setUniqueValues(map);
    } catch (err) {
      console.error("Error loading unique values:", err);
    }
  };

  const loadTabCounts = async () => {
    const counts: Record<string, number> = {};
    const tabs = ["Bid Documents", "Lost"];

    try {
      await Promise.all(
        tabs.map(async (tab) => {
          try {
            let allItems: any[] = [];
            let statuses: string[] = [];

            if (tab === 'Bid Documents') {
              statuses = ['Potential'];
            } else {
              statuses = ['Dead', 'Deleted', 'DeletedLead', 'DeadLead'];
            }

            for (const status of statuses) {
              try {
                const { items } = await getProjectsPage(status, 1, 5000, {});
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
          } catch (err) {
            console.error(`Error loading count for tab ${tab}:`, err);
            counts[tab] = 0;
          }
        })
      );
      setTabCounts(counts);
    } catch (err) {
      console.error("Error loading tab counts:", err);
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

  const visibleColumns = orderedColumnKeys
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
              void loadBids(1, filters, searchTerm);
            }}
            onResetFilters={handleResetFilters}
            activeFilters={serverFilters}
            onRowClick={(item) => {
              window.location.href = `/sites/Projects/SitePages/PipelineDetails.aspx?projectId=${item.ID}`;
            }}
            showFilterIcon
            onSettingsClick={() => {
              setTempColumnKeys(orderedColumnKeys);
              setShowColumnPopup(true);
            }}
            tabCounts={tabCounts}
            sortColumn={sortColumn}
            sortAscending={sortAscending}
            onSort={handleSort}
            searchTerm={searchTerm}
            onSearch={handleSearch}
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
                  setTempColumnKeys(orderedColumnKeys);
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
                  setOrderedColumnKeys(tempColumnKeys);
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
