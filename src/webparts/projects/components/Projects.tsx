
import * as React from "react";
import { useEffect, useState, useCallback } from "react";
import type { IProjectsProps } from "./IProjectsProps";
import { toast, ToastContainer } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";

import Hero from "../../../shared/component/Hero/Hero";
import heroBg from "../../../shared/assets/hero_new.jpg";

import TrackerTable from "../../../shared/component/TrackerTable/TrackerTable";
import { TableColumn } from "../../../shared/component/DataTable/DataTable";

import styles from "./Projects.module.scss";
import GlobalLoader from "../../../shared/component/GlobalLoader";
import CompletedCell from "../../../shared/component/CompletedCell/CompletedCell";
import "../../../shared/globalcss/globalcss.scss";
import {
  getProjectsByStatus,
  getProjectColumns,
  getProjectById,
  enhanceProjectColumns,
  DEFAULT_COLUMN_KEYS,
  getProjectsPage,
  getUniqueValuesForColumn,
  addProject,
} from "../../../shared/services/projectService";
import { Search } from "lucide-react";

const Projects: React.FC<IProjectsProps> = (props) => {
  const STORAGE_KEY = "projects_active_tab";

  const dragFromIndex = React.useRef<number | null>(null);
  const rowRefs = React.useRef<HTMLDivElement[]>([]);

  /* ===================== STATE ===================== */
  const [activeTab, setActiveTab] = useState(
    sessionStorage.getItem(STORAGE_KEY) || "Live"
  );

  const [rows, setRows] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  const [allColumns, setAllColumns] = useState<TableColumn[]>([]);
  const [orderedColumnKeys, setOrderedColumnKeys] = useState<string[]>(
    () => DEFAULT_COLUMN_KEYS
  );

  const [showColumnPopup, setShowColumnPopup] = useState(false);
  const [showFilters, setShowFilters] = useState(false);
  const [columnSearchQuery, setColumnSearchQuery] = useState("");

  //new
  const [selectedProject, setSelectedProject] = useState<any | null>(null);
  const [showProjectPopup, setShowProjectPopup] = useState(false);

  /* ===================== CREATE PROJECT MODAL STATE ===================== */
  // Controls whether the modal is visible
  const [showCreateModal, setShowCreateModal] = useState(false);
  
  // Stores all fields from SharePoint (field name, type, choices, required, etc.)
  const [formFields, setFormFields] = useState<any[]>([]);
  
  // Stores user input for each field (e.g., { Title: "Project Name", Status: "Live" })
  const [formValues, setFormValues] = useState<Record<string, any>>({});
  
  // Stores validation errors (e.g., { Title: "Title is required" })
  const [formErrors, setFormErrors] = useState<Record<string, string>>({});
  
  // True when saving to show "Saving..." button text
  const [isSaving, setIsSaving] = useState(false);
  
  // Toggle for create new project vs use existing URLs
  const [isCreateNew, setIsCreateNew] = useState(true);

  /* Server-side pagination state */
  const [currentPage, setCurrentPage] = useState(1);
  const [totalItems, setTotalItems] = useState(0);
  const PAGE_SIZE = 12;
  const [serverFilters, setServerFilters] = useState<Record<string, string>>(
    {}
  );
  const [uniqueValues, setUniqueValues] = useState<Record<string, string[]>>(
    {}
  );
  const [tabCounts, setTabCounts] = useState<Record<string, number>>({});

  /* ===================== LOAD COLUMNS ===================== */
  useEffect(() => {
    const loadColumns = async () => {
      const cols = await getProjectColumns();
      const enhanced = enhanceProjectColumns(cols);

      // setAllColumns(
      //   enhanced.map(col =>
      //     col.key === "PercentComplete"
      //       ? {
      //           ...col,
      //           render: (v: number) => <CompletedCell value={v ?? 0} />,
      //         }
      //       : col
      //   )
      // );
      //new
      setAllColumns(
        enhanced.map((col) => {
          // ✅ Project Title click
          if (col.key === "Title") {
            return {
              ...col,
              render: (_: any, row: any) => (
                <span
                  className={styles.projectLink}
                  onClick={() => {
                    // Set flag to indicate navigation from Projects page - Dashboard should be default
                    sessionStorage.setItem('projectDetails_navigation_source', 'projects_page');
                    // Set Dashboard as default for this specific project
                    sessionStorage.setItem(`projectDetails_tab_${row.ID}`, 'Dashboard');
                    window.location.href = `/sites/Projects/SitePages/ProjectDetails.aspx?projectId=${row.ID}`;
                  }}
                >
                  {row.Title}
                </span>
              ),
            };
          }

          // ✅ Percent Complete bar
          if (col.key === "PercentComplete") {
            return {
              ...col,
              render: (v: number) => <CompletedCell value={v ?? 0} />,
            };
          }

          // ✅ Budget → convert $ to £
          if (col.key === "TotalProjectValue") {
            return {
              ...col,
              render: (value: any) => {
                if (!value) return "—";

                // "$315,000.00" → 315000
                const usdNumber =
                  typeof value === "string"
                    ? Number(value.replace(/[^0-9.-]+/g, ""))
                    : value;

                const USD_TO_GBP = 0.79;
                const gbpValue = usdNumber * USD_TO_GBP;

                return `£${gbpValue.toLocaleString("en-GB", {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                })}`;
              },
            };
          }

          return col;
        })
      );
    };

    void loadColumns();
  }, []);

  /* ===================== LOAD DATA ===================== */
  const mapTabToStatus = (tab: string): string => {
    const statusMap: Record<string, string> = {
      Pipeline: "Potential",
      Live: "Project",
      Closed: "Closed",
      "Non-CMAP": "", // Non-CMAP: no Status filter, only NonCmap=1
    };
    return statusMap[tab] ?? "Project";
  };

  const loadProjects = async (
    page = 1,
    filters: Record<string, string> = {},
    status?: string,
    tab?: string
  ) => {
    setIsLoading(true);
    try {
      const currentTab = tab || activeTab;
      const statusToUse = status || mapTabToStatus(currentTab);
      
      // Add noncmap filter: include only NonCmap items for Non-CMAP tab, exclude them for Live/Closed
      const finalFilters = currentTab === "Non-CMAP"
        ? { ...filters, NonCmap: "1" }
        : { ...filters, NonCmap: "0" };
      
      console.log("Loading projects - Page:", page, "Status:", statusToUse, "Filters:", finalFilters);
      
      const { items, totalCount } = await getProjectsPage(
        statusToUse,
        page,
        PAGE_SIZE,
        finalFilters
      );
      
      console.log("Loaded projects - Count:", items.length, "Total:", totalCount, "Page:", page,items, "items:");
      
      // No need for client-side filtering - now done on server
      setRows(items);
      setTotalItems(totalCount);
      setCurrentPage(page);
    } catch (err) {
      console.error("Error loading projects:", err);
      setRows([]);
      setTotalItems(0);
    } finally {
      setIsLoading(false);
    }
  };

  const loadUniqueValues = async (tab?: string) => {
    const currentTab = tab || activeTab;
    const status = mapTabToStatus(currentTab);
    const tabFilters = currentTab === "Non-CMAP"
      ? { NonCmap: "1" }
      : { NonCmap: "0" };

    try {
      // Fetch ALL items for this tab using the exact same filters as the table
      const { items } = await getProjectsPage(status, 1, 5000, tabFilters);

      const filterableKeys = allColumns.filter((c) => c.filterable).map((c) => c.key);
      const map: Record<string, string[]> = {};

      filterableKeys.forEach((k) => {
        const values = Array.from(
          new Set(items.map((item: any) => String(item[k] ?? "")).filter((v: string) => v !== ""))
        ).sort() as string[];
        map[k] = values;
      });

      setUniqueValues(map);
    } catch (err) {
      console.error("Error loading unique values:", err);
    }
  };

  const loadTabCounts = async () => {
    const counts: Record<string, number> = {};
    const tabs = ["Live", "Closed", "Non-CMAP"];
    
    try {
      await Promise.all(
        tabs.map(async (tab) => {
          try {
            const status = mapTabToStatus(tab); // "" for Non-CMAP
            const filters = tab === "Non-CMAP" ? { NonCmap: "1" } : { NonCmap: "0" };
            const { totalCount } = await getProjectsPage(status, 1, 1, filters);
            counts[tab] = totalCount;
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

  /**
   * Opens the Create Project modal and prepares the form
   * Step 1: Load all available fields from SharePoint
   * Step 2: Filter out "Potential" from Status choices
   * Step 3: Set default values for the form
   */
  const openCreateModal = async () => {
    try {
      // Step 1: Get all fields from the Projects list
      const cols = await getProjectColumns();
      
      // Step 2: Remove "Potential", dead and deleted statuses from Status dropdown options
      const excludedStatuses = ['Potential', 'DeletedLead', 'Deleted', 'DeadLead', 'Dead', 'Lead'];
      const processedCols = cols.map((f: any) => {
        if (f.key === "Status" && Array.isArray(f.choices)) {
          return {
            ...f,
            choices: f.choices.filter((c: string) => !excludedStatuses.includes(c))
          };
        }
        return f;
      });
      setFormFields(processedCols);

      // Step 3: Set default values for each field
      const defaults: Record<string, any> = {};
      processedCols.forEach((f: any) => {
        // Status defaults to current tab (Live → Project, Closed → Closed, etc.)
        if (f.key === "Status") {
          const tab = activeTab === "Non-CMAP" ? "Live" : activeTab;
          defaults[f.key] = mapTabToStatus(tab);
        }
        // Percent Complete defaults to 0
        else if (f.key === "PercentComplete") {
          defaults[f.key] = 0;
        }
        // NonCmap always defaults to true
        else if (f.key === "NonCmap") {
          defaults[f.key] = true;
        }
        // All other fields start empty
        else {
          defaults[f.key] = "";
        }
      });
      
      setFormValues(defaults);
      setFormErrors({});
      setIsCreateNew(true); // Reset to default state
      setShowCreateModal(true);
    } catch (e) {
      console.error("Failed to load fields", e);
      alert("Couldn't load fields. Please try again.");
    }
  };

  /**
   * Handle toggle change - clear URL fields when switching to "Create New"
   */
  useEffect(() => {
    if (isCreateNew && formValues) {
      // Clear URL fields when switching to "Create New" mode
      const clearedValues = { ...formValues };
      delete clearedValues.ProjectDocumentsUrl;
      delete clearedValues.ContractsDocumentsUrl;
      delete clearedValues.BidDocumentsUrl;
      setFormValues(clearedValues);
    }
  }, [isCreateNew]);

  /**
   * Validates the form before submission
   * Checks that all required fields have values
   * Returns true if valid, false if there are errors
   */
  const validateForm = () => {
    const errs: Record<string, string> = {};
    
    // Check each field
    formFields.forEach((f: any) => {
      if (f.required) {
        const value = formValues[f.key];
        
        // Check if field is empty
        const isEmpty = 
          value === undefined || 
          value === null || 
          value === "" || 
          (Array.isArray(value) && value.length === 0);
        
        if (isEmpty) {
          errs[f.key] = `${f.label} is required`;
        }
      }
    });
    
    setFormErrors(errs);
    return Object.keys(errs).length === 0; // Return true if no errors
  };

  /**
   * Converts form input values to SharePoint-compatible formats
   * Different field types need different formats (numbers, dates, multi-choice, etc.)
   */
  const coerceValue = (type: string | undefined, allowMultiple: boolean | undefined, value: any) => {
    // Skip empty values
    if (value === "" || value === null || value === undefined) {
      return undefined;
    }
    
    switch (type) {
      // Convert to number for Number and Currency fields
      case "Number":
      case "Currency":
        const num = Number(value);
        return isNaN(num) ? undefined : num;
      
      // Convert to Date object for DateTime fields
      case "DateTime":
        try {
          const date = new Date(value);
          return isNaN(date.getTime()) ? undefined : date;
        } catch {
          return undefined;
        }
      
      // Convert to boolean for Boolean fields
      case "Boolean":
        return Boolean(value);
      
      // MultiChoice needs special format: { results: ["choice1", "choice2"] }
      case "MultiChoice":
        if (Array.isArray(value)) {
          return { results: value };
        }
        if (typeof value === "string") {
          const choices = value.split(",").map((s) => s.trim()).filter(Boolean);
          return { results: choices };
        }
        return undefined;
      
      // Default: return as-is or wrap in results array if multiple values allowed
      default:
        if (allowMultiple && Array.isArray(value)) {
          return { results: value };
        }
        return value;
    }
  };

  /**
   * Submits the form to create a new project
   * Step 1: Validate all required fields
   * Step 2: Convert form values to SharePoint format
   * Step 3: Add NonCmap=true to the data
   * Step 4: Save to SharePoint and refresh the list
   */
  const submitCreate = async () => {
    // Step 1: Check if form is valid
    if (!validateForm()) {
      return; // Stop if validation fails
    }
    
    setIsSaving(true);
    try {
      // Step 2: Build the data object for SharePoint
      const payload: Record<string, any> = {};
      
      // Convert each form value to the correct format
      formFields.forEach((f: any) => {
        const convertedValue = coerceValue(f.type, f.allowMultiple, formValues[f.key]);
        if (convertedValue !== undefined) {
          payload[f.key] = convertedValue;
        }
      });

      // Make sure Status has a value (default to Live if missing)
      const hasStatusField = formFields.some((f: any) => f.key === "Status");
      if (!payload.Status && hasStatusField) {
        const tab = activeTab === "Non-CMAP" ? "Live" : activeTab;
        payload.Status = mapTabToStatus(tab);
      }

      // Step 3: Always set NonCmap to true for all new projects
      payload.NonCmap = true;

      // Step 4: Save to SharePoint
      await addProject(payload);
      
      // Close modal and refresh the project list
      setShowCreateModal(false);
      void loadProjects(1, serverFilters);
      
      // Show success toast
      toast.success("Project created successfully!", {
        position: "top-right",
        autoClose: 3000,
      });
    } catch (e: any) {
      console.error("Create failed", e);
      // Show error toast
      toast.error(e?.message || "Failed to create project", {
        position: "top-right",
        autoClose: 5000,
      });
    } finally {
      setIsSaving(false);
    }
  };

  useEffect(() => {
    let savedTab = sessionStorage.getItem(STORAGE_KEY);

    // ❌ Remove Pipeline completely
    if (!savedTab || savedTab === "Pipeline" || savedTab === "Potential") {
      savedTab = "Live";
      sessionStorage.setItem(STORAGE_KEY, "Live");
    }

    setActiveTab(savedTab);
    if (savedTab === "Live") {
      setOrderedColumnKeys(DEFAULT_COLUMN_KEYS);
    }
    const status = mapTabToStatus(savedTab);
    void loadProjects(1, {}, status);
  }, []);

  useEffect(() => {
    if (allColumns.length > 0) {
      void loadUniqueValues();
    }
  }, [allColumns, activeTab]);

  useEffect(() => {
    void loadTabCounts();
  }, [allColumns]);

  /* ===================== TAB CHANGE ===================== */
  const onTabChange = (tab: string) => {
    setActiveTab(tab);
    sessionStorage.setItem(STORAGE_KEY, tab);
    setShowFilters(false);
    setServerFilters({});
    setCurrentPage(1);
    setUniqueValues({}); // clear stale dropdown values immediately
    if (tab === "Live") {
      setOrderedColumnKeys(DEFAULT_COLUMN_KEYS);
    }
    const status = mapTabToStatus(tab);
    void loadProjects(1, {}, status, tab);
    void loadUniqueValues(tab);
  };

  /* ===================== PAGE CHANGE HANDLER ===================== */
  const handlePageChange = useCallback((p: number) => {
    console.log("handlePageChange called with page:", p, "Current page state:", currentPage);
    setCurrentPage(p);
    const status = mapTabToStatus(activeTab);
    console.log("Calling loadProjects with page:", p, "status:", status, "filters:", serverFilters);
    void loadProjects(p, serverFilters, status, activeTab);
  }, [activeTab, serverFilters]);

  const handleResetFilters = () => {
    setServerFilters({});
    setCurrentPage(1);
    const status = mapTabToStatus(activeTab);
    void loadProjects(1, {}, status, activeTab);
  };

  /* ===================== ORDERED COLUMNS ===================== */
  const visibleColumns = activeTab === "Non-CMAP" 
    ? allColumns.filter(col => col.key === "Title")  // Show only Project Title for Non-CMAP
    : orderedColumnKeys
        .map((k) => allColumns.find((c) => c.key === k))
        .filter(Boolean) as TableColumn[];

  const PROJECT_SECTION_RULES = [
    {
      title: "Overview",
      match: (key: string) =>
        ["Title", "Company", "Stage", "Owner", "Status"].includes(key),
    },
    {
      title: "Timeline",
      match: (key: string) => key.toLowerCase().includes("date"),
    },
    {
      title: "Financials",
      match: (key: string) =>
        key.toLowerCase().includes("value") ||
        key.toLowerCase().includes("percent"),
    },
    {
      title: "Additional Details",
      match: (_key: string) => true, // fallback group
    },
  ];

  const SECTION_ORDER = [
    "Overview",
    "Timeline",
    "Financials",
    "Additional Details", // 👈 ALWAYS LAST
  ];

  const HIDDEN_SYSTEM_FIELDS = [
    "odata.metadata",
    "odata.id",
    "odata.type",
    "odata.etag",
    "odata.editLink",
    "ID",
    "GUID",
    "AuthorId",
    "EditorId",
    "Attachments",
    "ContentType",
    "ContentTypeId",
    "_UIVersionString",
    "ComplianceAssetId",
    "ProjectId",
    "ProjectID",
  ];

  const groupProjectFields = (project: any) => {
    const cleanEntries = Object.entries(project).filter(
      ([key, value]) =>
        value && !HIDDEN_SYSTEM_FIELDS.includes(key) && key !== "MoreInfo"
    );

    const grouped: Record<string, [string, any][]> = {};

    cleanEntries.forEach(([key, value]) => {
      const section =
        PROJECT_SECTION_RULES.find((rule) => rule.match(key))?.title ||
        "Additional Details";

      if (!grouped[section]) grouped[section] = [];
      grouped[section].push([key, value]);
    });

    return grouped;
  };

  const formatFieldValue = (value: any) => {
    if (value === null || value === undefined || value === "") return null;

    // ✅ Handle ISO date strings (SharePoint)
    if (
      typeof value === "string" &&
      /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(value)
    ) {
      return new Date(value).toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
    }

    // ✅ Handle Date objects
    if (value instanceof Date) {
      return value.toLocaleDateString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
      });
    }

    // ✅ Handle user / lookup / taxonomy
    if (typeof value === "object") {
      return value.Title || value.Label || JSON.stringify(value);
    }

    return String(value);
  };

  const prettifyLabel = (key: string) =>
    key
      .replace(/([A-Z])/g, " $1")
      .replace(/^./, (s) => s.toUpperCase())
      .trim();

  /* ===================== RENDER ===================== */
  return (
    <>
      <ToastContainer />
      {/* <Hero
        title={`Hey ${props.userDisplayName},`}
        subtitle="Step into your project control hub!"
        bg={heroBg}
      /> */}

      <div className={styles.wrapper}>
        {isLoading ? (
          <GlobalLoader variant="content" />
        ) : (
          <TrackerTable
            title="Project Tracker"
            tabs={["Live", "Closed", "Non-CMAP"]}
            activeTab={activeTab}
            onTabChange={onTabChange}
            columns={visibleColumns}
            rows={rows}
            serverSide
            totalItems={totalItems}
            currentPage={currentPage}
            pageSize={PAGE_SIZE}
            onPageChange={handlePageChange}
            uniqueValues={uniqueValues}
            onFilterChange={(filters) => {
              setServerFilters(filters);
              const status = mapTabToStatus(activeTab);
              void loadProjects(1, filters, status, activeTab);
            }}
            onResetFilters={handleResetFilters}
            activeFilters={serverFilters}
            showDivider={showFilters}
            showFilterIcon
            isFilterActive={showFilters}
            onFilterClick={() => setShowFilters((p) => !p)}
            onSettingsClick={() => setShowColumnPopup(true)}
            onCreateClick={activeTab === "Non-CMAP" ? openCreateModal : undefined}
            showFilterControls={showFilters}
            tabCounts={tabCounts}
          />
        )}

        {!isLoading && rows.length === 0 && (
          <div className={styles.emptyState}>No projects found.</div>
        )}
      </div>

      {/* ================= CREATE PROJECT MODAL ================= */}
      {showCreateModal && (
        <div className={styles.columnPopupOverlay}>
          <div className={styles.columnPopup} style={{ width: 720, height: isCreateNew ? '400px' : '400px' }}>
            <div className={styles.columnPopupHeader}>
              <div>
                <h3>Add Non-CMAP Project</h3>
                <p>Fill in details and save</p>
              </div>
              <div>
                <strong>
                  <span
                    onClick={() => {
                      setShowCreateModal(false);
                      setIsCreateNew(true); // Reset toggle state
                    }}
                    style={{ cursor: "pointer" }}
                  >
                    X
                  </span>
                </strong>
              </div>
            </div>

            <div className={styles.columnList}>
              {/* Toggle for Create New vs Use Existing */}
              <div style={{ 
                marginBottom: '20px', 
                padding: '15px', 
                backgroundColor: '#f8f9fa', 
                border: '1px solid #dee2e6', 
                borderRadius: '6px' 
              }}>
                <div style={{ 
                  marginBottom: '12px', 
                  fontWeight: '600', 
                  fontSize: '14px',
                  color: '#495057' 
                }}>
                  Project Setup:
                </div>
                <div style={{ display: 'flex', gap: '25px' }}>
                  <label style={{ 
                    display: 'flex', 
                    alignItems: 'center', 
                    gap: '8px', 
                    cursor: 'pointer',
                    fontSize: '13px',
                    fontWeight: '500'
                  }}>
                    <input 
                      type="radio" 
                      name="projectSetup" 
                      checked={isCreateNew} 
                      onChange={() => setIsCreateNew(true)}
                      style={{ marginRight: '4px' }}
                    />
                    Create New Project
                  </label>
                  <label style={{ 
                    display: 'flex', 
                    alignItems: 'center', 
                    gap: '8px', 
                    cursor: 'pointer',
                    fontSize: '13px',
                    fontWeight: '500'
                  }}>
                    <input 
                      type="radio" 
                      name="projectSetup" 
                      checked={!isCreateNew} 
                      onChange={() => setIsCreateNew(false)}
                      style={{ marginRight: '4px' }}
                    />
                    Use Existing URLs
                  </label>
                </div>
              </div>
              
              {/* Render only specified fields for create project form */}
              <div className={styles.formGrid}>
                {formFields.filter((f: any) => {
                  if (isCreateNew) {
                    // When "Create New" is selected, only show Title field
                    return ["Title"].includes(f.key);
                  } else {
                    // When "Use Existing" is selected, show all fields
                    return ["Title", "ProjectDocumentsUrl", "ContractsDocumentsUrl", "BidDocumentsUrl"].includes(f.key);
                  }
                }).map((f: any) => (
                  <div key={f.key} className={styles.formField}>
                    {/* Field Label with red * for required fields */}
                    <label className={styles.formLabel}>
                      {f.label}{f.required ? <span className={styles.required}> *</span> : ""}
                    </label>
                    
                    {/* Dropdown for single choice fields (e.g., Status, Stage) */}
                    {f.type === "Choice" && Array.isArray(f.choices) ? (
                      <select
                        value={formValues[f.key] ?? ""}
                        onChange={(e) => setFormValues(v => ({ ...v, [f.key]: e.target.value }))}
                        className={styles.formInput}
                      >
                        <option value="">-- Select --</option>
                        {f.choices.map((c: string) => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    ) : f.type === "MultiChoice" && Array.isArray(f.choices) ? (
                      /* Multi-select dropdown for multi-choice fields */
                      <select
                        multiple
                        value={Array.isArray(formValues[f.key]) ? formValues[f.key] : []}
                        onChange={(e) => {
                          const opts = Array.from(e.target.selectedOptions).map(o => o.value);
                          setFormValues(v => ({ ...v, [f.key]: opts }));
                        }}
                        className={styles.formInputMulti}
                      >
                        {f.choices.map((c: string) => (
                          <option key={c} value={c}>{c}</option>
                        ))}
                      </select>
                    ) : f.type === "DateTime" ? (
                      /* Date picker for date fields (StartDate, EndDate, etc.) */
                      <input
                        type="date"
                        value={formValues[f.key] ?? ""}
                        onChange={(e) => setFormValues(v => ({ ...v, [f.key]: e.target.value }))}
                        className={styles.formInput}
                      />
                    ) : f.type === "Boolean" ? (
                      /* Checkbox for true/false fields */
                      <input
                        type="checkbox"
                        checked={!!formValues[f.key]}
                        onChange={(e) => setFormValues(v => ({ ...v, [f.key]: e.target.checked }))}
                      />
                    ) : f.type === "Number" || f.type === "Currency" ? (
                      /* Number input for numeric and currency fields */
                      <input
                        type="number"
                        value={formValues[f.key] ?? ""}
                        onChange={(e) => setFormValues(v => ({ ...v, [f.key]: e.target.value }))}
                        className={styles.formInput}
                      />
                    ) : (
                      /* Text input for all other field types (default) */
                      <input
                        type="text"
                        value={formValues[f.key] ?? ""}
                        onChange={(e) => setFormValues(v => ({ ...v, [f.key]: e.target.value }))}
                        placeholder={(f.type?.includes("User") || f.type?.includes("Lookup")) ? "Enter ID or value" : undefined}
                        className={styles.formInput}
                      />
                    )}
                    {/* Show error message if field has validation error */}
                    {formErrors[f.key] && (
                      <div className={styles.formError}>{formErrors[f.key]}</div>
                    )}
                  </div>
                ))}
              </div>
            </div>

            <div className={styles.popupActions}>
              <button onClick={submitCreate} disabled={isSaving}>
                {isSaving ? "Saving..." : "Save"}
              </button>
              <button onClick={() => {
                setShowCreateModal(false);
                setIsCreateNew(true); // Reset toggle state
              }}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {/* ================= COLUMN CUSTOMISER ================= */}
      {showColumnPopup && (
        <div className={styles.columnPopupOverlay}>
          <div className={styles.columnPopup}>
            <div className={styles.columnPopupHeader}>
              <div>
                {" "}
                <h3>Customise Columns</h3>
                <p>Reorder or toggle columns</p>
              </div>
              <div>
                {" "}
                <strong>
                  <span
                    onClick={() => {
                      setColumnSearchQuery("");
                      setShowColumnPopup(false);
                    }}
                    style={{ cursor: "pointer" }}
                  >
                    X
                  </span>
                </strong>
              </div>
            </div>

            {/* Search Input */}
            <div className={styles.columnSearch}>
              <div className={styles.iconcss}>
                <Search />
              </div>
              <input
                type="text"
                placeholder="Search columns..."
                value={columnSearchQuery}
                onChange={(e) => setColumnSearchQuery(e.target.value)}
              />
            </div>

            <div className={styles.columnList}>
              <h4>Visible Columns</h4>

              {orderedColumnKeys
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
                      ref={(el) => {
                        if (el) rowRefs.current[index] = el;
                      }}
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

                        setOrderedColumnKeys((prev) => {
                          const copy = [...prev];
                          const [moved] = copy.splice(from, 1);
                          copy.splice(index, 0, moved);
                          return copy;
                        });

                        dragFromIndex.current = null;
                      }}
                      /* ===== MOBILE ===== */
                      onTouchStart={() => {
                        dragFromIndex.current = index;
                      }}
                      onTouchEnd={(e) => {
                        const from = dragFromIndex.current;
                        if (from === null) return;

                        const y = e.changedTouches[0].clientY;
                        const targetIndex = rowRefs.current.findIndex((el) => {
                          const r = el.getBoundingClientRect();
                          return y >= r.top && y <= r.bottom;
                        });

                        if (targetIndex === -1 || targetIndex === from) {
                          dragFromIndex.current = null;
                          return;
                        }

                        setOrderedColumnKeys((prev) => {
                          const updated = [...prev];
                          const [moved] = updated.splice(from, 1);

                          const insertIndex =
                            targetIndex > from ? targetIndex - 1 : targetIndex;

                          updated.splice(insertIndex, 0, moved);
                          return updated;
                        });

                        dragFromIndex.current = null;
                      }}
                    >
                      <input
                        type="checkbox"
                        checked
                        onChange={() =>
                          setOrderedColumnKeys((prev) =>
                            prev.filter((k) => k !== key)
                          )
                        }
                      />

                      <div className={styles.columnContent}>
                        <span className={styles.columnLabel}>{col.label}</span>
                        <span className={styles.dragHandle}>⋮⋮</span>
                      </div>
                    </div>
                  );
                })}

              <h4 style={{ marginTop: 12 }}>Available Columns</h4>

              {allColumns
                .filter((c) => !orderedColumnKeys.includes(c.key) && c.label.toLowerCase().includes(columnSearchQuery.toLowerCase()))
                .map((col) => (
                  <label key={col.key} className={styles.columnItem}>
                    <input
                      type="checkbox"
                      onChange={() =>
                        setOrderedColumnKeys((p) => [...p, col.key])
                      }
                    />
                    {col.label}
                  </label>
                ))}
            </div>

            <div className={styles.popupActions}>
              <button onClick={() => setShowColumnPopup(false)}>Apply</button>
              <button
                onClick={() => {
                  setOrderedColumnKeys(DEFAULT_COLUMN_KEYS);
                  setShowColumnPopup(false);
                }}
              >
                Reset
              </button>
            </div>
          </div>
        </div>
      )}

      {showProjectPopup && (
        <div className={styles.projectPopupOverlay}>
          <div className={styles.projectPopup}>
            {/* 🔒 FIXED HEADER */}
            <div className={styles.projectHeader}>
              <h3>{selectedProject?.Title}</h3>
            </div>

            {/* 🔽 SCROLLABLE BODY */}
            <div className={styles.projectBody}>
              {!selectedProject ? (
                <GlobalLoader variant="content" />
              ) : (
                <div className={styles.projectDocument}>
                  {SECTION_ORDER.map((section) => {
                    const fields = groupProjectFields(selectedProject)[section];
                    if (!fields || fields.length === 0) return null;

                    return (
                      <div key={section} className={styles.sectionCard}>
                        <h4 className={styles.sectionTitle}>{section}</h4>

                        <div className={styles.sectionGrid}>
                          {fields.map(([key, value]) => {
                            const formatted = formatFieldValue(value);
                            if (!formatted) return null;

                            return (
                              <div key={key} className={styles.fieldRow}>
                                <span className={styles.fieldLabel}>
                                  {prettifyLabel(key)}
                                </span>
                                <span className={styles.fieldValue}>
                                  {formatted}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* FOOTER */}
            <div className={styles.popupActions}>
              <button onClick={() => setShowProjectPopup(false)}>Close</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default Projects;