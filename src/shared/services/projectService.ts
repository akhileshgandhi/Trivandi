/* eslint-disable no-useless-catch */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/search";
import { spfi, SPFI, SPFx } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";

let sp: any = null;
// simple in-memory cache for counts to avoid repeated heavy queries
export interface PagedResult<T> {
  items: T[];
  totalCount: number;
}

export interface IDashboardProject {
  Id: number;
  Title: string;
  Company?: string;
  Office?: string;
}
// Initialize PnP
export const initializePnP = (context: any) => {
  try {
    sp = spfi().using(SPFx(context));
    console.log("PnPjs initialized successfully");
    return sp;
  } catch (error) {
    console.error("Error initializing PnPjs:", error);
    throw error;
  }
};
/* ================= INIT ================= */
export const initProjectService = (spInstance: SPFI) => {
  sp = spInstance;
};

/* ================= CONSTANTS ================= */
const EXCLUDED_FIELDS = [
  "ID",
  "GUID",
  "Attachments",
  "Editor",
  "Created",
  "Modified",
  "ContentType",
  "ContentTypeId",
  "_UIVersionString",
  "ComplianceAssetId",
  "ProjectId",
  "ProjectID",
];

/**
 * Default visible columns (ORDER MATTERS)
 */
export const DEFAULT_COLUMN_KEYS = [
  "Title",
  "Company",
  "Owner",
  "BusinessUnit",
  "Sector",
  "ContractedEntity",
  "Country",
  "EndDate",
];

/* ================= GET LIST COLUMNS ================= */
export const getProjectColumns = async () => {
  if (!sp) throw new Error("PnPjs not initialized");

  const fields = await sp.web.lists
    .getByTitle("ProjectsNew")
    .fields
    .filter("Hidden eq false and ReadOnlyField eq false")
    .select("InternalName", "Title", "TypeAsString", "Required", "Choices", "AllowMultipleValues")();


  return fields
    .filter((f: any) => !EXCLUDED_FIELDS.includes(f.InternalName))
    .map((f: any) => ({
      key: f.InternalName,
      label: f.Title,
      filterable: true,
      sortable: true,
      type: f.TypeAsString,
      required: !!f.Required,
      choices: Array.isArray(f.Choices) ? f.Choices : undefined,
      allowMultiple: !!f.AllowMultipleValues,
    }));
};

/* ================= ENHANCE COLUMNS ================= */
export const enhanceProjectColumns = (columns: any[]) =>
  columns.map(col => {
    if (col.key === "Title") {
      return { ...col, label: "Project Title" };
    }
    if (col.key === "PercentComplete") {
      return { ...col, label: "% Completed" };
    }
    if (col.key === "TotalProjectValue") {
      return { ...col, label: "Budget" };
    }
    if (col.key === "Owner") {
      return { ...col, label: "Owner" };
    }
    return col;
  });

/* ================= GET PROJECT DATA ================= */
export const getProjectsByStatus = async (status: string) => {
  if (!sp) throw new Error("PnPjs not initialized");

  const fields = await getProjectColumns();
  const selectFields = fields.map(f => f.key);

  // ✅ ADD ID BEFORE QUERY
  if (!selectFields.includes("ID")) {
    selectFields.push("ID");
  }

  let items: any[] = [];

  items = await sp.web.lists
    .getByTitle("ProjectsNew")
    .items
    .select(...selectFields) // 👈 ID IS NOW INCLUDED
    .filter(`startswith(Status,'${status}')`)
    .top(5000)();



  return items.map((item: any, index: number) => {
    const row: any = {};

    selectFields.forEach(key => {
      let value = item[key];

      if (value && key.toLowerCase().includes("date")) {
        value = new Date(value).toLocaleDateString();
      }

      if (key === "PercentComplete" && typeof value === "number") {
        value = Math.round(value);
      }

      if (key === "TotalProjectValue" && typeof value === "number") {
        value = `$${value.toLocaleString()}`;
      }

      row[key] = value ?? "";
    });

    return row;
  });
};

/**
 * Fetch a paged list of projects with optional exact-match filters.
 * Returns items mapped to the same row shape and the total count.
 */
export const getProjectsPage = async (
  status: string,
  page: number,
  pageSize: number,
  filters: Record<string, string> = {},
  sortColumn?: string,
  sortAscending: boolean = true,
  searchTerm?: string
): Promise<PagedResult<any>> => {
  if (!sp) throw new Error("PnPjs not initialized");

  const fields = await getProjectColumns();

  // Separate regular fields from expanded fields
  const regularFields: string[] = [];
  const expandFields = new Set<string>();

  fields.forEach((f) => {
    if (f.key.includes("/")) {
      const [expandField] = f.key.split("/");
      expandFields.add(expandField);
    }
    regularFields.push(f.key);
  });

  if (!regularFields.includes("ID")) regularFields.push("ID");

  // Build filter string (exact matches for provided filters)
  const filterParts: string[] = [];
  // Only add Status filter if status is provided (empty string = no status filter, e.g. Non-CMAP tab)
  if (status) {
    filterParts.push(`startswith(Status,'${status}')`);
  }

  Object.entries(filters).forEach(([k, v]) => {
    if (!v) return; // Skip empty filter values

    // Handle boolean fields - SharePoint requires numeric comparison without quotes
    if (k === "NonCmap" || k === "Noncmap" || k.toLowerCase() === "noncmap") {
      const boolValue = (v === "1" || v === "true") ? 1 : 0;
      filterParts.push(`${k} eq ${boolValue}`);
    } else if (v === "1" || v === "0" || v === "true" || v === "false") {
      // Other boolean-like fields
      const boolValue = (v === "1" || v === "true") ? 1 : 0;
      filterParts.push(`${k} eq ${boolValue}`);
    } else {
      // Regular text fields - escape single quotes
      const safe = String(v).replace(/'/g, "''");
      filterParts.push(`${k} eq '${safe}'`);
    }
  });

  // Global search across Project Title, Company and Code
  if (searchTerm && searchTerm.trim()) {
    const safeSearch = searchTerm.replace(/'/g, "''");
    // User requested to search by Project Title, Company and Code
    filterParts.push(`(substringof('${safeSearch}',Title) or substringof('${safeSearch}',Company) or substringof('${safeSearch}',Code))`);
  }

  const filterString = filterParts.join(" and ");
  console.log("Generated OData filter:", filterString);

  // Check expand fields from filters too
  Object.keys(filters).forEach((k) => {
    if (k.includes("/")) {
      expandFields.add(k.split("/")[0]);
    }
  });

  try {
    // SharePoint's .skip() doesn't work reliably with filters
    // Instead, fetch all matching items and paginate client-side
    // Use a reasonable limit to avoid performance issues
    const MAX_FETCH = 5000;

    // Use passed sortColumn or default to "ID"
    const finalSortCol = sortColumn || "ID";

    let query = sp.web.lists.getByTitle("ProjectsNew").items
      .select(...regularFields)
      .filter(filterString)
      .orderBy(finalSortCol, sortAscending)
      .top(MAX_FETCH);

    // Add expand if needed
    if (expandFields.size > 0) {
      query = query.expand(...Array.from(expandFields));
    }

    const allItems = await query();

    // Get total count
    const totalCount = Array.isArray(allItems) ? allItems.length : 0;

    // Paginate client-side
    const skip = Math.max(0, (page - 1) * pageSize);
    const paged = Array.isArray(allItems)
      ? allItems.slice(skip, skip + pageSize)
      : [];

    console.log(`getProjectsPage: page=${page}, pageSize=${pageSize}, skip=${skip}, total=${totalCount}, returning=${paged.length} items`);

    const mapped = (paged || []).map((item: any) => {
      const row: any = { ID: item.ID };

      regularFields.forEach((key) => {
        if (key === "ID") return;

        let value = item[key];

        // Handle expanded fields like Owner/Title
        if (key.includes("/")) {
          const [expandField, selectField] = key.split("/");
          value = item[expandField] ? item[expandField][selectField] : undefined;
        }

        // Format dates
        if (value && key.toLowerCase().includes("date")) {
          try {
            value = new Date(value).toLocaleDateString();
          } catch {
            // Keep original value if date parsing fails
          }
        }

        // Format percent complete
        if (key === "PercentComplete" && typeof value === "number") {
          value = Math.round(value);
        }

        // Format currency
        if (key === "TotalProjectValue" && typeof value === "number") {
          value = `$${value.toLocaleString()}`;
        }

        row[key] = value ?? "";
      });

      return row;
    });

    return { items: mapped, totalCount };

  } catch (error) {
    console.error("Error fetching projects:", error);
    throw new Error(`Failed to fetch projects: ${error.message || error}`);
  }
};

/**
 * Fetch unique values for a column within a status (used to populate filter dropdowns)
 */
export const getUniqueValuesForColumn = async (
  status: string,
  columnKey: string,
  filters?: Record<string, string>
): Promise<string[]> => {
  if (!sp) throw new Error("PnPjs not initialized");
  // lightweight cache to avoid repeated expensive requests during a session
  // const filterStr = filters ? JSON.stringify(filters) : "";
  // const cacheKey = `${status}::${columnKey}::${filterStr}`;
  // (getUniqueValuesForColumn as any)._cache = (getUniqueValuesForColumn as any)._cache || {};
  // const cache = (getUniqueValuesForColumn as any)._cache as Record<string, string[]>;
  // if (cache[cacheKey]) return cache[cacheKey];

  const selectFields = [columnKey];
  if (columnKey === "Title") selectFields.push("Code");
  if (!selectFields.includes("ID")) selectFields.push("ID");

  // limit how many items we fetch for unique values to avoid heavy queries
  const MAX_TOP = 500;

  // try with expand if needed, but fallback to non-expanded request on failure
  const tryQuery = async (useExpand: boolean) => {
    try {
      // Build filter parts list for correct grouping
      const filterParts: string[] = [];
      // Only add Status filter if status is provided (empty string = no status filter, e.g. Non-CMAP tab)
      if (status) {
        filterParts.push(`startswith(Status,'${status}')`);
      }

      // Apply additional filters if provided (uses same logic as getProjectsPage)
      if (filters) {
        Object.entries(filters).forEach(([key, value]) => {
          if (key.toLowerCase() === "noncmap") {
            // NonCmap is a boolean - use numeric 1/0 (matches getProjectsPage)
            const boolValue = (value === "1" || value === "true") ? 1 : 0;
            filterParts.push(`${key} eq ${boolValue}`);
          } else {
            const safe = String(value).replace(/'/g, "''");
            filterParts.push(`${key} eq '${safe}'`);
          }
        });
      }

      const filterCondition = filterParts.join(" and ");
      const query = sp.web.lists.getByTitle("ProjectsNew").items.select(...selectFields).filter(filterCondition).top(MAX_TOP);
      const items = useExpand && columnKey.includes("/") ? await query.expand(columnKey.split("/")[0])() : await query();
      const set = new Set<string>();
      (items || []).forEach((it: any) => {
        let v: any = undefined;
        if (columnKey.includes("/")) {
          const parts = columnKey.split("/");
          v = it[parts[0]] ? it[parts[0]][parts[1]] : undefined;
        } else {
          v = it[columnKey];
        }
        if (v !== undefined && v !== null && String(v).trim() !== "") {
          if (columnKey === "Title" && it.Code) {
            set.add(`${it.Code} - ${v}`);
          } else {
            set.add(String(v));
          }
        }
      });
      const result = Array.from(set).sort();
      // cache[cacheKey] = result;
      return result;
    } catch (err) {
      // bubble up for caller to handle
      throw err;
    }
  };

  try {
    return await tryQuery(true);
  } catch (_) {
    try {
      return await tryQuery(false);
    } catch (err) {
      console.error("getUniqueValuesForColumn failed for", columnKey, err);
      return [];
    }
  }
};

export const getProjectById = async (id: number) => {
  if (!sp) throw new Error("PnPjs not initialized");

  const item = await sp.web.lists
    .getByTitle("ProjectsNew")
    .items.getById(id)
    .select("*", "ProjectDocumentsUrl", "BidDocumentsUrl", "ContractsDocumentsUrl")();

  return item;
};

/**
 * Create a new project item in the Projects list
 */
export const addProject = async (data: Record<string, any>) => {
  if (!sp) throw new Error("PnPjs not initialized");
  const res = await sp.web.lists.getByTitle("ProjectsNew").items.add(data);
  return res;
};


/**
 * Convert full URL to server relative URL and decode URL encoding
 * @param url - Full URL (https://...) or server relative URL (/sites/...)
 * @returns Server relative URL with proper decoding
 */
const toServerRelativeUrl = (url: string): string => {
  if (!url) return '';

  // If already server relative (starts with /), decode and return
  if (url.startsWith('/')) {
    return decodeURIComponent(url);
  }

  // If full URL, extract the path part and decode
  try {
    const urlObj = new URL(url);
    const decodedPath = decodeURIComponent(urlObj.pathname);
    console.log('🔗 Converted URL:', url, '->', decodedPath);
    return decodedPath;
  } catch (error) {
    console.warn('Failed to parse URL:', url, error);
    // Try direct decoding if URL parsing fails
    return decodeURIComponent(url);
  }
};

/**
 * Extract site URL and folder path from server relative URL
 * @param serverRelativeUrl - Server relative URL like /sites/sitename/library/folder
 * @returns Object with siteUrl and folderPath
 */
const parseSiteAndFolder = (serverRelativeUrl: string): { siteUrl: string; folderPath: string } => {
  // Format: /sites/sitename/library/folder
  const parts = serverRelativeUrl.split('/').filter(Boolean);

  if (parts.length < 2) {
    return { siteUrl: '', folderPath: serverRelativeUrl };
  }

  // Check if it's a /sites/ URL
  if (parts[0] === 'sites' && parts.length >= 2) {
    const siteUrl = `/${parts[0]}/${parts[1]}`;
    const folderPath = parts.slice(2).join('/');
    return { siteUrl, folderPath: folderPath ? `/${folderPath}` : '' };
  }

  // Default: treat first part as site
  return { siteUrl: `/${parts[0]}`, folderPath: `/${parts.slice(1).join('/')}` };
};

/**
 * Get documents from a server relative URL directly
 * @param serverRelativeUrl - The full or server relative URL to the folder
 * @param subFolderPath - Optional subfolder navigation within the URL
 */
export const getDocumentsByServerRelativeUrl = async (
  serverRelativeUrl: string,
  subFolderPath: string = "",
  page: number = 1,
  pageSize: number = 10
): Promise<{ items: any[]; totalCount: number }> => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    console.log("🔍 Input URL:", serverRelativeUrl);
    console.log("🔍 SubFolder:", subFolderPath);

    // Convert to server relative URL if it's a full URL
    const relativeUrl = toServerRelativeUrl(serverRelativeUrl);
    const fullPath = subFolderPath ? `${relativeUrl}/${subFolderPath}` : relativeUrl;

    console.log("🔗 Converted to server relative:", relativeUrl);
    console.log("📁 Full path with subfolder:", fullPath);

    // Parse site and folder path
    const { siteUrl, folderPath } = parseSiteAndFolder(fullPath);
    console.log("🌐 Site URL:", siteUrl);
    console.log("📂 Folder path:", folderPath);

    // Get the web context for the target site using absolute URL
    let absoluteSiteUrl = `${window.location.protocol}//${window.location.host}${siteUrl}`;
    
    // Ensure the URL is valid and doesn't have double slashes (except after protocol)
    absoluteSiteUrl = absoluteSiteUrl.replace(/([^:]\/)\/+/g, "$1");
    
    console.log("🌍 Absolute site URL:", absoluteSiteUrl);

    // Fallback to current web if siteUrl is empty or root
    const targetWeb = (!siteUrl || siteUrl === "/") ? sp.web : Web([sp.web, absoluteSiteUrl]);

    // Get files from the folder
    const items = await targetWeb
      .getFolderByServerRelativePath(fullPath)
      .files
      .select(
        "Name",
        "TimeLastModified",
        "ServerRelativeUrl",
        "Length"
      )();

    // Get subfolders
    const folders = await targetWeb
      .getFolderByServerRelativePath(fullPath)
      .folders
      .select("Name", "ServerRelativeUrl")
      .filter("Name ne 'Forms'")();

    console.log("📁 Raw files:", items.length);
    console.log("📂 Raw folders:", folders.length);

    // Combine files and folders
    const allItems = [
      ...folders.map((folder: any, index: number) => ({
        Id: Math.random() * 1000000 + index,
        FileLeafRef: folder.Name,
        Modified: new Date().toISOString(),
        Editor: { Title: "" },
        File_x0020_Type: "",
        FileRef: folder.ServerRelativeUrl,
        FSObjType: 1, // 1 = Folder
        VersionLabel: "",
        ServerRelativeUrl: folder.ServerRelativeUrl
      })),
      ...items.map((item: any, index: number) => ({
        Id: Math.random() * 1000000 + index + 100000,
        FileLeafRef: item.Name,
        Modified: item.TimeLastModified || new Date().toISOString(),
        Editor: { Title: "System" },
        File_x0020_Type: item.Name?.split('.').pop() || "",
        FileRef: item.ServerRelativeUrl,
        FSObjType: 0, // 0 = File
        VersionLabel: "1.0",
        ServerRelativeUrl: item.ServerRelativeUrl
      }))
    ];

    console.log("✅ Combined items:", allItems.length);

    const totalCount = allItems.length;
    const skip = (page - 1) * pageSize;
    const pagedItems = allItems.slice(skip, skip + pageSize);

    return { items: pagedItems, totalCount };
  } catch (error) {
    console.error("❌ Error fetching documents by server relative URL:", error);
    console.error("📍 Failed URL:", serverRelativeUrl);
    console.error("📍 SubFolder:", subFolderPath);

    // Provide more specific error information
    if (error instanceof Error) {
      console.error("📍 Error message:", error.message);
      if (error.message.includes('404') || error.message.includes('Not Found')) {
        console.error("💡 Folder may not exist or access denied. Check URL format and permissions.");
      }
    }

    return { items: [], totalCount: 0 };
  }
};

// Update getProjectDocuments - Remove 'Name' field
export const getProjectDocuments = async (
  libraryName: string,
  folderPath: string = "",
  page: number = 1,
  pageSize: number = 10
): Promise<{ items: any[]; totalCount: number }> => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    console.log("Library:", libraryName, "Folder:", folderPath || "root");

    // If no folder path, get items from root
    if (!folderPath) {
      const items = await sp.web.lists
        .getByTitle(libraryName)
        .items
        .select(
          "Id",
          "FileLeafRef",
          "Modified",
          "Editor/Title",
          "File_x0020_Type",
          "FileRef",
          "FSObjType",
          "OData__UIVersionString",
        )
        .expand("Editor")
        .filter("FSObjType eq 1 or FSObjType eq 0")
        .orderBy("FSObjType", false)
        .orderBy("FileLeafRef", true)
        .top(500)();

      // Filter to only root level items
      const rootPath = await sp.web.lists.getByTitle(libraryName).rootFolder.select("ServerRelativeUrl")();
      const rootItems = items.filter(item => {
        const itemPath = item.FileRef || "";
        const relativePath = itemPath.replace(rootPath.ServerRelativeUrl + "/", "");
        return !relativePath.includes("/");
      });

      // Map SharePoint internal version field to expected VersionLabel
      const mappedRootItems = rootItems.map((item: any) => ({
        ...item,
        VersionLabel: item.OData__UIVersionString,
      }));

      console.log("Root items found:", mappedRootItems.length);

      const totalCount = mappedRootItems.length;
      const skip = (page - 1) * pageSize;
      const pagedItems = mappedRootItems.slice(skip, skip + pageSize);

      return { items: pagedItems, totalCount };
    } else {
      // Get items from specific folder
      // Construct the full server-relative path accurately
      const rootFolder = await sp.web.lists.getByTitle(libraryName).rootFolder.select("ServerRelativeUrl")();
      const rootFolderPath = rootFolder.ServerRelativeUrl;
      
      let cleanFolderPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
      
      // Deduplicate library name if it's already in the folderPath
      const libNameClean = libraryName.replace(/\s/g, '');
      const folderPathClean = cleanFolderPath.replace(/\s/g, '');
      if (folderPathClean.startsWith(libNameClean)) {
        const slashIndex = cleanFolderPath.indexOf('/');
        cleanFolderPath = slashIndex !== -1 ? cleanFolderPath.substring(slashIndex + 1) : "";
      }

      const fullFolderPath = `${rootFolderPath}/${cleanFolderPath}`.replace(/\/+/g, '/').replace(/\/$/, "");
      console.log("📂 Sanitized Fetch Path:", fullFolderPath);

      // Get files and folders separately from the folder object (more reliable than filtering list items)
      const folderObj = sp.web.getFolderByServerRelativePath(fullFolderPath);
      
      const [files, folders] = await Promise.all([
        folderObj.files.select("Name", "TimeLastModified", "Length", "ServerRelativeUrl", "UniqueId").expand("ListItemAllFields")(),
        folderObj.folders.select("Name", "ServerRelativeUrl", "UniqueId").expand("ListItemAllFields").filter("Name ne 'Forms'")()
      ]);

      const allItems = [
        ...folders.map((f: any) => ({
          Id: f.ListItemAllFields?.Id || Math.random(),
          FileLeafRef: f.Name,
          Modified: f.ListItemAllFields?.Modified || new Date().toISOString(),
          Editor: { Title: f.ListItemAllFields?.Editor?.Title || "System" },
          File_x0020_Type: "",
          FileRef: f.ServerRelativeUrl,
          FSObjType: 1,
          VersionLabel: f.ListItemAllFields?.OData__UIVersionString || "1.0"
        })),
        ...files.map((f: any) => ({
          Id: f.ListItemAllFields?.Id || Math.random(),
          FileLeafRef: f.Name,
          Modified: f.TimeLastModified || new Date().toISOString(),
          Editor: { Title: f.ListItemAllFields?.Editor?.Title || "System" },
          File_x0020_Type: f.Name.split('.').pop() || "",
          FileRef: f.ServerRelativeUrl,
          FSObjType: 0,
          VersionLabel: f.ListItemAllFields?.OData__UIVersionString || "1.0"
        }))
      ];

      console.log("✅ Items found in folder:", allItems.length);

      const totalCount = allItems.length;
      const skip = (page - 1) * pageSize;
      const pagedItems = allItems.slice(skip, skip + pageSize);

      return { items: pagedItems, totalCount };
    }
  } catch (error) {
    console.error("Error fetching documents:", error);
    return { items: [], totalCount: 0 };
  }
};

/**
 * Create folder using server relative URL
 * @param serverRelativeUrl - The full or server relative URL to the parent folder
 * @param folderName - Name of the folder to create
 * @param subFolderPath - Optional subfolder within the URL
 */
export const createFolderByServerRelativeUrl = async (serverRelativeUrl: string, folderName: string, subFolderPath: string = "") => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    // Convert to server relative URL if it's a full URL
    const relativeUrl = toServerRelativeUrl(serverRelativeUrl);
    const parentPath = subFolderPath ? `${relativeUrl}/${subFolderPath}` : relativeUrl;
    console.log("📁 Creating folder:", folderName, "in:", parentPath);

    // Parse site and folder path
    const { siteUrl } = parseSiteAndFolder(parentPath);
    let absoluteSiteUrl = `${window.location.protocol}//${window.location.host}${siteUrl}`;
    absoluteSiteUrl = absoluteSiteUrl.replace(/([^:]\/)\/+/g, "$1");
    
    const targetWeb = (!siteUrl || siteUrl === "/") ? sp.web : Web([sp.web, absoluteSiteUrl]);

    const result = await targetWeb.getFolderByServerRelativePath(parentPath)
      .folders
      .addUsingPath(folderName);

    console.log("✅ Folder created successfully");
    return result;
  } catch (error) {
    console.error("❌ Error creating folder by server relative URL:", error);
    throw error;
  }
};

/**
 * Ensure folder path exists using server relative URL (creates nested folders if needed)
 * @param serverRelativeUrl - The full or server relative URL to the base folder
 * @param folderPath - The nested folder path to ensure exists
 */
export const ensureFolderPathByServerRelativeUrl = async (serverRelativeUrl: string, folderPath: string): Promise<void> => {
  if (!sp || !folderPath) return;

  try {
    // Convert to server relative URL if it's a full URL
    const relativeUrl = toServerRelativeUrl(serverRelativeUrl);
    console.log("🔍 Ensuring folder path exists:", folderPath, "in:", relativeUrl);

    // Parse site and folder path
    const { siteUrl } = parseSiteAndFolder(relativeUrl);
    let absoluteSiteUrl = `${window.location.protocol}//${window.location.host}${siteUrl}`;
    absoluteSiteUrl = absoluteSiteUrl.replace(/([^:]\/)\/+/g, "$1");
    
    const targetWeb = (!siteUrl || siteUrl === "/") ? sp.web : Web([sp.web, absoluteSiteUrl]);

    // Split the folder path and create each folder in the hierarchy
    const pathParts = folderPath.split("/").filter(part => part.length > 0);
    let currentPath = relativeUrl;

    for (const folderName of pathParts) {
      const newPath = `${currentPath}/${folderName}`;

      try {
        // Check if folder exists
        await targetWeb.getFolderByServerRelativePath(newPath).select("Name")();
        console.log("📁 Folder exists:", newPath);
      } catch (error) {
        // Folder doesn't exist, create it
        console.log("📁 Creating folder:", folderName, "in:", currentPath);
        await targetWeb.getFolderByServerRelativePath(currentPath)
          .folders
          .addUsingPath(folderName);
        console.log("✅ Created folder:", newPath);
      }

      currentPath = newPath;
    }

    console.log("✅ Folder path ensured:", folderPath);
  } catch (error) {
    console.error("❌ Error ensuring folder path:", error);
    throw error;
  }
};

/**
 * Upload document using server relative URL
 * @param serverRelativeUrl - The full or server relative URL to the folder
 * @param file - The file to upload
 * @param subFolderPath - Optional subfolder within the URL
 */
export const uploadDocumentByServerRelativeUrl = async (serverRelativeUrl: string, file: File, subFolderPath: string = "") => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    // Convert to server relative URL if it's a full URL
    const relativeUrl = toServerRelativeUrl(serverRelativeUrl);
    const fullPath = subFolderPath ? `${relativeUrl}/${subFolderPath}` : relativeUrl;
    console.log("📤 Uploading to server relative URL:", fullPath);

    // Parse site and folder path
    const { siteUrl } = parseSiteAndFolder(fullPath);
    let absoluteSiteUrl = `${window.location.protocol}//${window.location.host}${siteUrl}`;
    absoluteSiteUrl = absoluteSiteUrl.replace(/([^:]\/)\/+/g, "$1");
    
    const targetWeb = (!siteUrl || siteUrl === "/") ? sp.web : Web([sp.web, absoluteSiteUrl]);

    const result = await targetWeb.getFolderByServerRelativePath(fullPath)
      .files
      .addUsingPath(file.name, file, { Overwrite: true });

    console.log("File uploaded successfully");
    return result;
  } catch (error) {
    console.error("Error uploading file by server relative URL:", error);
    throw error;
  }
};

/**
 * Upload document with folder creation using server relative URL (for folder uploads)
 * @param serverRelativeUrl - The full or server relative URL to the folder
 * @param file - The file to upload
 * @param subFolderPath - Subfolder path that needs to be created
 */
export const uploadDocumentWithFolderCreation = async (serverRelativeUrl: string, file: File, subFolderPath: string) => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    // First ensure the folder path exists
    if (subFolderPath) {
      await ensureFolderPathByServerRelativeUrl(serverRelativeUrl, subFolderPath);
    }

    // Then upload the file using the original upload function
    return await uploadDocumentByServerRelativeUrl(serverRelativeUrl, file, subFolderPath);
  } catch (error) {
    console.error("Error uploading document with folder creation:", error);
    throw error;
  }
};

/**
 * Delete document using server relative URL
 * @param fileServerRelativeUrl - The full or server relative URL to the file
 */
export const deleteDocumentByServerRelativeUrl = async (fileServerRelativeUrl: string) => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    // Convert to server relative URL if it's a full URL
    const relativeUrl = toServerRelativeUrl(fileServerRelativeUrl);
    console.log("🗑️ Deleting file at:", relativeUrl);

    // Parse site URL
    const { siteUrl } = parseSiteAndFolder(relativeUrl);
    let absoluteSiteUrl = `${window.location.protocol}//${window.location.host}${siteUrl}`;
    absoluteSiteUrl = absoluteSiteUrl.replace(/([^:]\/)\/+/g, "$1");
    
    const targetWeb = (!siteUrl || siteUrl === "/") ? sp.web : Web([sp.web, absoluteSiteUrl]);

    await targetWeb.getFileByServerRelativePath(relativeUrl).delete();
    console.log("File deleted successfully");
  } catch (error) {
    console.error("Error deleting file by server relative URL:", error);
    throw error;
  }
};

// Update uploadProjectDocument - same as before
export const uploadProjectDocument = async (libraryName: string, file: File, folderPath: string = "") => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    console.log("Uploading to library:", libraryName, "folder:", folderPath || "root");

    let result;

    if (!folderPath) {
      // Upload to root of library
      result = await sp.web.lists
        .getByTitle(libraryName)
        .rootFolder
        .files
        .addUsingPath(file.name, file, { Overwrite: true });
    } else {
      // Ensure folder exists
      await ensureFolderPath(libraryName, folderPath);

      // Upload to specific folder
      result = await sp.web.lists
        .getByTitle(libraryName)
        .rootFolder
        .folders
        .getByUrl(folderPath)
        .files
        .addUsingPath(file.name, file, { Overwrite: true });
    }

    console.log("File uploaded successfully");

    // Optionally update metadata if required by your library
    // (Removed Status update: the field does not exist in this library)

    return result;
  } catch (error) {
    console.error("Error uploading file:", error);
    throw error;
  }
};

// Helper function to ensure folder path exists
const ensureFolderPath = async (libraryName: string, folderPath: string): Promise<void> => {
  if (!sp) throw new Error("PnPjs not initialized");

  const pathParts = folderPath.split("/");
  let currentPath = "";

  for (const part of pathParts) {
    if (part) {
      currentPath += (currentPath ? "/" : "") + part;

      try {
        await sp.web.lists
          .getByTitle(libraryName)
          .rootFolder
          .folders
          .getByUrl(currentPath)();
        console.log("Folder exists:", currentPath);
      } catch {
        console.log("Creating folder:", currentPath);
        const parentPath = currentPath.substring(0, currentPath.lastIndexOf("/"));
        if (parentPath) {
          await sp.web.lists
            .getByTitle(libraryName)
            .rootFolder
            .folders
            .getByUrl(parentPath)
            .folders
            .addUsingPath(part);
        } else {
          await sp.web.lists
            .getByTitle(libraryName)
            .rootFolder
            .folders
            .addUsingPath(part);
        }
      }
    }
  }
};

export const getProjectsByOwnerAndStatus = async (
  ownerEmail: string,
  status: "Project" | "Potential" | string[]
): Promise<IDashboardProject[]> => {
  if (!sp) throw new Error("PnPjs not initialized");

  console.log("[ProjectService] Fetching data", {
    ownerEmail,
    status,
  });

  // Handle both single status and multiple statuses
  const statusFilter = Array.isArray(status)
    ? status.map(s => `Status eq '${s}'`).join(' or ')
    : `Status eq '${status}'`;

  const items = await sp.web.lists
    .getByTitle("ProjectsNew")
    .items
    .select(
      "ID",
      "Title",
      "Company",
      "Office",
      "OwnerEmail",
      "Status"
    )
    .filter(
      `OwnerEmail eq '${ownerEmail}' and (${statusFilter})`
    )
    .top(20)();

  console.log(`[ProjectService] Raw ${status} items`, items);

  return items.map(item => ({
    Id: item.ID,
    Title: item.Title,
    Company: item.Company,
    Office: item.Office,
  }));
};
// Add this function to your projectService.ts

export const getRecentDocuments = async (limit: number = 10) => {
  if (!sp) throw new Error("PnPjs not initialized");

  try {
    // Get current user
    const currentUser = await sp.web.currentUser();
    const userEmail = currentUser.Email;

    console.log("Fetching recent documents for user via Search API:", currentUser.Title);

    // Using Search API is much more efficient as it retrieves documents across all libraries in a single call
    const results = await sp.search({
      Querytext: `IsDocument:1 AND (AuthorOWSUSER:${userEmail} OR EditorOWSUSER:${userEmail})`,
      SelectProperties: ["Title", "Path", "LastModifiedTime", "FileExtension", "ListItemID", "SiteTitle", "Filename"],
      SortList: [{ Property: "LastModifiedTime", Direction: 1 }], // 1 = Descending
      RowLimit: limit
    });

    if (!results || !results.PrimarySearchResults || results.PrimarySearchResults.length === 0) {
      console.log("No recent documents found via search.");
      return [];
    }

    const mappedDocs = results.PrimarySearchResults.map(res => {
      // Clean up Path to be relative if possible
      let relativePath = res.Path || "";
      try {
        if (relativePath.startsWith("http")) {
          relativePath = new URL(relativePath).pathname;
        }
      } catch (e) {
        // Fallback to original path
      }

      return {
        Id: parseInt(res.ListItemID) || Math.floor(Math.random() * 100000),
        FileLeafRef: res.Filename || res.Title,
        Modified: res.LastModifiedTime,
        FileRef: relativePath,
        File_x0020_Type: res.FileExtension,
        ProjectTitle: res.SiteTitle,
        LibraryTitle: res.SiteTitle
      };
    });

    console.log(`Found ${mappedDocs.length} recent documents via search.`);
    return mappedDocs;

  } catch (error) {
    console.error("Error fetching recent documents:", error);
    // Return mock data for testing
    return [
      {
        Id: 1,
        FileLeafRef: "Project_Kickoff_Deck_v2.pptx",
        Modified: new Date(Date.now() - 2 * 60 * 60 * 1000).toISOString(), // 2 hours ago
        FileRef: "/sites/test/docs/Project_Kickoff_Deck_v2.pptx",
        File_x0020_Type: "pptx",
        ProjectTitle: "Event Readiness Programme"
      },
      {
        Id: 2,
        FileLeafRef: "Budget_Tracking_Oct25.xlsx",
        Modified: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(), // Yesterday
        FileRef: "/sites/test/docs/Budget_Tracking_Oct25.xlsx",
        File_x0020_Type: "xlsx",
        ProjectTitle: "London Gatwick Airport"
      },
      {
        Id: 3,
        FileLeafRef: "Meeting_Minutes_SteerCo.docx",
        Modified: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(), // Yesterday
        FileRef: "/sites/test/docs/Meeting_Minutes_SteerCo.docx",
        File_x0020_Type: "docx",
        ProjectTitle: "Driving Council Steering"
      },
      {
        Id: 4,
        FileLeafRef: "RFP_Response_Draft.docx",
        Modified: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString(), // 2 days ago
        FileRef: "/sites/test/docs/RFP_Response_Draft.docx",
        File_x0020_Type: "docx",
        ProjectTitle: "Wembley Security Review"
      }
    ];
  }
};



//new
// Add these functions to your projectService.ts file

/**
 * Delete a document from SharePoint
 * @param libraryName - The document library name
 * @param fileServerRelativeUrl - The server relative URL of the file/folder
 * 
 */

const getRequestDigest = async (): Promise<string> => {
  try {
    const siteUrl = window.location.origin;
    const response = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      }
    });

    if (!response.ok) {
      throw new Error('Failed to get request digest');
    }

    const data = await response.json();
    return data.d.GetContextWebInformation.FormDigestValue;
  } catch (error) {
    console.error('Error getting request digest:', error);
    throw error;
  }
};
export const deleteProjectDocument = async (
  libraryName: string,
  fileServerRelativeUrl: string
): Promise<void> => {
  try {
    const siteUrl = window.location.origin;
    const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')`;

    const response = await fetch(apiUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': await getRequestDigest(), // You'll need to implement this
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to delete document: ${response.statusText}`);
    }

    console.log('Document deleted successfully:', fileServerRelativeUrl);
  } catch (error) {
    console.error('Error deleting document:', error);
    throw error;
  }
};

/**
 * Download a document from SharePoint
 * @param fileServerRelativeUrl - The server relative URL of the file
 * @param fileName - The name of the file
 */
export const downloadProjectDocument = async (
  fileServerRelativeUrl: string,
  fileName: string
): Promise<void> => {
  try {
    const siteUrl = window.location.origin;
    const fileUrl = `${siteUrl}${fileServerRelativeUrl}`;

    // Create a temporary anchor element to trigger download
    const link = document.createElement('a');
    link.href = fileUrl;
    link.download = fileName;
    link.target = '_blank';

    // Append to body, click, and remove
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    console.log('Document download initiated:', fileName);
  } catch (error) {
    console.error('Error downloading document:', error);
    throw error;
  }
};

/**
 * Get the SharePoint request digest for API calls
 * This is needed for POST/DELETE operations
 */


// Alternative implementation for downloading files (using fetch + blob)
export const downloadProjectDocumentBlob = async (
  fileServerRelativeUrl: string,
  fileName: string
): Promise<void> => {
  try {
    const siteUrl = window.location.origin;
    const fileUrl = `${siteUrl}${fileServerRelativeUrl}`;

    const response = await fetch(fileUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/octet-stream'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to download file: ${response.statusText}`);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Clean up
    window.URL.revokeObjectURL(url);

    console.log('Document downloaded successfully:', fileName);
  } catch (error) {
    console.error('Error downloading document:', error);
    throw error;
  }
};