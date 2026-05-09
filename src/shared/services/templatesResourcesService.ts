/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import "@pnp/sp/views";

import {
  IDocumentItem,
  IFolderItem,
  ITemplateDocument,
  IResourceFolder,
  IDocumentLibraryResponse,
  IFolderStructure,
} from "../interfaces/ITemplatesResourcesInterfaces";

// Helper method to generate hash codes for string IDs
declare global {
  interface String {
    hashCode(): number;
  }
}

String.prototype.hashCode = function(): number {
  let hash = 0;
  if (this.length === 0) {
      return hash;
  }
  for (let i = 0; i < this.length; i++) {
      const char = this.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash);
};

let sp: SPFI;

export const initTemplatesResourcesService = (spInstance: SPFI): void => {
  sp = spInstance;
};

/**
 * Get data from the specific external SharePoint site and folder
 */
export const getDataFromExternalSharePointSite = async (): Promise<IDocumentLibraryResponse> => {
  try {
    
    
    // The specific external SharePoint site and folder path
    const externalSiteUrl = "https://trivandildn.sharepoint.com/sites/TrivandiLondon";
    const folderPath = "/Shared Documents/6 Projects/_2024 Project Folder Template";
    
    // Build the REST API URL to get folder contents
    const apiUrl = `${externalSiteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')?$expand=Files,Folders&$select=Name,ServerRelativeUrl,Files/Name,Files/ServerRelativeUrl,Files/Length,Files/TimeLastModified,Folders/Name,Folders/ServerRelativeUrl,Folders/ItemCount`;
    
    // Get current context for authentication
    // Note: Cross-site access may require different authentication approach
    const webInfo = await sp.web.select("Url", "Title")();
    
    
    // Make REST API call with proper headers
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      credentials: 'include'
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const data = await response.json();
    const folderData = data.d;
    
    const allFiles: IDocumentItem[] = [];
    const allFolders: IFolderItem[] = [];
    const folderStructure: IFolderStructure[] = [];
    
    // Process files
    if (folderData.Files && folderData.Files.results) {
      folderData.Files.results.forEach((file: any) => {
        const fileExtension = file.Name.split('.').pop()?.toLowerCase() || '';
        allFiles.push({
          Id: file.ServerRelativeUrl.hashCode(),
          Title: file.Name,
          Name: file.Name,
          FileRef: file.ServerRelativeUrl,
          FileLeafRef: file.Name,
          File_x0020_Type: fileExtension,
          Modified: file.TimeLastModified,
          Created: file.TimeLastModified
        });
      });
    }
    
    // Process folders
    if (folderData.Folders && folderData.Folders.results) {
      for (const folder of folderData.Folders.results) {
        if (folder.Name && !folder.Name.startsWith('_')) { // Skip system folders
          allFolders.push({
            Id: folder.ServerRelativeUrl.hashCode(),
            Title: folder.Name,
            Name: folder.Name,
            FileRef: folder.ServerRelativeUrl,
            FileLeafRef: folder.Name,
            FSObjType: 1,
            ItemChildCount: folder.ItemCount || 0,
            Modified: new Date().toISOString(),
            Created: new Date().toISOString()
          });
          
          // Get the complete folder structure recursively
          try {
            const subFolderStructure = await getExternalFolderStructureRecursiveREST(
              externalSiteUrl,
              folder.ServerRelativeUrl, 
              folder.Name, 
              0
            );
            folderStructure.push(subFolderStructure);
          } catch (subError) {
            
            // Add basic folder structure without subfolders
            folderStructure.push({
              Id: folder.ServerRelativeUrl.hashCode(),
              Name: folder.Name,
              Title: folder.Name,
              FileRef: folder.ServerRelativeUrl,
              FSObjType: 1,
              Modified: new Date().toISOString(),
              Created: new Date().toISOString(),
              ItemChildCount: folder.ItemCount || 0,
              files: [],
              subFolders: [],
              level: 0
            });
          }
        }
      }
    }
    
    
    
    return {
      standardDocuments: allFiles.filter(f => getFileCategory(f.File_x0020_Type) === 'template'),
      invoicingDocuments: [],
      resourceFolders: allFolders.map(f => ({ ...f, Description: '', Color: '#0078d4', IconType: 'Folder' })),
      allFiles,
      allFolders,
      folderStructure
    };
    
  } catch (error) {
    
    
    // Return empty response on error
    return getEmptyResponse();
  }
};

export const getTemplatesResourcesData = async (): Promise<IDocumentLibraryResponse> => {
  try {
    
    
    // Try to get data from the specific external SharePoint location first
    const externalData = await getDataFromExternalSharePointSite();
    
    // If we got some data from external site, return it
    if (externalData.allFiles.length > 0 || externalData.allFolders.length > 0 || externalData.folderStructure.length > 0) {
      
      return externalData;
    }
    
    // Fallback: Try to get complete folder structure from the current site's document library
    const sharePointData = await getCompletefolderStructure("Documents");
    
    // If we got some data, return it
    if (sharePointData.allFiles.length > 0 || sharePointData.allFolders.length > 0 || sharePointData.folderStructure.length > 0) {
      
      return sharePointData;
    }
    
    // Try alternative library names
    const alternativeLibraries = ["Shared Documents", "Site Assets", "Style Library"];
    for (const libName of alternativeLibraries) {
      try {
        const altData = await getCompletefolderStructure(libName);
        if (altData.allFiles.length > 0 || altData.folderStructure.length > 0) {
          
          return altData;
        }
      } catch (altError) {
        
      }
    }
    
    // If no data was found, return empty response
    
    return getEmptyResponse();
    
  } catch (error) {
    
    
    // Return empty data on error
    return getEmptyResponse();
  }
};
/**
 * Get complete folder structure with files from SharePoint document library
 */
export const getCompletefolderStructure = async (libraryName: string = "Documents"): Promise<IDocumentLibraryResponse> => {
  try {
    

    // Get all items from the document library with folder information
    const libraryItems = await sp.web.lists
      .getByTitle(libraryName)
      .items
      .select(
        "Id,Title,FileRef,FileLeafRef,File_x0020_Type,Modified,Created,FSObjType," +
        "FileDirRef,FileSizeDisplay,Author/Title,Author/EMail,Editor/Title,Editor/EMail"
      )
      .expand("Author,Editor")
      .filter("FSObjType eq 0 or FSObjType eq 1") // Include both files and folders
      .top(5000)();

    

    // Separate files and folders
    const allFiles: IDocumentItem[] = libraryItems
      .filter((item: any) => item.FSObjType === 0)
      .map((item: any) => ({
        Id: item.Id,
        Title: item.Title || item.FileLeafRef,
        Name: item.FileLeafRef,
        FileRef: item.FileRef,
        FileLeafRef: item.FileLeafRef,
        File_x0020_Type: item.File_x0020_Type,
        Modified: item.Modified,
        Created: item.Created,
        FileSizeDisplay: item.FileSizeDisplay,
        Author: item.Author ? {
          Title: item.Author.Title,
          EMail: item.Author.EMail,
        } : undefined,
        Editor: item.Editor ? {
          Title: item.Editor.Title,
          EMail: item.Editor.EMail,
        } : undefined,
      }));

    const allFolders: IFolderItem[] = libraryItems
      .filter((item: any) => item.FSObjType === 1)
      .map((item: any) => ({
        Id: item.Id,
        Title: item.Title || item.FileLeafRef,
        Name: item.FileLeafRef,
        FileRef: item.FileRef,
        FileLeafRef: item.FileLeafRef,
        FSObjType: 1,
        Modified: item.Modified,
        Created: item.Created,
      }));

    // Build hierarchical folder structure
    const folderStructure = buildFolderHierarchy(allFolders, allFiles, libraryName);

    const standardDocuments = categorizeStandardDocuments(allFiles);
    const invoicingDocuments = categorizeInvoicingDocuments(allFiles);
    const resourceFolders = categorizeResourceFolders(allFolders);

    return {
      standardDocuments,
      invoicingDocuments,
      resourceFolders,
      allFiles,
      allFolders,
      folderStructure,
    };
  } catch (error) {
    
    return getEmptyResponse();
  }
};

/**
 * Build hierarchical folder structure from flat arrays
 */
const buildFolderHierarchy = (folders: IFolderItem[], files: IDocumentItem[], libraryName: string): IFolderStructure[] => {
  const folderMap = new Map<string, IFolderStructure>();
  const rootFolders: IFolderStructure[] = [];

  // Create folder structure objects
  folders.forEach(folder => {
    const folderStructure: IFolderStructure = {
      Id: folder.Id,
      Name: folder.Name,
      Title: folder.Title,
      FileRef: folder.FileRef,
      FSObjType: folder.FSObjType,
      Modified: folder.Modified,
      Created: folder.Created,
      files: [],
      subFolders: [],
      level: 0,
      Color: getRandomFolderColor(),
    };
    
    folderMap.set(folder.FileRef, folderStructure);
  });

  // Assign files to their parent folders
  files.forEach(file => {
    const fileDirRef = file.FileRef.substring(0, file.FileRef.lastIndexOf('/'));
    const parentFolder = folderMap.get(fileDirRef);
    
    if (parentFolder) {
      parentFolder.files.push(file);
    }
  });

  // Build hierarchy - find root folders and organize children
  folderMap.forEach((folder, path) => {
    const pathParts = path.split('/');
    const level = pathParts.length - 3; // Adjust based on SharePoint path structure
    folder.level = Math.max(0, level);
    
    // Find parent folder
    const parentPath = pathParts.slice(0, -1).join('/');
    const parentFolder = folderMap.get(parentPath);
    
    if (parentFolder) {
      parentFolder.subFolders.push(folder);
      folder.parentPath = parentPath;
    } else {
      // This is a root folder
      rootFolders.push(folder);
    }
  });

  return rootFolders.sort((a, b) => a.Name.localeCompare(b.Name));
};

/**
 * Get random color for folders
 */
const getRandomFolderColor = (): string => {
  const colors = ["#0b0f6b", "#f58220", "#8fa3e8", "#1327ff", "#ec2f91", "#c9a36a", "#10b981", "#8b5cf6"];
  return colors[Math.floor(Math.random() * colors.length)];
};

/**
 * Alternative method to get data from the current site's document library
 * Use this if the direct URL approach above doesn't work
 */
export const getTemplatesResourcesDataFromCurrentSite = async (libraryName: string = "Documents"): Promise<IDocumentLibraryResponse> => {
  try {
    

    // First, let's try to find documents in a specific folder if it exists
    // Look for template-related folders or documents
    let libraryItems: any[] = [];
    
    try {
      // Try to get items from a specific folder path within the library
      libraryItems = await sp.web.lists
        .getByTitle(libraryName)
        .items
        .select(
          "Id,Title,FileRef,FileLeafRef,File_x0020_Type,Modified,Created,FSObjType," +
          "FileSizeDisplay,Author/Title,Author/EMail,Editor/Title,Editor/EMail"
        )
        .expand("Author,Editor")
        .filter("FSObjType eq 0 or FSObjType eq 1") // Include both files and folders
        .top(5000)();
    } catch (libraryError) {
      
      
      // Try alternative library names
      const alternativeNames = ["Shared Documents", "Site Assets", "Style Library"];
      
      for (const altName of alternativeNames) {
        try {
          libraryItems = await sp.web.lists
            .getByTitle(altName)
            .items
            .select(
              "Id,Title,FileRef,FileLeafRef,File_x0020_Type,Modified,Created,FSObjType," +
              "FileSizeDisplay,Author/Title,Author/EMail,Editor/Title,Editor/EMail"
            )
            .expand("Author,Editor")
            .filter("FSObjType eq 0 or FSObjType eq 1")
            .top(1000)()
          
          
          break;
        } catch (altError) {
          
        }
      }
    }

    

    // Separate files and folders
    const allFiles: IDocumentItem[] = libraryItems
      .filter((item: any) => item.FSObjType === 0)
      .map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Name: item.FileLeafRef,
        FileRef: item.FileRef,
        FileLeafRef: item.FileLeafRef,
        File_x0020_Type: item.File_x0020_Type,
        Modified: item.Modified,
        Created: item.Created,
        FileSizeDisplay: item.FileSizeDisplay,
        Author: item.Author ? {
          Title: item.Author.Title,
          EMail: item.Author.EMail,
        } : undefined,
        Editor: item.Editor ? {
          Title: item.Editor.Title,
          EMail: item.Editor.EMail,
        } : undefined,
      }));

    const allFolders: IFolderItem[] = libraryItems
      .filter((item: any) => item.FSObjType === 1)
      .map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Name: item.FileLeafRef,
        FileRef: item.FileRef,
        FileLeafRef: item.FileLeafRef,
        FSObjType: 1,
        Modified: item.Modified,
        Created: item.Created,
      }));

    const standardDocuments = categorizeStandardDocuments(allFiles);
    const invoicingDocuments = categorizeInvoicingDocuments(allFiles);
    const resourceFolders = categorizeResourceFolders(allFolders);

    return {
      standardDocuments,
      invoicingDocuments,
      resourceFolders,
      allFiles,
      allFolders,
      folderStructure: [], // Empty for old method
    };
  } catch (error) {
    
    
    // Return empty data structure - let the main function handle fallback to mock data
    return {
      standardDocuments: [],
      invoicingDocuments: [],
      resourceFolders: [],
      allFiles: [],
      allFolders: [],
      folderStructure: [],
    };
  }
};

/**
 * Categorize files as standard documents based on naming patterns
 */
const categorizeStandardDocuments = (files: IDocumentItem[]): ITemplateDocument[] => {
  const standardKeywords = [
    "NDA",
    "Template",
    "Charter",
    "Brief",
    "Risk",
    "Register",
    "Issues",
    "Status",
    "Report",
    "SteerCo",
    "Presentation",
    "Change",
    "Request",
    "CRF",
    "Meeting",
    "Minutes",
    "Feedback",
    "Form",
  ];

  return files
    .filter((file) =>
      standardKeywords.some((keyword) =>
        file.Name.toLowerCase().includes(keyword.toLowerCase()) ||
        (file.Title && file.Title.toLowerCase().includes(keyword.toLowerCase()))
      )
    )
    .map((file) => ({
      ...file,
      Category: "Standard Documents",
      DocumentType: determineDocumentType(file.Name),
    }));
};

/**
 * Categorize files as invoicing documents based on naming patterns
 */
const categorizeInvoicingDocuments = (files: IDocumentItem[]): ITemplateDocument[] => {
  const invoicingKeywords = ["invoice", "billing", "payment", "UK", "AUS", "KSA", "UAE"];

  return files
    .filter((file) =>
      invoicingKeywords.some((keyword) =>
        file.Name.toLowerCase().includes(keyword.toLowerCase()) ||
        (file.Title && file.Title.toLowerCase().includes(keyword.toLowerCase()))
      )
    )
    .map((file) => ({
      ...file,
      Category: "Invoicing & Reporting",
      DocumentType: "book" as const,
    }));
};

/**
 * Categorize folders as resource folders with default colors
 */
const categorizeResourceFolders = (folders: IFolderItem[]): IResourceFolder[] => {
  const folderColors = [
    "#0b0f6b",
    "#f58220",
    "#8fa3e8",
    "#1327ff",
    "#ec2f91",
    "#c9a36a",
  ];

  return folders.map((folder, index) => ({
    ...folder,
    Description: `Resources and materials in ${folder.Name}`,
    Color: folderColors[index % folderColors.length],
    IconType: "folder",
  }));
};

/**
 * Determine document type based on file extension or name
 */
const determineDocumentType = (fileName: string): "book" | "folder" => {
  const extension = fileName.split('.').pop()?.toLowerCase();
  
  switch (extension) {
    case 'pdf':
    case 'doc':
    case 'docx':
    case 'xls':
    case 'xlsx':
    case 'ppt':
    case 'pptx':
      return 'book';
    default:
      return 'book'; // Default to book icon
  }
};

/**
 * Get direct download URL for a file
 */
export const getFileDownloadUrl = (fileRef: string): string => {
  // Check if fileRef already contains the full URL
  if (fileRef.startsWith('http')) {
    return fileRef;
  }
  
  // Use the external SharePoint site URL
  const baseUrl = "https://trivandildn.sharepoint.com";
  
  // If the fileRef doesn't start with a slash, assume it's a relative path
  if (!fileRef.startsWith('/')) {
    return `${baseUrl}/sites/TrivandiLondon/Shared Documents/${fileRef}`;
  }
  
  return `${baseUrl}${fileRef}`;
};

/**
 * Get file preview URL for supported file types
 */
export const getFilePreviewUrl = (fileRef: string, fileType: string): string => {
  const baseUrl = "https://trivandildn.sharepoint.com";
  let fullFileRef = fileRef;
  
  // Check if fileRef already contains the full URL
  if (!fileRef.startsWith('http')) {
    if (!fileRef.startsWith('/')) {
      // Relative path - add the full SharePoint path
      fullFileRef = `/sites/TrivandiLondon/Shared Documents/${fileRef}`;
    } else {
      // Absolute path from SharePoint root
      fullFileRef = fileRef;
    }
  } else {
    // If it's already a full URL, extract the path for encoding
    try {
      const url = new URL(fileRef);
      fullFileRef = url.pathname;
    } catch {
      fullFileRef = fileRef;
    }
  }
  
  // For Office documents, return the preview URL with proper encoding
  if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'].includes(fileType.toLowerCase())) {
    const encodedPath = encodeURIComponent(fullFileRef);
    return `${baseUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodedPath}&action=view`;
  }
  
  // For other files, return direct download URL
  return getFileDownloadUrl(fullFileRef.startsWith('http') ? fullFileRef : `${baseUrl}${fullFileRef}`);
};

/**
 * Get the direct URL to the external SharePoint Templates folder
 */

export const getExternalTemplatesFolderUrl = (): string => {
  return "https://trivandildn.sharepoint.com/sites/TrivandiLondon/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FTrivandiLondon%2FShared%20Documents%2F6%20Projects%2F%5F2024%20Project%20Folder%20Template&viewid=2c1c2bdd-31e0-4de6-a073-ba8b5b07f4e2";
};

/**
 * Check if the current user has access to the external SharePoint site
 */
export const checkExternalSiteAccess = async (): Promise<boolean> => {
  try {
    // This is a simple check - in a real implementation you might ping the external site
    // For now, we'll assume access is not available and return false
    return false;
  } catch (error) {
    
    return false;
  }
};

/**
 * Recursively get folder structure from external SharePoint site using REST API
 */
export const getExternalFolderStructureRecursiveREST = async (
  siteUrl: string,
  folderPath: string, 
  folderName: string, 
  level: number
): Promise<IFolderStructure> => {
  try {
    // Build REST API URL for this folder
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderPath)}')?$expand=Files,Folders&$select=Name,ServerRelativeUrl,Files/Name,Files/ServerRelativeUrl,Files/Length,Files/TimeLastModified,Folders/Name,Folders/ServerRelativeUrl,Folders/ItemCount`;
    
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      credentials: 'include'
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const data = await response.json();
    const folderData = data.d;
    
    const files: IDocumentItem[] = [];
    const subFolders: IFolderStructure[] = [];
    
    // Process files in this folder
    if (folderData.Files && folderData.Files.results) {
      folderData.Files.results.forEach((file: any) => {
        const fileExtension = file.Name.split('.').pop()?.toLowerCase() || '';
        files.push({
          Id: file.ServerRelativeUrl.hashCode(),
          Title: file.Name,
          Name: file.Name,
          FileRef: file.ServerRelativeUrl,
          FileLeafRef: file.Name,
          File_x0020_Type: fileExtension,
          Modified: file.TimeLastModified,
          Created: file.TimeLastModified
        });
      });
    }
    
    // Process subfolders recursively (limit depth to avoid infinite recursion)
    if (folderData.Folders && folderData.Folders.results && level < 3) {
      for (const subFolder of folderData.Folders.results) {
        if (subFolder.Name && !subFolder.Name.startsWith('_')) {
          try {
            const subFolderStructure = await getExternalFolderStructureRecursiveREST(
              siteUrl,
              subFolder.ServerRelativeUrl,
              subFolder.Name,
              level + 1
            );
            subFolders.push(subFolderStructure);
          } catch (subError) {
            
          }
        }
      }
    }
    
    return {
      Id: folderPath.hashCode(),
      Name: folderName,
      Title: folderName,
      FileRef: folderPath,
      FSObjType: 1,
      Modified: new Date().toISOString(),
      Created: new Date().toISOString(),
      files: files,
      subFolders: subFolders,
      level: level
    };
    
  } catch (error) {
    
    return {
      Id: folderPath.hashCode(),
      Name: folderName,
      Title: folderName,
      FileRef: folderPath,
      FSObjType: 1,
      Modified: new Date().toISOString(),
      Created: new Date().toISOString(),
      files: [],
      subFolders: [],
      level: level
    };
  }
};

/**
 * Get file category based on file extension
 */
export const getFileCategory = (fileType: string): 'template' | 'resource' | 'other' => {
  const templateExtensions = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'pdf'];
  const resourceExtensions = ['jpg', 'png', 'gif', 'svg', 'mp4', 'mp3', 'zip', 'rar'];
  
  if (templateExtensions.includes(fileType.toLowerCase())) {
    return 'template';
  } else if (resourceExtensions.includes(fileType.toLowerCase())) {
    return 'resource';
  } else {
    return 'other';
  }
};

/**
 * Get templates and documents from the TemplatesandDocuments list
 */
export const getTemplatesAndDocumentsListData = async (): Promise<Array<{
  id: string;
  title: string;
  description: string;
  link: string | null;
  hasChildren: boolean;
}>> => {
  try {
    
    
    // Fetch all items including ParentName
    const allItems = await sp.web.lists
      .getByTitle("TemplatesandDocuments")
      .items
      .select("ID,Title,Description,Link,ParentName").orderBy("SortOrder", true)();
    
    

    // Separate parents (no ParentName) from children
    const parentItems = allItems.filter((item: any) => !item.ParentName || item.ParentName.trim() === "");
    const childItems = allItems.filter((item: any) => item.ParentName && item.ParentName.trim() !== "");

    // Build set of parent titles that have children
    const parentTitlesWithChildren = new Set(childItems.map((c: any) => c.ParentName.trim()));

    return parentItems.map((item: any) => ({
      id: `template-${item.ID}`,
      title: item.Title || "",
      description: item.Description || "",
      link: item.Link || null,
      hasChildren: parentTitlesWithChildren.has((item.Title || "").trim())
    }));
  } catch (error) {
    
    return [];
  }
};

export const getTemplatesChildItemsByParent = async (parentTitle: string): Promise<Array<{
  id: string;
  title: string;
  link: string | null;
}>> => {
  try {
    const items = await sp.web.lists
      .getByTitle("TemplatesandDocuments")
      .items
      .select("ID,Title,Link,ParentName")
      .filter(`ParentName eq '${parentTitle}'`)
      .orderBy("SortOrder", true)();

    return items.map((item: any) => ({
      id: `child-${item.ID}`,
      title: item.Title || "",
      link: item.Link || null
    }));
  } catch (error) {
    
    return [];
  }
};

/**
 * Get root folders from a document library with item counts
 */
export const getRootFoldersWithCounts = async (libraryName: string = "Documents"): Promise<Array<{
  id: number;
  name: string;
  itemCount: number;
  modified: string;
  serverRelativeUrl: string;
  color: string;
}>> => {
  try {
    

    // Get all folders from the library
    const folders = await sp.web.lists
      .getByTitle(libraryName)
      .items
      .select("Id,FileLeafRef,FileRef,Modified,FSObjType,FileDirRef")
      .filter("FSObjType eq 1")
      .top(5000)();

    

    // Get the library root path
    const library = await sp.web.lists.getByTitle(libraryName).rootFolder();
    const libraryRootPath = library.ServerRelativeUrl;

    // Filter to get only root folders
    const rootFolders = folders.filter((folder: any) => {
      const parentPath = folder.FileDirRef;
      return parentPath === libraryRootPath;
    });

    

    // Get item counts for each folder
    const foldersWithCounts = await Promise.all(
      rootFolders.map(async (folder: any) => {
        try {
          // Get all items in this folder (not recursive)
          const folderPath = folder.FileRef;
          const folderItems = await sp.web.getFolderByServerRelativePath(folderPath).files();
          const subfolders = await sp.web.getFolderByServerRelativePath(folderPath).folders();
          
          const totalCount = folderItems.length + subfolders.length;
          
          return {
            id: folder.Id,
            name: folder.FileLeafRef,
            itemCount: totalCount,
            modified: folder.Modified,
            serverRelativeUrl: folder.FileRef,
            color: getRandomFolderColor()
          };
        } catch (error) {
          
          return {
            id: folder.Id,
            name: folder.FileLeafRef,
            itemCount: 0,
            modified: folder.Modified,
            serverRelativeUrl: folder.FileRef,
            color: getRandomFolderColor()
          };
        }
      })
    );

    return foldersWithCounts.sort((a, b) => a.name.localeCompare(b.name));
  } catch (error) {
    
    return [];
  }
};

/**
 * Get empty response when no data is found
 */
export const getEmptyResponse = (): IDocumentLibraryResponse => {
  return {
    standardDocuments: [],
    invoicingDocuments: [],
    resourceFolders: [],
    allFiles: [],
    allFolders: [],
    folderStructure: []
  };
};