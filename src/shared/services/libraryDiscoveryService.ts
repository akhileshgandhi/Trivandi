/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

let sp: SPFI;
const libraryCache = new Map<string, string>();

export const initLibraryDiscoveryService = (spInstance: SPFI): void => {
  sp = spInstance;
};

/**
 * Find project library by searching for libraries that match the project
 * Searches by Code, Title, or Code-Title combinations
 * Results are cached to avoid repeated API calls
 */
export const findProjectLibraryByProject = async (
  projectCode?: string,
  projectTitle?: string
): Promise<string | null> => {
  try {
    // Create cache key
    const cacheKey = `${projectCode || ''}-${projectTitle || ''}`;

    // Check cache first
    if (libraryCache.has(cacheKey)) {
      
      return libraryCache.get(cacheKey)!;
    }

    // Get all document libraries
    const lists = await sp.web.lists
      .filter("BaseTemplate eq 101 and Hidden eq false")
      .select("Title")
      .top(5000)();

    

    // Search patterns to try (in order of preference)
    const searchPatterns = [
      projectCode && projectTitle ? `${projectCode}-${projectTitle}` : null,
      projectCode && projectTitle ? `${projectCode} ${projectTitle}` : null,
      projectCode && projectTitle ? `${projectCode}_${projectTitle}` : null,
      projectTitle,
      projectCode
    ].filter(Boolean) as string[];

    

    // Try exact match first
    for (const pattern of searchPatterns) {
      const exactMatch = lists.find(l => l.Title === pattern);
      if (exactMatch) {
        
        libraryCache.set(cacheKey, exactMatch.Title);
        return exactMatch.Title;
      }
    }

    // Try partial match (library contains the pattern)
    for (const pattern of searchPatterns) {
      const partialMatch = lists.find(l => l.Title.includes(pattern));
      if (partialMatch) {
        
        libraryCache.set(cacheKey, partialMatch.Title);
        return partialMatch.Title;
      }
    }

    
    return null;
  } catch (error) {
    
    return null;
  }
};

/**
 * Get or construct library name with fallback
 */
export const getProjectLibraryName = async (
  projectCode?: string,
  projectTitle?: string
): Promise<string> => {
  const foundLibrary = await findProjectLibraryByProject(projectCode, projectTitle);

  if (foundLibrary) {
    return foundLibrary;
  }

  // Fallback to constructed name
  const fallbackName = projectCode && projectTitle
    ? `${projectCode}-${projectTitle}`
    : projectTitle || projectCode || '';

  
  return fallbackName;
};

/**
 * Clear the library cache (useful for testing or refresh scenarios)
 */
export const clearLibraryCache = (): void => {
  libraryCache.clear();
};

export const getProjectById = async (projectId: any) => {
  try {
    const project = await sp.web.lists.getByTitle("ProjectsNew").items.getById(projectId).select("*")();
    let arr = {
      "ProjectDocumentsUrl": project.ProjectDocumentsUrl,
      "BidDocumentsUrl": project.BidDocumentsUrl,
      "ContractsDocumentsUrl": project.ContractsDocumentsUrl,
      "ProjectTitle": project.ProjectTitle,
      "Code": project.Code,
      "Status": project.Status,
      "NonCmap": project.NonCmap,
      "ProjectID": project.ProjectID
    }
    return arr;
  } catch (error) {
    
    return '';
  }
};