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
      console.log('Using cached library name:', libraryCache.get(cacheKey));
      return libraryCache.get(cacheKey)!;
    }

    // Get all document libraries
    const lists = await sp.web.lists
      .filter("BaseTemplate eq 101 and Hidden eq false")
      .select("Title")
      .top(5000)();

    console.log('Available libraries:', lists.map(l => l.Title));

    // Search patterns to try (in order of preference)
    const searchPatterns = [
      projectCode && projectTitle ? `${projectCode}-${projectTitle}` : null,
      projectCode && projectTitle ? `${projectCode} ${projectTitle}` : null,
      projectCode && projectTitle ? `${projectCode}_${projectTitle}` : null,
      projectTitle,
      projectCode
    ].filter(Boolean) as string[];

    console.log('Search patterns:', searchPatterns);

    // Try exact match first
    for (const pattern of searchPatterns) {
      const exactMatch = lists.find(l => l.Title === pattern);
      if (exactMatch) {
        console.log('Found exact match:', exactMatch.Title);
        libraryCache.set(cacheKey, exactMatch.Title);
        return exactMatch.Title;
      }
    }

    // Try partial match (library contains the pattern)
    for (const pattern of searchPatterns) {
      const partialMatch = lists.find(l => l.Title.includes(pattern));
      if (partialMatch) {
        console.log('Found partial match:', partialMatch.Title);
        libraryCache.set(cacheKey, partialMatch.Title);
        return partialMatch.Title;
      }
    }

    console.warn('No matching library found for project');
    return null;
  } catch (error) {
    console.error('Error finding project library:', error);
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

  console.warn('Library not found, using fallback:', fallbackName);
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
    const project = await sp.web.lists.getByTitle("ProjectsNew").items.getById(projectId)();
    let arr = {
      "ProjectDocumentsUrl": project.ProjectDocumentsUrl,
      "BidDocumentsUrl": project.BidDocumentsUrl,
      "ContractsDocumentsUrl": project.ContractsDocumentsUrl,
      "ProjectTitle": project.ProjectTitle,
      "Status": project.Status,
      "NonCmap": project.NonCmap
    }
    return arr;
  } catch (error) {
    console.error('Error getting project by ID:', error);
    return '';
  }
};