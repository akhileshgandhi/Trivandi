import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import { Web } from '@pnp/sp/webs';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/sharing';
import { ISharedDocument, IGuestUser, IGuestAccessReportEntry, ISharedFileAccessLog } from '../components/IProjectExternalPortalState';
import { getProjectDocuments } from '../../../shared/services/projectService';
import { findProjectLibraryByProject } from '../../../shared/services/libraryDiscoveryService';
export interface IInviteGuestData {
  name: string;
  email: string;
  company?: string;
  role?: 'Viewer' | 'Editor';
  accessDuration?: number;
  phone?: string;
  projectId?: string;
}
export class ProjectExternalPortalService {
  private context: WebPartContext;
  private siteUrl: string;
  private sp: SPFI;
  private projectLibraryName: string | null = null;
  private externalLibraryName: string = 'ExternalShareDocument';
  private currentProjectId: string | null = null;
  private cache: Map<string, { data: any; timestamp: number }> = new Map();
  private pendingRequests: Map<string, Promise<any>> = new Map();
  private readonly CACHE_DURATION = 60000; // 1 minute cache
  private readonly RETRY_DELAY = 2000; // 2 seconds retry delay for throttle
  private readonly MAX_RETRIES = 3;
  private _isRestrictedGuestCache: boolean | null = null;
  private _isSiteOwnerCache: boolean | null = null;
  private _libraryCheckCache: { exists: boolean; hasAccess: boolean; error?: string } | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
    this.sp = spfi().using(SPFx(context as any));
  }

  /**
   * Set the current project ID for scoping all portal operations
   */
  public setProjectId(projectId: string | number): void {
    this.currentProjectId = String(projectId);
    // Clear cache so next load picks up the project-scoped folder
    this.cache.clear();
    
  }

  /**
   * Get the project-scoped root folder path (e.g. "Project-1109")
   */
  public getProjectFolderPath(): string {
    return this.currentProjectId ? `Project-${this.currentProjectId}` : '';
  }

  /**
   * Check if ExternalShareDocument library exists and user has access
   */
  public async checkExternalLibraryAccess(): Promise<{ exists: boolean; hasAccess: boolean; error?: string }> {
    if (this._libraryCheckCache !== null) return this._libraryCheckCache;

    try {
      const libraryName = this.externalLibraryName;

      // First check if library exists
      const lists = await this.sp.web.lists();
      const libraryExists = lists.some(list => list.Title === libraryName);

      if (!libraryExists) {
        
        const result = { exists: false, hasAccess: false, error: 'Library does not exist' };
        this._libraryCheckCache = result;
        return result;
      }

      // Check if user has access to the library
      try {
        const list = this.sp.web.lists.getByTitle(libraryName);
        await list.getCurrentUserEffectivePermissions();
        
        const result = { exists: true, hasAccess: true };
        this._libraryCheckCache = result;
        return result;
      } catch (accessError) {
        
        const result = { exists: true, hasAccess: false, error: 'Access denied' };
        this._libraryCheckCache = result;
        return result;
      }
    } catch (error) {
      
      return { exists: false, hasAccess: false, error: (error as any)?.message };
    }
  }

  /**
   * Find project library by searching for libraries that match the project
   * Delegates to centralized library discovery service
   */
  public async findProjectLibrary(projectCode?: string, projectTitle?: string): Promise<string | null> {
    return await findProjectLibraryByProject(projectCode, projectTitle);
  }

  /**
   * Set project library name for external portal operations
   */
  public setProjectLibraryName(libraryName: string): void {
    this.externalLibraryName = libraryName;
    
  }

  private _toWebRelativePath(path: string): string {
    return path.startsWith('/') ? path.substring(1) : path;
  }

  /**
   * Get cached data or fetch new data
   */
  private async _getCachedOrFetch<T>(cacheKey: string, fetchFn: () => Promise<T>, skipCache: boolean = false): Promise<T> {
    // Check cache first
    if (!skipCache && this.cache.has(cacheKey)) {
      const cached = this.cache.get(cacheKey);
      if (Date.now() - cached.timestamp < this.CACHE_DURATION) {
        
        return cached.data as T;
      }
    }

    // Check if request is already pending
    if (this.pendingRequests.has(cacheKey)) {
      
      return this.pendingRequests.get(cacheKey) as Promise<T>;
    }

    // Make new request
    const promise = this._executeWithRetry(fetchFn)
      .then((data) => {
        this.cache.set(cacheKey, { data, timestamp: Date.now() });
        this.pendingRequests.delete(cacheKey);
        return data;
      })
      .catch((error) => {
        this.pendingRequests.delete(cacheKey);
        throw error;
      });

    this.pendingRequests.set(cacheKey, promise);
    return promise;
  }

  /**
   * Execute function with retry logic for throttle errors
   */
  private async _executeWithRetry<T>(fn: () => Promise<T>, retryCount: number = 0): Promise<T> {
    try {
      return await fn();
    } catch (error) {
      // Check if it's a throttle error (429)
      const isThrottle = (error as any)?.status === 429 ||
        (error as any)?.message?.includes('429') ||
        (error as any)?.message?.includes('throttle') ||
        (error as any)?.message?.includes('Too Many Requests');

      if (isThrottle && retryCount < this.MAX_RETRIES) {
        const delay = this.RETRY_DELAY * Math.pow(2, retryCount); // Exponential backoff
        
        await this._delay(delay);
        return this._executeWithRetry(fn, retryCount + 1);
      }

      throw error;
    }
  }

  /**
   * Delay helper for retry logic
   */
  private _delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Parse server relative URL to extract site and folder path
   * @param serverRelativeUrl - URL like /sites/sitename/library/folder
   */
  private _parseSiteAndFolder(serverRelativeUrl: string): { siteUrl: string; folderPath: string } {
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
  }

  /**
   * Check if a server relative URL is from the current site
   */
  private _isCurrentSiteUrl(serverRelativeUrl: string): boolean {
    if (!serverRelativeUrl) return false;

    const currentSiteUrl = this.siteUrl;
    const urlObj = new URL(currentSiteUrl);
    const currentPathname = urlObj.pathname;

    // If current site is root (/), only root URLs should match
    if (currentPathname === '/') {
      return serverRelativeUrl.startsWith('/') && !serverRelativeUrl.startsWith('/sites/');
    }

    // Otherwise, URL should start with the current site's path
    return serverRelativeUrl.startsWith(currentPathname);
  }

  /**
   * Download file from the correct context (handling both current and external sites)
   */
  private async _downloadFileBuffer(serverRelativeUrl: string, fileName: string): Promise<ArrayBuffer> {
    try {
      if (this._isCurrentSiteUrl(serverRelativeUrl)) {
        // File is from current site - use current context
        
        const sourceFile = this.sp.web.getFileByServerRelativePath(serverRelativeUrl);
        return await sourceFile.getBuffer();
      } else {
        // File is from external site - need to create context for that site
        

        // Parse the URL to get the site
        const { siteUrl } = this._parseSiteAndFolder(serverRelativeUrl);

        if (!siteUrl) {
          throw new Error(`Cannot determine site URL from: ${serverRelativeUrl}`);
        }

        // Build absolute URL for the external site
        const urlObj = new URL(this.siteUrl);
        const absoluteSiteUrl = `${urlObj.protocol}//${urlObj.hostname}${siteUrl}`;

        

        // Get the web context for the external site
        const externalWeb = Web([this.sp.web, absoluteSiteUrl]);

        // Download the file from the external site using its server relative URL
        const sourceFile = externalWeb.getFileByServerRelativePath(serverRelativeUrl);
        return await sourceFile.getBuffer();
      }
    } catch (error) {
      
      throw error;
    }
  }

  /**
   * Clear cache for a specific key or all cache
   */
  public clearCache(cacheKey?: string): void {
    if (cacheKey) {
      
      this.cache.delete(cacheKey);
      this.pendingRequests.delete(cacheKey);
    } else {
      
      this.cache.clear();
      this.pendingRequests.clear();
    }
  }

  /**
   * Get shared documents from External_Shared_Area library
   */
  public async getSharedDocuments(folderPath: string): Promise<ISharedDocument[]> {
    try {
      const webRelPath = this._toWebRelativePath(folderPath);

      const folder = this.sp.web.getFolderByServerRelativePath(webRelPath);
      const subFolders = await folder.folders();
      const files = await folder.files();

      const documents: ISharedDocument[] = [];

      // Folders
      subFolders.forEach((f: any) => {
        if (!f.Name?.startsWith('.')) {
          documents.push({
            id: f.UniqueId,
            name: f.Name,
            modified: new Date(f.TimeLastModified),
            modifiedBy: '',
            path: `${folderPath}/${f.Name}`,
            fileRef: f.ServerRelativeUrl,
            isFolder: true
          });
        }
      });

      // Files
      for (const file of files) {
        const itemApi: any = this.sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).getItem();
        const item = await itemApi
          .select('Id', 'Modified', 'Editor/Title', 'ExternalSharingType', 'SharedWithGuests')
          .expand('Editor')();

        documents.push({
          id: item.Id,
          name: file.Name,
          modified: new Date(item.Modified),
          modifiedBy: item.Editor?.Title || 'Unknown',
          path: `${folderPath}/${file.Name}`,
          fileRef: file.ServerRelativeUrl,
          isFolder: false,
          externalSharingType: item.ExternalSharingType,
          sharedWithGuests: item.SharedWithGuests
        });
      }

      return documents;
    } catch (error) {
      
      throw new Error(`Failed to load documents: ${(error as any)?.message}`);
    }
  }

  /**
   * Get active guest users (with caching)
   */
  public async getActiveGuests(projectId?: number, skipCache: boolean = false): Promise<IGuestUser[]> {
    const cacheKey = `guests_${projectId || 'all'}`;

    return this._getCachedOrFetch(cacheKey, async () => {
      try {
        // Build filter: always filter by Active status, optionally by ProjectId
        let filterStr = "Status eq 'Active'";
        const pid = projectId || (this.currentProjectId ? Number(this.currentProjectId) : undefined);
        if (pid) {
          filterStr += ` and ProjectId eq '${pid}'`;
        }

        const items = await this.sp.web.lists.getByTitle('ExternalGuestAccess').items
          .select('Id', 'Title', 'Email', 'Company', 'Role', 'Status', 'LastAccessDate', 'AccessExpiryDate', 'InvitedDate', 'ProjectId')
          .filter(filterStr)
          .orderBy('LastAccessDate', false)
          .top(5000)();

        return items.map((item: any) => ({
          id: item.Id,
          title: item.Title,
          email: item.Email,
          company: item.Company,
          role: item.Role,
          status: item.Status,
          lastAccessDate: item.LastAccessDate ? new Date(item.LastAccessDate) : undefined,
          accessExpiryDate: item.AccessExpiryDate ? new Date(item.AccessExpiryDate) : undefined,
          invitedDate: new Date(item.InvitedDate)
        }));
      } catch (error) {
        
        throw new Error(`Failed to load guests: ${(error as any)?.message}`);
      }
    }, skipCache);
  }

  /**
   * Get project documents for import
   */
  //  public async getProjectDocuments(): Promise<any[]> {
  //    try {
  //      const items = await this.sp.web.lists.getByTitle('ProjectDocuments').items
  //        .select('Id', 'FileLeafRef', 'FileDirRef', 'AvailableForExternalSharing')
  //        .filter('AvailableForExternalSharing eq 1')
  //        .top(50)();

  //      return items.map((item: any) => ({
  //        id: item.Id,
  //        name: item.FileLeafRef,
  //        location: item.FileDirRef.split('/').pop() || '/'
  //      }));
  //    } catch (error) {
  //      
  //      // Mock data fallback
  //      return [
  //        { id: 1, name: 'TRO Sustainability Training', location: '/Project_Docs' },
  //        { id: 2, name: 'TRO Sustainability Training', location: '/Project_Docs' },
  //        { id: 3, name: 'TRO Sustainability Training', location: '/Project_Docs' },
  //        { id: 4, name: 'TRO Sustainability Training', location: '/All_Admin' },
  //        { id: 5, name: 'TRO Sustainability Training', location: '/All_Admin' }
  //      ];
  //    }
  //  }
  public async getProjectDocuments(
    libraryName: string,
    folderPath: string = "",
    page: number = 1,
    pageSize: number = 10
  ): Promise<any[]> {
    if (!this.sp) throw new Error("PnPjs not initialized");

    try {
      

      if (!folderPath) {
        const result = await getProjectDocuments(libraryName, folderPath, page, pageSize);
        const mappedRootItems = result.items;

        
        return mappedRootItems;
      } else {
        const result = await getProjectDocuments(libraryName, folderPath, page, pageSize);
        const mappedFolderItems = result.items;

        
        return mappedFolderItems;
      }
    } catch (error) {
      
      return [];
    }
  }
  /**
   * Get project documents and folders for import (with caching)
   * Uses project-specific document library
   */
  public async getProjectDocumentsForImport(
    folderPath: string = '',
    page: number = 1,
    pageSize: number = 10
  ): Promise<any[]> {
    if (!this.projectLibraryName) {
      
      return [];
    }

    const cacheKey = `project_documents_import_${this.projectLibraryName}_${folderPath}`;

    return this._getCachedOrFetch(cacheKey, async () => {
      try {
        const items = await this.getProjectDocuments(this.projectLibraryName, folderPath, page, pageSize);

        return items.map((item: any) => ({
          id: item.Id,
          name: item.FileLeafRef,
          location: folderPath || '/',
          isFolder: item.FSObjType === 1,
          serverRelativeUrl: item.FileRef
        }));
      } catch (error) {
        
        return [];
      }
    });
  }

  /**
   * Get documents from all 3 document types for import (Project, Bid, Contract Documents)
   * This method fetches documents from ProjectDocumentsUrl, BidDocumentsUrl, and ContractsDocumentsUrl
   */
  public async getAllProjectDocumentsForImport(
    projectId: string,
    folderPath: string = '',
    page: number = 1,
    pageSize: number = 10
  ): Promise<any[]> {
    const cacheKey = `all_project_documents_import_${projectId}_${folderPath}`;

    return this._getCachedOrFetch(cacheKey, async () => {
      try {
        // Get project details with URL fields
        const projectData = await this.getProjectById(projectId);

        if (!projectData || typeof projectData !== 'object') {
          
          return [];
        }

        // Import from libraryDiscoveryService to get URLs
        const libraryService = await import('../../../shared/services/libraryDiscoveryService');
        const projectDetails = await libraryService.getProjectById(parseInt(projectId));

        if (!projectDetails || typeof projectDetails !== 'object') {
          
          return [];
        }

        const projectDocUrl = projectDetails.ProjectDocumentsUrl || '';
        const bidDocUrl = projectDetails.BidDocumentsUrl || '';
        const contractDocUrl = projectDetails.ContractsDocumentsUrl || '';

        

        let allDocuments: any[] = [];

        // Import getDocumentsByServerRelativeUrl from projectService
        const projectService = await import('../../../shared/services/projectService');
        const getDocumentsByServerRelativeUrl = projectService.getDocumentsByServerRelativeUrl;

        // Load documents from Project Documents source
        if (projectDocUrl) {
          try {
            const result = await getDocumentsByServerRelativeUrl(projectDocUrl, folderPath, page, pageSize);
            const projectDocs = result.items;
            const mappedProjectDocs = projectDocs.map((item: any) => ({
              id: item.Id,
              name: item.FileLeafRef,
              location: folderPath || '/',
              isFolder: item.FSObjType === 1,
              serverRelativeUrl: item.FileRef,
              documentType: 'Project Documents'
            }));
            allDocuments.push(...mappedProjectDocs);
          } catch (error) {
            
          }
        } else {
          // Fallback: Try to discover library name if URL is missing
          try {
            const discoveredLibrary = await libraryService.findProjectLibraryByProject(projectDetails.Code, projectDetails.ProjectTitle);
            if (discoveredLibrary) {
              const result = await projectService.getProjectDocuments(discoveredLibrary, folderPath, page, pageSize);
              const projectDocs = result.items;
              const mappedProjectDocs = projectDocs.map((item: any) => ({
                id: item.Id,
                name: item.FileLeafRef,
                location: folderPath || '/',
                isFolder: item.FSObjType === 1,
                serverRelativeUrl: item.FileRef,
                documentType: 'Project Documents'
              }));
              allDocuments.push(...mappedProjectDocs);
            }
          } catch (fallbackError) {
            
          }
        }

        // Load documents from Bid Documents URL
        if (bidDocUrl) {
          try {
            const result = await getDocumentsByServerRelativeUrl(bidDocUrl, folderPath, page, pageSize);
            const bidDocs = result.items;
            const mappedBidDocs = bidDocs.map((item: any) => ({
              id: item.Id + 10000, // Offset ID to avoid conflicts
              name: item.FileLeafRef,
              location: folderPath || '/',
              isFolder: item.FSObjType === 1,
              serverRelativeUrl: item.FileRef,
              documentType: 'Bid Documents'
            }));
            allDocuments.push(...mappedBidDocs);
          } catch (error) {
            
          }
        }

        // Load documents from Contract Documents URL
        if (contractDocUrl) {
          try {
            const result = await getDocumentsByServerRelativeUrl(contractDocUrl, folderPath, page, pageSize);
            const contractDocs = result.items;
            const mappedContractDocs = contractDocs.map((item: any) => ({
              id: item.Id + 20000, // Offset ID to avoid conflicts
              name: item.FileLeafRef,
              location: folderPath || '/',
              isFolder: item.FSObjType === 1,
              serverRelativeUrl: item.FileRef,
              documentType: 'Contract Documents'
            }));
            allDocuments.push(...mappedContractDocs);
          } catch (error) {
            
          }
        }

        // If no documents found, return empty array
        if (allDocuments.length === 0) {
          
          return [];
        }

        
        return allDocuments;

      } catch (error) {
        
        return [];
      }
    });
  }



  /**
   * Get project by ID
   */
  public async getProjectById(projectId: string): Promise<any> {
    try {
      const item = await this.sp.web.lists.getByTitle('ProjectsNew').items
        .getById(parseInt(projectId))
        .select('Id', 'Title', 'Company', 'Status')();

      return {
        id: item.Id,
        title: item.Title,
        company: item.Company,
        status: item.Status
      };
    } catch (error) {
      
      return null;
    }
  }

  /**
   * Upload and share a document
   */
  public async uploadAndShareDocument(
    targetPath: string,
    file: File,
    externalSharing: string,
    sharedWith: string[],
    accessDuration: number
  ): Promise<void> {
    try {
      const folderRelPath = this._toWebRelativePath(targetPath);

      const uploadResult = await this.sp.web
        .getFolderByServerRelativePath(folderRelPath)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      const item = await uploadResult.file.getItem();

      const expiryDate = accessDuration > 0
        ? new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000))
        : null;

      await item.update({
        ExternalSharingType: externalSharing,
        AccessExpiryDate: expiryDate
      });

      await this._logShareActivity(
        'Uploaded',
        file.name,
        uploadResult.data.ServerRelativeUrl,
        sharedWith,
        accessDuration
      );
    } catch (error) {
      
      throw new Error(`Failed to upload document: ${(error as any)?.message}`);
    }
  }

  /**
   * Upload document without sharing
   */
  public async uploadDocument(
    targetPath: string,
    file: File
  ): Promise<void> {
    try {
      const folderRelPath = this._toWebRelativePath(targetPath);

      const uploadResult = await this.sp.web
        .getFolderByServerRelativePath(folderRelPath)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      await this._logShareActivity(
        'Uploaded',
        file.name,
        uploadResult.data.ServerRelativeUrl,
        [],
        0
      );
    } catch (error) {
      
      throw new Error(`Failed to upload document: ${(error as any)?.message}`);
    }
  }

  /**
   * Share access to existing documents (placeholder)
   */
  public async shareAccess(
    documentPath: string,
    users: string[],
    accessDuration: number
  ): Promise<void> {
    try {
      await this._logShareActivity(
        'Access Granted',
        documentPath,
        documentPath,
        users,
        accessDuration
      );
    } catch (error) {
      
      throw new Error(`Failed to share access: ${(error as any)?.message}`);
    }
  }

  /**
   * Share document access with specific users
   */
  public async shareDocumentAccess(
    fileRef: string,
    guestEmails: string[],
    accessDuration: number,
    customExpiryDate?: Date,
    permission: 'Read' | 'Review' | 'Edit' | 'Admin' = 'Read'
  ): Promise<void> {
    try {
      const expiryDate = customExpiryDate || (accessDuration > 0
        ? new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000))
        : null);

      // Update item with expiry date (item-level permission time)
      if (expiryDate) {
        try {
          const file = this.sp.web.getFileByServerRelativePath(fileRef);
          const item = await file.getItem();
          await item.update({
            AccessExpiryDate: expiryDate.toISOString()
          });
        } catch (itemError) {
          
        }
      }

      // Log the share activity with permission and computed expiry date
      await this._logShareActivity(
        'Shared',
        fileRef.split('/').pop() || '',
        fileRef,
        guestEmails,
        accessDuration,
        permission,
        expiryDate || undefined
      );

      
    } catch (error) {
      
      throw new Error(`Failed to share document: ${(error as any)?.message}`);
    }
  }

  /**
   * Share folder access with specific users.
   * Guests who already have active (non-expired) access to this folder OR
   * to any individual file inside it are automatically skipped — no duplicate
   * log entries are created for them.
   */
  public async shareFolderAccess(
    folderPath: string,
    guestEmails: string[],
    accessDuration: number,
    customExpiryDate?: Date,
    permission: 'Read' | 'Review' | 'Edit' | 'Admin' = 'Read'
  ): Promise<void> {
    try {
      const expiryDate = customExpiryDate || (accessDuration > 0
        ? new Date(Date.now() + (accessDuration * 24 * 60 * 60 * 1000))
        : null);

      // Determine which guests already have active access to this folder or any
      // file within it (individually shared), so we can skip them.
      const alreadyCovered = await this._getGuestsWithActiveAccessToFolder(folderPath, guestEmails);
      const guestsToShare = guestEmails.filter(
        e => !alreadyCovered.includes(e.trim().toLowerCase())
      );

      if (guestsToShare.length === 0) {
        
        return;
      }

      if (alreadyCovered.length > 0) {
        
      }

      // Log the share activity only for guests who don't already have active access
      await this._logShareActivity(
        'Shared',
        folderPath.split('/').pop() || '',
        folderPath,
        guestsToShare,
        accessDuration,
        permission,
        expiryDate || undefined
      );

      
    } catch (error) {
      
      throw new Error(`Failed to share folder: ${(error as any)?.message}`);
    }
  }

  /**
   * Returns the lowercase emails (from the supplied list) that already have
   * active, non-expired access to the given folder path OR to any individual
   * file stored under that folder path in SharedDocumentLog.
   */
  private async _getGuestsWithActiveAccessToFolder(
    folderPath: string,
    guestEmails: string[]
  ): Promise<string[]> {
    if (!guestEmails || guestEmails.length === 0) return [];

    const now = new Date();
    const escaped = folderPath.replace(/'/g, "''");

    try {
      // Match: exact folder log entry  OR  any per-file log entry inside the folder
      const logs = await this.sp.web.lists
        .getByTitle('SharedDocumentLog')
        .items
        .select('Document', 'GuestEmail', 'ActionDate', 'AccessDuration', 'ExpiryDate')
        .filter(
          `(Document eq '${escaped}' or startswith(Document, '${escaped}/')) and ActionType eq 'Shared'`
        )
        .orderBy('ActionDate', false)
        .top(5000)();

      if (!logs || logs.length === 0) return [];

      // Keep only the MOST RECENT entry per (email, document) pair
      const latestMap = new Map<string, any>();
      for (const log of logs) {
        const docUrl = (log.Document || '').toLowerCase();
        const emails: string[] = (log.GuestEmail || '')
          .split(';')
          .map((e: string) => e.trim().toLowerCase())
          .filter(Boolean);

        for (const email of emails) {
          const key = `${email}||${docUrl}`;
          if (!latestMap.has(key)) {
            latestMap.set(key, log);
          }
        }
      }

      // For each input guest, check whether at least one log entry is still active
      const guestsWithAccess: string[] = [];
      for (const rawEmail of guestEmails) {
        const email = rawEmail.trim().toLowerCase();

        for (const [key, entry] of latestMap) {
          if (!key.startsWith(`${email}||`)) continue;

          let isActive = false;
          if (entry.ExpiryDate) {
            isActive = new Date(entry.ExpiryDate) > now;
          } else if (entry.AccessDuration > 0 && entry.ActionDate) {
            const expiry = new Date(
              new Date(entry.ActionDate).getTime() + entry.AccessDuration * 24 * 60 * 60 * 1000
            );
            isActive = expiry > now;
          } else if (entry.AccessDuration === 0) {
            isActive = true; // permanent
          }

          if (isActive) {
            guestsWithAccess.push(email);
            break; // one active entry is enough to skip this guest
          }
        }
      }

      return guestsWithAccess;
    } catch (err) {
      
      return []; // fail-open: if check fails, allow the share to proceed
    }
  }

  /**
   * Import documents from source library to external library (supports files and folders recursively)
   */
  public async importDocuments(
    documents: any[],
    targetPath: string
  ): Promise<void> {
    try {
      
      
      const libraryName = this.externalLibraryName;
      let cleanTargetPath = targetPath.startsWith('/') ? targetPath.slice(1) : targetPath;
      if (cleanTargetPath.startsWith(libraryName)) {
        const slashIndex = cleanTargetPath.indexOf('/');
        cleanTargetPath = slashIndex !== -1 ? cleanTargetPath.substring(slashIndex + 1) : "";
      }
      
      const fullTargetPath = `${libraryName}/${cleanTargetPath}`.replace(/\/+/g, '/').replace(/\/$/, "");

      for (const doc of documents) {
        await this._importItemRecursive(doc, fullTargetPath);
      }
      
      this.clearCache();
    } catch (error) {
      
      throw new Error(`Failed to import documents: ${(error as any)?.message}`);
    }
  }

  /**
   * Helper to import an item (file or folder) recursively
   */
  private async _importItemRecursive(item: any, targetParentPath: string): Promise<void> {
    const itemName = item.FileLeafRef || item.name || item.Title;
    const itemPath = item.FileRef || item.path || item.ServerRelativeUrl;
    
    // Improved folder detection: Check FSObjType or look for lack of extension
    const hasExtension = itemName.indexOf('.') !== -1;
    const isFolder = item.FSObjType === 1 || (!hasExtension && !item.File_x0020_Type);
    
    try {
      if (isFolder) {
        
        const newFolderPath = `${targetParentPath}/${itemName}`.replace(/\/+/g, '/');
        await this.ensureFolderPathExists(newFolderPath);
        
        // Fetch children from source
        
        
        // Parse source site and library from path
        const { siteUrl, folderPath } = this._parseSiteAndFolder(itemPath);
        const sourceWeb = siteUrl ? Web([this.sp.web, `${window.location.protocol}//${window.location.host}${siteUrl}`]) : this.sp.web;
        
        const [files, folders] = await Promise.all([
          sourceWeb.getFolderByServerRelativePath(itemPath).files.select("Name", "ServerRelativeUrl", "UniqueId").expand("ListItemAllFields")(),
          sourceWeb.getFolderByServerRelativePath(itemPath).folders.select("Name", "ServerRelativeUrl", "UniqueId").expand("ListItemAllFields").filter("Name ne 'Forms'")()
        ]);
        
        const children = [
          ...folders.map(f => ({ ...f, name: f.Name, path: f.ServerRelativeUrl, FSObjType: 1 })),
          ...files.map(f => ({ ...f, name: f.Name, path: f.ServerRelativeUrl, FSObjType: 0 }))
        ];

        

        if (children.length > 0) {
          for (const child of children) {
            await this._importItemRecursive(child, newFolderPath);
          }
        }
      } else {
        
        const fileBuffer = await this._downloadFileBuffer(itemPath, itemName);
        await this.ensureFolderPathExists(targetParentPath);
        
        await this.sp.web.getFolderByServerRelativePath(targetParentPath)
          .files.addUsingPath(itemName, fileBuffer, { Overwrite: true });
          
        const targetUrl = `${this.siteUrl}/${targetParentPath}/${itemName}`.replace(/\/+/g, '/');
        const serverRelativeTargetUrl = new URL(targetUrl).pathname;
        await this._logShareActivity('Imported', itemName, serverRelativeTargetUrl, [], 0);
      }
    } catch (err) {
      
    }
  }

  /**
   * Ensures a full folder path exists in the external library
   */
  public async ensureFolderPathExists(fullPath: string): Promise<void> {
    const parts = fullPath.split('/').filter(Boolean);
    let currentPath = "";
    
    for (const part of parts) {
      currentPath = currentPath ? `${currentPath}/${part}` : part;
      await this.ensureFolderExists(currentPath);
    }
  }

  /**
   * Add a user (by email) to a named SharePoint site group.
   * Uses PnP ensureUser so it works for both internal and B2B guest accounts.
   */
  public async addUserToSharePointGroup(userEmail: string, groupName: string): Promise<void> {
    try {
      // ensureUser registers the account in the site user info list and returns the
      // resolved SP user object including their claim-based LoginName.
      const userResult = await this.sp.web.ensureUser(userEmail);
      const loginName: string = userResult.data?.LoginName;
      if (!loginName) {
        
        return;
      }
      

      await this.sp.web.siteGroups.getByName(groupName).users.add(loginName);
      
    } catch (error) {
      
      // Non-fatal — the B2B invitation email was already sent
    }
  }

  /**
   * Invite guest user
   */
  public async inviteGuest(guestData: IInviteGuestData): Promise<void> {
    try {
      const normalizedEmail = guestData.email.trim().toLowerCase();
      const safeEmail = normalizedEmail.replace(/'/g, "''");

      // Scope the duplicate check to the same project so the same guest can be
      // invited to multiple different projects independently.
      let dupFilter = `Email eq '${safeEmail}' and Status eq 'Active'`;
      if (guestData.projectId) {
        const safePid = String(guestData.projectId).replace(/'/g, "''");
        dupFilter += ` and ProjectId eq '${safePid}'`;
      }

      const existingGuests = await this.sp.web.lists
        .getByTitle('ExternalGuestAccess')
        .items
        .select('Id', 'Email', 'ProjectId')
        .filter(dupFilter)
        .top(1)();

      if (existingGuests && existingGuests.length > 0) {
        throw new Error('This email is already invited to this project.');
      }

      const itemData: any = {
        Title: guestData.name,
        Email: guestData.email,
        Status: 'Active',
        InvitedDate: new Date().toISOString(),
        InvitedById: this.context.pageContext.legacyPageContext.userId
      };

      // Role and AccessDuration are optional at invite time — they are set when documents are shared
      if (guestData.role) itemData.Role = guestData.role;
      if (guestData.accessDuration && guestData.accessDuration > 0) {
        itemData.AccessDuration = guestData.accessDuration;
        itemData.AccessExpiryDate = new Date(Date.now() + (guestData.accessDuration * 24 * 60 * 60 * 1000)).toISOString();
      }

      if (guestData.company) itemData.Company = guestData.company;
      if (guestData.phone) itemData.PhoneNumber = guestData.phone;
      if (guestData.projectId) itemData.ProjectId = guestData.projectId;

      await this.sp.web.lists.getByTitle('ExternalGuestAccess').items.add(itemData);

      // Clear all guests cache entries to ensure fresh data is fetched
      this.clearCache(`guests_${guestData.projectId || 'all'}`);
      this.clearCache(`guests_all`); // Also clear the 'all' cache
      
    } catch (error) {
      
      throw new Error(`Failed to invite guest: ${(error as any)?.message}`);
    }
  }

  private async _logShareActivity(
    actionType: string,
    documentName: string,
    documentUrl: string,
    sharedWith: string[],
    accessDuration: number,
    permission: 'Read' | 'Review' | 'Edit' | 'Admin' = 'Read',
    expiryDate?: Date
  ): Promise<void> {
    try {
      

      // Enhanced document ID retrieval from ExternalShareDocument library
      let docId: number | null = null;

      try {
        // First, try to get as a file
        if (!documentUrl || documentUrl.includes('undefined')) {
          
          return;
        }

        
        // Ensure path is server relative for these calls if it's currently a full URL
        const serverRelativeUrl = documentUrl.startsWith('http') ? new URL(documentUrl).pathname : documentUrl;
        const file = this.sp.web.getFileByServerRelativePath(serverRelativeUrl);
        const fileItem = await file.getItem();

        if (fileItem) {
          docId = (fileItem as any).Id || (fileItem as any).ID || (fileItem as any).id;
          if (docId) {
            
          }
        }
      } catch (fileError) {
        

        // If file lookup fails, try to get as a folder
        try {
          const folder = this.sp.web.getFolderByServerRelativePath(documentUrl);
          const folderItem = await folder.listItemAllFields();

          if (folderItem) {
            docId = (folderItem as any).Id || (folderItem as any).ID || (folderItem as any).id;
            if (docId) {
              
            }
          }
        } catch (folderError) {
          

          // If both fail, try to find by searching in the ExternalShareDocument library
          try {
            const libraryName = this.externalLibraryName;
            const items = await this.sp.web.lists
              .getByTitle(libraryName)
              .items
              .select('Id', 'FileRef', 'FileDirRef', 'FileLeafRef')
              .filter(`FileRef eq '${documentUrl.replace(/'/g, "''")}'`)();

            if (items && items.length > 0) {
              docId = items[0].Id;
              if (docId) {
                
              }
            } else {
              // Try searching by filename if full path search fails
              const fileName = documentUrl.substring(documentUrl.lastIndexOf('/') + 1);
              if (fileName) {
                const fileItems = await this.sp.web.lists
                  .getByTitle(libraryName)
                  .items
                  .select('Id', 'FileRef', 'FileLeafRef')
                  .filter(`FileLeafRef eq '${fileName.replace(/'/g, "''")}'`)();

                if (fileItems && fileItems.length > 0) {
                  // If multiple files with same name, try to match the path
                  const matchingItem = fileItems.find(item => item.FileRef === documentUrl) || fileItems[0];
                  docId = matchingItem.Id;
                  if (docId) {
                    
                  }
                }
              }
            }
          } catch (searchError) {
            
          }
        }
      }

      // Prepare log data
      const logData: any = {
        Title: `${actionType} - ${documentName}`,
        Document: documentUrl,
        DocumentName: documentName,
        ActionType: actionType,
        SharedBy: this.context.pageContext.user.displayName,
        ActionDate: new Date().toISOString(),
        AccessDuration: accessDuration,
        GuestEmail: sharedWith.join('; ')
      };

      // Populate GuestUserName — resolve display names from ExternalGuestAccess
      // for each guest email being shared with. Falls back to the email address.
      if (sharedWith.length > 0) {
        try {
          const emailFilters = sharedWith
            .map(e => `Email eq '${e.replace(/'/g, "''")}'`)
            .join(' or ');
          const guestRecords = await this.sp.web.lists
            .getByTitle('ExternalGuestAccess')
            .items
            .select('Title', 'Email')
            .filter(emailFilters)
            .top(sharedWith.length + 10)();
          const nameMap = new Map<string, string>(
            (guestRecords as any[]).map((r: any) => [r.Email?.toLowerCase(), r.Title])
          );
          const resolvedNames = sharedWith.map(
            e => nameMap.get(e.toLowerCase()) || e
          );
          logData.GuestUserName = resolvedNames.join('; ');
        } catch (nameErr: any) {
          
          logData.GuestUserName = sharedWith.join('; ');
        }
      }

      // Populate ProjectName from ProjectsNew list
      if (this.currentProjectId) {
        try {
          const project = await this.sp.web.lists
            .getByTitle('ProjectsNew')
            .items
            .getById(parseInt(this.currentProjectId))
            .select('Title')();
          if (project?.Title) {
            logData.ProjectName = project.Title;
          }
        } catch (projErr: any) {
          
        }
      }

      // Store computed expiry date for precise time-based access checks
      if (expiryDate) {
        logData.ExpiryDate = expiryDate.toISOString();
      }

      // Store project ID for project-scoped filtering
      if (this.currentProjectId) {
        logData.ProjectId = this.currentProjectId;
      }

      // Add Permission field - will be ignored if field doesn't exist in list
      try {
        logData.Permission = permission;
      } catch (e) {
        
      }

      // Add DocId if we found it
      if (docId) {
        logData.DocId = docId;
        
      } else {
        
      }

      // Add the log entry
      let result;
      try {
        result = await this.sp.web.lists.getByTitle('SharedDocumentLog').items.add(logData);
      } catch (addError: any) {
        // If Permission field doesn't exist, retry without it
        if (addError.message && addError.message.includes('Permission') && addError.message.includes('does not exist')) {
          
          delete logData.Permission;
          result = await this.sp.web.lists.getByTitle('SharedDocumentLog').items.add(logData);
        } else {
          throw addError;
        }
      }

      // Clear shared documents cache to ensure updated access is reflected immediately
      // This is especially important for restricted external guests who see documents based on SharedDocumentLog
      // Clear all cached shared documents entries for different folder paths
      const cacheKeysToDelete: string[] = [];
      for (const [key] of this.cache) {
        if (key.startsWith(`shared_docs_${this.externalLibraryName}_`)) {
          cacheKeysToDelete.push(key);
        }
      }
      cacheKeysToDelete.forEach(key => this.cache.delete(key));

      return result.data?.Id;
    } catch (error) {
      
      throw new Error(`Failed to log sharing activity: ${(error as any)?.message}`);
    }
  }

  /**
   * Get shared documents from ExternalShareDocument library
   */
  /**
   * Get shared documents from ExternalShareDocument library (with caching)
   * This is used for the Documents tab
   * Only returns documents the current user has access to
   * For users in ExternalGuestAccess, only shows documents explicitly shared with them
   */






  public async getSharedDocumentsFromExternalLibrary(
    folderPath: string = '',
    page: number = 1,
    pageSize: number = 10
  ): Promise<{ items: ISharedDocument[]; totalCount: number }> {
    const libraryName = this.externalLibraryName;
    const cacheKey = `shared_docs_${libraryName}_${folderPath}_p${page}_s${pageSize}`;

    // Always skip cache for first page or when refreshing
    const shouldSkipCache = page === 1;

    return this._getCachedOrFetch(cacheKey, async () => {
      try {
        // Parallelize initial checks
        const [libraryCheck, isRestrictedGuest] = await Promise.all([
          this.checkExternalLibraryAccess(),
          this.isCurrentUserRestrictedGuest()
        ]);

        // Ensure the project-specific folder exists (admin users only; guests can't create folders)
        if (!isRestrictedGuest && folderPath && folderPath.trim() !== '') {
          try {
            await this.ensureFolderExists(folderPath);
          } catch (folderErr) {
            
          }
        }

        const result = await getProjectDocuments(libraryName, folderPath, page, pageSize);
        const items = result.items;
        const totalCount = result.totalCount;
        

        const documents: ISharedDocument[] = items.map((item: any) => ({
          id: item.Id,
          name: item.FileLeafRef,
          modified: new Date(item.Modified),
          modifiedBy: item.Editor?.Title || 'Unknown',
          path: folderPath ? `${folderPath}/${item.FileLeafRef}` : item.FileLeafRef,
          fileRef: item.FileRef,
          isFolder: item.FSObjType === 1
        }));

        // If user is a restricted external guest, filter documents based on SharedDocumentLog
        if (isRestrictedGuest) {
          
          const filteredDocuments = await this.filterDocumentsForRestrictedUser(documents);
          
          return { items: filteredDocuments, totalCount: result.totalCount };
        }

        // For internal users and unrestricted guests, proceed with normal permission checks
        const accessibleDocuments: ISharedDocument[] = [];

        // Cache library-level permissions once
        const [isOwner, libraryPermission] = await Promise.all([
          this._checkIfUserIsSiteOwner(),
          this._checkSharePointPermissions()
        ]);

        const currentUserPermission = isOwner ? 'Admin' : libraryPermission;
        const ALL_PERMISSIONS = ['Read', 'Review', 'Edit', 'Admin'] as const;

        if (currentUserPermission && ALL_PERMISSIONS.includes(currentUserPermission as any)) {
          
          for (const doc of documents) {
            doc.permission = currentUserPermission as typeof ALL_PERMISSIONS[number];
            accessibleDocuments.push(doc);
          }
        } else {
          // Check individual document permissions in parallel
          
          const permissionPromises = documents.map(async (doc) => {
            try {
              // We already checked owner/library permissions, so we can use a more direct method or pass them
              // For now, let's just make sure getUserPermissionForDocument is as fast as possible
              const perm = await this.getUserPermissionForDocument(doc.fileRef);
              if (perm && ALL_PERMISSIONS.includes(perm as any)) {
                doc.permission = perm as typeof ALL_PERMISSIONS[number];
                return doc;
              }
              return null;
            } catch (err) {
              
              return null;
            }
          });

          const results = await Promise.all(permissionPromises);
          accessibleDocuments.push(...results.filter((d): d is ISharedDocument => d !== null));
        }

        
        return { items: accessibleDocuments, totalCount: result.totalCount };
      } catch (error) {
        

        // Check if it's a "list not found" error
        if ((error as any)?.message && (error as any).message.includes('does not exist')) {
          
          return { items: [], totalCount: 0 };
        }

        throw error;
      }
    }, shouldSkipCache);
  }

  public async getSharedDocumentsFromExternalArea(
    folderPath: string,
    page: number = 1,
    pageSize: number = 10
  ): Promise<{ items: ISharedDocument[]; totalCount: number }> {
    try {
      const { getProjectDocuments } = await import('../../../shared/services/projectService');
      const libraryName = this.externalLibraryName;
      const result = await getProjectDocuments(libraryName, folderPath, page, pageSize);
      const items = result.items;

      if (!items || items.length === 0) {
        return { items: [], totalCount: 0 };
      }

      const documents = items.map((item: any) => ({
        id: item.Id,
        name: item.FileLeafRef,
        modified: new Date(item.Modified),
        modifiedBy: item.Editor?.Title || 'Unknown',
        path: folderPath ? `${folderPath}/${item.FileLeafRef}` : item.FileLeafRef,
        fileRef: item.FileRef,
        isFolder: item.FSObjType === 1
      }));

      return { items: documents, totalCount: result.totalCount };
    } catch (error) {
      
      return { items: [], totalCount: 0 };
    }
  }

  /**
   * Ensure folder exists in ExternalShareDocument library, create if it doesn't exist
   */
  public async ensureFolderExists(folderPath: string): Promise<void> {
    try {
      const libraryName = this.externalLibraryName;
      const rootFolder = await this.sp.web.lists.getByTitle(libraryName).rootFolder();
      const rootFolderPath = rootFolder.ServerRelativeUrl;

      if (!folderPath || folderPath === '/' || folderPath.trim() === '') {
        return; // Root folder always exists
      }

      // Clean the folder path and deduplicate library name
      let cleanPath = folderPath.startsWith('/') ? folderPath.substring(1) : folderPath;
      // libraryName is already declared above at line 1444
      
      if (cleanPath.startsWith(libraryName)) {
        const slashIndex = cleanPath.indexOf('/');
        cleanPath = slashIndex !== -1 ? cleanPath.substring(slashIndex + 1) : "";
      }

      const fullFolderPath = `${rootFolderPath}/${cleanPath}`.replace(/\/+/g, '/').replace(/\/$/, "");

      try {
        // Try to get the folder first
        await this.sp.web.getFolderByServerRelativePath(fullFolderPath)();
        
      } catch (error) {
        if ((error as any)?.status === 404) {
          // Folder doesn't exist, create it recursively
          

          const folders = cleanPath.split('/');
          let currentPath = rootFolderPath;

          for (const folderName of folders) {
            if (folderName.trim() === '') continue;

            currentPath = `${currentPath}/${folderName}`;
            try {
              await this.sp.web.getFolderByServerRelativePath(currentPath)();
              
            } catch (innerError) {
              if ((innerError as any)?.status === 404) {
                
                const parentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
                await this.sp.web.getFolderByServerRelativePath(parentPath).folders.addUsingPath(folderName);
              } else {
                throw innerError;
              }
            }
          }
        } else {
          throw error;
        }
      }
    } catch (error) {
      
      throw new Error(`Failed to create folder structure: ${(error as any)?.message}`);
    }
  }

  /**
   * Upload document to ExternalShareDocument library
   */
  public async uploadToExternalShareDocument(targetPath: string, file: File): Promise<string> {
    try {
      const libraryName = this.externalLibraryName;

      // Get the root folder of the library to get its server relative URL
      const rootFolder = await this.sp.web.lists.getByTitle(libraryName).rootFolder();
      const rootFolderPath = rootFolder.ServerRelativeUrl;

      // If targetPath is provided and not root, upload to that folder
      if (targetPath && targetPath !== '/' && targetPath.trim() !== '') {
        // Ensure the target folder exists first
        await this.ensureFolderExists(targetPath);

        // Construct full server-relative path
        const fullFolderPath = `${rootFolderPath}${targetPath.startsWith('/') ? targetPath : '/' + targetPath}`;

        

        const folder = await this.sp.web.getFolderByServerRelativePath(fullFolderPath);
        const uploadResult = await folder.files.addUsingPath(file.name, file, { Overwrite: true });

        

        // Clear cache for this specific folder AND all related folders
        const cacheKeyForFolder = `shared_docs_${libraryName}_${targetPath}`;
        this.cache.delete(cacheKeyForFolder);
        

        // Also clear parent folder caches if in a subfolder
        let currentPath = targetPath;
        while (currentPath && currentPath !== '/') {
          const parentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
          const parentCacheKey = `shared_docs_${libraryName}_${parentPath || '/'}`;
          this.cache.delete(parentCacheKey);
          
          currentPath = parentPath;
        }

        return uploadResult.data.ServerRelativeUrl;
      } else {
        // Upload to root folder of the library
        const uploadResult = await this.sp.web.lists
          .getByTitle(libraryName)
          .rootFolder
          .files
          .addUsingPath(file.name, file, { Overwrite: true });

        

        // Clear cache for root and all folders
        const rootCacheKey = `shared_docs_${libraryName}_`;
        for (const key of Array.from(this.cache.keys())) {
          if (key.startsWith(rootCacheKey)) {
            this.cache.delete(key);
            
          }
        }

        return uploadResult.data.ServerRelativeUrl;
      }
    } catch (error) {
      
      throw new Error(`Failed to upload document: ${(error as any)?.message}`);
    }
  }

  /**
   * Check whether a list of guests already have active (non-expired) access to each of the
   * given document paths.  Returns an array of conflict descriptions so the caller can
   * surface a precise error to the user.
   *
   * A conflict exists when the MOST RECENT "Shared" log entry for a (guest, document) pair
   * has a permission AND its access has not yet expired.
   */
  public async checkExistingActiveAccess(
    documentPaths: string[],
    guestEmails: string[]
  ): Promise<Array<{ guestEmail: string; documentName: string; expiresOn: string }>> {
    const conflicts: Array<{ guestEmail: string; documentName: string; expiresOn: string }> = [];
    const now = new Date();

    try {
      // Build an OData filter for the relevant documents and guests
      const docFilters = documentPaths
        .map(p => `Document eq '${p.replace(/'/g, "''")}'`)
        .join(' or ');

      const logs = await this.sp.web.lists
        .getByTitle('SharedDocumentLog')
        .items
        .select('Document', 'DocumentName', 'GuestEmail', 'Permission', 'ActionType', 'ActionDate', 'AccessDuration', 'ExpiryDate')
        .filter(`(${docFilters}) and ActionType eq 'Shared'`)
        .orderBy('ActionDate', false)
        .top(5000)();

      if (!logs || logs.length === 0) return conflicts;

      // Keep only the MOST RECENT entry per (guest, document) pair
      const latestMap = new Map<string, any>();
      for (const log of logs) {
        const docUrl = (log.Document || '').toLowerCase();
        const emails: string[] = (log.GuestEmail || '')
          .split(';')
          .map((e: string) => e.trim().toLowerCase())
          .filter(Boolean);

        for (const email of emails) {
          const key = `${email}||${docUrl}`;
          if (!latestMap.has(key)) {
            latestMap.set(key, { ...log, _resolvedEmail: email });
          }
        }
      }

      // Check each (guest, document) pair that we care about
      const normalizedGuests = guestEmails.map(e => e.trim().toLowerCase());

      for (const docPath of documentPaths) {
        const docPathLower = docPath.toLowerCase();
        const docName = docPath.split('/').pop() || docPath;

        for (const guestEmail of normalizedGuests) {
          const key = `${guestEmail}||${docPathLower}`;
          const entry = latestMap.get(key);
          if (!entry || !entry.Permission) continue;

          // Determine whether current access is still valid
          let isStillActive = false;
          let expiresOn = 'No expiry';

          if (entry.ExpiryDate) {
            const expiry = new Date(entry.ExpiryDate);
            if (expiry > now) {
              isStillActive = true;
              expiresOn = expiry.toLocaleString('en-GB', {
                day: '2-digit', month: 'short', year: 'numeric',
                hour: '2-digit', minute: '2-digit', hour12: true
              });
            }
          } else if (entry.AccessDuration && entry.AccessDuration > 0 && entry.ActionDate) {
            const expiry = new Date(
              new Date(entry.ActionDate).getTime() + entry.AccessDuration * 24 * 60 * 60 * 1000
            );
            if (expiry > now) {
              isStillActive = true;
              expiresOn = expiry.toLocaleString('en-GB', {
                day: '2-digit', month: 'short', year: 'numeric',
                hour: '2-digit', minute: '2-digit', hour12: true
              });
            }
          } else if (entry.AccessDuration === 0) {
            // Permanent access — always active
            isStillActive = true;
            expiresOn = 'No expiry (permanent)';
          }

          if (isStillActive) {
            conflicts.push({ guestEmail, documentName: docName, expiresOn });
          }
        }
      }
    } catch (err) {
      
    }

    return conflicts;
  }

  /**
   * Get active guests for sharing
   */
  public async getActiveGuestsForSharing(): Promise<Array<{ email: string; name: string; }>> {
    try {
      // Scope to the current project so share dialogs only show guests for this project
      let filter = "Status eq 'Active'";
      if (this.currentProjectId) {
        filter += ` and ProjectId eq '${this.currentProjectId}'`;
      }

      const items = await this.sp.web.lists.getByTitle('ExternalGuestAccess').items
        .select('Title', 'Email', 'ProjectId')
        .filter(filter)
        .top(200)();

      return items.map((item: any) => ({
        email: item.Email,
        name: item.Title
      }));
    } catch (error) {
      
      return [];
    }
  }

  /**
   * Copy documents from Project Documents to ExternalShareDocument library
   */
  public async copyDocumentsToExternalLibrary(documents: any[], targetPath: string): Promise<void> {
    try {
      const libraryName = this.externalLibraryName;

      // Get root folder to construct proper paths
      const rootFolder = await this.sp.web.lists.getByTitle(libraryName).rootFolder();
      const rootFolderPath = rootFolder.ServerRelativeUrl;

      for (const doc of documents) {
        try {
          // Construct the target folder path
          let destFolderPath = rootFolderPath;
          if (targetPath && targetPath !== '/' && targetPath.trim() !== '') {
            destFolderPath = `${rootFolderPath}${targetPath.startsWith('/') ? targetPath : '/' + targetPath}`;
          }

          if (doc.isFolder) {
            // Create folder in the target path
            const targetFolderPath = `${destFolderPath}/${doc.name}`;
            try {
              const destFolder = await this.sp.web.getFolderByServerRelativePath(destFolderPath);
              await destFolder.folders.addUsingPath(doc.name);
              
            } catch (e) {
              
            }
          } else {
            // Copy file to target path - handle both current site and external site URLs
            

            const fileBuffer = await this._downloadFileBuffer(doc.serverRelativeUrl, doc.name);

            const targetFolder = await this.sp.web.getFolderByServerRelativePath(destFolderPath);
            await targetFolder.files.addUsingPath(doc.name, fileBuffer, { Overwrite: true });

            
          }

          // Log the import activity
          await this._logShareActivity(
            'Imported',
            doc.name,
            `${destFolderPath}/${doc.name}`,
            [],
            0
          );
        } catch (error) {
          
          throw error; // Re-throw so caller knows about individual file failures
        }
      }

      // Clear cache for the target path and all related paths
      const cacheKeyForPath = `shared_docs_${libraryName}_${targetPath}`;
      this.cache.delete(cacheKeyForPath);
      

      // Also clear parent folder caches
      let currentPath = targetPath;
      while (currentPath && currentPath !== '/') {
        const parentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
        const parentCacheKey = `shared_docs_${libraryName}_${parentPath || '/'}`;
        this.cache.delete(parentCacheKey);
        
        currentPath = parentPath;
      }

      // Clear all root level caches
      const rootCacheKey = `shared_docs_${libraryName}_`;
      for (const key of Array.from(this.cache.keys())) {
        if (key.startsWith(rootCacheKey)) {
          this.cache.delete(key);
          
        }
      }
    } catch (error) {
      
      throw new Error(`Failed to copy documents: ${(error as any)?.message}`);
    }
  }

  /**
   * Get user's permission for a specific document from SharePoint permissions and SharedDocumentLog
   */
  public async getUserPermissionForDocument(documentUrl: string): Promise<'Read' | 'Review' | 'Edit' | 'Admin' | null> {
    try {
      // Prioritize user.email so B2B guest email matches the plain email stored in logs
      const currentUserEmail = this.context.pageContext.user.email ||
        this.context.pageContext.user.loginName ||
        this.context.pageContext.user.displayName;

      // Check if user is site owner or admin - they should have full access
      const isOwner = await this._checkIfUserIsSiteOwner();
      if (isOwner) {
        
        return 'Admin'; // Site owners/admins always have Admin permission
      }

      // First, check actual SharePoint permissions on the ExternalShareDocument library
      try {
        const sharePointPermission = await this._checkSharePointPermissions();
        if (sharePointPermission) {
          
          return sharePointPermission;
        }
      } catch (permError) {
        
      }

      // Fallback: Query SharedDocumentLog to find this user's permission for external sharing
      const escapedUrl = documentUrl.replace(/'/g, "''");
      const logs = await this.sp.web.lists
        .getByTitle('SharedDocumentLog')
        .items
        .select('GuestEmail', 'Permission', 'ActionType', 'ActionDate', 'AccessDuration', 'ExpiryDate')
        .filter(`Document eq '${escapedUrl}' and ActionType eq 'Shared'`)
        .orderBy('ActionDate', false)();

      if (!logs || logs.length === 0) {
        
        return null;
      }

      // Look for current user in the GuestEmail field
      // Check both email formats (with and without domain)
      const userEmailVariants = [
        currentUserEmail.toLowerCase(),
        currentUserEmail.toLowerCase().split('@')[0], // Just the username part
        currentUserEmail.toLowerCase().replace('i:0#.f|membership|', '') // Remove SharePoint prefix if present
      ];

      let userPermission: 'Read' | 'Review' | 'Edit' | 'Admin' | null = null;
      const now = new Date();

      for (const log of logs) {
        const guestEmails = log.GuestEmail ?
          log.GuestEmail.split(';').map((e: string) => e.trim().toLowerCase()) :
          [];

        const hasAccess = userEmailVariants.some(variant =>
          guestEmails.some(email =>
            email === variant ||
            email.includes(variant) ||
            variant.includes(email.split('@')[0])
          )
        );

        if (hasAccess && log.Permission) {
          // Check access expiry: prefer stored ExpiryDate, fall back to ActionDate + AccessDuration
          let isAccessValid = true;
          if (log.ExpiryDate) {
            const expiryDate = new Date(log.ExpiryDate);
            isAccessValid = now <= expiryDate;
            if (!isAccessValid) {
              
              continue; // Try next log entry; skip expired grants
            }
          } else if (log.AccessDuration && log.AccessDuration > 0 && log.ActionDate) {
            const actionDate = new Date(log.ActionDate);
            const expiryDate = new Date(actionDate.getTime() + (log.AccessDuration * 24 * 60 * 60 * 1000));
            isAccessValid = now <= expiryDate;
            if (!isAccessValid) {
              
              continue; // Skip expired grants
            }
          }

          userPermission = log.Permission as 'Read' | 'Review' | 'Edit' | 'Admin';
          
          break; // Take the most recent non-expired permission
        }
      }

      if (!userPermission) {
        
      }

      return userPermission;
    } catch (error) {
      
      return null;
    }
  }

  /**
   * Check SharePoint permissions for the current user on the ExternalShareDocument library
   */
  private async _checkSharePointPermissions(): Promise<'Read' | 'Review' | 'Edit' | 'Admin' | null> {
    try {
      const libraryName = this.externalLibraryName;

      // First check if library exists and is accessible
      const libraryCheck = await this.checkExternalLibraryAccess();
      if (!libraryCheck.exists || !libraryCheck.hasAccess) {
        
        return null;
      }

      // Check permissions on the ExternalShareDocument library
      const list = this.sp.web.lists.getByTitle(libraryName);
      const listPerms = await list.getCurrentUserEffectivePermissions();

      // Check for specific permission levels
      // Full Control: High bit 0, Low bit 0x40000001
      const hasFullControl = (listPerms.High & 0x1) !== 0;

      // Contribute: Low bit 0x40000000 (includes add, edit, delete items)
      const hasContribute = (listPerms.Low & 0x40000000) !== 0;

      // Edit: Low bit 0x6 (EditListItems + AddListItems)
      const hasEdit = (listPerms.Low & 0x6) === 0x6;

      // Read: Low bit 0x1
      const hasRead = (listPerms.Low & 0x1) !== 0;

      

      if (hasFullControl) {
        
        return 'Admin';
      } else if (hasContribute || hasEdit) {
        
        return 'Edit';
      } else if (hasRead) {
        
        return 'Read';
      }

      
      return null;
    } catch (error) {
      

      // For external users or when library doesn't exist, return null
      return null;
    }
  }

  /**
   * Check if current user is site owner or has full control permissions
   */
  private async _checkIfUserIsSiteOwner(): Promise<boolean> {
    if (this._isSiteOwnerCache !== null) return this._isSiteOwnerCache;

    try {
      // Try to check current user's groups
      const web = this.sp.web;
      const currentUser = await web.currentUser();

      // Get site owners group
      const ownerGroup = await web.associatedOwnerGroup();
      if (ownerGroup) {
        const siteGroups = await web.siteGroups.getById(ownerGroup.Id).users();
        const isOwner = siteGroups.some((member: any) =>
          member.Id === currentUser.Id ||
          member.LoginName === currentUser.LoginName
        );

        if (isOwner) {
          return true;
        }
      }

      // Check if user is site admin by trying to get site admin role
      const userPerms = await web.getCurrentUserEffectivePermissions();
      // Check if has manage web permission (bit 12)
      const hasManageWeb = (userPerms.Low & 268435456) !== 0 || (userPerms.High & 1) !== 0;

      this._isSiteOwnerCache = hasManageWeb;
      return hasManageWeb;
    } catch (error) {
      
      // If we cannot verify, assume they're not owner (safer default)
      return false;
    }
  }

  /**
   * Get user's permission for the current folder path in ExternalShareDocument
   */
  public async getUserPermissionForFolderPath(folderPath: string): Promise<'Read' | 'Review' | 'Edit' | 'Admin' | null> {
    try {
      const libraryName = this.externalLibraryName;

      // First check if user has library-level permissions
      const libraryPermission = await this._checkSharePointPermissions();
      if (libraryPermission) {
        
        return libraryPermission;
      }

      // If no library permission, check specific folder path
      if (folderPath && folderPath !== '' && folderPath !== '/') {
        const rootFolder = await this.sp.web.lists.getByTitle(libraryName).rootFolder();
        const rootFolderPath = rootFolder.ServerRelativeUrl;
        const fullFolderPath = `${rootFolderPath}${folderPath.startsWith('/') ? folderPath : '/' + folderPath}`;
        return this.getUserPermissionForDocument(fullFolderPath);
      }

      return null;
    } catch (error) {
      
      return null;
    }
  }

  /**
   * Get access activity log for a specific shared document or folder.
   */
  public async getDocumentAccessLog(documentUrl: string, guestEmail?: string): Promise<ISharedFileAccessLog[]> {
    try {
      const escapedUrl = (documentUrl || '').replace(/'/g, "''");
      const emailFilter = guestEmail
        ? ` and GuestEmail eq '${guestEmail.replace(/'/g, "''")}'`
        : '';
      const logs = await this.sp.web.lists
        .getByTitle('SharedDocumentLog')
        .items
        .select('Title', 'ActionType', 'Document', 'DocumentName', 'SharedBy', 'GuestEmail', 'Permission', 'AccessDuration', 'ActionDate', 'ExpiryDate')
        .filter(`Document eq '${escapedUrl}'${emailFilter}`)
        .orderBy('ActionDate', false)
        .top(500)();

      return logs.map((log: any) => ({
        title: log.Title || '',
        actionType: log.ActionType || '',
        document: log.Document || '',
        documentName: log.DocumentName || '',
        sharedBy: log.SharedBy || '',
        guestEmail: log.GuestEmail || '',
        permission: (['Read', 'Review', 'Edit', 'Admin'] as const).includes(log.Permission) ? log.Permission as 'Read' | 'Review' | 'Edit' | 'Admin' : undefined,
        accessDuration: typeof log.AccessDuration === 'number' ? log.AccessDuration : undefined,
        expiryDate: log.ExpiryDate ? new Date(log.ExpiryDate) : undefined,
        actionDate: log.ActionDate ? new Date(log.ActionDate) : undefined
      }));
    } catch (error) {
      
      throw new Error(`Failed to load access log: ${(error as any)?.message}`);
    }
  }

  /**
   * Build access report for active guests and shared file permissions.
   * Uses latest "Shared" log entry per guest + file.
   */
  public async getGuestAccessReport(projectId?: number, skipCache: boolean = false): Promise<IGuestAccessReportEntry[]> {
    // Check if current user is restricted - external users should not have access to full reports
    const isRestricted = await this.isCurrentUserRestrictedGuest();
    if (isRestricted) {
      
      throw new Error('You do not have permission to access guest access reports.');
    }

    const cacheKey = `guest_access_report_${projectId || 'all'}`;

    return this._getCachedOrFetch(cacheKey, async () => {
      const guests = await this.getActiveGuests(projectId, skipCache);
      const guestMap = new Map<string, IGuestUser>();

      guests.forEach((guest) => {
        guestMap.set((guest.email || '').trim().toLowerCase(), guest);
      });

      // Scope SharedDocumentLog to this project so report only shows this project's entries
      const reportPid = projectId || (this.currentProjectId ? Number(this.currentProjectId) : undefined);
      let logFilter = "ActionType eq 'Shared'";
      if (reportPid) {
        logFilter += ` and ProjectId eq '${reportPid}'`;
      }

      const logs = await this.sp.web.lists
        .getByTitle('SharedDocumentLog')
        .items
        .select('Document', 'DocumentName', 'GuestEmail', 'Permission', 'ActionType', 'ActionDate', 'ProjectId')
        .filter(logFilter)
        .orderBy('ActionDate', false)
        .top(5000)();

      const latestByGuestAndDoc = new Map<string, IGuestAccessReportEntry>();

      logs.forEach((log: any) => {
        const guestEmails = (log.GuestEmail || '')
          .split(';')
          .map((email: string) => email.trim().toLowerCase())
          .filter((email: string) => !!email);

        guestEmails.forEach((email: string) => {
          const guest = guestMap.get(email);
          if (!guest) {
            return;
          }

          const documentPath = log.Document || '';
          const key = `${email}|${documentPath}`;
          if (latestByGuestAndDoc.has(key)) {
            return;
          }

          latestByGuestAndDoc.set(key, {
            guestName: guest.title || '',
            guestEmail: guest.email || '',
            guestRole: guest.role,
            guestCompany: guest.company || '',
            guestStatus: guest.status,
            fileName: log.DocumentName || (documentPath ? documentPath.split('/').pop() : '') || '',
            filePath: documentPath,
            permission: (['Read', 'Review', 'Edit', 'Admin'] as const).includes(log.Permission)
              ? (log.Permission as 'Read' | 'Review' | 'Edit' | 'Admin')
              : 'Read',
            sharedOn: log.ActionDate ? new Date(log.ActionDate) : undefined,
            guestLastAccess: guest.lastAccessDate
          });
        });
      });

      const rows = Array.from(latestByGuestAndDoc.values());
      const guestsWithoutSharedFiles = guests
        .filter((guest) => !rows.some((row) => row.guestEmail.toLowerCase() === guest.email.toLowerCase()))
        .map((guest) => ({
          guestName: guest.title || '',
          guestEmail: guest.email || '',
          guestRole: guest.role,
          guestCompany: guest.company || '',
          guestStatus: guest.status,
          fileName: '',
          filePath: '',
          permission: 'No Access' as const,
          sharedOn: undefined,
          guestLastAccess: guest.lastAccessDate
        }));

      return [...rows, ...guestsWithoutSharedFiles].sort((a, b) => {
        const emailCompare = a.guestEmail.localeCompare(b.guestEmail);
        if (emailCompare !== 0) {
          return emailCompare;
        }

        return a.fileName.localeCompare(b.fileName);
      });
    }, skipCache);
  }

  /**
   * Get the Role ('Viewer' | 'Editor') of the current user from ExternalGuestAccess.
   * Returns null if the user is not an external guest or not found.
   */
  public async getCurrentGuestRole(): Promise<'Viewer' | 'Editor' | undefined> {
    try {
      const currentUserEmail = this.context.pageContext.user.loginName ||
        this.context.pageContext.user.email ||
        this.context.pageContext.user.displayName;

      if (!currentUserEmail) return null;

      const normalizedEmail = currentUserEmail
        .toLowerCase()
        .replace('i:0#.f|membership|', '')
        .trim();

      const items = await this.sp.web.lists.getByTitle('ExternalGuestAccess').items
        .select('Id', 'Email', 'Role', 'Status', 'AccessExpiryDate')
        .filter(`Email eq '${normalizedEmail.replace(/'/g, "''")}' and Status eq 'Active'`)
        .top(1)();

      if (!items || items.length === 0) return null;

      // Honour expiry
      const expiryDate = items[0].AccessExpiryDate ? new Date(items[0].AccessExpiryDate) : null;
      if (expiryDate && new Date() > expiryDate) return null;

      const role = items[0].Role as string;
      if (role === 'Editor') return 'Editor';
      if (role === 'Viewer') return 'Viewer';
      return null;
    } catch (error) {
      
      return null;
    }
  }

  /**
   * Check if the current logged-in user is in the ExternalGuestAccess list
   * Returns true if the user is a restricted external guest
   */
  public async isCurrentUserRestrictedGuest(): Promise<boolean> {
    if (this._isRestrictedGuestCache !== null) return this._isRestrictedGuestCache;

    try {
      // Prioritize user.email — for B2B external guests, loginName is their UPN
      // (e.g. spweb94_gmail.com#EXT#@tenant.onmicrosoft.com) which does NOT match
      // the plain email stored in ExternalGuestAccess.
      const currentUserEmail = this.context.pageContext.user.email ||
        this.context.pageContext.user.loginName ||
        this.context.pageContext.user.displayName;

      if (!currentUserEmail) {
        
        return false;
      }

      // Normalize email — strip SharePoint claims prefix if still present
      const normalizedEmail = currentUserEmail
        .toLowerCase()
        .replace(/^i:0#\.f\|membership\|/i, '')
        .trim();

      const items = await this.sp.web.lists.getByTitle('ExternalGuestAccess').items
        .select('Id', 'Email', 'Status', 'AccessExpiryDate')
        .filter(`Email eq '${normalizedEmail.replace(/'/g, "''")}' and Status eq 'Active'`)
        .top(1)();

      if (!items || items.length === 0) {
        
        return false;
      }

      // Check if the user's overall access has expired
      const userRecord = items[0];
      const accessExpiryDate = userRecord.AccessExpiryDate;

      if (accessExpiryDate) {
        const expiryDate = new Date(accessExpiryDate);
        const currentDate = new Date();

        if (currentDate > expiryDate) {
          
          return false; // Expired users should not be considered restricted (they get no access at all)
        }
      }

      const isRestricted = true;


      this._isRestrictedGuestCache = isRestricted;
      return isRestricted;
    } catch (_error) {
      // Fail safely - if we cannot check, allow access
      return false;
    }
  }

  /**
   * Filter documents for restricted external guest users based on SharedDocumentLog
   * Only shows documents that have been explicitly shared with the current user
   */
  private async filterDocumentsForRestrictedUser(documents: ISharedDocument[]): Promise<ISharedDocument[]> {
    try {
      // Prioritize user.email so B2B guest email matches the plain email stored in SharedDocumentLog
      const currentUserEmail = this.context.pageContext.user.email ||
        this.context.pageContext.user.loginName ||
        this.context.pageContext.user.displayName;

      if (!currentUserEmail) {
        
        return [];
      }

      // Normalize email for comparison
      const normalizedEmail = currentUserEmail
        .toLowerCase()
        .replace('i:0#.f|membership|', '') // Remove SharePoint prefix if present
        .trim();

      

      // Get all SharedDocumentLog entries for this user
      const sharedLogEntries = await this.sp.web.lists.getByTitle('SharedDocumentLog').items
        .select('Id', 'Document', 'DocumentName', 'GuestEmail', 'Permission', 'ActionType', 'ActionDate', 'AccessDuration', 'ExpiryDate')
        .filter(`substringof('${normalizedEmail.replace(/'/g, "''")}', GuestEmail) and ActionType eq 'Shared'`)
        .orderBy('ActionDate', false) // Get the most recent entries first
        .top(1000)();

      if (!sharedLogEntries || sharedLogEntries.length === 0) {
        
        return [];
      }

      

      // Create a map of accessible documents based on SharedDocumentLog with access duration check
      const accessibleDocumentsMap = new Map<string, { permission: 'Read' | 'Review' | 'Edit' | 'Admin'; isValid: boolean }>();
      const currentDate = new Date();

      for (const logEntry of sharedLogEntries) {
        const documentUrl = logEntry.Document;
        const permission = logEntry.Permission as 'Read' | 'Review' | 'Edit' | 'Admin' | undefined;

        if (!documentUrl || !permission || !(['Read', 'Review', 'Edit', 'Admin'] as const).includes(permission)) {
          continue;
        }

        const actionDate = logEntry.ActionDate ? new Date(logEntry.ActionDate) : null;
        const accessDuration = logEntry.AccessDuration !== undefined ? logEntry.AccessDuration : 0;

        // Determine access validity — prefer stored ExpiryDate, fall back to ActionDate + Duration
        let isAccessValid = true;
        if (logEntry.ExpiryDate) {
          const expiryDate = new Date(logEntry.ExpiryDate);
          isAccessValid = currentDate <= expiryDate;
          if (!isAccessValid) {
            
          } else {
            
          }
        } else if (actionDate && accessDuration > 0) {
          const expiryDate = new Date(actionDate.getTime() + (accessDuration * 24 * 60 * 60 * 1000));
          isAccessValid = currentDate <= expiryDate;
          if (!isAccessValid) {
            
          } else {
            
          }
        } else {
          // accessDuration === -1 (permanent) or 0 or no actionDate — permanent access
          
        }

        // Key 1: full server-relative URL (most reliable match)
        const normalizedDocUrl = documentUrl.toLowerCase().trim();
        if (!accessibleDocumentsMap.has(normalizedDocUrl)) {
          accessibleDocumentsMap.set(normalizedDocUrl, { permission, isValid: isAccessValid });
        }

        // Key 2: filename only (fallback)
        const urlParts = documentUrl.split('/');
        const documentName = urlParts[urlParts.length - 1];
        const docKey = documentName.toLowerCase();
        if (!accessibleDocumentsMap.has(docKey)) {
          accessibleDocumentsMap.set(docKey, { permission, isValid: isAccessValid });
        }
      }

      // Build a set of ancestor folder server-relative URLs that lead to valid accessible documents.
      // This ensures that parent folders of shared files are shown to the guest so they can
      // navigate into them. Without this, a file shared inside a subfolder would be invisible
      // because the subfolder itself has no log entry.
      const accessibleFolderUrls = new Set<string>();
      for (const [docUrl, { isValid }] of accessibleDocumentsMap) {
        if (!isValid) continue;
        // Only process full server-relative URL keys (they start with '/')
        if (!docUrl.startsWith('/')) continue;
        const parts = docUrl.split('/');
        parts.pop(); // remove the filename
        while (parts.length > 1) {
          accessibleFolderUrls.add(parts.join('/'));
          parts.pop();
        }
      }

      // Filter documents based on what's accessible to the user
      const filteredDocuments: ISharedDocument[] = [];

      for (const doc of documents) {
        let hasAccess = false;
        let docPermission: 'Read' | 'Review' | 'Edit' | 'Admin' = 'Read';
        let isAccessValid = false;

        if (doc.isFolder) {
          // For folders: show the folder if it is an ancestor of any valid shared document.
          // This lets guests navigate into subfolders that contain their shared files.
          const folderRef = doc.fileRef.toLowerCase().trim();
          if (accessibleFolderUrls.has(folderRef)) {
            hasAccess = true;
            docPermission = 'Read';
            isAccessValid = true;
            
          }
        } else {
          // For files: match by full fileRef first (exact), then filename fallback
          const matchKeys = [
            doc.fileRef.toLowerCase().trim(),
            doc.fileRef.split('/').pop()?.toLowerCase() || '',
            doc.name.toLowerCase(),
            doc.path.toLowerCase()
          ];

          for (const key of matchKeys) {
            const accessInfo = accessibleDocumentsMap.get(key);
            if (accessInfo) {
              hasAccess = true;
              docPermission = accessInfo.permission;
              isAccessValid = accessInfo.isValid;
              
              break;
            }
          }
        }

        // Only include items that have valid (non-expired) access
        if (hasAccess && isAccessValid) {
          doc.permission = docPermission;
          filteredDocuments.push(doc);
        }
      }

      return filteredDocuments;
    } catch (_error) {
      // Fail safely - return empty array if filtering fails
      return [];
    }
  }
}
