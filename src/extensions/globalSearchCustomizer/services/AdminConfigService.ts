import { SPHttpClient } from '@microsoft/sp-http';

export interface ISearchSiteConfigItem {
  title: string;
  siteSlug: string;
  siteUrl: string;
  isActive: boolean;
  sortOrder: number;
}

export class AdminConfigService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private static _cache: ISearchSiteConfigItem[] | null = null;
  private static _fileTypesCache: string[] | null = null;
  private static _dateFiltersCache: string[] | null = null;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  public static clearCache(): void {
    AdminConfigService._cache = null;
    AdminConfigService._fileTypesCache = null;
    AdminConfigService._dateFiltersCache = null;
  }

  public async getSites(): Promise<ISearchSiteConfigItem[]> {
    if (AdminConfigService._cache) {
      return AdminConfigService._cache;
    }

    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchSitesConfig')/items` +
        `?$orderby=SortOrder asc`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const json = await response.json();
        const items = json.value || [];
        const sites: ISearchSiteConfigItem[] = items
          .filter((item: any) => {
            const val = item.IsActive;
            if (val === undefined || val === null) return false;
            if (typeof val === 'boolean') return val;
            if (typeof val === 'string') {
              const lower = val.toLowerCase();
              return lower === 'yes' || lower === 'true' || lower === '1';
            }
            if (typeof val === 'number') return val === 1;
            return !!val;
          })
          .map((item: any) => ({
            title: item.Title,
            siteSlug: item.SiteSlug,
            siteUrl: item.SiteUrl,
            isActive: true,
            sortOrder: item.SortOrder || 0
          }));
        AdminConfigService._cache = sites;
        return sites;
      }
    } catch (e) {
      console.error('[AdminConfigService] getSites failed:', e);
    }
    return [];
  }

  public async addSite(
    label: string, 
    siteSlug: string
  ): Promise<void> {
    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchSitesConfig')/items`;

      // Get current max sort order
      const existingSites = await this.getSites();
      const maxOrder = existingSites.reduce(
        (max, s) => Math.max(max, s.sortOrder), 0
      );

      const response = await this._spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            Title: label,
            SiteSlug: siteSlug,
            SiteUrl: `/sites/${siteSlug}`,
            IsActive: 'Yes',
            SortOrder: maxOrder + 1,
          }),
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to add site to SharePoint list (Status ${response.status}): ${errText}`);
      }

      // Clear cache so next getSites() fetches fresh
      AdminConfigService.clearCache();

    } catch (error) {
      console.error('[AdminConfigService] addSite failed:', error);
      throw error;
    }
  }

  public async removeSite(
    label: string
  ): Promise<void> {
    try {
      // First find the item ID by Title
      const findUrl = `${this._siteUrl}/_api/web/lists/getbytitle('SearchSitesConfig')/items` +
        `?$filter=Title eq '${label.replace(/'/g, "''")}'` +
        `&$select=Id&$top=1`;

      const findResponse = await this._spHttpClient.get(
        findUrl,
        SPHttpClient.configurations.v1
      );
      const findJson = await findResponse.json();
      const items = findJson.value || [];
      
      if (items.length === 0) return;
      
      const itemId = items[0].Id;

      // Delete item
      const response = await this._spHttpClient.fetch(
        `${this._siteUrl}/_api/web/lists/getbytitle('SearchSitesConfig')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          method: 'POST',
          headers: {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to remove site from SharePoint list (Status ${response.status}): ${errText}`);
      }

      AdminConfigService.clearCache();

    } catch (error) {
      console.error(
        '[AdminConfigService] removeSite failed:', 
        error
      );
      throw error;
    }
  }

  public async getFileTypes(): Promise<string[]> {
    if (AdminConfigService._fileTypesCache) {
      return AdminConfigService._fileTypesCache;
    }

    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchFileTypesConfig')/items` +
        `?$orderby=SortOrder asc`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const json = await response.json();
        const items = json.value || [];
        const fileTypes: string[] = items
          .filter((item: any) => {
            const val = item.IsActive;
            if (val === undefined || val === null) return false;
            if (typeof val === 'boolean') return val;
            if (typeof val === 'string') {
              const lower = val.toLowerCase();
              return lower === 'yes' || lower === 'true' || lower === '1';
            }
            if (typeof val === 'number') return val === 1;
            return !!val;
          })
          .map((item: any) => item.Title);
        AdminConfigService._fileTypesCache = fileTypes;
        return fileTypes;
      }
    } catch (e) {
      console.error('[AdminConfigService] getFileTypes failed:', e);
    }
    return [];
  }

  public async addFileType(fileType: string): Promise<void> {
    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchFileTypesConfig')/items`;

      // Get current max sort order
      const existing = await this.getFileTypes();
      const maxOrder = existing.length;

      const response = await this._spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            Title: fileType.toUpperCase(),
            IsActive: 'Yes',
            SortOrder: maxOrder + 1,
          }),
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to add file type to SharePoint list (Status ${response.status}): ${errText}`);
      }

      AdminConfigService.clearCache();

    } catch (error) {
      console.error('[AdminConfigService] addFileType failed:', error);
      throw error;
    }
  }

  public async removeFileType(fileType: string): Promise<void> {
    try {
      // First find the item ID by Title
      const findUrl = `${this._siteUrl}/_api/web/lists/getbytitle('SearchFileTypesConfig')/items` +
        `?$filter=Title eq '${fileType.toUpperCase().replace(/'/g, "''")}'` +
        `&$select=Id&$top=1`;

      const findResponse = await this._spHttpClient.get(
        findUrl,
        SPHttpClient.configurations.v1
      );
      const findJson = await findResponse.json();
      const items = findJson.value || [];
      
      if (items.length === 0) return;
      
      const itemId = items[0].Id;

      // Delete item
      const response = await this._spHttpClient.fetch(
        `${this._siteUrl}/_api/web/lists/getbytitle('SearchFileTypesConfig')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          method: 'POST',
          headers: {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to remove file type from SharePoint list (Status ${response.status}): ${errText}`);
      }

      AdminConfigService.clearCache();

    } catch (error) {
      console.error(
        '[AdminConfigService] removeFileType failed:', 
        error
      );
      throw error;
    }
  }

  public async getDateFilters(): Promise<string[]> {
    if (AdminConfigService._dateFiltersCache) {
      return AdminConfigService._dateFiltersCache;
    }

    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchDateFiltersConfig')/items` +
        `?$orderby=SortOrder asc`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (response.ok) {
        const json = await response.json();
        const items = json.value || [];
        const dateFilters: string[] = items
          .filter((item: any) => {
            const val = item.IsActive;
            if (val === undefined || val === null) return false;
            if (typeof val === 'boolean') return val;
            if (typeof val === 'string') {
              const lower = val.toLowerCase();
              return lower === 'yes' || lower === 'true' || lower === '1';
            }
            if (typeof val === 'number') return val === 1;
            return !!val;
          })
          .map((item: any) => item.Title);
        AdminConfigService._dateFiltersCache = dateFilters;
        return dateFilters;
      }
    } catch (e) {
      console.error('[AdminConfigService] getDateFilters failed:', e);
    }
    return [];
  }

  public async addDateFilter(dateFilter: string): Promise<void> {
    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('SearchDateFiltersConfig')/items`;

      // Get current max sort order
      const existing = await this.getDateFilters();
      const maxOrder = existing.length;

      const response = await this._spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            Title: dateFilter,
            IsActive: 'Yes',
            SortOrder: maxOrder + 1,
          }),
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to add date filter to SharePoint list (Status ${response.status}): ${errText}`);
      }

      AdminConfigService.clearCache();

    } catch (error) {
      console.error('[AdminConfigService] addDateFilter failed:', error);
      throw error;
    }
  }

  public async removeDateFilter(dateFilter: string): Promise<void> {
    try {
      // First find the item ID by Title
      const findUrl = `${this._siteUrl}/_api/web/lists/getbytitle('SearchDateFiltersConfig')/items` +
        `?$filter=Title eq '${dateFilter.replace(/'/g, "''")}'` +
        `&$select=Id&$top=1`;

      const findResponse = await this._spHttpClient.get(
        findUrl,
        SPHttpClient.configurations.v1
      );
      const findJson = await findResponse.json();
      const items = findJson.value || [];
      
      if (items.length === 0) return;
      
      const itemId = items[0].Id;

      // Delete item
      const response = await this._spHttpClient.fetch(
        `${this._siteUrl}/_api/web/lists/getbytitle('SearchDateFiltersConfig')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          method: 'POST',
          headers: {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE',
          },
        }
      );

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Failed to remove date filter from SharePoint list (Status ${response.status}): ${errText}`);
      }

      AdminConfigService.clearCache();

    } catch (error) {
      console.error(
        '[AdminConfigService] removeDateFilter failed:', 
        error
      );
      throw error;
    }
  }

  public async searchSharePointSites(query: string): Promise<{ title: string; slug: string }[]> {
    if (!query || query.trim().length < 2) return [];
    try {
      const escaped = query.replace(/'/g, "''");
      const url = `${this._siteUrl}/_api/search/query?querytext='(contentclass:STS_Site OR contentclass:STS_Web) AND (Title:${escaped}* OR Path:http*sites/${escaped}*)'&selectproperties='Title,Path'&rowlimit=15`;
      const response = await this._spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (response.ok) {
        const json = await response.json();
        const rows = json.PrimaryQueryResult?.RelevantResults?.Table?.Rows || 
                     json.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
        
        const results: { title: string; slug: string }[] = [];
        rows.forEach((row: any) => {
          const cells = row.Cells || [];
          let title = '';
          let path = '';
          cells.forEach((cell: any) => {
            if (cell.Key === 'Title') title = cell.Value;
            if (cell.Key === 'Path') path = cell.Value;
          });
          if (title && path) {
            const pathSegments = path.split('/').filter(Boolean);
            const sitesIndex = pathSegments.indexOf('sites');
            let slug = '';
            if (sitesIndex !== -1 && pathSegments[sitesIndex + 1]) {
              slug = pathSegments[sitesIndex + 1];
            } else if (pathSegments.length > 0) {
              slug = pathSegments[pathSegments.length - 1];
            }
            if (slug && slug.toLowerCase() !== 'sites') {
              if (!results.some(r => r.slug.toLowerCase() === slug.toLowerCase())) {
                results.push({ title, slug });
              }
            }
          }
        });
        return results;
      }
    } catch (e) {
      console.error('[AdminConfigService] searchSharePointSites failed:', e);
    }
    return [];
  }
}
