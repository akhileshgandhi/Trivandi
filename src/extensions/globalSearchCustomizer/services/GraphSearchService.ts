import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import { ISearchResult } from '../../../models/ISearchResult';

export class GraphSearchService {
  private _msGraphClientFactory: MSGraphClientFactory;

  constructor(msGraphClientFactory: MSGraphClientFactory) {
    this._msGraphClientFactory = msGraphClientFactory;
  }

  public async search(
    query: string, 
    pageSize: number = 20, 
    from: number = 0,
    fileTypes: string[] = [],
    activeTopTab: string = 'All',
    date: string = '',
    selectedAuthors: string[] = [],
    sortBy: string = 'relevance'
  ): Promise<{ results: ISearchResult[]; totalCount: number }> {
    const client: any = await this._msGraphClientFactory.getClient('3');

    const rawQuery = query.trim();
    // Build the query string using a safer KQL fallback and server-side Title boosting via XRANK
    let boostedQuery = '';
    if (rawQuery) {
      if (rawQuery.includes(' ')) {
        boostedQuery = `(${rawQuery} XRANK(cb=10000) Title:"${rawQuery}*")`;
      } else {
        boostedQuery = `(${rawQuery} XRANK(cb=10000) Title:${rawQuery}*)`;
      }
    }

    let queryString = boostedQuery || 'IsDocument:1';

    if (activeTopTab === 'Folders') {
      queryString = boostedQuery ? `(${boostedQuery}) AND (IsContainer:true OR contentclass:STS_Folder)` : '(IsContainer:true OR contentclass:STS_Folder)';
    } else if (activeTopTab === 'Files') {
      queryString = boostedQuery ? `(${boostedQuery}) AND IsContainer:false AND NOT contentclass:STS_Folder` : 'IsContainer:false AND NOT contentclass:STS_Folder';
    } else if (activeTopTab === 'Images') {
      const imgFilter = '(filetype:png OR filetype:jpg OR filetype:jpeg OR filetype:gif OR filetype:svg) AND IsContainer:false AND NOT contentclass:STS_Folder';
      queryString = boostedQuery ? `(${boostedQuery}) AND (${imgFilter})` : imgFilter;
    } else if (activeTopTab === 'Videos') {
      const vidFilter = '(filetype:mp4 OR filetype:mov OR filetype:avi) AND IsContainer:false AND NOT contentclass:STS_Folder';
      queryString = boostedQuery ? `(${boostedQuery}) AND (${vidFilter})` : vidFilter;
    }
    
    // Add file type filters if present and not "All" (Only if not in Folders tab)
    const actualTypes = fileTypes.filter(t => t !== 'All');
    if (actualTypes.length > 0 && activeTopTab !== 'Folders') {
      const typeQueries = actualTypes.map(t => `filetype:${t.toLowerCase()}`);
      queryString += ` AND (${typeQueries.join(' OR ')})`;
    }

    // Add native SharePoint KQL date filters using standard KQL colon syntax
    if (date) {
      const now = new Date();
      const formatKQLDate = (d: Date): string => {
        return d.toISOString().split('T')[0];
      };

      if (date === 'Today') {
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(now)}`;
      } else if (date === 'Yesterday') {
        const dYesterday = new Date(now);
        dYesterday.setDate(dYesterday.getDate() - 1);
        queryString += ` AND LastModifiedTime:${formatKQLDate(dYesterday)}`;
      } else if (date === 'This Week') {
        const dWeek = new Date(now);
        dWeek.setDate(dWeek.getDate() - 7);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dWeek)}`;
      } else if (date === 'This Month') {
        const dMonth = new Date(now);
        dMonth.setDate(dMonth.getDate() - 30);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dMonth)}`;
      } else if (date === 'This Year') {
        const dYear = new Date(now);
        dYear.setDate(dYear.getDate() - 365);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dYear)}`;
      } else if (date.match(/^\d{4}-\d{2}-\d{2}$/)) {
        queryString += ` AND LastModifiedTime:${date}`;
      } else {
        try {
          const formattedDate = new Date(date).toISOString().split('T')[0];
          queryString += ` AND LastModifiedTime:${formattedDate}`;
        } catch (e) {
          // ignore invalid date formats
        }
      }
    }

    // Add native SharePoint KQL author filters
    if (selectedAuthors && selectedAuthors.length > 0) {
      const authorQueries = selectedAuthors.map(a => `(Author:"${a}" OR AuthorOWSUSER:"${a}")`);
      queryString += ` AND (${authorQueries.join(' OR ')})`;
    }

    // Build Graph Search POST payload according to Microsoft Graph Search API guidelines
    const searchPayload: any = {
      requests: [
        {
          entityTypes: ['driveItem'],
          query: {
            queryString: queryString
          },
          from: from,
          size: pageSize,
          fields: [
            'id',
            'name',
            'size',
            'webUrl',
            'lastModifiedDateTime',
            'createdBy',
            'parentReference',
            'file',
            'folder'
          ]
        }
      ]
    };

    if (sortBy && sortBy !== 'relevance') {
      let sortField = 'lastModifiedDateTime';
      let desc = true;
      if (sortBy === 'dateAsc') {
        desc = false;
      } else if (sortBy === 'sizeDesc') {
        sortField = 'size';
        desc = true;
      }
      searchPayload.requests[0].sortProperties = [
        {
          name: sortField,
          isDescending: desc
        }
      ];
    }

    console.log('--- [DEBUG] Raw Search Query input ---', query);
    console.log('--- [DEBUG] Active Top Tab ---', activeTopTab);
    console.log('--- [DEBUG] File Type Filters ---', fileTypes);
    console.log('--- [DEBUG] Date Filter ---', date);
    console.log('--- [DEBUG] Final Constructed KQL QueryString ---', queryString);
    console.log('--- [DEBUG] Full MS Graph Search Request Payload ---');
    console.log(JSON.stringify(searchPayload, null, 2));

    let response: any;
    try {
      response = await client
        .api('/search/query')
        .version('v1.0')
        .post(searchPayload);
      console.log('--- [DEBUG] Raw MS Graph Search Response ---');
      console.log(JSON.stringify(response, null, 2));
    } catch (error) {
      console.error('--- [DEBUG] Error in MS Graph API Call ---', error);
      throw error;
    }

    // Parse response objects
    const searchResponse = response.value?.[0];
    const hitsContainer = searchResponse?.hitsContainers?.[0];
    const totalCount = hitsContainer?.total || 0;
    const hits = hitsContainer?.hits || [];
    console.log('--- [DEBUG] Parsed Hits Count from MS Graph ---', hits.length);
    console.log('--- [DEBUG] Total Results Count from MS Graph Index ---', totalCount);

    const results: ISearchResult[] = hits.map((hit: any) => {
      const resource = hit.resource || {};
      
      // Determine file extension cleanly
      let fileType = 'doc';
      const name = resource.name || '';
      const rawUrl = resource.webUrl || '';
      
      const hasExtension = name.includes('.') && name.split('.').pop()?.toLowerCase() !== name.toLowerCase();
      const ext = hasExtension ? name.split('.').pop()?.toLowerCase() : '';
      
      const isFolder = 
        !!resource.folder || 
        rawUrl.includes('/:f:/') || 
        !hasExtension; // No extension means it is a folder

      if (isFolder) {
        fileType = 'folder';
      } else if (ext) {
        if (ext === 'pdf') fileType = 'pdf';
        else if (['xls', 'xlsx'].includes(ext)) fileType = 'xlsx';
        else if (['ppt', 'pptx'].includes(ext)) fileType = 'pptx';
        else if (['png', 'jpg', 'jpeg', 'gif', 'svg'].includes(ext)) fileType = 'png';
        else if (['mp4', 'mov', 'avi', 'wmv', 'mkv', 'flv', 'webm'].includes(ext)) fileType = 'mp4';
        else fileType = ext;
      } else if (resource.file?.mimeType) {
        const mime = resource.file.mimeType.toLowerCase();
        if (mime.includes('pdf')) fileType = 'pdf';
        else if (mime.includes('excel') || mime.includes('spreadsheet')) fileType = 'xlsx';
        else if (mime.includes('presentation') || mime.includes('powerpoint')) fileType = 'pptx';
        else if (mime.includes('image')) fileType = 'png';
        else if (mime.includes('video')) fileType = 'mp4';
      }

      const getCleanSharePointUrl = (webUrl: string, title: string): string => {
        if (!webUrl) return '';
        try {
          let cleaned = webUrl;
          if (cleaned.includes('/Forms/DispForm.aspx')) {
            cleaned = cleaned.replace(/\/Forms\/DispForm\.aspx.*/i, `/${title}`);
          }
          cleaned = cleaned.replace(/\/:[a-z]:\/[a-z]\//i, '/');
          cleaned = cleaned.replace(/\/:[a-z]:\/r\//i, '/');
          cleaned = cleaned.replace(/\/:[a-z]:\/g\//i, '/');
          cleaned = cleaned.split('?')[0]; // Strip query parameters

          // Force files to open in the browser instead of downloading by appending ?web=1
          const parts = title.split('.');
          if (parts.length > 1) {
            const ext = parts.pop()?.toLowerCase();
            if (ext && ext !== 'folder') {
              cleaned = `${cleaned}?web=1`;
            }
          }
          return cleaned;
        } catch (e) {
          return webUrl;
        }
      };

      return {
        id: resource.id || hit.hitId || Math.random().toString(),
        title: resource.name || 'Untitled Document',
        webUrl: getCleanSharePointUrl(resource.webUrl || '', resource.name || ''),
        fileType: fileType,
        lastModified: resource.lastModifiedDateTime || new Date().toISOString(),
        author: resource.createdBy?.user?.displayName || 'SharePoint User',
        size: resource.size || 0,
        summary: hit.summary || resource.description || 'No description preview available.',
        siteUrl: resource.parentReference?.siteId || '',
        siteName: resource.parentReference?.sharepointIds?.siteUrl?.split('/').pop() || 'SharePoint Portal'
      };
    });

    console.log('--- [DEBUG] Mapped results array returning to caller ---');
    console.log(JSON.stringify(results.map(r => ({ id: r.id, title: r.title, fileType: r.fileType })), null, 2));

    return { results, totalCount };
  }

  private sortAuthorsList(names: string[]): string[] {
    const uniqueNames = Array.from(new Set(names));
    const alphabetic = uniqueNames.filter(name => /^[a-zA-Z]/.test(name.trim().charAt(0)));
    const nonAlphabetic = uniqueNames.filter(name => !/^[a-zA-Z]/.test(name.trim().charAt(0)));
    
    alphabetic.sort((a, b) => a.localeCompare(b));
    nonAlphabetic.sort((a, b) => a.localeCompare(b));
    
    return [...alphabetic, ...nonAlphabetic];
  }

  public async getAuthors(query: string = ''): Promise<string[]> {
    try {
      const client: any = await this._msGraphClientFactory.getClient('3');
      
      let url = '/users?$select=displayName&$top=999';
      if (query.trim()) {
        const cleanQuery = query.trim().replace(/'/g, "''");
        url += `&$filter=startsWith(displayName,'${cleanQuery}') or startsWith(givenName,'${cleanQuery}') or startsWith(surname,'${cleanQuery}')`;
      }
      
      const response = await client.api(url).version('v1.0').get();
      const users = response.value || [];
      const names = users
        .map((u: any) => u.displayName)
        .filter((name: string) => name && name.trim().length > 0);
        
      return this.sortAuthorsList(names);
    } catch (error) {
      console.error('Error fetching authors from Graph:', error);
      // Graceful fallback to search for people using MS Graph Search query API
      try {
        const client: any = await this._msGraphClientFactory.getClient('3');
        const searchPayload = {
          requests: [
            {
              entityTypes: ['person'],
              query: {
                queryString: query.trim() ? `${query.trim()}*` : '*'
              },
              size: 500
            }
          ]
        };
        const response = await client.api('/search/query').version('v1.0').post(searchPayload);
        const hits = response.value?.[0]?.hitsContainers?.[0]?.hits || [];
        const names = hits
          .map((hit: any) => hit.resource?.displayName)
          .filter((name: string) => name && name.trim().length > 0);
        return this.sortAuthorsList(names);
      } catch (e) {
        console.error('Fallback author search failed:', e);
        return [];
      }
    }
  }
}

