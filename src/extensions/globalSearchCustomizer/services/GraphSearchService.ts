import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import { ISearchResult } from '../../../models/ISearchResult';
import { expandQuery } from '../../../utils/synonymDictionary';

const userPhotoCache = new Map<string, string | null>();

export class GraphSearchService {
  private _msGraphClientFactory: MSGraphClientFactory;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _spHttpClient: any;
  private _siteUrl: string;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  constructor(msGraphClientFactory: MSGraphClientFactory, spHttpClient?: any, siteUrl?: string) {
    this._msGraphClientFactory = msGraphClientFactory;
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  public async search(
    query: string, 
    pageSize: number = 20, 
    from: number = 0,
    fileTypes: string[] = [],
    activeTopTab: string = 'All',
    date: string = '',
    selectedAuthors: string[] = [],
    sortBy: string = 'dateDesc'
  ): Promise<{ results: ISearchResult[]; totalCount: number; suggestedQuery?: string }> {
    const client: any = await this._msGraphClientFactory.getClient('3');

    const rawQuery = query.trim();
    // Build the query string using a safer KQL fallback and server-side Title boosting via XRANK
    let boostedQuery = '';
    if (rawQuery) {
      boostedQuery = rawQuery.replace(/["\\]/g, '\\$&');
    }

    let queryString = boostedQuery || 'IsDocument:1';

    if (activeTopTab === 'Folders') {
      queryString = boostedQuery 
        ? `(${boostedQuery}) AND (ContentTypeId:0x0120*)` 
        : 'ContentTypeId:0x0120*';
    } else if (activeTopTab === 'Files') {
      const docTypes = 'filetype:pdf OR filetype:doc OR filetype:docx OR filetype:xls OR filetype:xlsx OR filetype:csv OR filetype:ppt OR filetype:pptx OR filetype:txt OR filetype:rtf OR filetype:msg OR filetype:zip';
      const filesFilter = `NOT(ContentTypeId:0x0120*) AND (${docTypes})`;
      queryString = boostedQuery ? `(${boostedQuery}) AND (${filesFilter})` : filesFilter;
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
      const typeQueries = actualTypes.map(t => {
        const lowerT = t.toLowerCase();
        if (lowerT === 'xls') return '(filetype:xls OR filetype:xlsx OR filetype:csv)';
        if (lowerT === 'doc') return '(filetype:doc OR filetype:docx)';
        if (lowerT === 'ppt') return '(filetype:ppt OR filetype:pptx)';
        return `filetype:${lowerT}`;
      });
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
      // Strictly filter by SharePoint uploader/creator, ignoring embedded computer file authors
      const authorQueries = selectedAuthors.map(a => `AuthorOWSUSER:"${a}"`);
      queryString += ` AND (${authorQueries.join(' OR ')})`;
    }

    // Globally exclude developer and system files from all searches
    const excludedTypes = ['md', 'ts', 'jsx', 'json', 'cmd', 'js', 'java', 'css', 'html', 'scss', 'xml', 'yml', 'yaml', 'env', 'sh', 'bat', 'py', 'sql'];
    queryString += ` ${excludedTypes.map(ext => `-filetype:${ext}`).join(' ')}`;

    const expandedQuery = expandQuery(queryString);
    console.log('--- [DEBUG] Expanded Query ---', expandedQuery);

    // Build Graph Search POST payload according to Microsoft Graph Search API guidelines
    const searchPayload: any = {
      requests: [
        {
          entityTypes: ['driveItem'],
          query: {
            queryString: expandedQuery
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
            'lastModifiedBy',
            'parentReference',
            'file',
            'folder',
            'Author'
          ],
          queryAlterationOptions: {
            enableSuggestion: true,
            enableModification: true
          }
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

    let results: ISearchResult[] = hits.map((hit: any) => {
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
        // Microsoft Graph's webUrl is natively designed to open the file in the browser.
        // Stripping query parameters or modifying the path (like replacing DispForm.aspx)
        // destroys the link integrity and causes 404s. We must return it exactly as-is.
        return webUrl || '';
      };

      const getSiteName = (url: string, fallbackUrl: string): string => {
        if (!url) return 'SharePoint Portal';
        
        // Handle OneDrive URLs
        if (url.includes('-my.sharepoint.com') || url.includes('/personal/')) {
          return 'OneDrive';
        }

        // Try extracting from /sites/SiteName
        const siteMatch = url.match(/\/sites\/([^/]+)/i);
        if (siteMatch && siteMatch[1]) {
          try {
            return decodeURIComponent(siteMatch[1]);
          } catch (e) {
            return siteMatch[1];
          }
        }
        
        // Handle Root SharePoint Site (e.g. moreyahs.sharepoint.com/Shared Documents)
        if (url.includes('.sharepoint.com')) {
          const hostname = url.split('/')[2];
          if (hostname && !hostname.includes('-my.sharepoint.com')) {
            return 'Main Portal';
          }
        }

        // Fallback
        return fallbackUrl?.split('/').pop() || 'SharePoint Portal';
      };

      let indexedAuthor = '';
      if (resource.Author) {
        indexedAuthor = Array.isArray(resource.Author) ? resource.Author[0] : resource.Author;
      }
      
      const createdByName = resource.createdBy?.user?.displayName || '';
      const modifiedByName = resource.lastModifiedBy?.user?.displayName || '';
      
      // author is the SharePoint uploader
      const finalAuthor = createdByName || indexedAuthor || modifiedByName || 'SharePoint User';
      
      // Try to find the original author from either indexedAuthor or lastModifiedBy
      let potentialOriginal = indexedAuthor || modifiedByName;
      
      // Smart Fallback: If MS Graph hides the internal Author property (common for driveItems), 
      // but the user explicitly filtered by an author and this file was returned, 
      // we can deduce that the filtered author MUST be the original embedded author!
      if (selectedAuthors && selectedAuthors.length > 0) {
        const potentialMatchesFilter = potentialOriginal ? selectedAuthors.some(a => a.toLowerCase() === potentialOriginal.toLowerCase()) : false;
        const uploaderMatchesFilter = selectedAuthors.some(a => a.toLowerCase() === createdByName.toLowerCase());
        
        // If neither the known potential original nor the uploader matches the filter, 
        // the filter MUST have matched the hidden embedded author.
        if (!potentialMatchesFilter && !uploaderMatchesFilter) {
          potentialOriginal = selectedAuthors[0];
        }
      }
      
      let origAuthor = '';
      if (potentialOriginal && createdByName && potentialOriginal.trim().toLowerCase() !== createdByName.trim().toLowerCase()) {
        origAuthor = potentialOriginal.trim();
      }

      return {
        id: resource.id || hit.hitId || Math.random().toString(),
        title: resource.name || 'Untitled Document',
        webUrl: getCleanSharePointUrl(resource.webUrl || '', resource.name || ''),
        fileType: fileType,
        lastModified: resource.lastModifiedDateTime || new Date().toISOString(),
        author: finalAuthor,
        originalAuthor: origAuthor,
        authorEmail: resource.createdBy?.user?.email || '',
        driveId: resource.parentReference?.driveId || '',
        itemId: resource.id || '',
        size: resource.size || 0,
        summary: hit.summary || resource.description || 'No description preview available.',
        siteUrl: resource.parentReference?.siteId || '',
        siteName: getSiteName(resource.webUrl || '', resource.parentReference?.sharepointIds?.siteUrl || ''),
        relevanceScore: hit.rank ?? 0
      };
    });

    // Fetch thumbnails + library URL for each result using MS Graph.
    // driveUrlCache avoids duplicate /drives/{driveId} calls within a single search.
    const driveUrlCache = new Map<string, string>();

    results = await Promise.all(results.map(async (res) => {
      if (res.driveId) {
        // ── Library URL (needed for correct AllItems.aspx viewer links) ──────────
        if (!driveUrlCache.has(res.driveId)) {
          try {
            const driveInfo = await client.api(`/drives/${res.driveId}`).select('webUrl').get();
            driveUrlCache.set(res.driveId, driveInfo?.webUrl || '');
          } catch (e) {
            driveUrlCache.set(res.driveId, ''); // cache the failure so we don't retry
          }
        }
        res.libraryUrl = driveUrlCache.get(res.driveId) || '';

        // ── Thumbnail (images & videos only) ─────────────────────────────────────
        const isMedia = ['png', 'jpg', 'jpeg', 'gif', 'svg', 'mp4', 'mov', 'avi'].includes(res.fileType?.toLowerCase() || '');
        if (isMedia && res.itemId) {
          try {
            // Fetch the actual DriveItem to get the true physical webUrl (bypassing DispForm.aspx)
            const itemResponse = await client.api(`/drives/${res.driveId}/items/${res.itemId}`).select('webUrl,name,parentReference').get();
            if (itemResponse) {
              if (itemResponse.webUrl && !itemResponse.webUrl.includes('DispForm.aspx')) {
                res.webUrl = itemResponse.webUrl;
              } else if (res.libraryUrl && itemResponse.name && itemResponse.parentReference?.path) {
                // If webUrl is the ugly DispForm properties page, mathematically reconstruct the true file URL
                const pathParts = itemResponse.parentReference.path.split('/drive/root:');
                const subPath = pathParts.length > 1 && pathParts[1] ? pathParts[1] : '';
                res.webUrl = `${res.libraryUrl}${subPath}/${itemResponse.name}`;
              }
            }

            // Request large explicitly to force Graph to generate it if it's not cached
            const thumbResponse = await client.api(`/drives/${res.driveId}/items/${res.itemId}/thumbnails/0/large`).get();
            if (thumbResponse?.url) {
              res.thumbnailUrlLarge = thumbResponse.url;
              res.thumbnailUrl = thumbResponse.url;
            }
          } catch (e) {
            try {
              const smallThumb = await client.api(`/drives/${res.driveId}/items/${res.itemId}/thumbnails/0/small`).get();
              if (smallThumb?.url) {
                res.thumbnailUrl = smallThumb.url;
                res.thumbnailUrlLarge = smallThumb.url;
              }
            } catch (e2) {
              console.warn(`[GraphSearchService] Thumbnail fetch failed for ${res.itemId}`, e2);
            }
          }
        }
      }
      return res;
    }));

    // Fetch author photo as blob to avoid browser-level 404 noise
    results = await Promise.all(results.map(async (res) => {
      if (res.authorEmail) {
        const cachedPhoto = userPhotoCache.get(res.authorEmail);
        if (cachedPhoto !== undefined) {
          res.authorPhotoUrl = cachedPhoto || undefined;
        } else {
          try {
            const photoResponse = await client
              .api(`/users/${res.authorEmail}/photo/$value`)
              .responseType('blob' as any)
              .get();
            if (photoResponse) {
              const blobUrl = URL.createObjectURL(photoResponse);
              userPhotoCache.set(res.authorEmail, blobUrl);
              res.authorPhotoUrl = blobUrl;
            } else {
              userPhotoCache.set(res.authorEmail, null);
            }
          } catch {
            userPhotoCache.set(res.authorEmail, null);
            // No photo — ResultCard will show initials
          }
        }
      }
      return res;
    }));

    console.log('--- [DEBUG] Mapped results array returning to caller ---');
    console.log('--- [DEBUG] Results before re-ranking ---', results.map(r => ({ id: r.id, title: r.title, relevanceScore: r.relevanceScore })));

    const alteration = searchResponse?.queryAlterationResponse?.queryAlteration;
    let suggestedQuery = alteration?.alteredQueryString || undefined;

    if (suggestedQuery) {
      // Clean up programmatically appended filters from the suggested query
      suggestedQuery = suggestedQuery.split(' -filetype:')[0];
      suggestedQuery = suggestedQuery.split(' AND (')[0];
      suggestedQuery = suggestedQuery.split(' AND LastModifiedTime')[0];
      suggestedQuery = suggestedQuery.replace(/^\((.*)\)$/, '$1');
      suggestedQuery = suggestedQuery.trim();
    }

    // Fallback: If MS Graph did not return a spelling suggestion, but we have 0 results, 
    // try fetching suggestions from the native SharePoint Search API.
    if (!suggestedQuery && totalCount === 0 && this._spHttpClient && this._siteUrl && rawQuery) {
      try {
        const restUrl = `${this._siteUrl}/_api/search/query?querytext='${rawQuery.replace(/'/g, "''")}'&enablequerysuggestions=true&selectproperties='Title'`;
        const restResponse = await this._spHttpClient.get(restUrl, {
          headers: {
            'Accept': 'application/json;odata=nometadata'
          }
        });
        if (restResponse.ok) {
          const json = await restResponse.json();
          const spellingSuggestion = json.SpellingSuggestion || json.d?.query?.SpellingSuggestion;
          if (spellingSuggestion) {
            suggestedQuery = spellingSuggestion;
            console.log('--- [DEBUG] Native SharePoint Search Spelled Suggestion ---', suggestedQuery);
          }
        }
      } catch (err) {
        console.error('--- [DEBUG] Error fetching native SP spelling suggestions ---', err);
      }
    }

    return { results, totalCount, suggestedQuery };
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

  public async fetchIndexedTerms(): Promise<string[]> {
    try {
      const client: any = await this._msGraphClientFactory.getClient('3');
      const searchPayload = {
        requests: [
          {
            entityTypes: ['driveItem', 'listItem'],
            query: {
              queryString: '*'
            },
            fields: ['title', 'filename'],
            size: 200
          }
        ]
      };
      
      const response = await client.api('/search/query').version('v1.0').post(searchPayload);
      const hits = response.value?.[0]?.hitsContainers?.[0]?.hits || [];
      
      const wordsSet = new Set<string>();
      
      hits.forEach((hit: any) => {
        const resource = hit.resource || {};
        const title = resource.title || '';
        const filename = resource.filename || resource.name || '';
        
        const titleParts = title.split(/[\s\-_().]+/);
        const fileParts = filename.split(/[\s\-_().]+/);
        
        titleParts.forEach((word: string) => {
          const cleanWord = word.trim();
          if (cleanWord.length >= 3) {
            wordsSet.add(cleanWord);
          }
        });
        
        fileParts.forEach((word: string) => {
          const cleanWord = word.trim();
          if (cleanWord.length >= 3) {
            wordsSet.add(cleanWord);
          }
        });
      });
      
      console.log('--- [DEBUG] Dynamic terms built ---', wordsSet.size, Array.from(wordsSet).slice(0, 20));
      return Array.from(wordsSet);
    } catch (error) {
      console.error('Error fetching indexed terms:', error);
      return [];
    }
  }

  public async fetchPeopleNames(): Promise<string[]> {
    try {
      const client: any = await this._msGraphClientFactory.getClient('3');
      const response = await client
        .api('/me/people')
        .version('v1.0')
        .select('displayName')
        .top(100)
        .get();
      
      const people = response.value || [];
      const names = people
        .map((p: any) => p.displayName)
        .filter((name: string) => name && name.trim().length > 0);
        
      return names;
    } catch (error) {
      console.error('Error fetching people names:', error);
      return [];
    }
  }

  public async searchFileSuggestions(query: string): Promise<string[]> {
    if (!query || query.length < 2) return [];
    
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const client: any = await this._msGraphClientFactory.getClient('3');
      const searchPayload = {
        requests: [
          {
            entityTypes: ['driveItem'],
            query: { queryString: query },
            fields: ['name', 'title'],
            size: 5
          }
        ]
      };
      
      const response = await client.api('/search/query').version('v1.0').post(searchPayload);
      const hits = response.value?.[0]?.hitsContainers?.[0]?.hits || [];
      
      const suggestions = new Set<string>();
      
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      hits.forEach((hit: any) => {
        const resource = hit.resource || {};
        const filename = resource.name || resource.title || '';
        if (filename) {
          // Strip extensions
          let cleanName = filename.replace(/\.(pdf|docx?|xlsx?|pptx?|png|jpg|jpeg|gif|zip|txt|csv|md)$/i, '');
          // Replace underscores and hyphens with spaces
          cleanName = cleanName.replace(/[_-]/g, ' ').trim();
          
          if (cleanName && cleanName.length > 1) {
            // Deduplicate by lowercased value but store original case
            const lower = cleanName.toLowerCase();
            const exists = Array.from(suggestions).some(s => s.toLowerCase() === lower);
            if (!exists) {
              suggestions.add(cleanName);
            }
          }
        }
      });
      
      return Array.from(suggestions).slice(0, 5);
    } catch (error) {
      console.error('Error fetching file suggestions:', error);
      return [];
    }
  }
}

