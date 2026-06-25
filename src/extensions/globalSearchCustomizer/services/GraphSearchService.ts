import { MSGraphClientFactory, MSGraphClient, SPHttpClient } from '@microsoft/sp-http';
import { ISearchResult } from '../../../models/ISearchResult';
import { expandQuery } from '../../../utils/synonymDictionary';
import { usePermissionStore } from '../../Permission/PermissionStore';


import { CacheService } from './CacheService';
import { RetryService } from './RetryService';

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

  private _getTenantUrl(): string {
    let tenantUrl = 'https://trivandildn.sharepoint.com';
    // let tenantUrl = 'https://moreyahs.sharepoint.com';
    try {
      if (this._siteUrl) {
        tenantUrl = new URL(this._siteUrl).origin;
      }
    } catch (e) {
      // ignore
    }
    return tenantUrl;
  }

  private _buildBoostedQuery(rawQuery: string): string {
    if (!rawQuery || !rawQuery.trim()) {
      return '';
    }

    const escaped = rawQuery.replace(/["\\]/g, '\\$&');
    const trimmed = escaped.trim();

    const words = trimmed.split(/\s+/).filter(Boolean);
    let boostedQuery = '';

    if (words.length > 1) {
      // For multi-word queries, separate terms with AND inside field parentheses to avoid KQL syntax errors.
      const wildcardGroup = words.map(w => `${w}*`).join(' AND ');
      boostedQuery = [
        `title:(${wildcardGroup})`,
        `Filename:(${wildcardGroup})`,
        `name:(${wildcardGroup})`,
        `path:(${wildcardGroup})`,
        `description:(${wildcardGroup})`,
        `(${wildcardGroup})`
      ].join(' OR ');
    } else {
      // Single-word query
      boostedQuery = [
        `title:${trimmed}*`,
        `Filename:${trimmed}*`,
        `name:${trimmed}*`,
        `path:${trimmed}*`,
        `description:${trimmed}*`,
        `${trimmed}*`
      ].join(' OR ');
    }

    // Also add exact phrase matching for better precision
    boostedQuery = `(${boostedQuery}) OR ("${trimmed}")`;

    // Handle leading zero variants
    if (rawQuery.match(/^0\d/)) {
      const withoutLeadingZero = rawQuery.replace(/^0+/, '').trim();
      const escapedAlt = withoutLeadingZero.replace(/["\\]/g, '\\$&');
      boostedQuery = `(${boostedQuery}) OR title:"${escapedAlt}" OR "${escapedAlt}"`;
    }

    return boostedQuery;
  }

  private _cleanPathSegment(segment: string): string {
    const withoutExtension = segment.replace(/\.[a-z0-9]{2,5}$/i, '');
    return withoutExtension
      .replace(/^\d{4,}\s*/, '')
      .replace(/[_-]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private _getFolderContext(webUrl: string, folderName: string): { folderTitle: string; parentFolderName?: string } {
    if (!webUrl || !folderName) return { folderTitle: folderName || 'Untitled Folder' };

    try {
      const parsed = new URL(webUrl);
      let path = parsed.pathname;
      const idParam = parsed.searchParams.get('id') || parsed.searchParams.get('RootFolder');
      if (idParam) {
        path = idParam;
      }

      const pathSegments = path
        .split('/')
        .filter(Boolean)
        .map(segment => this._cleanPathSegment(decodeURIComponent(segment)))
        .filter(Boolean);

      const folderTitle = this._cleanPathSegment(folderName);
      const folderIndex = pathSegments.map(segment => segment.toLowerCase()).lastIndexOf(folderTitle.toLowerCase());
      const siteIndex = pathSegments.findIndex(segment => segment.toLowerCase() === 'sites');
      const siteName = siteIndex >= 0 && pathSegments[siteIndex + 1] ? pathSegments[siteIndex + 1] : '';
      const excludedSegments = new Set([
        'sites',
        siteName.toLowerCase(),
        'shared documents',
        'documents',
        'forms',
        'allitems.aspx',
        'bids'
      ].filter(Boolean));

      const contextSegments = pathSegments
        .slice(siteIndex >= 0 ? siteIndex + 2 : 0, folderIndex >= 0 ? folderIndex : pathSegments.length - 1)
        .filter(segment => !excludedSegments.has(segment.toLowerCase()));

      const projectContext = contextSegments[0] || contextSegments[contextSegments.length - 1];
      if (!projectContext || projectContext.toLowerCase() === folderTitle.toLowerCase()) {
        return { folderTitle: folderTitle || folderName };
      }
      return { folderTitle: `${folderTitle || folderName}`, parentFolderName: `${projectContext}` };
    } catch (e) {
      return { folderTitle: folderName || 'Untitled Folder' };
    }
  }

  public async search(
    query: string,
    pageSize: number = 20,
    from: number = 0,
    fileTypes: string[] = [],
    activeTopTab: string = 'All',
    date: string = '',
    selectedAuthors: string[] = [],
    selectedSites: string[] = [],
    sortBy: string = 'dateDesc'
  ): Promise<{ results: ISearchResult[]; totalCount: number; suggestedQuery?: string }> {
    const cacheKey = CacheService.buildKey(
      query, fileTypes, date, selectedAuthors, selectedSites, from, activeTopTab, sortBy
    );
    // [DISABLED CACHE AS PER USER REQUEST]
    // const cachedResult = CacheService.get(cacheKey);
    // if (cachedResult) {
    //   return cachedResult;
    // }

    const client: any = await this._msGraphClientFactory.getClient('3');

    const rawQuery = query.trim();
    // Build the query string using a safer KQL fallback and server-side Title boosting via XRANK
    const boostedQuery = this._buildBoostedQuery(rawQuery);

    let queryString = boostedQuery;
    if (!queryString) {
      // Graph Search API does NOT support 'IsDocument:1' or 'IsContainer:true' managed properties.
      // Use '*' as the base query — site Path scoping limits results to the correct sites.
      queryString = '*';
    }

    if (activeTopTab === 'All' && boostedQuery) {
      // Search broadly across files and folders matching the query.
      // IsContainer and IsDocument are NOT valid in Graph driveItem search — removed.
      queryString = `(${boostedQuery})`;
    }

    if (activeTopTab === 'Folders') {
      queryString = boostedQuery ? `(${boostedQuery}) AND ContentTypeId:0x0120*` : 'ContentTypeId:0x0120*';
    } else if (activeTopTab === 'Files') {
      const docTypes = 'filetype:pdf OR filetype:doc OR filetype:docx OR filetype:xls OR filetype:xlsx OR filetype:csv OR filetype:ppt OR filetype:pptx OR filetype:txt OR filetype:rtf OR filetype:msg OR filetype:zip';
      queryString = boostedQuery ? `(${boostedQuery}) AND (${docTypes})` : `(${docTypes})`;
    } else if (activeTopTab === 'Images') {
      const imgFilter = '(filetype:png OR filetype:jpg OR filetype:jpeg OR filetype:gif OR filetype:svg)';
      queryString = boostedQuery ? `(${boostedQuery}) AND ${imgFilter}` : imgFilter;
    } else if (activeTopTab === 'Videos') {
      const vidFilter = '(filetype:mp4 OR filetype:mov OR filetype:avi)';
      queryString = boostedQuery ? `(${boostedQuery}) AND ${vidFilter}` : vidFilter;
    }

    // Add file type filters if present and not "All" (Only if not in Folders tab)
    const actualTypes = fileTypes.filter(t => t !== 'All');
    if (actualTypes.length > 0 && activeTopTab !== 'Folders') {
      const typeQueries = actualTypes.map(t => {
        const lowerT = t.toLowerCase();
        if (lowerT === 'xls' || lowerT === 'xlsx' || lowerT === 'excel') {
          return '(filetype:xls OR filetype:xlsx OR filetype:csv)';
        }
        if (lowerT === 'doc' || lowerT === 'docx' || lowerT === 'word') {
          return '(filetype:doc OR filetype:docx)';
        }
        if (lowerT === 'ppt' || lowerT === 'pptx' || lowerT === 'powerpoint') {
          return '(filetype:ppt OR filetype:pptx)';
        }
        if (lowerT === 'text' || lowerT === 'txt') {
          return '(filetype:txt OR filetype:rtf)';
        }
        if (lowerT === 'image' || lowerT === 'img') {
          return '(filetype:png OR filetype:jpg OR filetype:jpeg OR filetype:gif OR filetype:svg)';
        }
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
      } else if (date === 'Last 7 Days') {
        const dWeek = new Date(now);
        dWeek.setDate(dWeek.getDate() - 7);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dWeek)}`;
      } else if (date === 'Last 30 Days') {
        const dMonth = new Date(now);
        dMonth.setDate(dMonth.getDate() - 30);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dMonth)}`;
      } else if (date === 'This Year') {
        const dYear = new Date(now);
        dYear.setDate(dYear.getDate() - 365);
        queryString += ` AND LastModifiedTime:>=${formatKQLDate(dYear)}`;
      } else if (date.includes('_')) {
        const [start, end] = date.split('_');
        if (start && end) {
          queryString += ` AND LastModifiedTime:>=${start} AND LastModifiedTime:<=${end}`;
        } else if (start) {
          queryString += ` AND LastModifiedTime:>=${start}`;
        } else if (end) {
          queryString += ` AND LastModifiedTime:<=${end}`;
        }
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

    // Enforce global site scope: Search only these 4 MoreYahs sites if no specific filter is chosen
    const tenantUrl = this._getTenantUrl();
    const defaultSites = [
      'TrivandiHub',
      'PeopleHub',
      'CompanyHub',
      'BrandingMarketing',
      'Projects',
      'TrivandiLondon',
      'TDMCC',
      'TrivandiUSA',
      'TrivandiAustralia',
      'TrivandiKSA'
    ];

    // const defaultSites = [
    //   'OperationsHub',
    //   'Freudiger',
    //   'moreYeahsdepartmentsDMS',
    //   'PembePortal'
    // ];

    let activeSites = (selectedSites && selectedSites.length > 0) ? selectedSites : defaultSites;

    // Filter activeSites by allowedSites from the permission store (security trimming)
    // Only apply this trimming when no specific site is selected (letting SharePoint natively handle security trimming for explicitly selected sites)
    if (!selectedSites || selectedSites.length === 0) {
      const allowedSites = usePermissionStore.getState().allowedSites || [];
      if (allowedSites.length > 0) {
        activeSites = activeSites.filter(s => allowedSites.includes(s));
      }
    }

    if (activeSites.length > 0) {
      const siteQueries = activeSites.map(s => {
        if (s === 'TrivandiAustralia') {
          return `(Path:"${tenantUrl}/sites/TrivandiAustralia*" OR Path:"${tenantUrl}/teams/TrivandiAustralia*" OR Path:"${tenantUrl}/sites/TrivandiAus*" OR Path:"${tenantUrl}/teams/TrivandiAus*" OR Path:"${tenantUrl}/sites/TrivandiAU*" OR Path:"${tenantUrl}/teams/TrivandiAU*" OR Path:"${tenantUrl}/sites/Australia*" OR Path:"${tenantUrl}/teams/Australia*" OR Path:"${tenantUrl}/sites/Trivandi-Australia*" OR Path:"${tenantUrl}/teams/Trivandi-Australia*")`;
        }
        return `(Path:"${tenantUrl}/sites/${s}*" OR Path:"${tenantUrl}/teams/${s}*")`;
      });
      queryString += ` AND (${siteQueries.join(' OR ')})`;
    } else {
      // Exclude all results safely if no sites are allowed
      queryString += ` AND Path:"https://nonexistent.sharepoint.com/sites/none*"`;
    }

    // NOTE: Do NOT use -filetype: or -Path: exclusions here.
    // The Microsoft Graph /search/query API does NOT support KQL minus-prefix exclusion syntax
    // (e.g. -filetype:json or -Path:...). Using them causes a FanoutDownstreamContradiction 500 error.
    // Site scoping via AND (Path:"...sites/X*") already limits results to the correct sites,
    // which naturally excludes OneDrive and unrelated content.

    const expandedQuery = expandQuery(queryString);

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
          trimDuplicates: false,
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
          // NOTE: queryAlterationOptions (enableSuggestion/enableModification) removed intentionally.
          // These options cause 500 FanoutDownstreamContradiction errors in Graph API
          // when combined with site-scoped Path: filters on driveItem searches.
        }
      ]
    };

    if (sortBy && sortBy !== 'relevance') {
      // Apply server-side sort whenever user explicitly selects a sort option.
      // When sortBy === 'relevance', skip sortProperties so Graph API uses its own ranking.
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




    let response: any;
    try {
      response = await client
        .api('/search/query')
        .version('v1.0')
        .post(searchPayload);
    } catch (error: any) {
      return { results: [], totalCount: 0 };
    }

    
    // Parse response objects
    const searchResponse = response.value?.[0];
    const hitsContainer = searchResponse?.hitsContainers?.[0];
    let totalCount = hitsContainer?.total || 0;
    const hits = hitsContainer?.hits || [];
    // (probe search removed — was a dead debug call that added extra API calls)
    // console.log('--- [DEBUG] Parsed Hits Count from MS Graph ---', hits.length);
    // console.log('--- [DEBUG] Total Results Count from MS Graph Index ---', totalCount);

    let results: ISearchResult[] = hits.map((hit: any) => {
      const resource = hit.resource || {};

      // Determine file extension cleanlyssss
      let fileType = 'doc';
      const name = resource.name || '';
      const rawUrl = resource.webUrl || '';

      // Extract potential extension and validate it looks like a real file extension
      // Real extensions: 2-5 alphanumeric chars, no spaces (e.g., "pdf", "xlsx", "docx")
      // Fake extensions: longer than 5 chars OR contain spaces (e.g., "Working Document" from "6.Working Document")
      let potentialExt = '';
      if (name.includes('.')) {
        potentialExt = name.split('.').pop()?.toLowerCase() || '';
      }
      const isRealExtension = potentialExt && potentialExt.length >= 2 && potentialExt.length <= 5 && /^[a-z0-9]+$/.test(potentialExt);
      const ext = isRealExtension ? potentialExt : '';
      const hasExtension = ext.length > 0;

      const isFolder =
        !!resource.folder ||
        rawUrl.includes('/:f:/') ||
        !hasExtension; // No extension means it is a folder

      if (isFolder) {
        // folder detected
      }

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

        let rawSiteName = '';

        // Handle OneDrive URLs
        if (url.includes('-my.sharepoint.com') || url.includes('/personal/')) {
          return 'OneDrive';
        }

        // Try extracting from /sites/SiteName or /teams/SiteName
        const siteMatch = url.match(/\/(?:sites|teams)\/([^/]+)/i);
        if (siteMatch && siteMatch[1]) {
          try {
            rawSiteName = decodeURIComponent(siteMatch[1]);
          } catch (e) {
            rawSiteName = siteMatch[1];
          }
        } else if (url.includes('.sharepoint.com')) {
          // Handle Root SharePoint Site (e.g. moreyahs.sharepoint.com/Shared Documents)
          const hostname = url.split('/')[2];
          if (hostname && !hostname.includes('-my.sharepoint.com')) {
            rawSiteName = 'Main Portal';
          }
        }

        if (!rawSiteName) {
          rawSiteName = fallbackUrl?.split('/').pop() || 'SharePoint Portal';
        }

        // Map known site names to UI friendly names
        const siteNameMapping: { [key: string]: string } = {
          'companyhub': 'Company Hub',
          'trivandihub': 'Trivandi Hub',
          'peoplehub': 'People Hub',
          'brandingmarketing': 'Marketing',
          'trivandilondon': 'Trivandi London',
          'trivandiusa': 'Trivandi USA',
          'trivandiaustralia': 'Trivandi Australia',
          'trivandiaus': 'Trivandi Australia',
          'trivandiau': 'Trivandi Australia',
          'australia': 'Trivandi Australia',
          'trivandiksa': 'Trivandi KSA'
        };

        const friendlyName = siteNameMapping[rawSiteName.toLowerCase()];
        return friendlyName || rawSiteName;
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

      let cleanWebUrl = getCleanSharePointUrl(resource.webUrl || '', resource.name || '');
      const isMediaFile = ['png', 'jpg', 'jpeg', 'gif', 'svg', 'mp4', 'mov', 'avi'].includes(fileType);

      if (isMediaFile && cleanWebUrl.includes('DispForm.aspx')) {
        const parentPath = resource.parentReference?.path || '';
        const pathParts = parentPath.split('/drive/root:');
        const subPath = pathParts.length > 1 && pathParts[1] ? pathParts[1] : '';

        let libUrl = '';
        try {
          const urlObj = new URL(cleanWebUrl);
          const segments = urlObj.pathname.split('/').filter(Boolean);
          if (segments.length >= 3 && (segments[0].toLowerCase() === 'sites' || segments[0].toLowerCase() === 'teams')) {
            libUrl = `${urlObj.origin}/${segments[0]}/${segments[1]}/${segments[2]}`;
          } else {
            libUrl = urlObj.origin;
          }
        } catch (e) { }

        if (libUrl && resource.name) {
          cleanWebUrl = `${libUrl}${subPath}/${resource.name}`;
        }
      }

      let folderTitle = resource.name || 'Untitled Document';

      // Always calculate folder context to extract the root directory for BOTH files and folders
      const folderCtx = this._getFolderContext(cleanWebUrl || resource.webUrl || '', resource.name || 'Untitled Document');
      let parentFolder = folderCtx.parentFolderName;

      if (isFolder) {
        folderTitle = folderCtx.folderTitle;
      }

      return {
        id: resource.id || hit.hitId || Math.random().toString(),
        title: folderTitle,
        parentFolder: parentFolder,
        webUrl: cleanWebUrl,
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

    // Security trimming: filter out restricted libraries for non-Contributor/Owner users
    const siteRoles = usePermissionStore.getState().siteRoles || {};
    
    // 1. BrandingMarketing -> MarketingPrivate
    const marketingRole = siteRoles['BrandingMarketing'] || 'None';
    const isContributorOrOwnerOnMarketing = marketingRole === 'Owner' || marketingRole === 'Contributor';

    // 2. PeopleHub -> PeopleDocs
    const peopleRole = siteRoles['PeopleHub'] || 'None';
    const isContributorOrOwnerOnPeople = peopleRole === 'Owner' || peopleRole === 'Contributor';

    // 3. CompanyHub -> CompanyPrivate*
    const companyRole = siteRoles['CompanyHub'] || 'None';
    const isContributorOrOwnerOnCompany = companyRole === 'Owner' || companyRole === 'Contributor';

    const privateCompanyLibraries = [
      'companyprivateaustralia',
      'companyprivatelondon',
      'companyprivateksa',
      'companyprivatedubai',
      'companyprivateusa'
    ];

    if (!isContributorOrOwnerOnMarketing || !isContributorOrOwnerOnPeople || !isContributorOrOwnerOnCompany) {
      const originalLength = results.length;
      results = results.filter(r => {
        const pathLower = (r.webUrl || '').toLowerCase();
        const rootFolderLower = (r.parentFolder || '').toLowerCase();
        
        // Check BrandingMarketing - MarketingPrivate
        if (!isContributorOrOwnerOnMarketing) {
          const isMarketingSite = pathLower.includes('/sites/brandingmarketing') || pathLower.includes('/teams/brandingmarketing');
          const isPrivate = pathLower.includes('/marketingprivate/') || pathLower.endsWith('/marketingprivate') || rootFolderLower === 'marketingprivate' || rootFolderLower.includes('marketingprivate');
          if (isMarketingSite && isPrivate) {
            return false;
          }
        }
        
        // Check PeopleHub - PeopleDocs
        if (!isContributorOrOwnerOnPeople) {
          const isPeopleSite = pathLower.includes('/sites/peoplehub') || pathLower.includes('/teams/peoplehub');
          const isPrivatePeople = pathLower.includes('/peopledocs/') || pathLower.endsWith('/peopledocs') || rootFolderLower === 'peopledocs' || rootFolderLower.includes('peopledocs');
          if (isPeopleSite && isPrivatePeople) {
            return false;
          }
        }

        // Check CompanyHub - CompanyPrivate*
        if (!isContributorOrOwnerOnCompany) {
          const isCompanySite = pathLower.includes('/sites/companyhub') || pathLower.includes('/teams/companyhub');
          const isPrivateCompany = privateCompanyLibraries.some(lib => 
            pathLower.includes(`/${lib}/`) || pathLower.endsWith(`/${lib}`) || rootFolderLower === lib || rootFolderLower.includes(lib)
          );
          if (isCompanySite && isPrivateCompany) {
            return false;
          }
        }
        
        return true;
      });
      const removedCount = originalLength - results.length;
      totalCount = Math.max(0, totalCount - removedCount);
    }

    // Block 1 - Thumbnail + LibraryUrl (no API)
    results = results.map((res) => {
      // Thumbnail URLs via native SharePoint v2.0 API (Cookie Authenticated)
      // This is bulletproof because it relies on exact IDs, not string-math paths.
      if (res.fileType !== 'folder' && res.driveId && res.itemId) {
        const previewUrl = `${tenantUrl}/_api/v2.0/drives/${res.driveId}/items/${res.itemId}/thumbnails/0/large/content`;
        res.thumbnailUrl = previewUrl;
        res.thumbnailUrlLarge = previewUrl;
      }

      // Extract Library URL from webUrl without API call
      try {
        if (res.webUrl) {
          const urlObj = new URL(res.webUrl);
          const pathSegments = urlObj.pathname.split('/').filter(Boolean);
          if (pathSegments.length >= 3 && (pathSegments[0].toLowerCase() === 'sites' || pathSegments[0].toLowerCase() === 'teams')) {
            const libPath = `/${pathSegments[0]}/${pathSegments[1]}/${pathSegments[2]}`;
            res.libraryUrl = `${urlObj.origin}${libPath}`;
          } else {
            res.libraryUrl = urlObj.origin;
          }
        }
      } catch (e) {
        res.libraryUrl = '';
      }

      return res;
    });

    // Block 2 - Author Photo (CacheService with USER_PHOTOS TTL)
    results = results.map((res) => {
      if (res.authorEmail) {
        const photoCacheKey = `photo|${res.authorEmail}`;
        const cachedPhoto = CacheService.get(photoCacheKey);
        if (cachedPhoto) {
          res.authorPhotoUrl = cachedPhoto;
        } else {
          const photoUrl = `${tenantUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(res.authorEmail)}`;
          CacheService.set(photoCacheKey, photoUrl, CacheService.TTL.USER_PHOTOS);
          res.authorPhotoUrl = photoUrl;
        }
      }
      return res;
    });

    // console.log('--- [DEBUG] Mapped results array returning to caller ---');
    // console.log('--- [DEBUG] Results before re-ranking ---', results.map(r => ({ id: r.id, title: r.title, relevanceScore: r.relevanceScore })));

    let suggestedQuery: string | undefined = undefined;

    // Fetch clean, user-friendly spelling suggestions from the native SharePoint Search API
    // when the search yields 0 results. This completely avoids leaking KQL code.
    if (totalCount === 0 && this._spHttpClient && this._siteUrl && rawQuery && rawQuery.trim() !== '') {
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
          }
        }
      } catch (err) {
        // error fetching spelling suggestions
      }
    }

    // Fetch matched project titles for each result in parallel (with per-URL caching inside getTitleForUrl)
    try {
      await Promise.all(results.map(async (res) => {
        try {
          const matchedTitle = await this.getTitleForUrl(res.webUrl);
          if (matchedTitle) {
            res.matchedProjectTitle = matchedTitle;
          }
        } catch (itemErr) {
          // ignore per-item errors
        }
      }));
    } catch (e) {
      // ignore errors fetching project titles
    }

    CacheService.set(cacheKey, { results, totalCount, suggestedQuery });
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
        return [];
      }
    }
  }

  public async fetchIndexedTerms(): Promise<string[]> {
    if (!this._spHttpClient || !this._siteUrl) return [];
    try {
      const url = `${this._siteUrl}/_api/search/query?querytext='*'&selectproperties='Title,Filename'&rowlimit=100`;
      const response = await this._spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (response.ok) {
        const json = await response.json();
        const rows = json.PrimaryQueryResult?.RelevantResults?.Table?.Rows ||
          json.d?.query?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

        const wordsSet = new Set<string>();
        rows.forEach((row: any) => {
          const cells = row.Cells || [];
          let title = '';
          let filename = '';
          cells.forEach((cell: any) => {
            if (cell.Key === 'Title') title = cell.Value;
            if (cell.Key === 'Filename') filename = cell.Value;
          });

          if (title) {
            title.split(/[\s\-_().]+/).forEach((word: string) => {
              const cleanWord = word.trim();
              if (cleanWord.length >= 3) {
                wordsSet.add(cleanWord);
              }
            });
          }
          if (filename) {
            filename.split(/[\s\-_().]+/).forEach((word: string) => {
              const cleanWord = word.trim();
              if (cleanWord.length >= 3) {
                wordsSet.add(cleanWord);
              }
            });
          }
        });
        return Array.from(wordsSet);
      }
    } catch (error) {
      // error fetching indexed terms
    }
    return [];
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
      return [];
    }
  }

  public async searchFileSuggestions(query: string, selectedSites: string[] = []): Promise<Array<{ label: string; url: string }>> {
    if (!query || query.length < 2) return [];

    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const client: any = await this._msGraphClientFactory.getClient('3');
      // Use the same boosted KQL as main search so suggestions match main results
      let boostedQuery = this._buildBoostedQuery(query);
      if (!boostedQuery) return [];

      // Apply site filtering if selectedSites provided (filter to selected site only)
      if (selectedSites && selectedSites.length > 0) {
        const tenantUrl = this._getTenantUrl();
        const siteQueries = selectedSites.map(s => `Path:"${tenantUrl}/sites/${s}*"`);
        boostedQuery += ` AND (${siteQueries.join(' OR ')})`;
      }

      const searchPayload = {
        requests: [
          {
            entityTypes: ['driveItem'],
            query: { queryString: boostedQuery },
            fields: ['name', 'title', 'webUrl'],
            size: 5
          }
        ]
      };

      const response = await client.api('/search/query').version('v1.0').post(searchPayload);
      const hits = response.value?.[0]?.hitsContainers?.[0]?.hits || [];

      const suggestions: Array<{ label: string; url: string }> = [];
      const seen = new Set<string>();

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      hits.forEach((hit: any) => {
        const resource = hit.resource || {};
        const filename = resource.name || resource.title || '';
        const url = resource.webUrl || '';
        if (filename && url) {
          let cleanName = filename.replace(/\.(pdf|docx?|xlsx?|pptx?|png|jpg|jpeg|gif|zip|txt|csv|md)$/i, '');
          cleanName = cleanName.replace(/[_-]/g, ' ').trim();

          if (cleanName && cleanName.length > 1) {
            const lower = cleanName.toLowerCase();
            if (!seen.has(lower)) {
              seen.add(lower);
              suggestions.push({ label: cleanName, url });
            }
          }
        }
      });

      return suggestions.slice(0, 5);
    } catch (error) {
      return [];
    }
  }

  public async getTitleForUrl(webUrl: string): Promise<string | null> {

    if (!webUrl) return null;

    if (!this._spHttpClient) {
      return null;
    }

    try {
      const tenantUrl = this._getTenantUrl();
      const projectsSiteUrl = `${tenantUrl}/sites/Projects`;
      const listTitle = 'ProjectsNew';
      const absoluteUrl = webUrl.trim();
      let serverRelativeUrl = absoluteUrl;
      let urlWithoutQuery = absoluteUrl;
      const encodedAbsoluteUrl = encodeURI(absoluteUrl);
      const tryDecode = (value: string): string => {
        try {
          return decodeURIComponent(value);
        } catch (_err) {
          return value;
        }
      };
      const decodedAbsoluteUrl = tryDecode(absoluteUrl);

      try {
        const parsed = new URL(absoluteUrl);
        serverRelativeUrl = `${parsed.pathname}${parsed.search}`;
        urlWithoutQuery = `${parsed.origin}${parsed.pathname}`;
      } catch (e) {
        // If url is already relative, keep as-is
      }

      const decodedServerRelativeUrl = tryDecode(serverRelativeUrl);
      const decodedUrlWithoutQuery = tryDecode(urlWithoutQuery);
      const siteSuffix = serverRelativeUrl.replace(/^\/sites\/[^/]+/, '');
      const decodedSiteSuffix = tryDecode(siteSuffix);
      const pathAfterSites = siteSuffix.replace(/^\/+/, '');
      const decodedPathAfterSites = tryDecode(pathAfterSites);
      const pathAfterSharedDocuments = pathAfterSites.replace(/^Shared Documents\//i, 'Shared Documents/');
      const decodedPathAfterSharedDocuments = tryDecode(pathAfterSharedDocuments);

      const normalizeUrl = (value: string): string => {
        let normalized = tryDecode(value || '').trim().toLowerCase();
        normalized = normalized.split('#')[0].split('?')[0];
        normalized = normalized.replace(/^https?:\/\/[^/]+/i, '');
        normalized = normalized.replace(/\/forms\/allitems\.aspx$/i, '');
        normalized = normalized.replace(/\/+$/, '');
        normalized = normalized.replace(/%20/g, ' ');
        return normalized;
      };

      // Check per-URL cache to avoid repeated Projects list queries
      try {
        const cacheKeyForUrl = `projectTitle|${normalizeUrl(absoluteUrl || '')}`;
        const cachedTitle = CacheService.get(cacheKeyForUrl);
        if (cachedTitle) {
          return String(cachedTitle);
        }
      } catch (_cacheErr) {
        // ignore cache errors and continue
      }

      const exactMatchUrls = [
        absoluteUrl,
        encodedAbsoluteUrl,
        decodedAbsoluteUrl,
        urlWithoutQuery,
        decodedUrlWithoutQuery,
        serverRelativeUrl,
        decodedServerRelativeUrl,
        serverRelativeUrl.replace(/\/+$/, ''),
        urlWithoutQuery.replace(/\/+$/, '')
      ];

      const suffixMatchUrls = [
        siteSuffix,
        decodedSiteSuffix,
        pathAfterSites,
        decodedPathAfterSites,
        pathAfterSharedDocuments,
        decodedPathAfterSharedDocuments,
        `/${pathAfterSites}`,
        `/${decodedPathAfterSites}`,
        `/${pathAfterSharedDocuments}`,
        `/${decodedPathAfterSharedDocuments}`
      ];

      const targetMatches = Array.from(new Set(exactMatchUrls.concat(suffixMatchUrls)))
        .filter(Boolean)
        .map(normalizeUrl)
        .filter(Boolean) as string[];

      const projectIdMatches = Array.from(new Set(targetMatches.reduce((ids: string[], target) => {
        const projectsFolderMatch = target.match(/\/sites\/projects\/([^/]+)/i);
        const firstProjectFolderToken = projectsFolderMatch ? projectsFolderMatch[1].match(/^(\d{4,})/) : null;
        if (firstProjectFolderToken) {
          ids.push(firstProjectFolderToken[1]);
        }

        const numericPathTokens = target.match(/(?:^|\/)(\d{4,})(?=[\s/_-]|\/|$)/g) || [];
        numericPathTokens.forEach((token) => {
          const id = token.replace(/^\//, '').match(/^(\d{4,})/);
          if (id) ids.push(id[1]);
        });

        return ids;
      }, [])));

      const isMatchingUrl = (candidate: string): boolean => {
        const normalizedCandidate = normalizeUrl(candidate);
        if (!normalizedCandidate) return false;

        return targetMatches.some((target) => {
          if (!target) return false;
          return normalizedCandidate === target ||
            normalizedCandidate.indexOf(target) !== -1 ||
            target.indexOf(normalizedCandidate) !== -1;
        });
      };

      interface IListFieldInfo {
        InternalName?: string;
        Title?: string;
        TypeAsString?: string;
      }

      interface IListItemValueMap {
        [key: string]: unknown;
        Title?: string;
      }

      const getUrlValues = (value: unknown): string[] => {
        if (!value) return [];
        if (typeof value === 'string') return [value];
        if (Array.isArray(value)) {
          return value.reduce((acc: string[], item: unknown) => acc.concat(getUrlValues(item)), []);
        }
        if (typeof value === 'object') {
          const urlValue = value as { Url?: string; url?: string; Description?: string; description?: string };
          const values: string[] = [];
          if (urlValue.Url) values.push(urlValue.Url);
          if (urlValue.url) values.push(urlValue.url);
          if (urlValue.Description) values.push(urlValue.Description);
          if (urlValue.description) values.push(urlValue.description);
          return values;
        }
        return [];
      };

      const fieldsUrl = `${projectsSiteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields?$select=InternalName,Title,TypeAsString,Hidden&$filter=Hidden eq false`;
      const fieldsResponse = await this._spHttpClient.get(
        fieldsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata'
          }
        }
      );

      if (!fieldsResponse.ok) return null;

      const fieldsJson = await fieldsResponse.json();
      const fields = fieldsJson.value || fieldsJson.d?.results || [];
      const idFields = fields
        .filter((field: IListFieldInfo) => {
          const internalName = (field.InternalName || '').toLowerCase();
          const title = (field.Title || '').toLowerCase();
          return internalName === 'projectid' ||
            title === 'projectid' ||
            title === 'project id' ||
            internalName === 'code' ||
            title === 'code';
        })
        .map((field: IListFieldInfo) => field.InternalName)
        .filter((name: string) => !!name)
        .concat([
          'UncategorisedDocumentsUrl',
          'NonCmap',
          'ContractsDocumentsUrl',
          'ProjectDocumentsUrl',
          'BidDocumentsUrl'
        ]);

      const candidateFields = fields
        .filter((field: IListFieldInfo) => {
          const internalName = (field.InternalName || '').toLowerCase();
          const title = (field.Title || '').toLowerCase();
          const type = (field.TypeAsString || '').toLowerCase();
          const isSupportedType = type === 'url' || type === 'text' || type === 'note';

          return isSupportedType && (
            type === 'url' ||
            internalName.indexOf('url') !== -1 ||
            title.indexOf('url') !== -1 ||
            internalName.indexOf('document') !== -1 ||
            title.indexOf('document') !== -1 ||
            internalName.indexOf('noncmap') !== -1 ||
            title.indexOf('noncmap') !== -1 ||
            internalName.indexOf('location') !== -1 ||
            title.indexOf('location') !== -1
          );
        })
        .map((field: IListFieldInfo) => field.InternalName)
        .filter((name: string) => !!name)
        .concat([
          'UncategorisedDocumentsUrl',
          'NonCmap',
          'ContractsDocumentsUrl',
          'ProjectDocumentsUrl',
          'BidDocumentsUrl'
        ]);

      if (candidateFields.length === 0 && idFields.length === 0) return null;

      const selectFields = Array.from(new Set(['Title'].concat(candidateFields).concat(idFields)));
      const itemsUrl = `${projectsSiteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=${selectFields.join(',')}&$top=5000`;
      const itemsResponse = await this._spHttpClient.get(
        itemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata'
          }
        }
      );

      if (itemsResponse.ok) {
        const itemsJson = await itemsResponse.json();
        const items = itemsJson.value || itemsJson.d?.results || [];
        for (const item of items as IListItemValueMap[]) {
          const hasMatchingUrl = candidateFields.some((fieldName: string) => {
            const values = getUrlValues(item[fieldName]);
            return values.some(isMatchingUrl);
          });

          const hasMatchingProjectId = projectIdMatches.length > 0 && idFields.some((fieldName: string) => {
            const value = item[fieldName];
            return value !== undefined &&
              value !== null &&
              projectIdMatches.indexOf(String(value).trim()) !== -1;
          });

          if (hasMatchingUrl || hasMatchingProjectId) {
            try {
              const cacheKeyForUrl = `projectTitle|${normalizeUrl(absoluteUrl || '')}`;
              CacheService.set(cacheKeyForUrl, item.Title, CacheService.TTL.SITE_META);
            } catch (_setErr) {
              // ignore cache failures
            }
            return item.Title || null;
          }
        }
      }
    } catch (err) {
      // error querying ProjectsNew
    }

    return null;
  }
}


