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
    fileTypes: string[] = []
  ): Promise<{ results: ISearchResult[]; totalCount: number }> {
    const client: any = await this._msGraphClientFactory.getClient('3');

    // Build the query string using a safer KQL fallback
    let queryString = query.trim() || 'IsDocument:1';
    
    // Add file type filters if present and not "All"
    const actualTypes = fileTypes.filter(t => t !== 'All');
    if (actualTypes.length > 0) {
      const typeQueries = actualTypes.map(t => `filetype:${t.toLowerCase()}`);
      queryString += ` AND (${typeQueries.join(' OR ')})`;
    }

    // Build Graph Search POST payload according to Microsoft Graph Search API guidelines
    const searchPayload = {
      requests: [
        {
          entityTypes: ['driveItem'],
          query: {
            queryString: queryString
          },
          from: from,
          size: pageSize
        }
      ]
    };

    console.log('--- Sending MS Graph Search Payload ---');
    console.log(JSON.stringify(searchPayload, null, 2));

    const response = await client
      .api('/search/query')
      .version('v1.0')
      .post(searchPayload);

    // Parse response objects
    const searchResponse = response.value?.[0];
    const hitsContainer = searchResponse?.hitsContainers?.[0];
    const totalCount = hitsContainer?.total || 0;
    const hits = hitsContainer?.hits || [];

    const results: ISearchResult[] = hits.map((hit: any) => {
      const resource = hit.resource || {};
      
      // Determine file extension cleanly
      let fileType = 'doc';
      if (resource.file?.mimeType) {
        const mime = resource.file.mimeType.toLowerCase();
        if (mime.includes('pdf')) fileType = 'pdf';
        else if (mime.includes('excel') || mime.includes('spreadsheet')) fileType = 'xlsx';
        else if (mime.includes('presentation') || mime.includes('powerpoint')) fileType = 'pptx';
        else if (mime.includes('image')) fileType = 'png';
      } else {
        const ext = resource.name?.split('.').pop()?.toLowerCase();
        if (ext) fileType = ext;
      }

      return {
        id: resource.id || hit.hitId || Math.random().toString(),
        title: resource.name || 'Untitled Document',
        webUrl: resource.webUrl || '',
        fileType: fileType,
        lastModified: resource.lastModifiedDateTime || new Date().toISOString(),
        author: resource.createdBy?.user?.displayName || 'SharePoint User',
        size: resource.size || 0,
        summary: hit.summary || resource.description || 'No description preview available.',
        siteUrl: resource.parentReference?.siteId || '',
        siteName: resource.parentReference?.sharepointIds?.siteUrl?.split('/').pop() || 'SharePoint Portal'
      };
    });

    return { results, totalCount };
  }
}
