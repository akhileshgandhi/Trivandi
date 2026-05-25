import { IFileVersion, ISharePointFileVersion, ISelectedDocument } from './IVersionHistory';

export class VersionHistoryService {
    private context: any;

    constructor(context: any) {
        this.context = context;
    }

    /**
     * Fetch all versions of a file using PnPjs (or REST fallback)
     */
    public async getFileVersionsFromPnPjs(
        serverRelativeUrl: string
    ): Promise<IFileVersion[]> {
        // Use REST API as primary method since PnPjs has import issues
        return this.getFileVersionsFromREST(serverRelativeUrl);
    }

    /**
     * Fetch versions using SharePoint REST API (alternative)
     */
    public async getFileVersionsFromREST(
        serverRelativeUrl: string
    ): Promise<IFileVersion[]> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const encodedUrl = encodeURIComponent(serverRelativeUrl);

            const response = await fetch(
                `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativeUrl}')/Versions?$expand=CreatedBy`,
                {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-Type': 'application/json',
                        'X-RequestDigest': (window as any).__REQUESTDIGEST
                    }
                }
            );

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            console.log(response, 'response from rest api');
            const data = await response.json() as any;
            const versions: ISharePointFileVersion[] = data?.value || [];
            console.log(versions, 'versions')
            return versions.map((v: any, idx: number) => {
                const versionNum = parseInt((v.VersionNumber || '').split('.')[0] || `${versions.length - idx}`, 10);
                // SharePoint versions use CreatedBy (not ModifiedBy)
                const modifiedBy = v.CreatedBy?.Title || v.ModifiedBy?.Title || 'Unknown';
                const modifiedByEmail = v.CreatedBy?.EMail || v.ModifiedBy?.EMail || '';
                return {
                    id: v.ID,
                    versionNumber: versionNum,
                    displayNumber: v.VersionNumber || `${versionNum}.0`,
                    created: new Date(v.Created),
                    modifiedBy: modifiedBy,
                    modifiedByEmail: modifiedByEmail,
                    size: v.Size || 0,
                    comments: v.CheckInComment || '',
                    isCurrentVersion: idx === 0,
                    isMinorVersion: (v.VersionNumber || '').endsWith('.0') === false,
                    checkInComment: v.CheckInComment || '',
                    fileRef: serverRelativeUrl,
                    expiryInfo: {
                        expiryDate: undefined,
                        accessDuration: undefined,
                        status: 'Permanent'
                    }
                };
            });
        } catch (error) {
            console.error('Error fetching versions from REST:', error);
            throw new Error(`Failed to fetch file versions: ${(error as any).message}`);
        }
    }

    /**
     * Delete a specific version using REST API
     */
    public async deleteVersion(
        serverRelativeUrl: string,
        versionId: string
    ): Promise<void> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const numericId = versionId.replace('v', '');

            const response = await fetch(
                `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativeUrl}')/Versions(${numericId})/Delete`,
                {
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-RequestDigest': (window as any).__REQUESTDIGEST,
                        'X-HTTP-Method': 'DELETE'
                    }
                }
            );

            if (!response.ok && response.status !== 204) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
        } catch (error) {
            console.error('Error deleting version:', error);
            throw new Error(`Failed to delete version: ${(error as any).message}`);
        }
    }

    /**
     * Delete all versions except current using REST API
     */
    public async deleteAllVersions(serverRelativeUrl: string): Promise<void> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;

            // First get all versions
            const response = await fetch(
                `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativeUrl}')/Versions`,
                {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-RequestDigest': (window as any).__REQUESTDIGEST
                    }
                }
            );

            const data = await response.json() as any;
            const versions = data.value || [];

            // Delete all versions except the first one (current version)
            for (let i = 1; i < versions.length; i++) {
                try {
                    await this.deleteVersion(serverRelativeUrl, `v${versions[i].ID}`);
                    // Add small delay to prevent throttling
                    await new Promise(resolve => setTimeout(resolve, 100));
                } catch (versionError) {
                    console.warn(`Failed to delete version ${versions[i].ID}:`, versionError);
                }
            }
        } catch (error) {
            console.error('Error deleting all versions:', error);
            throw new Error(`Failed to delete all versions: ${(error as any).message}`);
        }
    }

    /**
     * Restore a specific version
     */
    public async restoreVersion(
        serverRelativeUrl: string,
        versionId: string
    ): Promise<void> {
        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const numericId = versionId.replace('v', '');

            // In SharePoint, restoring is done by deleting the current version
            // and the previous version becomes current
            const response = await fetch(
                `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativeUrl}')/Versions(${numericId})/Restore`,
                {
                    method: 'POST',
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json',
                        'X-RequestDigest': (window as any).__REQUESTDIGEST
                    }
                }
            );

            if (!response.ok && response.status !== 204) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
        } catch (error) {
            console.error('Error restoring version:', error);
            throw new Error(`Failed to restore version: ${(error as any).message}`);
        }
    }

    /**
     * Get file icon based on extension
     */
    public getFileIcon(fileName: string): string {
        const parts = fileName.split('.');
        const ext = parts.length > 0 ? parts[parts.length - 1].toLowerCase() : '';

        const iconMap: { [key: string]: string } = {
            'pdf': 'FilePDF',
            'doc': 'WordDocument',
            'docx': 'WordDocument',
            'xls': 'ExcelDocument',
            'xlsx': 'ExcelDocument',
            'ppt': 'PowerPointDocument',
            'pptx': 'PowerPointDocument',
            'png': 'ImagePixel',
            'jpg': 'ImagePixel',
            'jpeg': 'ImagePixel',
            'gif': 'ImagePixel',
            'bmp': 'ImagePixel',
            'txt': 'TextDocument',
            'zip': 'ZipFolder',
            'rar': 'ZipFolder',
            'exe': 'BranchFork',
            'csv': 'ExcelDocument'
        };

        return iconMap[ext] || 'Page';
    }

    /**
     * Format file size
     */
    public formatFileSize(bytes: number): string {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
    }

    /**
     * Format date to readable format
     */
    public formatDate(date: Date): string {
        return new Date(date).toLocaleString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        });
    }
}
