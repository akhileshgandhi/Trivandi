import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import { Web } from "@pnp/sp/webs";

export interface IPrivateItem {
    Id: string;
    Name: string;
    ServerRelativeUrl: string;
    Modified: string;
    ModifiedBy: string;
    SizeOrItems: string;
    IsFolder: boolean;
    FileExtension?: string;
}

export class PrivateService {
    private sp: any;
    private libraryName: string;

    constructor(context: any, libraryName: string = "Private") {
        this.sp = spfi().using(SPFx(context));
        this.libraryName = libraryName;
    }

    public async uploadFile(folderPath: string, file: File): Promise<any> {
        try {
            const folder = this.sp.web.getFolderByServerRelativePath(folderPath);
            if (file.size <= 10485760) {
                return await folder.files.addUsingPath(file.name, file, { Overwrite: true });
            } else {
                return await folder.files.addChunked(file.name, file, undefined, true);
            }
        } catch (error) {
            console.error("❌ Error uploading file:", error);
            throw error;
        }
    }

    public async createFolder(folderPath: string, folderName: string): Promise<any> {
        try {
            const folder = this.sp.web.getFolderByServerRelativePath(folderPath);
            return await folder.folders.addUsingPath(folderName);
        } catch (error) {
            console.error("❌ Error creating folder:", error);
            throw error;
        }
    }

    public async downloadFile(serverRelativeUrl: string, fileName: string): Promise<void> {
        try {
            const file = await this.sp.web.getFileByServerRelativePath(serverRelativeUrl).getBlob();
            const url = window.URL.createObjectURL(file);
            const a = document.createElement("a");
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error("❌ Error downloading file:", error);
        }
    }

    public async getLibraryContents(folderPath?: string): Promise<{ folders: any[]; files: any[]; currentPath: string }> {
        try {
            // 1. Get library root folder if no path provided
            let normalizedPath = folderPath;
            if (!normalizedPath) {
                const library = await this.sp.web.lists.getByTitle(this.libraryName).rootFolder.select("ServerRelativeUrl")();
                normalizedPath = library.ServerRelativeUrl;
            }

            // 2. Get current user email for filtering
            const user = await this.sp.web.currentUser();
            const userEmail = user.Email;

            // 3. Fetch using list items query with user filter
            const items = await this.sp.web.lists.getByTitle(this.libraryName).items
                .filter(`FileDirRef eq '${normalizedPath}' and Author/EMail eq '${userEmail}'`)
                .select('ID', 'FileLeafRef', 'FileRef', 'Modified', 'Editor/Title', 'Editor/ID', 'Author/Title', 'Author/ID', 'Author/EMail', 'FSObjType', 'File/Length', 'Folder/ItemCount', 'UniqueId')
                .expand('Editor', 'Author', 'File', 'Folder')();

            console.log(`📂 Private Content for ${userEmail}:`, items);

            const folders: any[] = [];
            const files: any[] = [];

            // 3. Separate into folders and files
            items.forEach((item: any) => {
                const baseItem = {
                    Id: item.UniqueId || item.ID,
                    Name: item.FileLeafRef,
                    ServerRelativeUrl: item.FileRef,
                    TimeLastModified: item.Modified,
                    Editor: { 
                        Title: item.Editor?.Title || "System",
                        ID: item.Editor?.ID 
                    },
                    Author: { 
                        Title: item.Author?.Title || "System",
                        ID: item.Author?.ID
                    }
                };

                if (item.FSObjType === 1) { // Folder
                    folders.push({
                        ...baseItem,
                        ItemCount: item.Folder?.ItemCount || 0,
                        IsFolder: true
                    });
                } else { // File
                    files.push({
                        ...baseItem,
                        Length: item.File?.Length || 0,
                        IsFolder: false
                    });
                }
            });

            return { 
                folders: folders.filter(f => f.Name !== "Forms"),
                files,
                currentPath: normalizedPath 
            };
        } catch (error) {
            console.error("❌ Error fetching library contents:", error);
            return { folders: [], files: [], currentPath: "" };
        }
    }

    public formatBytes(bytes: string | number, decimals = 1): string {
        const b = typeof bytes === 'string' ? parseInt(bytes) : bytes;
        if (!b || b === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(b) / Math.log(k));
        return parseFloat((b / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }
}
