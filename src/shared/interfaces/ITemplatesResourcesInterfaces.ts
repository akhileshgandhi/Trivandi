/* ================= DOCUMENT LIBRARY INTERFACES ================= */

export interface IDocumentItem {
  Id: number;
  Title: string;
  Name: string;
  FileRef: string;
  FileLeafRef: string;
  File_x0020_Type: string;
  Modified: string;
  Created: string;
  ServerRedirectedEmbedUri?: string;
  ServerRedirectedEmbedUrl?: string;
  FileSizeDisplay?: string;
  Author?: {
    Title: string;
    EMail: string;
  };
  Editor?: {
    Title: string;
    EMail: string;
  };
}

export interface IFolderItem {
  Id: number;
  Title: string;
  Name: string;
  FileRef: string;
  FileLeafRef: string;
  FSObjType: number; // 1 for folder, 0 for file
  ItemChildCount?: number;
  Modified: string;
  Created: string;
}

export interface ITemplateDocument extends IDocumentItem {
  Category?: string;
  DocumentType?: "book" | "folder";
  Tags?: string[];
}

export interface IResourceFolder extends IFolderItem {
  Description?: string;
  Color?: string;
  IconType?: string;
}

export interface IDocumentLibraryResponse {
  standardDocuments: ITemplateDocument[];
  invoicingDocuments: ITemplateDocument[];
  resourceFolders: IResourceFolder[];
  allFiles: IDocumentItem[];
  allFolders: IFolderItem[];
  folderStructure: IFolderStructure[];
}

export interface IFolderStructure {
  Id: number;
  Name: string;
  Title: string;
  FileRef: string;
  FSObjType: number;
  Modified: string;
  Created: string;
  ItemChildCount?: number;
  Description?: string;
  Color?: string;
  files: IDocumentItem[];
  subFolders: IFolderStructure[];
  parentPath?: string;
  level: number;
}