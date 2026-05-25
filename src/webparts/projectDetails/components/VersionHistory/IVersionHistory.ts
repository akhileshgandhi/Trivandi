export interface IFileVersion {
  id: string;
  versionNumber: number;
  displayNumber: string;
  created: Date;
  modifiedBy: string;
  modifiedByEmail?: string;
  size: number;
  comments: string;
  isCurrentVersion: boolean;
  isMinorVersion: boolean;
  checkInComment?: string;
  url?: string;
  fileRef?: string;
  expiryInfo?: {
    expiryDate?: Date;
    accessDuration?: number;
    status: 'Active' | 'Expired' | 'Permanent';
  };
}

export interface IVersionHistoryState {
  isOpen: boolean;
  selectedDocument?: ISelectedDocument;
  versions: IFileVersion[];
  loading: boolean;
  error?: string;
  sortBy: 'date' | 'version';
  sortOrder: 'asc' | 'desc';
  selectedVersionId?: string;
  deleteConfirmOpen: boolean;
  deleteTarget?: 'single' | 'all';
  deleteVersionId?: string;
  deletingAll: boolean;
  restoreLoading: boolean;
}

export interface ISelectedDocument {
  fileRef: string;
  name: string;
  serverRelativeUrl: string;
  listItemId?: number;
}

export interface IVersionHistoryProps {
  context: any;
  isOpen: boolean;
  document?: ISelectedDocument;
  onDismiss: () => void;
  onVersionRestored?: (version: IFileVersion) => void;
  onVersionDeleted?: () => void;
}

export interface IVersionResponse {
  value: ISharePointFileVersion[];
}

export interface ISharePointFileVersion {
  ID: string;
  VersionNumber: string;
  Created: string;
  ModifiedBy: {
    Title: string;
    EMail: string;
  };
  CheckInComment?: string;
  Size?: number;
}

export interface IFileVersionItem {
  ID: number;
  FileLeafRef: string;
  FileRef: string;
  File_x0020_Size: number;
  Modified: string;
  Author: {
    Title: string;
    EMail: string;
  };
}
