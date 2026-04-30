export interface IProjectExternalPortalState {
  activeTab: 'dashboard' | 'documents' | 'external-portal';
  loading: boolean;
  error?: string;
  projectInfo?: any;
  showInviteGuestDialog: boolean;
  showShareDocumentDialog: boolean;
  showImportDialog: boolean;
  currentFolderPath: string;
  breadcrumbs: IBreadcrumb[];
}

export interface IBreadcrumb {
  text: string;
  key: string;
  path: string;
}

export interface ISharedDocument {
  id: number;
  name: string;
  modified: Date;
  modifiedBy: string;
  path: string;
  fileRef: string;
  isFolder: boolean;
  externalSharingType?: string;
  sharedWithGuests?: string[];
  permission?: 'Read' | 'Review' | 'Edit' | 'Admin';
}

export interface IGuestUser {
  id: number;
  title: string;
  email: string;
  company: string;
  role: 'Viewer' | 'Editor';
  status: 'Active' | 'Inactive' | 'Expired' | 'Revoked';
  lastAccessDate?: Date;
  accessExpiryDate?: Date;
  invitedDate: Date;
}

export interface IGuestAccessReportEntry {
  guestName: string;
  guestEmail: string;
  guestRole: 'Viewer' | 'Editor';
  guestCompany?: string;
  guestStatus: 'Active' | 'Inactive' | 'Expired' | 'Revoked';
  fileName: string;
  filePath: string;
  permission: 'Read' | 'Review' | 'Edit' | 'Admin' | 'No Access';
  sharedOn?: Date;
  guestLastAccess?: Date;
}

export interface ISharedFileAccessLog {
  title: string;
  actionType: string;
  document: string;
  documentName: string;
  sharedBy: string;
  guestEmail: string;
  permission?: 'Read' | 'Review' | 'Edit' | 'Admin';
  accessDuration?: number;
  expiryDate?: Date;
  actionDate?: Date;
}

export interface IShareDocumentData {
  externalSharing: string;
  type?: string;
  name?: string;
  file?: File;
  selectedUsers?: string[];
  accessDuration?: number;
}

export interface IImportDocumentData {
  selectedDocuments: any[];
  targetFolder: string;
}
