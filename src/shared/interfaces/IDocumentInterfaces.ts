import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface DocumentItem {
  Id: number;
  FileLeafRef: string;
  Modified: string;
  Editor?: { Title: string };
  File_x0020_Type?: string;
  FileRef?: string;
  FSObjType?: number;
  VersionLabel?: string;
  Status?: string;
  ServerRelativeUrl?: string;
  FileDirRef?: string;
  ItemCount?: number;
  Length?: number;
}

export interface IProjectDocumentsProps {
  projectId: number;
  projectCode?: string;
  projectTitle?: string;
  businessProjectId?: string;
  hideNewButton?: boolean;
  onNewClick?: () => void;
  context: WebPartContext;
  onShowShareAccess?: (selectedDocs: any[]) => void;
  isUserRestricted?: boolean;
  isNonCmap?: boolean;
  location?: string;
  uncategorisedDocumentsUrl?: string;
}

export interface IBidDocumentsProps {
  projectId: number;
  projectCode?: string;
  projectTitle?: string;
  businessProjectId?: string;
  hideNewButton?: boolean;
  onNewClick?: () => void;
  context: WebPartContext;
  onShowShareAccess?: (selectedDocs: any[]) => void;
  isUserRestricted?: boolean;
  isNonCmap?: boolean;
  location?: string;
  uncategorisedDocumentsUrl?: string;
}

export interface IDocumentsProps {
  projectId: number;
  projectCode?: string;
  projectTitle?: string;
  businessProjectId?: string;
  hideNewButton?: boolean;
  onNewClick?: () => void;
  context: WebPartContext;
  onShowShareAccess?: (selectedDocs: any[]) => void;
  isUserRestricted?: boolean;
  isNonCmap?: boolean;
  location?: string;
  uncategorisedDocumentsUrl?: string;
}

export interface BreadcrumbItem {
  name: string;
  path: string;
}

export interface PaginationInfo {
  currentPage: number;
  pageSize: number;
  totalItems: number;
  totalPages: number;
}

export type DocumentTabType = 'project' | 'contract' | 'bid';
export type SortField = 'name' | 'size' | 'modified' | 'by';
export type SortDirection = 'asc' | 'desc';
