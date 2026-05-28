import * as React from 'react';

export interface IFiltersPanelProps {
  sidebarWidth: number;
  fileTypes: string[];
  selectedAuthors: string[];
  date: string;
  setFilters: (filters: { fileTypes: string[]; selectedAuthors: string[]; date: string }) => void;
  toggleFileType: (type: string) => void;
  startResizingSidebar: (e: React.MouseEvent) => void;
  isResizingSidebar: boolean;
  authorsList: string[];
  searchService?: any;
}
