import * as React from 'react';

export interface IFiltersPanelProps {
  sidebarWidth: number;
  fileTypes: string[];
  author: string;
  date: string;
  setFilters: (filters: { fileTypes: string[]; author: string; date: string }) => void;
  toggleFileType: (type: string) => void;
  isAuthorDropdownOpen: boolean;
  setIsAuthorDropdownOpen: (open: boolean) => void;
  startResizingSidebar: (e: React.MouseEvent) => void;
  isResizingSidebar: boolean;
  authorsList?: string[];
}
