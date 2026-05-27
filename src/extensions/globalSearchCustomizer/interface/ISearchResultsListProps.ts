import * as React from 'react';
import { ISearchResult } from '../../../models/ISearchResult';

export interface ISearchResultsListProps {
  resultsToRender: ISearchResult[];
  totalCountToRender: number;
  isLoading: boolean;
  selectedFileId: string | null;
  setSelectedFileId: (id: string | null) => void;
  starredIds: Set<string>;
  toggleStar: (e: React.MouseEvent, id: string) => void;
  from: number;
  setFrom: (from: number) => void;
  bottomTab: 'Search Results' | 'Recent Activities' | 'Starred Assets';
  setBottomTab: (tab: 'Search Results' | 'Recent Activities' | 'Starred Assets') => void;
  setIsHistoryOpen: (open: boolean) => void;
  isHistoryOpen: boolean;
  handleClearAllFilters: () => void;
  activeTopTab: string;
  setActiveTopTab: (tab: string) => void;
  filters: {
    fileTypes: string[];
    author: string;
    date: string;
  };
  setFilters: (filters: { fileTypes: string[]; author: string; date: string }) => void;
  setOpenMenuId: (id: string | null) => void;
  openMenuId: string | null;
}
