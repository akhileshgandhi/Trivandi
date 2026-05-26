import * as React from 'react';

export interface ISearchHistoryProps {
  isHistoryOpen: boolean;
  setIsHistoryOpen: (open: boolean) => void;
  previewWidth: number;
  searchHistory: string[];
  handleSearch: (query: string) => void;
  setSearchHistory: React.Dispatch<React.SetStateAction<string[]>>;
  removeHistoryItem: (item: string) => void;
  isResizingPreview: boolean;
  startResizingPreview: (e: React.MouseEvent) => void;
}
