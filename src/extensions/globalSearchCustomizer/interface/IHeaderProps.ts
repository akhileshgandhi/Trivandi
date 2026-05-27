export interface IHeaderProps {
  searchQuery: string;
  setSearchQuery: (query: string) => void;
  handleSearch: (query: string) => void;
  isSearchFocused: boolean;
  setIsSearchFocused: (focused: boolean) => void;
  searchHistory: string[];
  totalCountToRender: number;
  onDismiss: () => void;
}
