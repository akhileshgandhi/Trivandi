import { GraphSearchService } from '../services/GraphSearchService';

export interface IHeaderProps {
  searchQuery: string;
  setSearchQuery: (query: string) => void;
  handleSearch: (query: string) => void;
  isSearchFocused: boolean;
  setIsSearchFocused: (focused: boolean) => void;
  searchHistory: string[];
  totalCountToRender: number;
  onDismiss: () => void;
  service: GraphSearchService | null;
  isAdmin?: boolean;
  storeIsOwner?: boolean;
  isAdminPanelOpen?: boolean;
  setIsAdminPanelOpen?: (open: boolean) => void;
  selectedSites?: string[];
}
