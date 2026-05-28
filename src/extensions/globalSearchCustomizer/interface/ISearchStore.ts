export interface ISearchStore {
  searchHistory: string[];
  addHistoryItem: (query: string) => void;
  removeHistoryItem: (query: string) => void;
  clearHistory: () => void;
}
