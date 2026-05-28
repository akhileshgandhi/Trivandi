import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import { ISearchStore } from '../interface/ISearchStore';

export const useSearchStore = create<ISearchStore>()(
  persist(
    (set) => ({
      searchHistory: [], // Starts completely empty and 100% dynamic!
      addHistoryItem: (query) => set((state) => {
        const cleaned = query.trim();
        if (!cleaned) return state;
        const filtered = state.searchHistory.filter((item) => item.toLowerCase() !== cleaned.toLowerCase());
        return { searchHistory: [cleaned, ...filtered] }; // Keeps ALL history items indefinitely!
      }),
      removeHistoryItem: (query) => set((state) => ({
        searchHistory: state.searchHistory.filter((item) => item !== query)
      })),
      clearHistory: () => set({ searchHistory: [] })
    }),
    {
      name: 'trivandi_search_history_store_v2' // Incremented key to automatically clear stale browser cache!
    }
  )
);
