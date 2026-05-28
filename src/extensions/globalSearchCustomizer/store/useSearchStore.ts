import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import { ISearchStore } from '../interface/ISearchStore';

export const useSearchStore = create<ISearchStore>()(
  persist(
    (set) => ({
      searchHistory: [
        'Quarterly Strategy plans',
        'Financial sheets FY24',
        'Brand design guide PDF',
        'HR employee handbook v3'
      ],
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
      name: 'trivandi_search_history_store'
    }
  )
);
