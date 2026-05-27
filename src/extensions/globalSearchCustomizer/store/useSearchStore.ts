import { create } from 'zustand';
import { ISearchResult } from '../../../models/ISearchResult';

interface SearchStoreState {
  recentlyViewedFiles: ISearchResult[];
  starredFiles: ISearchResult[];
  starredIds: Set<string>;
  addRecentFile: (file: ISearchResult) => void;
  toggleStar: (file: ISearchResult) => void;
}

// Initial state helpers
const getSavedRecent = (): ISearchResult[] => {
  try {
    const saved = sessionStorage.getItem('trivandi_recent_files_data');
    return saved ? JSON.parse(saved) : [];
  } catch (e) {
    return [];
  }
};

const getSavedStarred = (): ISearchResult[] => {
  try {
    const saved = localStorage.getItem('trivandi_starred_files_data');
    return saved ? JSON.parse(saved) : [];
  } catch (e) {
    return [];
  }
};

export const useSearchStore = create<SearchStoreState>((set) => {
  const initialStarred = getSavedStarred();
  const initialRecent = getSavedRecent();
  const initialStarredIds = new Set(initialStarred.map(f => f.id));

  return {
    recentlyViewedFiles: initialRecent,
    starredFiles: initialStarred,
    starredIds: initialStarredIds,

    addRecentFile: (file: ISearchResult): void => {
      set((state) => {
        const isStarred = state.starredIds.has(file.id);
        const fileWithStar = { ...file, isStarred };
        const filtered = state.recentlyViewedFiles.filter(f => f.id !== file.id);
        const updated = [fileWithStar, ...filtered];

        try {
          sessionStorage.setItem('trivandi_recent_files_data', JSON.stringify(updated));
        } catch (e) {
          // ignored
        }

        return { recentlyViewedFiles: updated };
      });
    },

    toggleStar: (file: ISearchResult): void => {
      set((state) => {
        const hasStar = state.starredIds.has(file.id);
        const nextStarredIds = new Set(state.starredIds);
        let nextStarredFiles = [...state.starredFiles];

        if (hasStar) {
          nextStarredIds.delete(file.id);
          nextStarredFiles = nextStarredFiles.filter(f => f.id !== file.id);
        } else {
          nextStarredIds.add(file.id);
          nextStarredFiles = [...nextStarredFiles, { ...file, isStarred: true }];
        }

        // Sync starred state in recently viewed files too
        const nextRecent = state.recentlyViewedFiles.map(f => {
          if (f.id === file.id) {
            return { ...f, isStarred: !hasStar };
          }
          return f;
        });

        try {
          localStorage.setItem('trivandi_starred_files_data', JSON.stringify(nextStarredFiles));
          sessionStorage.setItem('trivandi_recent_files_data', JSON.stringify(nextRecent));
        } catch (e) {
          // ignored
        }

        return {
          starredIds: nextStarredIds,
          starredFiles: nextStarredFiles,
          recentlyViewedFiles: nextRecent
        };
      });
    }
  };
});
