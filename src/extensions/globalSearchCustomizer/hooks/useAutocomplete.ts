/* eslint-disable @rushstack/no-new-null */
import { useState, useEffect } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { SearchAnalyticsService } from '../services/SearchAnalyticsService';
import { useSearchStore } from '../store/useSearchStore';

export interface ISuggestion {
  label: string;
  type: 'history' | 'file' | 'suggested';
  url?: string;
}

export function useAutocomplete(query: string, service: GraphSearchService | null, selectedSites: string[] = []) {
  const [suggestions, setSuggestions] = useState<ISuggestion[]>([]);
  const [loading, setLoading] = useState(false);
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const searchHistory = useSearchStore((state: any) => state.searchHistory as string[]);

  const clearSuggestions = () => setSuggestions([]);

  useEffect(() => {
    const cleanQuery = query.trim();
    if (cleanQuery.length < 2) {
      setSuggestions([]);
      return;
    }

    const lowerQuery = cleanQuery.toLowerCase();
    
    // 1. Synchronous sources (History + Analytics)
    const historyMatches = searchHistory
      .filter(item => item.toLowerCase().includes(lowerQuery))
      .slice(0, 5);

    const analyticsQueries = SearchAnalyticsService.getTopQueries(20)
      .map(q => q.query)
      .filter(item => item.toLowerCase().includes(lowerQuery) || item.toLowerCase().startsWith(lowerQuery));

    const syncSuggestions: ISuggestion[] = [];
    const seen = new Set<string>();

    // Add History
    for (const match of historyMatches) {
      const lowerMatch = match.toLowerCase();
      if (!seen.has(lowerMatch)) {
        seen.add(lowerMatch);
        syncSuggestions.push({ label: match, type: 'history' });
      }
    }

    // Add Analytics
    for (const match of analyticsQueries) {
      const lowerMatch = match.toLowerCase();
      if (!seen.has(lowerMatch)) {
        seen.add(lowerMatch);
        syncSuggestions.push({ label: match, type: 'suggested' });
      }
    }

    setSuggestions(syncSuggestions.slice(0, 8));

    if (!service) return;

    // 2. Async source (Files)
    setLoading(true);
    const timeoutId = setTimeout(() => {
      service.searchFileSuggestions(cleanQuery, selectedSites).then(files => {
        setSuggestions(prev => {
          const newSuggestions: ISuggestion[] = [];
          const currentSeen = new Set<string>();

          // Keep history (top priority)
          const hist = prev.filter(s => s.type === 'history');
          hist.forEach(s => {
            currentSeen.add(s.label.toLowerCase());
            newSuggestions.push(s);
          });

          // Add files (second priority)
          for (const file of files) {
            const lowerFile = file.label.toLowerCase();
            if (!currentSeen.has(lowerFile)) {
              currentSeen.add(lowerFile);
              newSuggestions.push({ label: file.label, type: 'file', url: file.url });
            }
          }

          // Keep analytics (third priority)
          const sugg = prev.filter(s => s.type === 'suggested');
          for (const s of sugg) {
            const lowerSugg = s.label.toLowerCase();
            if (!currentSeen.has(lowerSugg)) {
              currentSeen.add(lowerSugg);
              newSuggestions.push(s);
            }
          }

          return newSuggestions.slice(0, 8);
        });
        setLoading(false);
      }).catch(err => {
        console.error('Error in useAutocomplete file search', err);
        setLoading(false);
      });
    }, 300);

    return () => clearTimeout(timeoutId);
  }, [query, searchHistory, service, selectedSites]);

  return { suggestions, loading, clearSuggestions };
}
