/* eslint-disable @rushstack/no-new-null */
import { useState, useEffect } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { STATIC_BRAND_TERMS } from '../../../utils/brandDictionary';

function getInitialTerms(): string[] {
  const initialSet = new Set<string>();
  
  // Seed with STATIC_BRAND_TERMS words
  STATIC_BRAND_TERMS.forEach(term => {
    initialSet.add(term);
    term.split(/[\s\-_().]+/).forEach(word => {
      const cleanWord = word.trim();
      if (cleanWord.length >= 3) {
        initialSet.add(cleanWord);
      }
    });
  });

  // Seed with history immediately
  try {
    const historyStr = localStorage.getItem('searchHistory');
    if (historyStr) {
      const historyTerms = JSON.parse(historyStr);
      if (Array.isArray(historyTerms)) {
        historyTerms.forEach((term: string) => {
          if (typeof term === 'string') {
            term.split(/[\s\-_().]+/).forEach(word => {
              const cleanWord = word.trim();
              if (cleanWord.length >= 3) {
                initialSet.add(cleanWord);
              }
            });
            initialSet.add(term);
          }
        });
      }
    }
  } catch (e) {
    console.error('Error parsing searchHistory on mount:', e);
  }

  return Array.from(initialSet);
}

export function useDynamicBrandTerms(service: GraphSearchService | null): string[] {
  const [terms, setTerms] = useState<string[]>(getInitialTerms);

  useEffect(() => {
    if (!service) return;

    const loadTerms = async (): Promise<void> => {
      try {
        const [indexedTerms, peopleNames] = await Promise.all([
          service.fetchIndexedTerms(),
          service.fetchPeopleNames()
        ]);

        let historyTerms: string[] = [];
        try {
          const historyStr = localStorage.getItem('searchHistory');
          if (historyStr) {
            historyTerms = JSON.parse(historyStr) || [];
            if (!Array.isArray(historyTerms)) {
              historyTerms = [];
            }
          }
        } catch (e) {
          console.error('Error reading search history from localStorage:', e);
        }

        const combined = new Set<string>([
          ...STATIC_BRAND_TERMS,
          ...indexedTerms,
          ...peopleNames,
          ...historyTerms
        ]);

        // Clean values, ensuring length >= 3
        const finalTerms = Array.from(combined).filter(
          (t) => typeof t === 'string' && t.trim().length >= 3
        );

        setTerms(finalTerms);
      } catch (err) {
        console.error('Error fetching dynamic brand terms:', err);
        // Keep initial STATIC_BRAND_TERMS on error
      }
    };

    loadTerms().catch(err => console.error('Unhandled error in loadTerms:', err));
  }, [service]);

  return terms;
}
