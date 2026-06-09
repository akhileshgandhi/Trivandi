import { ISearchResult } from '../models/ISearchResult';

export function rerankResults(
  results: ISearchResult[],
  getClickCount: (id: string) => number,
  query: string = ''
): ISearchResult[] {
  // Microsoft Graph API returns results already sorted by relevance.
  // The 'rank' property is an ordinal (1 = best, 2 = second best, etc.),
  // so a LARGER rank number actually means WORSE relevance.
  // To avoid reversing the array, we assign a base score derived from their natural index.
  const reranked = results.map((result, index) => {
    const clicks = getClickCount(result.id);
    // Base score starts high (e.g. 10000) and decreases by index.
    // Each click boosts the result significantly (e.g. +50 points).
    let score = (10000 - index) + (clicks * 50);

    const titleLower = result.title.toLowerCase();
    const queryLower = query.toLowerCase().trim();
    
    if (queryLower) {
      // Exact title match → massive boost
      if (titleLower === queryLower) {
        score += 100000;
      }
      // Title starts with query → very strong boost
      else if (titleLower.startsWith(queryLower)) {
        score += 80000;
      }
      // Title contains exact query as whole word → strong boost
      else if (titleLower.includes(queryLower)) {
        score += 50000;
      }
      // All query words present in title → moderate boost
      else {
        const queryWords = queryLower.split(/\s+/)
          .filter(w => w.length > 1);
        const allWordsInTitle = queryWords.every(
          w => titleLower.includes(w)
        );
        if (allWordsInTitle && queryWords.length > 0) {
          score += 20000;
        }
      }
      
      // Extra boost for folders when query matches folder name
      // Folders are navigation containers — exact match 
      // should always beat content-only file matches
      if (result.fileType === 'folder' && 
          titleLower.includes(queryLower)) {
        score += 30000;
      }
    }

    return { result, score, clicks };
  });

  // Sort descending by calculated score
  reranked.sort((a, b) => b.score - a.score);

  return reranked.map(item => item.result);
}
