import { ISearchResult } from '../models/ISearchResult';

export function rerankResults(
  results: ISearchResult[],
  getClickCount: (id: string) => number
): ISearchResult[] {
  // Microsoft Graph API returns results already sorted by relevance.
  // The 'rank' property is an ordinal (1 = best, 2 = second best, etc.),
  // so a LARGER rank number actually means WORSE relevance.
  // To avoid reversing the array, we assign a base score derived from their natural index.
  const reranked = results.map((result, index) => {
    const clicks = getClickCount(result.id);
    // Base score starts high (e.g. 10000) and decreases by index.
    // Each click boosts the result significantly (e.g. +50 points).
    const score = (10000 - index) + (clicks * 50);
    return { result, score, clicks };
  });

  // Sort descending by calculated score
  reranked.sort((a, b) => b.score - a.score);

  return reranked.map(item => item.result);
}
