import { ISearchResult } from '../models/ISearchResult';

export function rerankResults(
  results: ISearchResult[],
  getClickCount: (id: string) => number
): ISearchResult[] {
  const reranked = [...results].sort((a, b) => {
    const clicksA = getClickCount(a.id);
    const clicksB = getClickCount(b.id);
    const scoreA = (a.relevanceScore ?? 0) + clicksA * 10;
    const scoreB = (b.relevanceScore ?? 0) + clicksB * 10;
    return scoreB - scoreA;
  });

  console.log('--- [DEBUG] Re-ranked results ---', reranked.map(r => ({
    title: r.title,
    graphRank: r.relevanceScore,
    clicks: getClickCount(r.id),
    finalScore: (r.relevanceScore ?? 0) + getClickCount(r.id) * 10
  })));

  return reranked;
}
