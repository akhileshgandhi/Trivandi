/* eslint-disable @rushstack/no-new-null */
export const STATIC_BRAND_TERMS: string[] = ["Trivandi", "moreYeahs", "SharePoint"];


export function getLevenshteinDistance(a: string, b: string): number {
  const tmp: number[][] = [];
  for (let i = 0; i <= a.length; i++) {
    tmp[i] = [i];
  }
  for (let j = 0; j <= b.length; j++) {
    tmp[0][j] = j;
  }
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      tmp[i][j] = Math.min(
        tmp[i - 1][j] + 1, // deletion
        tmp[i][j - 1] + 1, // insertion
        tmp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1) // substitution
      );
    }
  }
  return tmp[a.length][b.length];
}

export function fuzzyBrandMatch(query: string, terms: string[]): string | null {
  if (!query) return null;
  const cleanQuery = query.trim();
  if (!cleanQuery) return null;

  // 1. Check full query match first
  for (const term of terms) {
    const dist = getLevenshteinDistance(cleanQuery.toLowerCase(), term.toLowerCase());
    if (dist > 0 && dist <= 2) {
      return term;
    }
  }

  // 2. Check word-by-word
  const words = cleanQuery.split(/\s+/);
  let changed = false;
  const correctedWords = words.map(word => {
    for (const term of terms) {
      const dist = getLevenshteinDistance(word.toLowerCase(), term.toLowerCase());
      if (dist > 0 && dist <= 2 && word.length >= 3) {
        changed = true;
        return term;
      }
    }
    return word;
  });

  return changed ? correctedWords.join(' ') : null;
}
