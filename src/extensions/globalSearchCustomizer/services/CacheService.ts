export class CacheService {
  private static _cache = new Map<string, { data: any; timestamp: number }>();
  private static TTL = 5 * 60 * 1000; // 5 minutes

  public static buildKey(
    query: string, 
    fileTypes: string[], 
    date: string, 
    selectedAuthors: string[], 
    selectedSites: string[],
    from: number,
    activeTopTab: string,
    sortBy: string
  ): string {
    return `${query}|${JSON.stringify(fileTypes)}|${date}|${JSON.stringify(selectedAuthors)}|${JSON.stringify(selectedSites)}|${from}|${activeTopTab}|${sortBy}`;
  }

  public static get(key: string): any | null {
    const entry = this._cache.get(key);
    if (!entry) return null;
    if (Date.now() - entry.timestamp > this.TTL) {
      this._cache.delete(key);
      return null;
    }
    return entry.data;
  }

  public static set(key: string, data: any): void {
    this._cache.set(key, { data, timestamp: Date.now() });
  }

  public static clear(): void {
    this._cache.clear();
  }
}
