import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";

export class CacheService {
  public static readonly serviceKey: ServiceKey<CacheService> =
    ServiceKey.create<CacheService>("Trivandi:CacheService", CacheService);

  private static _cache = new Map<string, { data: any; timestamp: number; ttl: number }>();
  private static DEFAULT_TTL = 5 * 60 * 1000;
  private static MAX_ENTRIES = 100;
  private static _hits = 0;
  private static _misses = 0;
  private static _cleanupTimer: number | null = null;

  public static TTL = {
    SEARCH:      5  * 60 * 1000,
    THUMBNAILS:  10 * 60 * 1000,
    USER_PHOTOS: 30 * 60 * 1000,
    SITE_META:   60 * 60 * 1000,
  } as const;

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
    return [
      query.trim().toLowerCase(),
      JSON.stringify([...fileTypes].sort()),
      date,
      JSON.stringify([...selectedAuthors].sort()),
      JSON.stringify([...selectedSites].sort()),
      from,
      activeTopTab,
      sortBy,
    ].join("|");
  }

  public static get(key: string): any | null {
    const entry = this._cache.get(key);
    if (!entry) {
      this._misses++;
      return null;
    }
    if (Date.now() - entry.timestamp > entry.ttl) {
      this._cache.delete(key);
      this._misses++;
      return null;
    }
    this._hits++;
    return entry.data;
  }

  public static set(key: string, data: any, ttl = this.DEFAULT_TTL): void {
    if (data === null || data === undefined) return;
    if (this._cache.size >= this.MAX_ENTRIES && !this._cache.has(key)) {
      this._evictOldest();
    }
    this._cache.set(key, { data, timestamp: Date.now(), ttl });
  }

  public static invalidate(key: string): void {
    this._cache.delete(key);
  }

  public static invalidateByPrefix(prefix: string): void {
    for (const k of this._cache.keys()) {
      if (k.startsWith(prefix)) this._cache.delete(k);
    }
  }

  public static clear(): void {
    this._cache.clear();
    this._hits = 0;
    this._misses = 0;
  }

  public static getStats(): { hits: number; misses: number; hitRate: string; size: number } {
    const total = this._hits + this._misses;
    return {
      hits: this._hits,
      misses: this._misses,
      hitRate: total === 0 ? "0%" : `${((this._hits / total) * 100).toFixed(1)}%`,
      size: this._cache.size,
    };
  }

  public static startCleanup(intervalMs = 2 * 60 * 1000): void {
    if (this._cleanupTimer !== null) return;
    this._cleanupTimer = window.setInterval(() => this._evictExpired(), intervalMs);
  }

  public static stopCleanup(): void {
    if (this._cleanupTimer !== null) {
      clearInterval(this._cleanupTimer);
      this._cleanupTimer = null;
    }
  }

  private static _evictExpired(): void {
    const now = Date.now();
    for (const [k, v] of this._cache.entries()) {
      if (now - v.timestamp > v.ttl) this._cache.delete(k);
    }
  }

  private static _evictOldest(): void {
    let oldestKey = "";
    let oldestTime = Infinity;
    for (const [k, v] of this._cache.entries()) {
      if (v.timestamp < oldestTime) {
        oldestTime = v.timestamp;
        oldestKey = k;
      }
    }
    if (oldestKey) this._cache.delete(oldestKey);
  }
}
