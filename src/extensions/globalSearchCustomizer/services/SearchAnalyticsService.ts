export interface IQueryAnalytic {
  query: string;
  count: number;
  lastSearched: string;
}

export interface IClickedDocumentAnalytic {
  id: string;
  title: string;
  url: string;
  clickCount: number;
  lastClicked: string;
}

export class SearchAnalyticsService {
  private static QUERIES_KEY = 'trivandi_search_analytics_queries';
  private static CLICKS_KEY = 'trivandi_search_analytics_clicks';

  /**
   * Tracks a successful search query by incrementing its count
   */
  public static trackSearchQuery(query: string): void {
    const clean = query.trim();
    if (!clean) return;

    try {
      const stored = localStorage.getItem(this.QUERIES_KEY);
      let list: IQueryAnalytic[] = stored ? JSON.parse(stored) : [];

      const existingIndex = list.findIndex(item => item.query.toLowerCase() === clean.toLowerCase());
      if (existingIndex !== -1) {
        list[existingIndex].count += 1;
        list[existingIndex].lastSearched = new Date().toISOString();
        // Keep correct casing if newer
        list[existingIndex].query = clean;
      } else {
        list.push({
          query: clean,
          count: 1,
          lastSearched: new Date().toISOString()
        });
      }

      // Sort by count descending
      list.sort((a, b) => b.count - a.count);
      localStorage.setItem(this.QUERIES_KEY, JSON.stringify(list));
    } catch (e) {
      console.error('Failed to track search query in analytics:', e);
    }
  }

  /**
   * Tracks a clicked search result card
   */
  public static trackCardClick(resultId: string, resultTitle: string, webUrl: string): void {
    if (!resultId || !resultTitle) return;

    try {
      const stored = localStorage.getItem(this.CLICKS_KEY);
      let list: IClickedDocumentAnalytic[] = stored ? JSON.parse(stored) : [];

      const existingIndex = list.findIndex(item => item.id === resultId);
      if (existingIndex !== -1) {
        list[existingIndex].clickCount += 1;
        list[existingIndex].lastClicked = new Date().toISOString();
      } else {
        list.push({
          id: resultId,
          title: resultTitle,
          url: webUrl,
          clickCount: 1,
          lastClicked: new Date().toISOString()
        });
      }

      // Sort by clickCount descending
      list.sort((a, b) => b.clickCount - a.clickCount);
      localStorage.setItem(this.CLICKS_KEY, JSON.stringify(list));
    } catch (e) {
      console.error('Failed to track result card click in analytics:', e);
    }
  }

  /**
   * Gets top search queries sorted by count
   */
  public static getTopQueries(limit: number = 5): IQueryAnalytic[] {
    try {
      const stored = localStorage.getItem(this.QUERIES_KEY);
      const list: IQueryAnalytic[] = stored ? JSON.parse(stored) : [];
      return list.slice(0, limit);
    } catch (e) {
      return [];
    }
  }

  /**
   * Gets top clicked documents
   */
  public static getTopClickedDocuments(limit: number = 5): IClickedDocumentAnalytic[] {
    try {
      const stored = localStorage.getItem(this.CLICKS_KEY);
      const list: IClickedDocumentAnalytic[] = stored ? JSON.parse(stored) : [];
      return list.slice(0, limit);
    } catch (e) {
      return [];
    }
  }

  /**
   * Clears all search analytics databases
   */
  public static clearAnalytics(): void {
    try {
      localStorage.removeItem(this.QUERIES_KEY);
      localStorage.removeItem(this.CLICKS_KEY);
    } catch (e) {
      console.error('Failed to clear analytics databases:', e);
    }
  }

  public static getClickCount(id: string): number {
    try {
      const stored = localStorage.getItem('searchClickAnalytics');
      if (!stored) return 0;
      const dict: Record<string, number> = JSON.parse(stored);
      return dict[id] || 0;
    } catch (e) {
      return 0;
    }
  }

  public static recordClick(id: string): void {
    if (!id) return;
    try {
      const stored = localStorage.getItem('searchClickAnalytics');
      const dict: Record<string, number> = stored ? JSON.parse(stored) : {};
      dict[id] = (dict[id] || 0) + 1;
      localStorage.setItem('searchClickAnalytics', JSON.stringify(dict));
    } catch (e) {
      console.error('Failed to record click for re-ranking:', e);
    }
  }
}
