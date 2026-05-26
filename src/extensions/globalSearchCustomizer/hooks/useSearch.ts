import { useState, useEffect, useCallback } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { ISearchResult } from '../../../models/ISearchResult';

export interface IUseSearchOptions {
  service: GraphSearchService | null;
  initialQuery?: string;
  pageSize?: number;
}

export function useSearch({ service, initialQuery = '', pageSize = 20 }: IUseSearchOptions) {
  const [query, setQuery] = useState(initialQuery);
  const [fileTypes, setFileTypes] = useState<string[]>(['All']);
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [from, setFrom] = useState(0);

  const executeSearch = useCallback(async () => {
    if (!service) {
      setError('Search service is not initialized.');
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await service.search(query, pageSize, from, fileTypes);
      setResults(res.results);
      setTotalCount(res.totalCount);
    } catch (err: any) {
      console.error('Error executing Graph Search in hook:', err);
      setError(err?.message || 'Failed to fetch search results from Microsoft Graph API.');
    } finally {
      setLoading(false);
    }
  }, [service, query, pageSize, from, fileTypes]);

  useEffect(() => {
    if (service) {
      executeSearch();
    }
  }, [executeSearch, service]);

  return {
    query,
    setQuery,
    fileTypes,
    setFileTypes,
    results,
    totalCount,
    loading,
    error,
    from,
    setFrom,
    pageSize,
    refresh: executeSearch
  };
}
