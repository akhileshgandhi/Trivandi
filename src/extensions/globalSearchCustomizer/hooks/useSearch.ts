import { useState, useEffect, useCallback } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { ISearchResult } from '../../../models/ISearchResult';

import { IUseSearchOptions } from '../interface/IUseSearchOptions';

export function useSearch({ service, initialQuery = '', pageSize = 20 }: IUseSearchOptions) {
  const [query, setQuery] = useState(initialQuery);
  const [fileTypes, setFileTypes] = useState<string[]>(['All']);
  const [activeTopTab, setActiveTopTab] = useState<string>('All');
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
      const res = await service.search(query, pageSize, from, fileTypes, activeTopTab);
      setResults(res.results);
      setTotalCount(res.totalCount);
    } catch (err: any) {
      console.error('Error executing Graph Search in hook:', err);
      if (err?.statusCode === 429 || (err?.message && err.message.includes('429'))) {
        setError('Too many requests. Please wait a moment before searching again.');
      } else {
        setError(err?.message || 'Failed to fetch search results from Microsoft Graph API.');
      }
    } finally {
      setLoading(false);
    }
  }, [service, query, pageSize, from, fileTypes, activeTopTab]);

  // Reset pagination offset to 0 whenever the query, activeTopTab, or filters change
  useEffect(() => {
    setFrom(0);
  }, [query, fileTypes, activeTopTab]);

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
    activeTopTab,
    setActiveTopTab,
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
