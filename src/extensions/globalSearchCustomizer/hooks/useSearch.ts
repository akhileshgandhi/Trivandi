import { useState, useEffect, useCallback, useRef } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { ISearchResult } from '../../../models/ISearchResult';

import { IUseSearchOptions } from '../interface/IUseSearchOptions';

export function useSearch({ service, initialQuery = '', pageSize = 20 }: IUseSearchOptions) {
  const [query, setQuery] = useState(initialQuery);
  const [fileTypes, setFileTypes] = useState<string[]>(['All']);
  const [activeTopTab, setActiveTopTab] = useState<string>('All');
  const [date, setDate] = useState<string>('');
  const [selectedAuthors, setSelectedAuthors] = useState<string[]>([]);
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [from, setFrom] = useState(0);

  // Keep track of the latest active parameters to prevent asynchronous race conditions
  const activeParamsRef = useRef({ query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors });

  useEffect(() => {
    activeParamsRef.current = { query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors };
  }, [query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors]);

  const executeSearch = useCallback(async () => {
    if (!service) {
      setError('Search service is not initialized.');
      return;
    }

    const currentParams = { query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors };

    console.log('--- [DEBUG hook] Triggering executeSearch in useSearch ---', currentParams);

    setLoading(true);
    setError(null);
    try {
      const res = await service.search(query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors);
      
      // Check if parameters have changed since this request was started
      const latest = activeParamsRef.current;
      const isStale = 
        latest.query !== currentParams.query ||
        latest.pageSize !== currentParams.pageSize ||
        latest.from !== currentParams.from ||
        latest.date !== currentParams.date ||
        latest.activeTopTab !== currentParams.activeTopTab ||
        JSON.stringify(latest.selectedAuthors) !== JSON.stringify(currentParams.selectedAuthors) ||
        JSON.stringify(latest.fileTypes) !== JSON.stringify(currentParams.fileTypes);

      if (isStale) {
        console.warn('--- [DEBUG hook] Stale search results discarded for query:', query);
        return;
      }

      console.log('--- [DEBUG hook] Successful search response ---', {
        resultsLength: res.results.length,
        totalCount: res.totalCount
      });
      setResults(res.results);
      setTotalCount(res.totalCount);
    } catch (err: any) {
      console.error('--- [DEBUG hook] Error executing Graph Search in hook ---', err);
      if (err?.statusCode === 429 || (err?.message && err.message.includes('429'))) {
        setError('Too many requests. Please wait a moment before searching again.');
      } else {
        setError(err?.message || 'Failed to fetch search results from Microsoft Graph API.');
      }
    } finally {
      // Only set loading to false if this request is not stale
      const latest = activeParamsRef.current;
      const isStale = 
        latest.query !== currentParams.query ||
        latest.pageSize !== currentParams.pageSize ||
        latest.from !== currentParams.from ||
        latest.date !== currentParams.date ||
        latest.activeTopTab !== currentParams.activeTopTab ||
        JSON.stringify(latest.selectedAuthors) !== JSON.stringify(currentParams.selectedAuthors) ||
        JSON.stringify(latest.fileTypes) !== JSON.stringify(currentParams.fileTypes);

      if (!isStale) {
        setLoading(false);
      }
    }
  }, [service, query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors]);

  // Reset pagination offset to 0 whenever the query, activeTopTab, or filters change
  useEffect(() => {
    setFrom(0);
  }, [query, fileTypes, activeTopTab, date, selectedAuthors]);

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
    date,
    setDate,
    selectedAuthors,
    setSelectedAuthors,
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
