import { useState, useEffect, useCallback, useRef } from 'react';
import { GraphSearchService } from '../services/GraphSearchService';
import { ISearchResult } from '../../../models/ISearchResult';
import { fuzzyBrandMatch } from '../../../utils/brandDictionary';

import { IUseSearchOptions } from '../interface/IUseSearchOptions';
import { rerankResults } from '../../../utils/rerankResults';
import { SearchAnalyticsService } from '../services/SearchAnalyticsService';
export function useSearch({ service, initialQuery = '', pageSize = 20, dynamicTerms = [] }: IUseSearchOptions) {
  const [query, setQuery] = useState(initialQuery);
  const [fileTypes, setFileTypes] = useState<string[]>(['All']);
  const [activeTopTab, setActiveTopTab] = useState<string>('All');
  const [date, setDate] = useState<string>('');
  const [selectedAuthors, setSelectedAuthors] = useState<string[]>([]);
  const [selectedSites, setSelectedSites] = useState<string[]>([]);
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [from, setFrom] = useState(0);
  const [sortBy, setSortBy] = useState<string>('dateDesc');
  const [suggestedQuery, setSuggestedQuery] = useState<string | undefined>(undefined);
  const [correctedQuery, setCorrectedQuery] = useState<string | null>(null);
  const [skipCorrection, setSkipCorrection] = useState<boolean>(false);
  const abortControllerRef = useRef<AbortController | null>(null);

  // Track query transitions for auto sorting switch
  const prevQueryRef = useRef(query);
  useEffect(() => {
    const prev = prevQueryRef.current.trim();
    const current = query.trim();
    if (prev === '' && current !== '') {
      setSortBy('relevance');
    } else if (prev !== '' && current === '') {
      setSortBy('dateDesc');
    }
    prevQueryRef.current = query;
  }, [query]);

  // Keep track of the latest active parameters to prevent asynchronous race conditions
  const activeParamsRef = useRef({ query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy, skipCorrection });

  useEffect(() => {
    activeParamsRef.current = { query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy, skipCorrection };
  }, [query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy, skipCorrection]);

  const executeSearch = useCallback(async () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
    abortControllerRef.current = new AbortController();

    if (!service) {
      setError('Search service is not initialized.');
      return;
    }

    const currentParams = { query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy, skipCorrection };

    console.log('--- [DEBUG hook] Triggering executeSearch in useSearch ---', currentParams);

    setLoading(true);
    setError(null);
    try {
      let queryToSearch = query;
      if (!skipCorrection) {
        console.log('--- [DEBUG] dynamicTerms received in executeSearch ---', dynamicTerms.length, dynamicTerms.slice(0, 20));
        const brandMatch = fuzzyBrandMatch(query, dynamicTerms);
        console.log('--- [DEBUG] fuzzyBrandMatch result ---', brandMatch);
        if (brandMatch) {
          console.log('--- [DEBUG] Silently using corrected brand query:', brandMatch);
          queryToSearch = brandMatch;
        }
      }

      const res = await service.search(queryToSearch, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy);
      
      // Check if parameters have changed since this request was started
      const latest = activeParamsRef.current;
      const isStale = 
        latest.query !== currentParams.query ||
        latest.pageSize !== currentParams.pageSize ||
        latest.from !== currentParams.from ||
        latest.date !== currentParams.date ||
        latest.sortBy !== currentParams.sortBy ||
        latest.activeTopTab !== currentParams.activeTopTab ||
        JSON.stringify(latest.selectedAuthors) !== JSON.stringify(currentParams.selectedAuthors) ||
        JSON.stringify(latest.selectedSites) !== JSON.stringify(currentParams.selectedSites) ||
        JSON.stringify(latest.fileTypes) !== JSON.stringify(currentParams.fileTypes);

      if (isStale) {
        console.warn('--- [DEBUG hook] Stale search results discarded for query:', query);
        return;
      }

      console.log('--- [DEBUG hook] Successful search response ---', {
        resultsLength: res.results.length,
        totalCount: res.totalCount
      });
      const reranked = rerankResults(res.results, SearchAnalyticsService.getClickCount.bind(SearchAnalyticsService));
      setResults(reranked);

      // Artificially cap results to 200 for completely blank wildcard searches to reduce pagination overload
      const isBlankSearch = queryToSearch.trim() === '' &&
        fileTypes.length === 1 && fileTypes[0] === 'All' &&
        selectedAuthors.length === 0 &&
        date === '';
      
      const cappedTotalCount = isBlankSearch ? Math.min(res.totalCount, 200) : res.totalCount;
      setTotalCount(cappedTotalCount);
      setSuggestedQuery(res.suggestedQuery);

      if (skipCorrection) {
        setCorrectedQuery(null);
      } else {
        setCorrectedQuery(res.suggestedQuery || null);
      }
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
        latest.sortBy !== currentParams.sortBy ||
        latest.activeTopTab !== currentParams.activeTopTab ||
        latest.skipCorrection !== currentParams.skipCorrection ||
        JSON.stringify(latest.selectedAuthors) !== JSON.stringify(currentParams.selectedAuthors) ||
        JSON.stringify(latest.selectedSites) !== JSON.stringify(currentParams.selectedSites) ||
        JSON.stringify(latest.fileTypes) !== JSON.stringify(currentParams.fileTypes);

      if (!isStale) {
        setLoading(false);
      }
    }
  }, [service, query, pageSize, from, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy, skipCorrection, dynamicTerms]);

  // Reset pagination offset to 0 whenever the query, activeTopTab, sortBy or filters change
  useEffect(() => {
    setFrom(0);
  }, [query, fileTypes, activeTopTab, date, selectedAuthors, selectedSites, sortBy]);

  useEffect(() => {
    setSkipCorrection(false);
    setCorrectedQuery(null);
  }, [query]);

  useEffect(() => {
    if (service) {
      executeSearch().catch(() => undefined);
    }
  }, [executeSearch, service]);

  // Re-run search when dynamicTerms updates (async load) so corrections apply to current query
  const prevDynamicTermsLenRef = useRef(0);
  useEffect(() => {
    const prevLen = prevDynamicTermsLenRef.current;
    prevDynamicTermsLenRef.current = dynamicTerms.length;
    if (dynamicTerms.length > prevLen && query.trim().length > 0 && service) {
      console.log('--- [DEBUG] dynamicTerms updated, re-running search ---', query, 'terms:', dynamicTerms.length);
      executeSearch().catch(() => undefined);
    }
  }, [dynamicTerms, query, service, executeSearch]);

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
    selectedSites,
    setSelectedSites,
    results,
    totalCount,
    loading,
    error,
    from,
    setFrom,
    sortBy,
    setSortBy,
    suggestedQuery,
    setSuggestedQuery,
    correctedQuery,
    setCorrectedQuery,
    skipCorrection,
    setSkipCorrection,
    pageSize,
    refresh: executeSearch
  };
}
