import { useDebounce as useDebounceLib } from 'use-debounce';

export function useDebounce<T>(value: T, delay: number): [T, { isPending: () => boolean; flush: () => void; cancel: () => void }] {
  return useDebounceLib(value, delay);
}
