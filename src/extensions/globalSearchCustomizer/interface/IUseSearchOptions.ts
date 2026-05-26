import { GraphSearchService } from '../services/GraphSearchService';

export interface IUseSearchOptions {
  service: GraphSearchService | null;
  initialQuery?: string;
  pageSize?: number;
}
