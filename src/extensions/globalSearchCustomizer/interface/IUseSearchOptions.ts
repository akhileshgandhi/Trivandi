/* eslint-disable @rushstack/no-new-null */
import { GraphSearchService } from '../services/GraphSearchService';

export interface IUseSearchOptions {
  service: GraphSearchService | null;
  initialQuery?: string;
  pageSize?: number;
  dynamicTerms?: string[];
}
