import { ISearchResult } from '../../../models/ISearchResult';

export interface IResultCardProps {
  result: ISearchResult;
  idx: number;
  onClick: (id: string) => void;
}
