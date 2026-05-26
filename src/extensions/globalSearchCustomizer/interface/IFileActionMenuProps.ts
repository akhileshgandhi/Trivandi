import { ISearchResult } from '../../../models/ISearchResult';

export interface IFileActionMenuProps {
  file: ISearchResult;
  onOpenChange?: (isOpen: boolean) => void;
}
