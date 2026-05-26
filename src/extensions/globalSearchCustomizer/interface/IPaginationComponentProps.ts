export interface IPaginationComponentProps {
  totalCount: number;
  pageSize: number;
  from: number;
  onPageChange: (newFrom: number) => void;
}
