import * as React from 'react';
import ReactPaginate from 'react-paginate';
import { ChevronLeft, ChevronRight } from 'lucide-react';
import styles from './Pagination.module.scss';

import { IPaginationComponentProps } from '../interface/IPaginationComponentProps';

export const PaginationComponent: React.FC<IPaginationComponentProps> = ({
  totalCount,
  pageSize,
  from,
  onPageChange
}) => {
  const pageCount = Math.ceil(totalCount / pageSize);
  const currentPage = Math.floor(from / pageSize);

  const handlePageClick = (event: { selected: number }) => {
    onPageChange(event.selected * pageSize);
  };

  if (pageCount <= 1) return null;

  return (
    <div className={styles.paginationContainer}>
      <ReactPaginate
        breakLabel="..."
        nextLabel={<ChevronRight size={16} strokeWidth={2.5} />}
        onPageChange={handlePageClick}
        pageRangeDisplayed={3}
        marginPagesDisplayed={1}
        pageCount={pageCount}
        previousLabel={<ChevronLeft size={16} strokeWidth={2.5} />}
        renderOnZeroPageCount={null}
        forcePage={currentPage}
        containerClassName={styles.pagination}
        pageClassName={styles.pageItem}
        pageLinkClassName={styles.pageLink}
        previousClassName={styles.pageItem}
        previousLinkClassName={styles.pageLink}
        nextClassName={styles.pageItem}
        nextLinkClassName={styles.pageLink}
        breakClassName={styles.pageItem}
        breakLinkClassName={styles.pageLink}
        activeClassName={styles.activePage}
        disabledClassName={styles.disabledPage}
      />
    </div>
  );
};
