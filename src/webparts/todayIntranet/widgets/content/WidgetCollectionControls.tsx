import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetListItem, WidgetItemsView } from './IWidgetContent';
import { WidgetItems } from './WidgetItems';
import { WidgetEmpty } from './WidgetStates';
import styles from './WidgetCollections.module.scss';

interface IDateToolbarProps {
  ranges: { key: string; text: string }[];
  activeRange: string;
  date: string;
  dateError?: string;
  isRefreshing?: boolean;
  onRangeChange(key: string): void;
  onDateChange(date: string): void;
  onRefresh(): void;
}

export const WidgetDateToolbar: React.FunctionComponent<IDateToolbarProps> = (props) => (
  <div className={styles.toolbar}>
    <div className={styles.rangeTabs} role="group" aria-label="Calendar date range">
      {props.ranges.map((range) => (
        <button
          type="button"
          key={range.key}
          className={`${styles.rangeTab} ${props.activeRange === range.key ? styles.rangeTabActive : ''}`}
          aria-pressed={props.activeRange === range.key}
          onClick={() => props.onRangeChange(range.key)}
        >
          {range.text}
        </button>
      ))}
    </div>
    <label className={styles.datePicker}>
      <span className={styles.srOnly}>Choose a calendar date</span>
      <input
        type="date"
        value={props.date}
        required={true}
        aria-invalid={!!props.dateError}
        onChange={(event) => props.onDateChange(event.currentTarget.value)}
      />
    </label>
    <button
      type="button"
      className={styles.syncButton}
      disabled={props.isRefreshing}
      onClick={props.onRefresh}
    >
      <Icon iconName="Sync" aria-hidden={true} />
      <span>{props.isRefreshing ? 'Syncing...' : 'Sync calendar'}</span>
    </button>
    {props.dateError && <span role="alert">{props.dateError}</span>}
  </div>
);

interface IFilteredItemsProps {
  items: IWidgetListItem[];
  view: WidgetItemsView;
  itemsPerPage: number;
  ariaLabel: string;
}

export const WidgetFilteredItems: React.FunctionComponent<IFilteredItemsProps> = ({
  items, view, itemsPerPage, ariaLabel
}) => {
  const [query, setQuery] = React.useState('');
  const [page, setPage] = React.useState(1);
  const needle = query.trim().toLowerCase();
  const filtered = items.filter((item) =>
    [item.title, item.description, item.badge?.text, ...(item.meta ?? [])].filter(Boolean).join(' ').toLowerCase()
      .indexOf(needle) >= 0
  );
  const totalPages = Math.max(1, Math.ceil(filtered.length / itemsPerPage));
  const currentPage = Math.min(page, totalPages);
  const firstPage = Math.max(1, Math.min(currentPage - 2, totalPages - 4));
  const pages: number[] = [];
  for (let index = firstPage; index <= Math.min(totalPages, firstPage + 4); index++) {
    pages.push(index);
  }

  return (
    <div className={styles.collectionStack}>
      <div className={styles.searchSurface}>
        <label className={styles.searchBox}>
          <span className={styles.srOnly}>Find your link</span>
          <input
            type="search"
            placeholder="Find your link"
            value={query}
            onChange={(event) => { setQuery(event.currentTarget.value); setPage(1); }}
          />
          <Icon iconName="Search" aria-hidden={true} />
        </label>
      </div>
      <div className={styles.srOnly} role="status" aria-live="polite">
        {filtered.length === 1 ? '1 link found' : `${filtered.length} links found`}. Page {currentPage} of {totalPages}.
      </div>
      {filtered.length ? (
        <WidgetItems
          items={filtered.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage)}
          view={view}
          ariaLabel={ariaLabel}
        />
      ) : <WidgetEmpty iconName="Search" text={`No links match "${query.trim()}".`} />}
      {totalPages > 1 && (
        <nav className={styles.pagination} aria-label="Links pages">
          <button type="button" className={styles.pageButton} disabled={currentPage === 1}
            aria-label="Previous links page" onClick={() => setPage(currentPage - 1)}>
            <Icon iconName="ChevronLeft" aria-hidden={true} />
          </button>
          {pages.map((pageNumber) => (
            <button type="button" key={pageNumber}
              className={`${styles.pageButton} ${pageNumber === currentPage ? styles.pageButtonActive : ''}`}
              aria-label={`Links page ${pageNumber}`}
              aria-current={pageNumber === currentPage ? 'page' : undefined}
              onClick={() => setPage(pageNumber)}>
              {pageNumber}
            </button>
          ))}
          <button type="button" className={styles.pageButton} disabled={currentPage === totalPages}
            aria-label="Next links page" onClick={() => setPage(currentPage + 1)}>
            <Icon iconName="ChevronRight" aria-hidden={true} />
          </button>
        </nav>
      )}
    </div>
  );
};
