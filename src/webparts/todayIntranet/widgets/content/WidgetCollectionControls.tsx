import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetListItem, WidgetItemsView } from './IWidgetContent';
import { WidgetItems } from './WidgetItems';
import { WidgetEmpty } from './WidgetStates';
import { useWidgetPagination, WidgetPagination } from './WidgetPagination';
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
      aria-label={props.isRefreshing ? 'Syncing calendar' : 'Sync calendar'}
      title={props.isRefreshing ? 'Syncing calendar' : 'Sync calendar'}
      disabled={props.isRefreshing}
      onClick={props.onRefresh}
    >
      <Icon iconName="Sync" aria-hidden={true} />
      <span>{props.isRefreshing ? 'Syncing...' : 'Sync calendar'}</span>
    </button>
    {props.dateError && <span role="alert" className={styles.dateError}>{props.dateError}</span>}
  </div>
);

interface IFilteredItemsProps {
  items: IWidgetListItem[];
  view: WidgetItemsView;
  itemsPerPage?: number;
  ariaLabel: string;
  emptyText?: string;
  emptyState?: React.ReactNode;
}

export const WidgetFilteredItems: React.FunctionComponent<IFilteredItemsProps> = ({
  items, view, itemsPerPage, ariaLabel, emptyText, emptyState
}) => {
  const [query, setQuery] = React.useState('');
  const needle = query.trim().toLowerCase();
  const filtered = items.filter((item) =>
    [item.title, item.href, item.description, item.badge?.text, ...(item.meta ?? [])].filter(Boolean).join(' ').toLowerCase()
      .indexOf(needle) >= 0
  );
  const pagination = useWidgetPagination(filtered, itemsPerPage, needle);

  return (
    <div className={styles.collectionStack}>
      <div className={styles.searchSurface}>
        <label className={styles.searchBox}>
          <span className={styles.srOnly}>Find your link</span>
          <input
            type="search"
            placeholder="Find your link"
            value={query}
            onChange={(event) => setQuery(event.currentTarget.value)}
          />
          <Icon iconName="Search" aria-hidden={true} />
        </label>
      </div>
      <div className={styles.srOnly} role="status" aria-live="polite">
        {filtered.length === 1 ? '1 link found' : `${filtered.length} links found`}.
        {' '}Page {pagination.page} of {pagination.totalPages}.
      </div>
      {filtered.length ? (
        <WidgetItems
          items={pagination.items}
          view={view}
          ariaLabel={ariaLabel}
        />
      ) : emptyState ?? <WidgetEmpty iconName="Search" text={emptyText ?? `No links match "${query.trim()}".`} />}
      <WidgetPagination page={pagination.page} totalPages={pagination.totalPages}
        onPageChange={pagination.setPage} ariaLabel="Links" />
    </div>
  );
};
