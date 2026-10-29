import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetContext } from '../IWidget';
import { booleanSetting, INumberSettingBounds, NumberSetting, numberSetting, ToggleSetting } from './WidgetSettings';
import styles from './WidgetCollections.module.scss';

const PAGE_SIZE_BOUNDS = { min: 1, max: 100 };

export function paginationSize(context: IWidgetContext, fallback = 6, enabled = false,
  bounds: INumberSettingBounds = PAGE_SIZE_BOUNDS): number | undefined {
  return booleanSetting(context, 'paginationEnabled', enabled)
    ? numberSetting(context, 'itemsPerPage', fallback, bounds) : undefined;
}

/** Pages of results a collection requests ahead of the one on screen. */
export const PAGE_WINDOW = 4;
const MAX_FETCH_LIMIT = 100;

export interface ICollectionPaging {
  /** Items per page; undefined when pagination is off. */
  pageSize?: number;
  /** How many items to request from Graph or search so every page can be shown. */
  fetchLimit: number;
}

/**
 * The widget's configured count is its page size, and pagination is on unless switched off.
 * The remote query asks for a bounded window of pages instead of the whole collection;
 * the page control appears only when more than one page comes back. A page size saved
 * explicitly by an earlier release (`itemsPerPage`) is still honoured.
 */
export function collectionPaging(context: IWidgetContext, displayCount: number,
  maxFetch = MAX_FETCH_LIMIT): ICollectionPaging {
  if (!booleanSetting(context, 'paginationEnabled', true)) {
    return { fetchLimit: displayCount };
  }
  const pageSize = context.settings.paginationEnabled === true
    ? numberSetting(context, 'itemsPerPage', displayCount, PAGE_SIZE_BOUNDS) : displayCount;
  return { pageSize, fetchLimit: Math.min(maxFetch, Math.max(displayCount, pageSize * PAGE_WINDOW)) };
}

export const PaginationSettings: React.FunctionComponent<{
  context: IWidgetContext; fallback?: number; enabled?: boolean; bounds?: INumberSettingBounds; label?: string;
}> = ({ context, fallback = 6, enabled = false, bounds = PAGE_SIZE_BOUNDS, label = 'Items per page' }) => (
  <>
    <ToggleSetting context={context} settingKey="paginationEnabled" label="Use pagination" fallback={enabled} />
    {booleanSetting(context, 'paginationEnabled', enabled) && (
      <NumberSetting context={context} settingKey="itemsPerPage" label={label} fallback={fallback} {...bounds} />
    )}
  </>
);

export function useWidgetPagination<T>(items: T[], pageSize?: number, resetKey?: string, additionalPages = 0): {
  items: T[]; page: number; totalPages: number; setPage(page: number): void;
} {
  const size = pageSize !== undefined && isFinite(pageSize) && pageSize > 0 ? Math.floor(pageSize) || 1 : undefined;
  const [selection, setSelection] = React.useState({ page: 1, size, resetKey });
  const totalPages = size ? Math.max(1, Math.ceil(items.length / size)) + additionalPages : 1;
  const page = selection.size === size && selection.resetKey === resetKey
    ? Math.min(selection.page, totalPages) : 1;
  // Persist clamping so a later refresh cannot unexpectedly return to an old page.
  if (selection.page !== page || selection.size !== size || selection.resetKey !== resetKey) {
    setSelection({ page, size, resetKey });
  }
  return {
    items: size ? items.slice((page - 1) * size, page * size) : items,
    page, totalPages,
    setPage: (next) => setSelection({ page: Math.max(1, Math.min(next, totalPages)), size, resetKey })
  };
}

/** Page numbers shown at once; the chevrons move the window through longer collections. */
export const VISIBLE_PAGE_COUNT = 4;

export const WidgetPagination: React.FunctionComponent<{
  page: number; totalPages: number; onPageChange(page: number): void; ariaLabel: string; disabled?: boolean;
}> = ({ page, totalPages, onPageChange, ariaLabel, disabled = false }) => {
  if (totalPages <= 1) { return null; }
  const first = Math.max(1, Math.min(page - 1, totalPages - VISIBLE_PAGE_COUNT + 1));
  const pages: number[] = [];
  for (let index = first; index <= Math.min(totalPages, first + VISIBLE_PAGE_COUNT - 1); index++) { pages.push(index); }
  return (
    <nav className={styles.pagination} aria-label={`${ariaLabel} pages`}>
      <button type="button" className={styles.pageButton} disabled={disabled || page === 1}
        aria-label={`Previous ${ariaLabel.toLowerCase()} page`} onClick={() => onPageChange(page - 1)}>
        <Icon iconName="ChevronLeft" aria-hidden={true} />
      </button>
      {pages.map((number) => (
        <button type="button" key={number} disabled={disabled}
          className={`${styles.pageButton} ${number === page ? styles.pageButtonActive : ''}`}
          aria-label={`${ariaLabel} page ${number}`} aria-current={number === page ? 'page' : undefined}
          onClick={() => onPageChange(number)}>{number}</button>
      ))}
      <button type="button" className={styles.pageButton} disabled={disabled || page === totalPages}
        aria-label={`Next ${ariaLabel.toLowerCase()} page`} onClick={() => onPageChange(page + 1)}>
        <Icon iconName="ChevronRight" aria-hidden={true} />
      </button>
      <span className={styles.srOnly} role="status" aria-live="polite">Page {page} of {totalPages}</span>
    </nav>
  );
};
