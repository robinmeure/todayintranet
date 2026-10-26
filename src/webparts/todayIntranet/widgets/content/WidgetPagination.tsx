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

export function useWidgetPagination<T>(items: T[], pageSize?: number, resetKey?: string): {
  items: T[]; page: number; totalPages: number; setPage(page: number): void;
} {
  const size = pageSize !== undefined && isFinite(pageSize) && pageSize > 0 ? Math.floor(pageSize) || 1 : undefined;
  const [selection, setSelection] = React.useState({ page: 1, size, resetKey });
  const totalPages = size ? Math.max(1, Math.ceil(items.length / size)) : 1;
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

export const WidgetPagination: React.FunctionComponent<{
  page: number; totalPages: number; onPageChange(page: number): void; ariaLabel: string;
}> = ({ page, totalPages, onPageChange, ariaLabel }) => {
  if (totalPages <= 1) { return null; }
  const first = Math.max(1, Math.min(page - 2, totalPages - 4));
  const pages: number[] = [];
  for (let index = first; index <= Math.min(totalPages, first + 4); index++) { pages.push(index); }
  return (
    <nav className={styles.pagination} aria-label={`${ariaLabel} pages`}>
      <button type="button" className={styles.pageButton} disabled={page === 1}
        aria-label={`Previous ${ariaLabel.toLowerCase()} page`} onClick={() => onPageChange(page - 1)}>
        <Icon iconName="ChevronLeft" aria-hidden={true} />
      </button>
      {pages.map((number) => (
        <button type="button" key={number}
          className={`${styles.pageButton} ${number === page ? styles.pageButtonActive : ''}`}
          aria-label={`${ariaLabel} page ${number}`} aria-current={number === page ? 'page' : undefined}
          onClick={() => onPageChange(number)}>{number}</button>
      ))}
      <button type="button" className={styles.pageButton} disabled={page === totalPages}
        aria-label={`Next ${ariaLabel.toLowerCase()} page`} onClick={() => onPageChange(page + 1)}>
        <Icon iconName="ChevronRight" aria-hidden={true} />
      </button>
      <span className={styles.pageStatus} role="status" aria-live="polite">Page {page} of {totalPages}</span>
    </nav>
  );
};
