import * as React from 'react';
import { IWidgetDataState, IWidgetEmptyState } from './IWidgetContent';
import { WidgetEmpty, WidgetErrorMessage, WidgetLoading, IWidgetLoadingProps } from './WidgetStates';
import styles from './WidgetContent.module.scss';

export interface IWidgetViewProps<T> {
  state: IWidgetDataState<T>;
  loading?: IWidgetLoadingProps;
  empty: IWidgetEmptyState;
  /** Overrides the default test, which treats an empty array as nothing to show. */
  isEmpty?(data: T): boolean;
  /** Called only when there is data worth rendering. */
  children(data: T): React.ReactElement;
}

/**
 * Turns one data state into the right thing on screen: skeleton, error, empty state
 * or content. Every widget that loads something goes through here, so the four
 * states look and behave the same everywhere and no widget has to re-invent them.
 */
export function WidgetView<T>(props: IWidgetViewProps<T>): React.ReactElement {
  const { state, loading, empty, isEmpty, children } = props;

  if (state.status === 'loading') {
    return <WidgetLoading {...loading} />;
  }

  if (state.status === 'error' && state.error) {
    return <WidgetErrorMessage error={state.error} onRetry={state.reload} />;
  }

  const data: T | undefined = state.data;
  if (data === undefined) {
    return <WidgetEmpty {...empty} />;
  }

  const hasNothing: boolean = isEmpty ? isEmpty(data) : Array.isArray(data) && data.length === 0;
  if (hasNothing) {
    return <WidgetEmpty {...empty} />;
  }

  const content: React.ReactElement = children(data);
  if (!state.isRefreshing) {
    return content;
  }

  // Re-reading data the user is already looking at: keep it, and say it is happening
  // rather than replacing a full tile with a skeleton again.
  return (
    <div className={styles.refreshing} aria-busy={true}>
      <div className={styles.refreshingBar} aria-hidden={true} />
      {content}
    </div>
  );
}
