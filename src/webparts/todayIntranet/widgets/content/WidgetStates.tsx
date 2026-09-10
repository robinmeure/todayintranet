import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Link } from '@fluentui/react/lib/Link';
import { Icon } from '@fluentui/react/lib/Icon';
import { ActionButton } from '@fluentui/react/lib/Button';
import { IWidgetEmptyState, IWidgetError } from './IWidgetContent';
import styles from './WidgetContent.module.scss';

export interface IWidgetLoadingProps {
  /** Announced to screen readers while the skeleton is on screen. */
  label?: string;
  /** Placeholder rows to draw. Match it to what the widget usually shows. */
  rows?: number;
}

/**
 * Skeleton placeholder. A spinner tells the user to wait; a skeleton tells them what
 * is coming, so the tile does not visibly jump when the data lands.
 */
export const WidgetLoading: React.FunctionComponent<IWidgetLoadingProps> = ({ label, rows }) => {
  const count: number = Math.max(1, rows ?? 3);
  const placeholders: number[] = [];
  for (let i: number = 0; i < count; i++) {
    placeholders.push(i);
  }

  return (
    <div className={styles.skeleton} role="status" aria-busy={true} aria-label={label ?? 'Loading'}>
      {placeholders.map((index) => (
        <div key={index} className={styles.skeletonRow}>
          <div className={`${styles.skeletonBlock} ${styles.skeletonIcon}`} />
          <div className={styles.skeletonLines}>
            <div className={`${styles.skeletonBlock} ${styles.skeletonLine}`} />
            <div className={`${styles.skeletonBlock} ${styles.skeletonLine} ${styles.skeletonLineShort}`} />
          </div>
        </div>
      ))}
    </div>
  );
};

/** Nothing to show, said in a way that does not read like a failure. */
export const WidgetEmpty: React.FunctionComponent<IWidgetEmptyState> = (props) => {
  const { iconName, text, actionText, actionHref, onAction } = props;

  return (
    <div className={styles.centered}>
      <Icon iconName={iconName} className={styles.emptyIcon} aria-hidden={true} />
      <div className={styles.emptyText}>{text}</div>
      {actionText && (actionHref || onAction) && (
        <ActionButton
          href={actionHref}
          target={actionHref ? '_blank' : undefined}
          rel={actionHref ? 'noreferrer' : undefined}
          onClick={onAction}
        >
          {actionText}
        </ActionButton>
      )}
    </div>
  );
};

export interface IWidgetErrorProps {
  error: IWidgetError;
  onRetry?: () => void;
}

/** A contained failure: the tile explains itself, the rest of the dashboard carries on. */
export const WidgetErrorMessage: React.FunctionComponent<IWidgetErrorProps> = ({ error, onRetry }) => (
  <MessageBar
    messageBarType={error.isActionRequired ? MessageBarType.warning : MessageBarType.error}
    isMultiline={true}
  >
    {error.message}
    {onRetry && !error.isActionRequired && (
      <>
        {' '}
        <Link onClick={onRetry}>Try again</Link>
      </>
    )}
  </MessageBar>
);
