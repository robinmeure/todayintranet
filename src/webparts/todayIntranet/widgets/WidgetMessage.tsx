import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Link } from '@fluentui/react/lib/Link';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './WidgetMessage.module.scss';
import { IGraphError } from './graph/useGraphData';

export const WidgetLoading: React.FunctionComponent<{ label?: string }> = ({ label }) => (
  <div className={styles.centered}>
    <Spinner size={SpinnerSize.medium} label={label} />
  </div>
);

export const WidgetEmpty: React.FunctionComponent<{ iconName: string; text: string }> = ({ iconName, text }) => (
  <div className={styles.centered}>
    <Icon iconName={iconName} className={styles.emptyIcon} />
    <div className={styles.emptyText}>{text}</div>
  </div>
);

export const WidgetError: React.FunctionComponent<{ error: IGraphError; onRetry?: () => void }> = ({
  error,
  onRetry
}) => (
  <MessageBar
    messageBarType={error.needsConsent ? MessageBarType.warning : MessageBarType.error}
    isMultiline={true}
  >
    {error.message}
    {onRetry && !error.needsConsent && (
      <>
        {' '}
        <Link onClick={onRetry}>Try again</Link>
      </>
    )}
  </MessageBar>
);
