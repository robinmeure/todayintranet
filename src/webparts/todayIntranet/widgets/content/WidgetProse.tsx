import * as React from 'react';
import styles from './WidgetContent.module.scss';

export interface IWidgetProseProps {
  /** Optional heading, rendered at the tile's body text scale. */
  heading?: string;
  children?: React.ReactNode;
}

/** Text content — a greeting, an announcement, a short piece of guidance. */
export const WidgetProse: React.FunctionComponent<IWidgetProseProps> = ({ heading, children }) => (
  <div className={styles.prose}>
    {heading && <div className={styles.proseHeading}>{heading}</div>}
    {children && <div className={styles.proseBody}>{children}</div>}
  </div>
);
