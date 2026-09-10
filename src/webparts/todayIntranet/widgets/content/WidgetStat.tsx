import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WidgetTone } from './IWidgetContent';
import { toneClass } from './tone';
import styles from './WidgetContent.module.scss';

export interface IWidgetStatProps {
  /** The number or short string the tile exists to show. */
  value: string;
  /** Small caps line above the value, e.g. "Unread". */
  label?: string;
  /** Supporting line below the value. */
  caption?: string;
  iconName?: string;
  tone?: WidgetTone;
  /** Makes the whole block a link. */
  href?: string;
}

/** A single headline figure, centred in the tile. */
export const WidgetStat: React.FunctionComponent<IWidgetStatProps> = (props) => {
  const { value, label, caption, iconName, tone, href } = props;
  // Without a tone the value takes the tile's own text colour rather than a muted one.
  const className: string = `${styles.stat} ${tone ? toneClass(tone) : ''}`;

  const body: React.ReactElement = (
    <>
      {iconName && <Icon iconName={iconName} className={styles.statIcon} aria-hidden={true} />}
      {label && <span className={styles.statLabel}>{label}</span>}
      <span className={styles.statValue}>{value}</span>
      {caption && <span className={styles.statCaption}>{caption}</span>}
    </>
  );

  return href ? (
    <a className={className} href={href} target="_blank" rel="noreferrer">
      {body}
    </a>
  ) : (
    <div className={className}>{body}</div>
  );
};
