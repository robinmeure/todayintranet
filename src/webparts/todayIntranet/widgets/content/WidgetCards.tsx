import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetListItem } from './IWidgetContent';
import { toneClass } from './tone';
import styles from './WidgetContent.module.scss';

export interface IWidgetCardsProps {
  items: IWidgetListItem[];
  /** `stack` puts one card per row, `grid` flows them side by side. */
  layout?: 'stack' | 'grid';
  ariaLabel?: string;
}

const WidgetCard: React.FunctionComponent<{ item: IWidgetListItem }> = ({ item }) => {
  const isInteractive: boolean = !!item.href || !!item.onClick;
  const className: string = [styles.card, isInteractive ? styles.cardInteractive : '', toneClass(item.tone)]
    .filter((part) => !!part)
    .join(' ');

  const meta: string = (item.meta ?? []).filter((part) => !!part).join(' · ');

  const body: React.ReactElement = (
    <>
      <span className={styles.cardTop}>
        {item.iconName && <Icon iconName={item.iconName} className={styles.cardIcon} aria-hidden={true} />}
        {item.badge && (
          <span className={`${styles.cardBadge} ${item.badge.tone ? toneClass(item.badge.tone) : ''}`}>
            {item.badge.text}
          </span>
        )}
      </span>
      <span className={`${styles.cardTitle} ${item.isEmphasized ? styles.cardTitleEmphasized : ''}`}>
        {item.title}
      </span>
      {meta && <span className={styles.cardMeta}>{meta}</span>}
      {item.description && <span className={styles.cardDescription}>{item.description}</span>}
    </>
  );

  if (item.href) {
    return (
      <a className={className} href={item.href} target="_blank" rel="noreferrer">
        {body}
      </a>
    );
  }
  if (item.onClick) {
    return (
      <button type="button" className={className} onClick={item.onClick}>
        {body}
      </button>
    );
  }
  return <div className={className}>{body}</div>;
};

/** The same items as `WidgetList`, drawn as separate card surfaces. */
export const WidgetCards: React.FunctionComponent<IWidgetCardsProps> = ({ items, layout, ariaLabel }) => (
  <ul
    className={`${styles.cards} ${layout === 'grid' ? styles.cardsGrid : ''}`}
    role="list"
    aria-label={ariaLabel}
  >
    {items.map((item, index) => (
      <li key={item.key ?? index} className={styles.cardsItem}>
        <WidgetCard item={item} />
      </li>
    ))}
  </ul>
);
