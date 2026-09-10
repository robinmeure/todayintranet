import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetListItem } from './IWidgetContent';
import { toneClass } from './tone';
import styles from './WidgetContent.module.scss';

export interface IWidgetListProps {
  items: IWidgetListItem[];
  /** `compact` tightens the rows, for tiles that are short or hold many items. */
  density?: 'comfortable' | 'compact';
  /** Names the list for screen readers, e.g. "Upcoming events". */
  ariaLabel?: string;
}

const WidgetListRow: React.FunctionComponent<{
  item: IWidgetListItem;
  density?: 'comfortable' | 'compact';
}> = ({ item, density }) => {
  const isInteractive: boolean = !!item.href || !!item.onClick;
  const className: string = [
    styles.item,
    density === 'compact' ? styles.itemCompact : '',
    isInteractive ? styles.itemInteractive : '',
    toneClass(item.tone)
  ]
    .filter((part) => !!part)
    .join(' ');

  const meta: string = (item.meta ?? []).filter((part) => !!part).join(' · ');

  const body: React.ReactElement = (
    <>
      {item.hasAccentBar && <span className={styles.itemAccent} aria-hidden={true} />}
      {item.iconName && <Icon iconName={item.iconName} className={styles.itemIcon} aria-hidden={true} />}
      <span className={styles.itemText}>
        <span className={`${styles.itemTitle} ${item.isEmphasized ? styles.itemTitleEmphasized : ''}`}>
          {item.title}
        </span>
        {meta && <span className={styles.itemMeta}>{meta}</span>}
        {item.description && <span className={styles.itemDescription}>{item.description}</span>}
      </span>
      {item.badge && (
        <span className={`${styles.itemBadge} ${item.badge.tone ? toneClass(item.badge.tone) : ''}`}>
          {item.badge.text}
        </span>
      )}
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

/**
 * The one list every widget uses. Feeding it `IWidgetListItem`s is what keeps a
 * calendar tile, a mail tile and anything added later looking like one product.
 */
export const WidgetList: React.FunctionComponent<IWidgetListProps> = ({ items, density, ariaLabel }) => (
  // Removing the bullets also removes list semantics in some browsers, hence role.
  <ul className={styles.list} role="list" aria-label={ariaLabel}>
    {items.map((item, index) => (
      <li key={item.key ?? index}>
        <WidgetListRow item={item} density={density} />
      </li>
    ))}
  </ul>
);
