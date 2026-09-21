import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetListItem } from './IWidgetContent';
import { toneClass } from './tone';
import styles from './WidgetCollections.module.scss';

interface ICollectionProps {
  items: IWidgetListItem[];
  ariaLabel?: string;
}

export const WidgetAgenda: React.FunctionComponent<ICollectionProps> = ({ items, ariaLabel }) => (
  <div className={styles.collection}>
    <ol className={styles.agenda} role="list" aria-label={ariaLabel}>
      {items.map((item, index) => (
        <li className={styles.event} key={item.key ?? index}>
          <div className={`${styles.dateTile} ${toneClass(item.leadingLabel?.tone, 'neutral')}`}>
            <span className={styles.datePrimary}>{item.leadingLabel?.text ?? item.badge?.text ?? ''}</span>
            {item.leadingLabel?.description && (
              <span className={styles.dateSecondary}>{item.leadingLabel.description}</span>
            )}
          </div>
          <div className={styles.eventDetails}>
            {item.href ? (
              <a className={styles.eventTitle} href={item.href} target="_blank" rel="noreferrer" title={item.title}>
                {item.title}
                <Icon iconName="OpenInNewWindow" className={styles.titleIcon} aria-hidden={true} />
              </a>
            ) : <span className={styles.eventTitle}>{item.title}</span>}
            {(item.meta ?? []).filter((text) => !!text).map((text, metaIndex) => (
              <span className={styles.eventMeta} key={metaIndex} title={text}>{text}</span>
            ))}
            {item.description && <span className={styles.eventMeta} title={item.description}>{item.description}</span>}
          </div>
          <div className={styles.eventActions}>
            {(item.actions ?? []).map((action) => (
              <a
                key={action.ariaLabel}
                className={`${styles.actionButton} ${action.isPrimary ? styles.actionPrimary : ''}`}
                href={action.href}
                target="_blank"
                rel="noreferrer"
                aria-label={action.ariaLabel}
              >
                {action.iconName && <Icon iconName={action.iconName} aria-hidden={true} />}
                <span>{action.text}</span>
              </a>
            ))}
          </div>
        </li>
      ))}
    </ol>
  </div>
);

const LinkTile: React.FunctionComponent<{ item: IWidgetListItem }> = ({ item }) => {
  const body = (
    <>
      <span className={`${styles.linkIcon} ${toneClass(item.tone)}`}>
        <Icon iconName={item.iconName || 'Link'} aria-hidden={true} />
      </span>
      <span className={styles.linkContent}>
        {item.badge && (
          <span className={`${styles.badge} ${toneClass(item.badge.tone)}`}>
            {item.badge.iconName && <Icon iconName={item.badge.iconName} aria-hidden={true} />}
            {item.badge.text}
          </span>
        )}
        <span className={styles.linkTitle} title={item.title}>{item.title}</span>
        {item.description && <span className={styles.linkDescription} title={item.description}>{item.description}</span>}
      </span>
      {(item.href || item.onClick) && <Icon iconName="ChevronRight" className={styles.chevron} aria-hidden={true} />}
    </>
  );
  if (item.href) {
    return <a className={styles.linkCard} href={item.href} target="_blank" rel="noreferrer">{body}</a>;
  }
  if (item.onClick) {
    return <button className={styles.linkCard} type="button" onClick={item.onClick}>{body}</button>;
  }
  return <div className={styles.linkCard}>{body}</div>;
};

export const WidgetLinkTiles: React.FunctionComponent<ICollectionProps> = ({ items, ariaLabel }) => (
  <div className={styles.collection}>
    <ul className={styles.linksGrid} role="list" aria-label={ariaLabel}>
      {items.map((item, index) => <li key={item.key ?? index}><LinkTile item={item} /></li>)}
    </ul>
  </div>
);
