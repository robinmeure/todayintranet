import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { WidgetRegistry, IWidgetCategoryGroup } from '../widgets/WidgetRegistry';
import { IWidgetDefinition } from '../widgets/IWidget';
import styles from './AddWidgetPanel.module.scss';

export interface IAddWidgetPanelProps {
  isOpen: boolean;
  /** Widget types already on the dashboard, so the catalogue can say which. */
  existingTypes: string[];
  onDismiss(): void;
  onAdd(definition: IWidgetDefinition): void;
}

/** Catalogue of everything registered in WidgetRegistry. */
export const AddWidgetPanel: React.FunctionComponent<IAddWidgetPanelProps> = (props) => {
  const { isOpen, existingTypes, onDismiss, onAdd } = props;
  const [query, setQuery] = React.useState<string>('');

  // Every visit starts from the full catalogue rather than the last search.
  React.useEffect(() => {
    if (isOpen) {
      setQuery('');
    }
  }, [isOpen]);

  const groups: IWidgetCategoryGroup[] = React.useMemo(
    () => WidgetRegistry.group(WidgetRegistry.list().filter((d) => WidgetRegistry.matches(d, query))),
    [query]
  );

  const countOf = (type: string): number => existingTypes.filter((existing) => existing === type).length;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Add a widget"
      closeButtonAriaLabel="Close"
    >
      <p className={styles.intro}>Widgets land at the bottom of your dashboard. Drag them anywhere from there.</p>

      <SearchBox
        placeholder="Search widgets"
        value={query}
        onChange={(_, next) => setQuery(next ?? '')}
        onClear={() => setQuery('')}
      />

      {groups.length === 0 && <div className={styles.noResults}>No widget matches “{query}”.</div>}

      {groups.map((group) => (
        <section key={group.category.id} className={styles.group} aria-label={group.category.name}>
          <h3 className={styles.groupTitle}>{group.category.name}</h3>
          <div className={styles.catalogue}>
            {group.widgets.map((definition) => {
              const count = countOf(definition.type);
              return (
                <button
                  key={definition.type}
                  type="button"
                  className={styles.card}
                  onClick={() => onAdd(definition)}
                  aria-label={`Add ${definition.displayName}`}
                >
                  <Icon iconName={definition.iconName} className={styles.cardIcon} aria-hidden={true} />
                  <span className={styles.cardText}>
                    <span className={styles.cardTitle}>{definition.displayName}</span>
                    <span className={styles.cardDescription}>{definition.description}</span>
                    <span className={styles.cardTags}>
                      {count > 0 && (
                        <span className={`${styles.tag} ${styles.tagAdded}`}>
                          {count === 1 ? 'On your dashboard' : `${count} on your dashboard`}
                        </span>
                      )}
                      {definition.requiredPermission && (
                        <span className={styles.tag} title="Needs tenant admin approval of this Graph permission">
                          Needs {definition.requiredPermission}
                        </span>
                      )}
                    </span>
                  </span>
                  <Icon iconName="Add" className={styles.cardAdd} aria-hidden={true} />
                </button>
              );
            })}
          </div>
        </section>
      ))}
    </Panel>
  );
};
