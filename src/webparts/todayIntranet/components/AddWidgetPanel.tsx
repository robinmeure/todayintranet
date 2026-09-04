import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { WidgetRegistry } from '../widgets/WidgetRegistry';
import { IWidgetDefinition } from '../widgets/IWidget';
import styles from './AddWidgetPanel.module.scss';

export interface IAddWidgetPanelProps {
  isOpen: boolean;
  onDismiss(): void;
  onAdd(definition: IWidgetDefinition): void;
}

/** Catalogue of everything registered in WidgetRegistry. */
export const AddWidgetPanel: React.FunctionComponent<IAddWidgetPanelProps> = (props) => {
  const { isOpen, onDismiss, onAdd } = props;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Add a widget"
      closeButtonAriaLabel="Close"
    >
      <div className={styles.catalogue}>
        {WidgetRegistry.list().map((definition) => (
          <div key={definition.type} className={styles.card}>
            <Icon iconName={definition.iconName} className={styles.cardIcon} />
            <div className={styles.cardText}>
              <div className={styles.cardTitle}>{definition.displayName}</div>
              <div className={styles.cardDescription}>{definition.description}</div>
            </div>
            <DefaultButton text="Add" onClick={() => onAdd(definition)} />
          </div>
        ))}
      </div>
    </Panel>
  );
};
