import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import styles from './WidgetFrame.module.scss';
import { IWidgetInstance } from '../model/IDashboardLayout';
import { IWidgetContext } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

export interface IWidgetFrameProps {
  instance: IWidgetInstance;
  widgetContext: IWidgetContext;
  isEditing: boolean;
  onRemove(instanceId: string): void;
}

/** Shared chrome (title bar, drag handle, remove button) around every widget. */
export const WidgetFrame: React.FunctionComponent<IWidgetFrameProps> = (props) => {
  const { instance, widgetContext, isEditing, onRemove } = props;
  const definition = WidgetRegistry.get(instance.type);

  return (
    <div className={`${styles.frame} ${isEditing ? styles.frameEditing : ''}`}>
      <div className={`${styles.header} ${isEditing ? `${styles.dragHandle} widget-drag-handle` : ''}`}>
        <Icon iconName={definition ? definition.iconName : 'Unknown'} className={styles.icon} />
        <div className={styles.title}>{definition ? definition.displayName : instance.type}</div>
        {isEditing && (
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title={`Remove ${definition ? definition.displayName : instance.type}`}
            ariaLabel={`Remove ${definition ? definition.displayName : instance.type}`}
            onClick={() => onRemove(instance.id)}
            onMouseDown={(e) => e.stopPropagation()}
          />
        )}
      </div>
      <div className={styles.body}>
        {definition ? (
          definition.render(widgetContext)
        ) : (
          <div className={styles.missing}>
            This widget (<code>{instance.type}</code>) is no longer available. Remove it in edit mode.
          </div>
        )}
      </div>
    </div>
  );
};
