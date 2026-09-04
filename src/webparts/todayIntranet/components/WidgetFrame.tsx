import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import styles from './WidgetFrame.module.scss';
import { WidgetErrorBoundary } from './WidgetErrorBoundary';
import { IWidgetInstance } from '../model/IDashboardLayout';
import { IWidgetContext } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

/** A keyboard driven move or resize, expressed in grid units. */
export interface INudge {
  dx?: number;
  dy?: number;
  dw?: number;
  dh?: number;
}

export interface IWidgetFrameProps {
  instance: IWidgetInstance;
  widgetContext: IWidgetContext;
  isEditing: boolean;
  onRemove(instanceId: string): void;
  onNudge(instanceId: string, nudge: INudge): void;
}

const ARROW_MOVES: Record<string, INudge> = {
  ArrowLeft: { dx: -1 },
  ArrowRight: { dx: 1 },
  ArrowUp: { dy: -1 },
  ArrowDown: { dy: 1 }
};

const ARROW_RESIZES: Record<string, INudge> = {
  ArrowLeft: { dw: -1 },
  ArrowRight: { dw: 1 },
  ArrowUp: { dh: -1 },
  ArrowDown: { dh: 1 }
};

/** Shared chrome (title bar, drag handle, settings, remove) around every widget. */
export const WidgetFrame: React.FunctionComponent<IWidgetFrameProps> = (props) => {
  const { instance, widgetContext, isEditing, onRemove, onNudge } = props;
  const definition = WidgetRegistry.get(instance.type);
  const name = definition ? definition.displayName : instance.type;

  const [isSettingsOpen, setIsSettingsOpen] = React.useState<boolean>(false);
  const settingsButtonId = `widget-settings-${instance.id}`;

  React.useEffect(() => {
    if (!isEditing) {
      setIsSettingsOpen(false);
    }
  }, [isEditing]);

  const handleKeyDown = (event: React.KeyboardEvent<HTMLDivElement>): void => {
    // Ignore keys aimed at the buttons inside the header.
    if (!isEditing || event.target !== event.currentTarget) {
      return;
    }
    const nudge = event.shiftKey ? ARROW_RESIZES[event.key] : ARROW_MOVES[event.key];
    if (!nudge) {
      return;
    }
    event.preventDefault();
    onNudge(instance.id, nudge);
  };

  return (
    <div className={`${styles.frame} ${isEditing ? styles.frameEditing : ''}`}>
      <div
        className={`${styles.header} ${isEditing ? `${styles.dragHandle} widget-drag-handle` : ''}`}
        tabIndex={isEditing ? 0 : undefined}
        role={isEditing ? 'group' : undefined}
        aria-label={
          isEditing
            ? `${name}. Use the arrow keys to move this widget, and shift with the arrow keys to resize it.`
            : undefined
        }
        onKeyDown={handleKeyDown}
      >
        <Icon iconName={definition ? definition.iconName : 'Unknown'} className={styles.icon} aria-hidden="true" />
        <div className={styles.title}>{name}</div>
        {isEditing && definition?.renderSettings && (
          <IconButton
            id={settingsButtonId}
            iconProps={{ iconName: 'Settings' }}
            title={`${name} settings`}
            ariaLabel={`${name} settings`}
            checked={isSettingsOpen}
            onClick={() => setIsSettingsOpen((open) => !open)}
            onMouseDown={(e) => e.stopPropagation()}
          />
        )}
        {isEditing && (
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title={`Remove ${name}`}
            ariaLabel={`Remove ${name}`}
            onClick={() => onRemove(instance.id)}
            onMouseDown={(e) => e.stopPropagation()}
          />
        )}
      </div>

      <div className={styles.body}>
        {definition ? (
          <WidgetErrorBoundary widgetName={name}>{definition.render(widgetContext)}</WidgetErrorBoundary>
        ) : (
          <div className={styles.missing}>
            This widget (<code>{instance.type}</code>) is no longer available. Remove it in edit mode.
          </div>
        )}
      </div>

      {isSettingsOpen && definition?.renderSettings && (
        <Callout
          target={`#${settingsButtonId}`}
          directionalHint={DirectionalHint.bottomRightEdge}
          onDismiss={() => setIsSettingsOpen(false)}
          setInitialFocus={true}
          role="dialog"
          ariaLabel={`${name} settings`}
        >
          <div className={styles.settingsCallout}>{definition.renderSettings(widgetContext)}</div>
        </Callout>
      )}
    </div>
  );
};
