import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { Link } from '@fluentui/react/lib/Link';
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
import styles from './WidgetFrame.module.scss';
import { WidgetErrorBoundary } from './WidgetErrorBoundary';
import { IWidgetInstance } from '../model/IDashboardLayout';
import { IWidgetContext, IWidgetHostContext } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';
import { ViewSetting, WidgetEmpty } from '../widgets/content';

/** A keyboard driven move or resize, expressed in grid units. */
export interface INudge {
  dx?: number;
  dy?: number;
  dw?: number;
  dh?: number;
}

export interface IWidgetFrameProps {
  instance: IWidgetInstance;
  /** Everything but the refresh state, which this frame owns. */
  widgetContext: IWidgetHostContext;
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

/** Fluent's own class beats a stylesheet rule here, so the size is set on the button. */
const ACTION_BUTTON_STYLES: { root: { width: number; height: number } } = {
  root: { width: 28, height: 28 }
};

/** Shared chrome (title bar, drag handle, refresh, settings, remove) around every widget. */
export const WidgetFrame: React.FunctionComponent<IWidgetFrameProps> = (props) => {
  const { instance, widgetContext, isEditing, onRemove, onNudge } = props;
  const definition = WidgetRegistry.get(instance.type);
  const name = definition ? definition.displayName : instance.type;

  const [isSettingsOpen, setIsSettingsOpen] = React.useState<boolean>(false);
  const [refreshToken, setRefreshToken] = React.useState<number>(0);
  const settingsButtonId = `widget-settings-${instance.id}`;

  // A widget offering more than one view gets the picker, and therefore a gear,
  // whether or not it has settings of its own.
  const views = definition?.supportedViews ?? [];
  const hasViewPicker = views.length > 1;
  const hasSettings = !!definition?.renderSettings || hasViewPicker;

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

  // The widget sees the refresh as a changed dependency, so it re-reads its data
  // without the frame having to know anything about where that data comes from.
  const context: IWidgetContext = { ...widgetContext, refreshToken };

  const footerSource = definition?.footerLink;
  const footerLink = typeof footerSource === 'function' ? footerSource(context) : footerSource;

  return (
    <section className={`${styles.frame} ${isEditing ? styles.frameEditing : ''}`} aria-label={name}>
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
        {isEditing && (
          <Icon iconName="GripperDotsVertical" className={styles.gripper} aria-hidden="true" />
        )}
        <Icon iconName={definition ? definition.iconName : 'Unknown'} className={styles.icon} aria-hidden="true" />
        <div className={styles.title}>{name}</div>

        {/* Quiet until the tile is hovered or focused, so a full dashboard stays calm. */}
        <div className={`${styles.actions} ${isEditing ? styles.actionsPinned : ''}`}>
          {definition?.isRefreshable && (
            <IconButton
              styles={ACTION_BUTTON_STYLES}
              iconProps={{ iconName: 'Refresh' }}
              title={`Refresh ${name}`}
              ariaLabel={`Refresh ${name}`}
              onClick={() => setRefreshToken((token) => token + 1)}
              onMouseDown={(e) => e.stopPropagation()}
            />
          )}
          {isEditing && hasSettings && (
            <IconButton
              id={settingsButtonId}
              styles={ACTION_BUTTON_STYLES}
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
              styles={ACTION_BUTTON_STYLES}
              iconProps={{ iconName: 'Delete' }}
              title={`Remove ${name}`}
              ariaLabel={`Remove ${name}`}
              onClick={() => onRemove(instance.id)}
              onMouseDown={(e) => e.stopPropagation()}
            />
          )}
        </div>
      </div>

      <div className={styles.body}>
        {definition ? (
          <WidgetErrorBoundary widgetName={name}>{definition.render(context)}</WidgetErrorBoundary>
        ) : (
          <WidgetEmpty
            iconName="Unknown"
            text={`This widget (${instance.type}) is no longer available. Remove it in edit mode.`}
          />
        )}
      </div>

      {footerLink && (
        <div className={styles.footer}>
          <Link href={footerLink.href} target="_blank" rel="noreferrer" className={styles.footerLink}>
            {footerLink.text}
            <Icon iconName="ChevronRightSmall" className={styles.footerIcon} aria-hidden="true" />
          </Link>
        </div>
      )}

      {isSettingsOpen && definition && hasSettings && (
        <Callout
          target={`#${settingsButtonId}`}
          directionalHint={DirectionalHint.bottomRightEdge}
          onDismiss={() => setIsSettingsOpen(false)}
          setInitialFocus={true}
          role="dialog"
          ariaLabel={`${name} settings`}
        >
          <div className={`${styles.settingsCallout} ${definition.isSettingsWide ? styles.settingsCalloutWide : ''}`}>
            <div className={styles.settingsTitle}>{name}</div>
            <div className={styles.settingsStack}>
              {hasViewPicker && (
                <ViewSetting
                  context={context}
                  views={views}
                  fallback={definition.defaultView ?? views[0]}
                />
              )}
              {definition.renderSettings && definition.renderSettings(context)}
            </div>
          </div>
        </Callout>
      )}
    </section>
  );
};
