import * as React from 'react';
import { SpinButton } from '@fluentui/react/lib/SpinButton';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { TextField } from '@fluentui/react/lib/TextField';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetContext } from '../IWidget';
import { IWidgetItemsViewInfo, WidgetItemsView, getItemsView } from './IWidgetContent';
import styles from './WidgetContent.module.scss';

/** Where the chosen view lives in a widget instance's settings. */
export const VIEW_SETTING_KEY: string = 'view';

/* Typed readers for the untyped settings bag every widget instance carries. */

export function numberSetting(context: IWidgetContext, key: string, fallback: number): number {
  const value: unknown = context.settings[key];
  return typeof value === 'number' && isFinite(value) ? value : fallback;
}

export function booleanSetting(context: IWidgetContext, key: string, fallback: boolean): boolean {
  const value: unknown = context.settings[key];
  return typeof value === 'boolean' ? value : fallback;
}

export function textSetting(context: IWidgetContext, key: string, fallback: string): string {
  const value: unknown = context.settings[key];
  return typeof value === 'string' ? value : fallback;
}

/** The view this tile is set to, falling back to the widget's own default. */
export function viewSetting(context: IWidgetContext, fallback: WidgetItemsView): WidgetItemsView {
  const value: unknown = context.settings[VIEW_SETTING_KEY];
  const known = typeof value === 'string' ? getItemsView(value) : undefined;
  return known ? known.id : fallback;
}

function write(context: IWidgetContext, key: string, value: unknown): void {
  context.updateSettings({ ...context.settings, [key]: value });
}

export interface IDraftSetting {
  value: string;
  setValue(next: string): void;
  /** Writes the draft into settings, if it differs from what is stored. */
  commit(): void;
}

/**
 * A text setting the user is still editing. The draft is only written on demand,
 * which keeps half-typed values — a partial JSON payload, an unfinished query — out
 * of the saved settings, and out of anything watching them.
 *
 * Dismissing the settings flyout unmounts the field before it can blur, so the draft
 * is also committed on the way out.
 */
export function useDraftSetting(context: IWidgetContext, key: string, fallback: string): IDraftSetting {
  const stored: string = textSetting(context, key, fallback);
  const [value, setValue] = React.useState<string>(stored);

  React.useEffect(() => setValue(stored), [stored]);

  const latest = React.useRef<() => void>();
  latest.current = () => {
    if (value !== stored) {
      write(context, key, value);
    }
  };

  React.useEffect(
    () => () => {
      if (latest.current) {
        latest.current();
      }
    },
    []
  );

  return {
    value,
    setValue,
    commit: () => {
      if (latest.current) {
        latest.current();
      }
    }
  };
}

export interface INumberSettingProps {
  context: IWidgetContext;
  settingKey: string;
  label: string;
  fallback: number;
  min: number;
  max: number;
}

export const NumberSetting: React.FunctionComponent<INumberSettingProps> = (props) => {
  const { context, settingKey, label, fallback, min, max } = props;
  const value: number = numberSetting(context, settingKey, fallback);

  const commit = (next: number): void => {
    write(context, settingKey, Math.max(min, Math.min(max, next)));
  };

  return (
    <SpinButton
      label={label}
      min={min}
      max={max}
      step={1}
      value={String(value)}
      onValidate={(raw) => {
        commit(parseInt(raw, 10) || fallback);
        return undefined;
      }}
      onIncrement={() => {
        commit(value + 1);
        return undefined;
      }}
      onDecrement={() => {
        commit(value - 1);
        return undefined;
      }}
    />
  );
};

export interface IToggleSettingProps {
  context: IWidgetContext;
  settingKey: string;
  label: string;
  fallback: boolean;
}

export const ToggleSetting: React.FunctionComponent<IToggleSettingProps> = (props) => {
  const { context, settingKey, label, fallback } = props;
  return (
    <Toggle
      label={label}
      checked={booleanSetting(context, settingKey, fallback)}
      onChange={(_, checked) => write(context, settingKey, !!checked)}
    />
  );
};

export interface ITextSettingProps {
  context: IWidgetContext;
  settingKey: string;
  label: string;
  fallback: string;
  maxLength?: number;
  multiline?: boolean;
  rows?: number;
  /**
   * `blur` keeps half-typed values out of the saved settings, which matters for
   * anything that has to parse — a JSON payload, for instance. Defaults to `change`.
   */
  commitOn?: 'change' | 'blur';
}

export const TextSetting: React.FunctionComponent<ITextSettingProps> = (props) => {
  const { context, settingKey, label, fallback, maxLength, multiline, rows, commitOn } = props;
  const draft = useDraftSetting(context, settingKey, fallback);

  return (
    <TextField
      label={label}
      value={draft.value}
      maxLength={maxLength}
      multiline={multiline}
      rows={rows}
      onChange={(_, next) => {
        const value = next ?? '';
        draft.setValue(value);
        if (commitOn !== 'blur') {
          write(context, settingKey, value);
        }
      }}
      onBlur={() => {
        if (commitOn === 'blur') {
          draft.commit();
        }
      }}
    />
  );
};

export interface IViewSettingProps {
  context: IWidgetContext;
  /** Views this widget supports, in the order they should be offered. */
  views: WidgetItemsView[];
  /** The widget's own default, used until the user picks something. */
  fallback: WidgetItemsView;
}

/**
 * The view picker every collection widget gets for free. The frame renders it, so
 * choosing "Gallery" or "Adaptive card" works the same way on every tile.
 */
export const ViewSetting: React.FunctionComponent<IViewSettingProps> = ({ context, views, fallback }) => {
  const current: WidgetItemsView = viewSetting(context, fallback);
  const labelId: string = `widget-view-${context.instanceId}`;
  const options: IWidgetItemsViewInfo[] = [];
  views.forEach((id) => {
    const info = getItemsView(id);
    if (info) {
      options.push(info);
    }
  });

  return (
    <div className={styles.viewSetting}>
      <div className={styles.viewSettingLabel} id={labelId}>
        View
      </div>
      <div className={styles.viewOptions} role="group" aria-labelledby={labelId}>
        {options.map((option) => {
          const isSelected = option.id === current;
          return (
            <button
              key={option.id}
              type="button"
              title={option.description}
              aria-pressed={isSelected}
              className={`${styles.viewOption} ${isSelected ? styles.viewOptionSelected : ''}`}
              onClick={() => write(context, VIEW_SETTING_KEY, option.id)}
            >
              <Icon iconName={option.iconName} className={styles.viewOptionIcon} aria-hidden={true} />
              <span className={styles.viewOptionText}>{option.name}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
};

export interface ISettingsSurfaceProps {
  /** One line explaining what these settings affect. */
  description?: string;
  children: React.ReactNode;
}

/** Consistent spacing for the fields inside a tile's settings flyout. */
export const SettingsSurface: React.FunctionComponent<ISettingsSurfaceProps> = ({ description, children }) => (
  <div className={styles.settings}>
    {description && <p className={styles.settingsDescription}>{description}</p>}
    {children}
  </div>
);
