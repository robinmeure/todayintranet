import * as React from 'react';
import { SpinButton } from '@fluentui/react/lib/SpinButton';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { IWidgetContext } from '../IWidget';
import styles from '../WidgetMessage.module.scss';

export function numberSetting(context: IWidgetContext, key: string, fallback: number): number {
  const value = context.settings[key];
  return typeof value === 'number' && isFinite(value) ? value : fallback;
}

export function booleanSetting(context: IWidgetContext, key: string, fallback: boolean): boolean {
  const value = context.settings[key];
  return typeof value === 'boolean' ? value : fallback;
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
  const value = numberSetting(context, settingKey, fallback);

  const commit = (next: number): void => {
    const clamped = Math.max(min, Math.min(max, next));
    context.updateSettings({ ...context.settings, [settingKey]: clamped });
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
      onChange={(_, checked) => context.updateSettings({ ...context.settings, [settingKey]: !!checked })}
    />
  );
};

export const SettingsSurface: React.FunctionComponent<{ children: React.ReactNode }> = ({ children }) => (
  <div className={styles.settings}>{children}</div>
);
