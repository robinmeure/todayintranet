import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import { SettingsSurface, TextSetting, WidgetProse, textSetting } from '../content';

const DEFAULT_GREETING: string = 'Good to see you';

/**
 * Reference implementation: it shows how a widget reads its own settings, and how
 * text content is laid out with the shared `WidgetProse` primitive.
 */
export const WelcomeWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const greeting = textSetting(context, 'greeting', DEFAULT_GREETING);
  const displayName = context.spContext.pageContext.user.displayName;

  return (
    <WidgetProse heading={`${greeting}, ${displayName.split(' ')[0]}.`}>
      <p>
        This tile is a placeholder. Real widgets (calendar, mail, tasks) plug in the same way:
        describe the content with the shapes in <code>widgets/content</code>, implement{' '}
        <code>IWidgetDefinition</code> and register it.
      </p>
    </WidgetProse>
  );
};

export const WelcomeWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
    <TextSetting
      context={context}
      settingKey="greeting"
      label="Greeting"
      fallback={DEFAULT_GREETING}
      maxLength={40}
    />
  </SettingsSurface>
);
