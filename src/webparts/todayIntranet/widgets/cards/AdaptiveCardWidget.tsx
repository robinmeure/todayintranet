import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import {
  AdaptiveCardPayload,
  SettingsSurface,
  TextSetting,
  WidgetAdaptiveCard,
  WidgetErrorMessage,
  textSetting
} from '../content';

/** Shown until someone pastes their own payload. */
const SAMPLE_CARD: AdaptiveCardPayload = {
  type: 'AdaptiveCard',
  $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
  version: '1.5',
  body: [
    { type: 'TextBlock', text: 'Adaptive Card', weight: 'Bolder', size: 'Medium', wrap: true },
    {
      type: 'TextBlock',
      text: 'This tile renders any Adaptive Card, themed from this site. Paste your own payload in the tile settings.',
      wrap: true,
      isSubtle: true
    },
    {
      type: 'FactSet',
      facts: [
        { title: 'Schema', value: '1.5 and lower' },
        { title: 'Colours', value: 'From the site theme' }
      ]
    }
  ],
  actions: [{ type: 'Action.OpenUrl', title: 'Open the designer', url: 'https://adaptivecards.io/designer/' }]
};

const SAMPLE_JSON: string = JSON.stringify(SAMPLE_CARD, undefined, 2);

/**
 * Renders an Adaptive Card payload the user owns. Useful on its own, and the
 * reference for any widget that wants to hand over a card instead of items.
 */
export const AdaptiveCardWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const json = textSetting(context, 'payload', SAMPLE_JSON);

  const card = React.useMemo<AdaptiveCardPayload | undefined>(() => {
    try {
      const parsed: unknown = JSON.parse(json);
      return parsed && typeof parsed === 'object' ? (parsed as AdaptiveCardPayload) : undefined;
    } catch {
      return undefined;
    }
  }, [json]);

  if (!card) {
    return (
      <WidgetErrorMessage
        error={{
          message: 'That payload is not valid JSON. Fix it in the tile settings.',
          isActionRequired: true
        }}
      />
    );
  }

  return <WidgetAdaptiveCard card={card} ariaLabel="Adaptive card" />;
};

export const AdaptiveCardWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Paste an Adaptive Card payload, schema 1.5 or lower. It is saved with your dashboard, not shared.">
    <TextSetting
      context={context}
      settingKey="payload"
      label="Card payload"
      fallback={SAMPLE_JSON}
      multiline={true}
      rows={12}
      commitOn="blur"
    />
  </SettingsSurface>
);
