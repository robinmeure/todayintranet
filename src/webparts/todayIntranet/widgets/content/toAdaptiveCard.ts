import { IWidgetListItem, WidgetTone } from './IWidgetContent';

/** An Adaptive Card payload. Kept loose because the schema is data, not types. */
export type AdaptiveCardPayload = Record<string, unknown>;

/** Our tones line up one for one with the Adaptive Card text colours. */
const TONE_COLOR: Record<WidgetTone, string> = {
  neutral: 'Default',
  accent: 'Accent',
  success: 'Good',
  warning: 'Warning',
  danger: 'Attention'
};

function toContainer(item: IWidgetListItem, index: number): AdaptiveCardPayload {
  const meta: string = (item.meta ?? []).filter((part) => !!part).join(' · ');

  const texts: AdaptiveCardPayload[] = [
    {
      type: 'TextBlock',
      text: item.title,
      wrap: true,
      weight: item.isEmphasized ? 'Bolder' : 'Default',
      color: TONE_COLOR[item.tone ?? 'neutral'],
      spacing: 'None'
    }
  ];
  if (meta) {
    texts.push({ type: 'TextBlock', text: meta, wrap: false, size: 'Small', isSubtle: true, spacing: 'None' });
  }
  if (item.description) {
    texts.push({
      type: 'TextBlock',
      text: item.description,
      wrap: true,
      size: 'Small',
      isSubtle: true,
      spacing: 'None'
    });
  }

  const columns: AdaptiveCardPayload[] = [
    { type: 'Column', width: 'stretch', verticalContentAlignment: 'Center', items: texts }
  ];
  if (item.badge) {
    columns.push({
      type: 'Column',
      width: 'auto',
      verticalContentAlignment: 'Center',
      items: [
        {
          type: 'TextBlock',
          text: item.badge.text,
          size: 'Small',
          weight: 'Bolder',
          wrap: false,
          spacing: 'None',
          color: TONE_COLOR[item.badge.tone ?? item.tone ?? 'neutral']
        }
      ]
    });
  }

  const container: AdaptiveCardPayload = {
    type: 'Container',
    spacing: index === 0 ? 'None' : 'Small',
    separator: index > 0,
    items: [{ type: 'ColumnSet', spacing: 'None', columns }]
  };
  if (item.href) {
    container.selectAction = { type: 'Action.OpenUrl', title: item.title, url: item.href };
  }
  return container;
}

/**
 * Renders the same items every other view shows, as an Adaptive Card payload. This is
 * what lets a tile switch to the Adaptive Card view without its widget knowing.
 */
export function itemsToAdaptiveCard(items: IWidgetListItem[]): AdaptiveCardPayload {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: items.map(toContainer)
  };
}
