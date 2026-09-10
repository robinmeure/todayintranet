import * as React from 'react';
import { IWidgetListItem, WidgetItemsView } from './IWidgetContent';
import { WidgetList } from './WidgetList';
import { WidgetCards } from './WidgetCards';
import { WidgetAdaptiveCard } from './WidgetAdaptiveCard';
import { itemsToAdaptiveCard } from './toAdaptiveCard';

export interface IWidgetItemsProps {
  items: IWidgetListItem[];
  /** Chosen by whoever owns the tile. Defaults to the plain list. */
  view?: WidgetItemsView;
  ariaLabel?: string;
}

/**
 * Draws a collection of items in whichever view the tile is set to. Widgets render
 * through here rather than picking a presentation themselves, which is what lets the
 * same calendar or inbox be a list, a set of cards, a gallery or an Adaptive Card
 * without a single widget knowing about it.
 */
export const WidgetItems: React.FunctionComponent<IWidgetItemsProps> = ({ items, view, ariaLabel }) => {
  switch (view) {
    case 'compact':
      return <WidgetList items={items} density="compact" ariaLabel={ariaLabel} />;
    case 'cards':
      return <WidgetCards items={items} layout="stack" ariaLabel={ariaLabel} />;
    case 'gallery':
      return <WidgetCards items={items} layout="grid" ariaLabel={ariaLabel} />;
    case 'adaptive':
      return <WidgetAdaptiveCard card={itemsToAdaptiveCard(items)} ariaLabel={ariaLabel} />;
    default:
      return <WidgetList items={items} ariaLabel={ariaLabel} />;
  }
};
