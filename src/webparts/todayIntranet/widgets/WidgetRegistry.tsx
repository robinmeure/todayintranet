import * as React from 'react';
import { IWidgetDefinition } from './IWidget';
import { WelcomeWidget } from './demo/WelcomeWidget';
import { ClockWidget } from './demo/ClockWidget';

const definitions: IWidgetDefinition[] = [
  {
    type: 'demo.welcome',
    displayName: 'Welcome',
    description: 'Placeholder tile that shows the signed in user and a configurable greeting.',
    iconName: 'Contact',
    defaultSize: { w: 6, h: 4 },
    minSize: { w: 3, h: 3 },
    render: (context) => <WelcomeWidget context={context} />
  },
  {
    type: 'demo.clock',
    displayName: 'Clock',
    description: 'Placeholder tile showing the current date and time.',
    iconName: 'Clock',
    defaultSize: { w: 3, h: 4 },
    minSize: { w: 2, h: 3 },
    render: (context) => <ClockWidget context={context} />
  }
];

const byType: Map<string, IWidgetDefinition> = new Map(definitions.map((d) => [d.type, d]));

export const WidgetRegistry = {
  /** All widgets offered in the "Add a widget" catalogue. */
  list(): IWidgetDefinition[] {
    return definitions.slice();
  },

  /** Undefined when a stored layout references a widget that no longer ships. */
  get(type: string): IWidgetDefinition | undefined {
    return byType.get(type);
  }
};
