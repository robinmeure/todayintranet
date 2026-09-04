import * as React from 'react';
import { IWidgetDefinition } from './IWidget';
import { WelcomeWidget } from './demo/WelcomeWidget';
import { ClockWidget } from './demo/ClockWidget';
import { CalendarWidget, CalendarWidgetSettings } from './graph/CalendarWidget';
import { MailWidget, MailWidgetSettings } from './graph/MailWidget';
import { TasksWidget, TasksWidgetSettings } from './graph/TasksWidget';

const definitions: IWidgetDefinition[] = [
  {
    type: 'm365.calendar',
    displayName: 'Calendar',
    description: 'Your upcoming Outlook meetings and appointments.',
    iconName: 'Calendar',
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <CalendarWidget context={context} />,
    renderSettings: (context) => <CalendarWidgetSettings context={context} />
  },
  {
    type: 'm365.mail',
    displayName: 'My mail',
    description: 'The latest messages in your Outlook inbox.',
    iconName: 'Mail',
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <MailWidget context={context} />,
    renderSettings: (context) => <MailWidgetSettings context={context} />
  },
  {
    type: 'm365.tasks',
    displayName: 'My tasks',
    description: 'Open items from your Microsoft To Do list.',
    iconName: 'CheckboxComposite',
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <TasksWidget context={context} />,
    renderSettings: (context) => <TasksWidgetSettings context={context} />
  },
  {
    type: 'demo.clock',
    displayName: 'Clock',
    description: 'The current date and time.',
    iconName: 'Clock',
    defaultSize: { w: 3, h: 4 },
    minSize: { w: 2, h: 3 },
    render: (context) => <ClockWidget context={context} />
  },
  {
    type: 'demo.welcome',
    displayName: 'Welcome',
    description: 'A greeting for the signed in user.',
    iconName: 'Contact',
    defaultSize: { w: 6, h: 4 },
    minSize: { w: 3, h: 3 },
    render: (context) => <WelcomeWidget context={context} />
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
