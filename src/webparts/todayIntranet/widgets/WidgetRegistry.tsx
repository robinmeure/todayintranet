import * as React from 'react';
import {
  IWidgetCategoryInfo,
  IWidgetDefinition,
  WIDGET_CATEGORIES,
  WidgetCategory
} from './IWidget';
import { ALL_ITEM_VIEWS } from './content';
import { WelcomeWidget, WelcomeWidgetSettings } from './demo/WelcomeWidget';
import { ClockWidget } from './demo/ClockWidget';
import { AdaptiveCardWidget, AdaptiveCardWidgetSettings } from './cards/AdaptiveCardWidget';
import { NewsWidget, NewsWidgetSettings, NEWS_DEFAULT_VIEW, newsFooterLink } from './search/NewsWidget';
import {
  SearchResultsWidget,
  SearchResultsWidgetSettings,
  SEARCH_DEFAULT_VIEW
} from './search/SearchResultsWidget';
import {
  CalendarWidget,
  CalendarWidgetSettings,
  CALENDAR_LINK,
  CALENDAR_DEFAULT_VIEW
} from './graph/CalendarWidget';
import { MailWidget, MailWidgetSettings, MAIL_LINK, MAIL_DEFAULT_VIEW } from './graph/MailWidget';
import { TasksWidget, TasksWidgetSettings, TODO_LINK, TASKS_DEFAULT_VIEW } from './graph/TasksWidget';

const definitions: IWidgetDefinition[] = [
  {
    type: 'm365.calendar',
    displayName: 'Calendar',
    description: 'Your upcoming Outlook meetings and appointments.',
    iconName: 'Calendar',
    category: 'microsoft365',
    keywords: ['outlook', 'agenda', 'meetings', 'events', 'schedule'],
    requiredPermission: 'Calendars.ReadBasic',
    isRefreshable: true,
    footerLink: CALENDAR_LINK,
    supportedViews: ALL_ITEM_VIEWS,
    defaultView: CALENDAR_DEFAULT_VIEW,
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
    category: 'microsoft365',
    keywords: ['outlook', 'inbox', 'email', 'messages'],
    requiredPermission: 'Mail.ReadBasic',
    isRefreshable: true,
    footerLink: MAIL_LINK,
    supportedViews: ALL_ITEM_VIEWS,
    defaultView: MAIL_DEFAULT_VIEW,
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
    category: 'microsoft365',
    keywords: ['to do', 'todo', 'planner', 'checklist'],
    requiredPermission: 'Tasks.Read',
    isRefreshable: true,
    footerLink: TODO_LINK,
    supportedViews: ALL_ITEM_VIEWS,
    defaultView: TASKS_DEFAULT_VIEW,
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <TasksWidget context={context} />,
    renderSettings: (context) => <TasksWidgetSettings context={context} />
  },
  {
    type: 'sp.news',
    displayName: 'News',
    description: 'News posts from this site, this hub, or everywhere you have access.',
    iconName: 'News',
    category: 'content',
    keywords: ['news', 'posts', 'announcements', 'articles', 'search'],
    isRefreshable: true,
    footerLink: newsFooterLink,
    isSettingsWide: true,
    supportedViews: ALL_ITEM_VIEWS,
    defaultView: NEWS_DEFAULT_VIEW,
    defaultSize: { w: 6, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <NewsWidget context={context} />,
    renderSettings: (context) => <NewsWidgetSettings context={context} />
  },
  {
    type: 'sp.search',
    displayName: 'Search results',
    description: 'Any SharePoint search query, with a preview while you build it.',
    iconName: 'Search',
    category: 'content',
    keywords: ['search', 'query', 'kql', 'documents', 'files', 'results'],
    isRefreshable: true,
    isSettingsWide: true,
    supportedViews: ALL_ITEM_VIEWS,
    defaultView: SEARCH_DEFAULT_VIEW,
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 4 },
    render: (context) => <SearchResultsWidget context={context} />,
    renderSettings: (context) => <SearchResultsWidgetSettings context={context} />
  },
  {
    type: 'card.adaptive',
    displayName: 'Adaptive card',
    description: 'Show any Adaptive Card. Paste a payload into the tile settings.',
    iconName: 'RectangleShape',
    category: 'content',
    keywords: ['adaptive', 'card', 'json', 'payload', 'announcement'],
    defaultSize: { w: 4, h: 6 },
    minSize: { w: 3, h: 3 },
    render: (context) => <AdaptiveCardWidget context={context} />,
    renderSettings: (context) => <AdaptiveCardWidgetSettings context={context} />
  },
  {
    type: 'demo.clock',
    displayName: 'Clock',
    description: 'The current date and time.',
    iconName: 'Clock',
    category: 'demo',
    keywords: ['time', 'date'],
    defaultSize: { w: 3, h: 4 },
    minSize: { w: 2, h: 3 },
    render: (context) => <ClockWidget context={context} />
  },
  {
    type: 'demo.welcome',
    displayName: 'Welcome',
    description: 'A greeting for the signed in user.',
    iconName: 'Contact',
    category: 'demo',
    keywords: ['greeting', 'hello', 'sample'],
    defaultSize: { w: 6, h: 4 },
    minSize: { w: 3, h: 3 },
    render: (context) => <WelcomeWidget context={context} />,
    renderSettings: (context) => <WelcomeWidgetSettings context={context} />
  }
];

const byType: Map<string, IWidgetDefinition> = new Map(definitions.map((d) => [d.type, d]));

/** A catalogue section: one category and the widgets in it. */
export interface IWidgetCategoryGroup {
  category: IWidgetCategoryInfo;
  widgets: IWidgetDefinition[];
}

export const WidgetRegistry = {
  /** All widgets offered in the "Add a widget" catalogue. */
  list(): IWidgetDefinition[] {
    return definitions.slice();
  },

  /** Undefined when a stored layout references a widget that no longer ships. */
  get(type: string): IWidgetDefinition | undefined {
    return byType.get(type);
  },

  /** True when the widget matches a free text search of the catalogue. */
  matches(definition: IWidgetDefinition, query: string): boolean {
    const needle = query.trim().toLowerCase();
    if (!needle) {
      return true;
    }
    const haystack = [definition.displayName, definition.description, ...(definition.keywords ?? [])]
      .join(' ')
      .toLowerCase();
    return haystack.indexOf(needle) >= 0;
  },

  /** Buckets widgets into catalogue sections, in category order. Empty ones are dropped. */
  group(widgets: IWidgetDefinition[]): IWidgetCategoryGroup[] {
    const groups: IWidgetCategoryGroup[] = [];
    WIDGET_CATEGORIES.forEach((category: IWidgetCategoryInfo) => {
      const members = widgets.filter((widget) => widget.category === category.id);
      if (members.length > 0) {
        groups.push({ category, widgets: members });
      }
    });

    // A widget registered with a category that is not declared still has to show up.
    const known: WidgetCategory[] = WIDGET_CATEGORIES.map((category) => category.id);
    const orphans = widgets.filter((widget) => known.indexOf(widget.category) < 0);
    if (orphans.length > 0) {
      groups.push({ category: { id: orphans[0].category, name: 'Other' }, widgets: orphans });
    }

    return groups;
  }
};
