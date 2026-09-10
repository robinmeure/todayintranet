import * as React from 'react';
import { Event } from '@microsoft/microsoft-graph-types';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  IWidgetListItem,
  NumberSetting,
  SettingsSurface,
  WidgetItems,
  WidgetItemsView,
  WidgetView,
  numberSetting,
  viewSetting
} from '../content';

const SCOPE: string = 'Calendars.ReadBasic';
const DEFAULT_DAYS: number = 7;
const DEFAULT_MAX_ITEMS: number = 5;
const CALENDAR_URL: string = 'https://outlook.office.com/calendar/';

export const CALENDAR_LINK: IWidgetLink = { text: 'Open calendar', href: CALENDAR_URL };
export const CALENDAR_DEFAULT_VIEW: WidgetItemsView = 'list';

function localTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}

function dayLabel(start: Date, now: Date, locale: string | undefined): string {
  const tomorrow = new Date(now);
  tomorrow.setDate(tomorrow.getDate() + 1);

  if (start.toDateString() === now.toDateString()) {
    return 'Today';
  }
  if (start.toDateString() === tomorrow.toDateString()) {
    return 'Tomorrow';
  }
  return start.toLocaleDateString(locale, { weekday: 'short', day: 'numeric', month: 'short' });
}

function formatWhen(event: Event, start: Date, now: Date, locale: string | undefined): string {
  const day = dayLabel(start, now, locale);
  if (event.isAllDay) {
    return `${day} · All day`;
  }

  const time = start.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' });
  const end = event.end?.dateTime
    ? new Date(event.end.dateTime).toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })
    : undefined;

  return end ? `${day} · ${time}–${end}` : `${day} · ${time}`;
}

/** One meeting, described in the shared list vocabulary. */
function toItem(event: Event, index: number, now: Date, locale: string | undefined): IWidgetListItem {
  // Graph returns an unzoned string already converted to the Prefer time zone.
  const start = event.start?.dateTime ? new Date(event.start.dateTime) : undefined;
  const end = event.end?.dateTime ? new Date(event.end.dateTime) : undefined;
  const isToday = !!start && start.toDateString() === now.toDateString();
  const isInProgress = !!start && start <= now && !!end && end > now;

  return {
    key: event.id ?? String(index),
    title: event.subject || '(No subject)',
    meta: [start ? formatWhen(event, start, now, locale) : undefined, event.location?.displayName ?? undefined],
    hasAccentBar: true,
    // Today stands out from the rest of the week, and whatever is running right now
    // stands out from today.
    tone: isInProgress ? 'success' : isToday ? 'accent' : 'neutral',
    isEmphasized: isToday,
    badge: isInProgress ? { text: 'Now' } : undefined,
    href: event.webLink ?? undefined
  };
}

export const CalendarWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const days = numberSetting(context, 'days', DEFAULT_DAYS);
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS);
  const view = viewSetting(context, CALENDAR_DEFAULT_VIEW);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const state = useGraphData<Event[]>(
    context,
    SCOPE,
    async (client) => {
      const start = new Date();
      const end = new Date(start.getTime() + days * 24 * 60 * 60 * 1000);

      const response = await client
        .api('/me/calendarView')
        .version('v1.0')
        .header('Prefer', `outlook.timezone="${localTimeZone()}"`)
        .query({ startDateTime: start.toISOString(), endDateTime: end.toISOString() })
        .select('subject,start,end,isAllDay,location,webLink,isCancelled')
        .orderby('start/dateTime')
        .top(maxItems + 5)
        .get();

      const events = (response.value ?? []) as Event[];
      return events.filter((e) => !e.isCancelled).slice(0, maxItems);
    },
    [days, maxItems]
  );

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Loading your calendar…', rows: Math.min(maxItems, 4) }}
      empty={{
        iconName: 'Calendar',
        text: `Nothing scheduled in the next ${days} days.`
      }}
    >
      {(events) => {
        const now = new Date();
        return (
          <WidgetItems
            view={view}
            ariaLabel="Upcoming events"
            items={events.map((event, index) => toItem(event, index, now, locale))}
          />
        );
      }}
    </WidgetView>
  );
};

export const CalendarWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
    <NumberSetting context={context} settingKey="days" label="Days ahead" fallback={DEFAULT_DAYS} min={1} max={30} />
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Events to show"
      fallback={DEFAULT_MAX_ITEMS}
      min={1}
      max={20}
    />
  </SettingsSurface>
);
