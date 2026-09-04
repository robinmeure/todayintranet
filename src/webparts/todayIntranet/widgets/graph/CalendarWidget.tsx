import * as React from 'react';
import { Event } from '@microsoft/microsoft-graph-types';
import { IWidgetContext } from '../IWidget';
import { useGraphData } from './useGraphData';
import { WidgetLoading, WidgetEmpty, WidgetError } from '../WidgetMessage';
import { NumberSetting, SettingsSurface, numberSetting } from './settings';
import styles from '../WidgetMessage.module.scss';

const SCOPE: string = 'Calendars.ReadBasic';
const DEFAULT_DAYS: number = 7;
const DEFAULT_MAX_ITEMS: number = 5;

function localTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}

function formatWhen(event: Event, locale: string | undefined): string {
  if (!event.start?.dateTime) {
    return '';
  }
  // Graph returns an unzoned string already converted to the Prefer time zone.
  const start = new Date(event.start.dateTime);
  const day = start.toLocaleDateString(locale, { weekday: 'short', day: 'numeric', month: 'short' });

  if (event.isAllDay) {
    return `${day} · All day`;
  }

  const time = start.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' });
  const end = event.end?.dateTime
    ? new Date(event.end.dateTime).toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })
    : undefined;

  return end ? `${day} · ${time}–${end}` : `${day} · ${time}`;
}

export const CalendarWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const days = numberSetting(context, 'days', DEFAULT_DAYS);
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const result = useGraphData<Event[]>(
    context.spContext,
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

  if (result.status === 'loading') {
    return <WidgetLoading label="Loading your calendar…" />;
  }
  if (result.status === 'error' && result.error) {
    return <WidgetError error={result.error} onRetry={result.reload} />;
  }
  if (!result.data || result.data.length === 0) {
    return <WidgetEmpty iconName="Calendar" text={`Nothing scheduled in the next ${days} days.`} />;
  }

  return (
    <ul className={styles.list}>
      {result.data.map((event) => (
        <li key={event.id}>
          <a className={styles.item} href={event.webLink ?? '#'} target="_blank" rel="noreferrer">
            <span className={styles.itemAccent} aria-hidden="true" />
            <span className={styles.itemText}>
              <span className={styles.itemTitle}>{event.subject || '(No subject)'}</span>
              <span className={styles.itemMeta}>
                {formatWhen(event, locale)}
                {event.location?.displayName ? ` · ${event.location.displayName}` : ''}
              </span>
            </span>
          </a>
        </li>
      ))}
    </ul>
  );
};

export const CalendarWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface>
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
