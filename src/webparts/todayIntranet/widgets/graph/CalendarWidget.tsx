import * as React from 'react';
import { Event } from '@microsoft/microsoft-graph-types';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  NumberSetting, SettingsSurface, WidgetDateToolbar, WidgetItems, WidgetItemsView, WidgetView,
  numberSetting, viewSetting
} from '../content';
import {
  CalendarRange, calendarDateKey, calendarItems, calendarWindow, loadCalendarEvents, parseCalendarDate
} from './calendarData';

const DEFAULT_DAYS: number = 7;
const DEFAULT_MAX_ITEMS: number = 5;
const DAYS_BOUNDS = { min: 1, max: 30 };
const MAX_ITEMS_BOUNDS = { min: 1, max: 20 };
const CALENDAR_URL: string = 'https://outlook.office.com/calendar/';
const REFRESH_INTERVAL: number = 5 * 60 * 1000;

export const CALENDAR_LINK: IWidgetLink = { text: 'Open Outlook', href: CALENDAR_URL };
export const CALENDAR_DEFAULT_VIEW: WidgetItemsView = 'agenda';

function useCalendarClock(): Date {
  const [now, setNow] = React.useState(() => new Date());
  React.useEffect(() => {
    const tick = (): void => setNow(new Date());
    const timer = window.setInterval(tick, 15000);
    window.addEventListener('focus', tick);
    document.addEventListener('visibilitychange', tick);
    return () => {
      window.clearInterval(timer);
      window.removeEventListener('focus', tick);
      document.removeEventListener('visibilitychange', tick);
    };
  }, []);
  return now;
}

interface ICalendarScheduleProps {
  context: IWidgetContext;
  now: Date;
  range: CalendarRange;
  start: string;
  end: string;
  days: number;
  maxItems: number;
  locale: string | undefined;
  dateError?: string;
  onRangeChange(key: string): void;
  onDateChange(value: string): void;
}

interface ICalendarResult {
  start: string;
  end: string;
  events: Event[];
}

const CalendarSchedule: React.FunctionComponent<ICalendarScheduleProps> = ({
  context, now, range, start, end, days, maxItems, locale, dateError, onRangeChange, onDateChange
}) => {
  const refreshWindow = Math.floor(now.getTime() / REFRESH_INTERVAL);
  const previousRefreshWindow = React.useRef(refreshWindow);
  const state = useGraphData<ICalendarResult>(
    context,
    'Calendars.ReadBasic',
    async (client) => ({ start, end, events: await loadCalendarEvents(client, start, end) }),
    [start, end]
  );
  React.useEffect(() => {
    if (previousRefreshWindow.current !== refreshWindow) {
      previousRefreshWindow.current = refreshWindow;
      state.reload();
    }
  }, [refreshWindow, state.reload]);

  // Keep the toolbar mounted for keyboard focus, but never label another range's data as this one.
  const hasPreviousRange = !!state.data && (state.data.start !== start || state.data.end !== end);
  const status = hasPreviousRange ? 'loading' : state.status;
  const items = state.data && !hasPreviousRange
    ? calendarItems(state.data.events, new Date(start), now, maxItems, locale) : undefined;

  return (
    <>
      <WidgetDateToolbar
        ranges={[{ key: 'today', text: 'Today' }, { key: 'tomorrow', text: 'Tomorrow' }, { key: 'upcoming', text: 'Upcoming' }]}
        activeRange={range}
        date={calendarDateKey(new Date(start))}
        dateError={dateError}
        isRefreshing={status === 'loading' || state.isRefreshing}
        onRangeChange={onRangeChange}
        onDateChange={onDateChange}
        onRefresh={state.reload}
      />
      <WidgetView
        state={{ ...state, status, data: items }}
        loading={{ label: 'Loading your calendar...', rows: Math.min(maxItems, 4) }}
        empty={{
          iconName: 'Calendar',
          text: range === 'today' ? 'Nothing else scheduled today.' : range === 'tomorrow'
            ? 'Nothing scheduled tomorrow.' : range === 'upcoming'
              ? `Nothing else scheduled in the next ${days} days.` : 'Nothing scheduled on this date.'
        }}
      >
        {(events) => <WidgetItems view={viewSetting(context, CALENDAR_DEFAULT_VIEW)} items={events} ariaLabel="Calendar events" />}
      </WidgetView>
    </>
  );
};

export const CalendarWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const days = numberSetting(context, 'days', DEFAULT_DAYS, DAYS_BOUNDS);
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, MAX_ITEMS_BOUNDS);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;
  const now = useCalendarClock();
  const [range, setRange] = React.useState<CalendarRange>('today');
  const [customDate, setCustomDate] = React.useState(() => calendarDateKey(now));
  const [dateError, setDateError] = React.useState<string>();
  const selected = calendarWindow(now, range, customDate, days);
  const start = selected.start.toISOString();
  const end = selected.end.toISOString();
  return (
    <CalendarSchedule
      context={context}
      now={now}
      range={range}
      start={start}
      end={end}
      days={days}
      maxItems={maxItems}
      locale={locale}
      dateError={dateError}
      onRangeChange={(key) => {
        if (key === 'today' || key === 'tomorrow' || key === 'upcoming') {
          setRange(key);
          setDateError(undefined);
        }
      }}
      onDateChange={(value) => {
        if (!parseCalendarDate(value)) {
          setDateError('Choose a valid calendar date.');
        } else {
          setCustomDate(value);
          setRange('date');
          setDateError(undefined);
        }
      }}
    />
  );
};

export const CalendarWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
    <NumberSetting context={context} settingKey="days" label="Days in upcoming" fallback={DEFAULT_DAYS} {...DAYS_BOUNDS} />
    <NumberSetting context={context} settingKey="maxItems" label="Events to show"
      fallback={DEFAULT_MAX_ITEMS} {...MAX_ITEMS_BOUNDS} />
  </SettingsSurface>
);
