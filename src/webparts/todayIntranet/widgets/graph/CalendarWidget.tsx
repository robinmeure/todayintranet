import * as React from 'react';
import { Event } from '@microsoft/microsoft-graph-types';
import { Icon } from '@fluentui/react/lib/Icon';
import { Slider } from '@fluentui/react/lib/Slider';
import { ChoiceGroup } from '@fluentui/react/lib/ChoiceGroup';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  NumberSetting, SettingsSurface, ToggleSetting, WidgetDateToolbar, WidgetItems, WidgetItemsView, WidgetView,
  booleanSetting, numberSetting, viewSetting, PAGE_WINDOW
} from '../content';
import {
  CalendarRange, calendarDateKey, calendarItems, calendarTimeFormat, calendarWindow, loadCalendarEvents, parseCalendarDate
} from './calendarData';
import styles from '../content/WidgetCollections.module.scss';

const DEFAULT_DAYS: number = 7;
const DEFAULT_MAX_ITEMS: number = 5;
const DAYS_BOUNDS = { min: 1, max: 30 };
const PAGE_SIZE_BOUNDS = { min: 3, max: 6 };
const CALENDAR_URL: string = 'https://outlook.office.com/calendar/';
const REFRESH_INTERVAL: number = 5 * 60 * 1000;

export const CALENDAR_LINK: IWidgetLink = { text: 'Open Outlook', href: CALENDAR_URL };
export const CALENDAR_DEFAULT_VIEW: WidgetItemsView = 'agenda';

function meetingsPerPage(context: IWidgetContext): number {
  const legacySize = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, PAGE_SIZE_BOUNDS);
  return numberSetting(context, 'itemsPerPage', legacySize, PAGE_SIZE_BOUNDS);
}

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
  pageSize: number;
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
  context, now, range, start, end, days, pageSize, locale, dateError, onRangeChange, onDateChange
}) => {
  const refreshWindow = Math.floor(now.getTime() / REFRESH_INTERVAL);
  const previousRefreshWindow = React.useRef(refreshWindow);
  const limit = booleanSetting(context, 'paginationEnabled', true) ? pageSize * PAGE_WINDOW : undefined;
  const state = useGraphData<ICalendarResult>(
    context,
    'Calendars.ReadBasic',
    async (client) => ({ start, end, events: await loadCalendarEvents(client, start, end, limit) }),
    [start, end, limit]
  );
  const { reload } = state;
  React.useEffect(() => {
    if (previousRefreshWindow.current !== refreshWindow) {
      previousRefreshWindow.current = refreshWindow;
      reload();
    }
  }, [refreshWindow, reload]);

  // Keep the toolbar mounted for keyboard focus, but never label another range's data as this one.
  const hasPreviousRange = !!state.data && (state.data.start !== start || state.data.end !== end);
  const status = hasPreviousRange ? 'loading' : state.status;
  const items = state.data && !hasPreviousRange
    ? calendarItems(state.data.events, new Date(start), now, state.data.events.length, locale,
      calendarTimeFormat(context.settings.timeFormat, locale)) : undefined;
  const period = range === 'today' ? 'today' : range === 'tomorrow' ? 'tomorrow'
    : range === 'upcoming' ? `the next ${days} days`
      : new Date(start).toLocaleDateString(locale, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });

  return (
    <div className={styles.calendar}>
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
        loading={{ label: 'Loading your calendar...', rows: Math.min(pageSize, 4) }}
        empty={{
          iconName: 'Calendar',
          text: range === 'today' ? 'Nothing else scheduled today.' : range === 'tomorrow'
            ? 'Nothing scheduled tomorrow.' : range === 'upcoming'
              ? `Nothing else scheduled in the next ${days} days.` : 'Nothing scheduled on this date.'
        }}
        emptyContent={(
          <div className={styles.calendarEmpty} role="status">
            <Icon iconName="InfoSolid" aria-hidden={true} />
            <div>
              <strong>No meetings scheduled</strong>
              <p>Looks like you don&apos;t have any meetings for <strong>{period}</strong>.</p>
            </div>
          </div>
        )}
      >
        {(events) => <WidgetItems key={`${start}:${end}`}
          pageSize={booleanSetting(context, 'paginationEnabled', true) ? pageSize : undefined}
          view={viewSetting(context, CALENDAR_DEFAULT_VIEW)} items={events} ariaLabel="Calendar events" />}
      </WidgetView>
    </div>
  );
};

export const CalendarWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const days = numberSetting(context, 'days', DEFAULT_DAYS, DAYS_BOUNDS);
  const pageSize = meetingsPerPage(context);
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
      pageSize={pageSize}
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

export const CalendarWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;
  return (
    <SettingsSurface description="Customize how your calendar is displayed and access your meetings for the day in a way that works best for you.">
      <Slider label="Number of meetings per page" min={3} max={6} step={1} showValue={true}
        disabled={!booleanSetting(context, 'paginationEnabled', true)}
        value={meetingsPerPage(context)}
        onChange={(value) => context.updateSettings({ ...context.settings, itemsPerPage: value })} />
      <ToggleSetting context={context} settingKey="paginationEnabled" label="Use pagination" fallback={true} />
      <ChoiceGroup label="Time format" selectedKey={calendarTimeFormat(context.settings.timeFormat, locale)}
        options={[{ key: '12', text: 'AM/PM' }, { key: '24', text: '24 hour' }]}
        onChange={(_, option) => {
          if (option) { context.updateSettings({ ...context.settings, timeFormat: option.key }); }
        }} />
      <NumberSetting context={context} settingKey="days" label="Days in upcoming" fallback={DEFAULT_DAYS} {...DAYS_BOUNDS} />
    </SettingsSurface>
  );
};
