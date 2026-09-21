import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Event } from '@microsoft/microsoft-graph-types';
import { IWidgetListItem } from '../content/IWidgetContent';

export type CalendarRange = 'today' | 'tomorrow' | 'upcoming' | 'date';

export function calendarDateKey(date: Date): string {
  return `${date.getFullYear()}-${`0${date.getMonth() + 1}`.slice(-2)}-${`0${date.getDate()}`.slice(-2)}`;
}

export function parseCalendarDate(value: string): Date | undefined {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return undefined;
  }
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(0);
  date.setFullYear(year, month - 1, day);
  date.setHours(0, 0, 0, 0);
  return calendarDateKey(date) === value ? date : undefined;
}

function addDays(date: Date, days: number): Date {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

export function calendarWindow(now: Date, range: CalendarRange, customDate: string, days: number): { start: Date; end: Date } {
  const start = range === 'date' ? parseCalendarDate(customDate) : new Date(now);
  if (!start) {
    throw new Error('Choose a valid calendar date.');
  }
  start.setHours(0, 0, 0, 0);
  const selected = range === 'tomorrow' ? addDays(start, 1) : start;
  return { start: selected, end: addDays(selected, range === 'upcoming' ? days : 1) };
}

export function calendarEventTime(value: Event['start']): Date | undefined {
  if (!value?.dateTime) {
    return undefined;
  }
  // Graph's UTC preference still returns dateTime without a trailing Z.
  const raw = value.timeZone === 'UTC' && !/(Z|[+-]\d{2}:\d{2})$/i.test(value.dateTime)
    ? `${value.dateTime}Z` : value.dateTime;
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? undefined : date;
}

interface ICalendarPage {
  value: Event[];
  '@odata.nextLink'?: string;
}

export async function loadCalendarEvents(client: MSGraphClientV3, start: string, end: string): Promise<Event[]> {
  const events: Event[] = [];
  const visited = new Set<string>();
  let next: string | undefined;
  do {
    let request = client.api(next ?? '/me/calendarView').version('v1.0').header('Prefer', 'outlook.timezone="UTC"');
    if (!next) {
      request = request.query({ startDateTime: start, endDateTime: end })
        .select('id,subject,start,end,isAllDay,location,webLink,isCancelled,onlineMeeting,onlineMeetingUrl')
        .orderby('start/dateTime').top(100);
    }
    const page: ICalendarPage = await request.get();
    if (!page || !Array.isArray(page.value)) {
      throw new Error('Microsoft 365 returned an invalid calendar response.');
    }
    events.push(...page.value.filter((event) => !event.isCancelled));
    next = page['@odata.nextLink'];
    if (next) {
      if (visited.has(next)) {
        throw new Error('Microsoft 365 returned a repeated calendar page.');
      }
      visited.add(next);
    }
  } while (next);
  return events;
}

function dayLabel(start: Date, now: Date, locale: string | undefined): string {
  const key = calendarDateKey(start);
  if (key === calendarDateKey(now)) {
    return 'Today';
  }
  if (key === calendarDateKey(addDays(now, 1))) {
    return 'Tomorrow';
  }
  return start.toLocaleDateString(locale, { month: 'short', day: 'numeric' });
}

export function calendarItems(
  events: Event[], rangeStart: Date, now: Date, maxItems: number, locale: string | undefined
): IWidgetListItem[] {
  const isHistory = calendarDateKey(rangeStart) < calendarDateKey(now);
  return events.filter((event) => {
    const end = calendarEventTime(event.end);
    return !event.isCancelled && (isHistory || !end || end > now);
  }).sort((first, second) =>
    (calendarEventTime(first.start)?.getTime() ?? Infinity) - (calendarEventTime(second.start)?.getTime() ?? Infinity)
  ).slice(0, maxItems).map((event, index) => {
    const title = event.subject || '(No subject)';
    const start = calendarEventTime(event.start);
    const end = calendarEventTime(event.end);
    const inProgress = !event.isAllDay && !!start && start <= now && !!end && end > now;
    const minutes = start ? Math.ceil((start.getTime() - now.getTime()) / 60000) : 0;
    const startsSoon = !event.isAllDay && minutes > 0 && minutes < 60;
    const day = start ? dayLabel(start, now, locale) : 'Date TBC';
    const timeOptions: Intl.DateTimeFormatOptions = { hour: 'numeric', minute: '2-digit' };
    const time = event.isAllDay ? 'All day' : start
      ? `${start.toLocaleTimeString(locale, timeOptions)}${end ? ` - ${end.toLocaleTimeString(locale, timeOptions)}` : ''}`
      : 'Time to be confirmed';
    const join = event.onlineMeeting?.joinUrl || event.onlineMeetingUrl;
    return {
      key: event.id ?? String(index),
      title,
      href: event.webLink ?? undefined,
      meta: [`${day} · ${time}`, event.location?.displayName ?? undefined],
      hasAccentBar: true,
      tone: inProgress ? 'success' : start && calendarDateKey(start) === calendarDateKey(now) ? 'accent' : 'neutral',
      isEmphasized: inProgress || startsSoon,
      badge: inProgress ? { text: 'Now', tone: 'success' } : undefined,
      leadingLabel: {
        text: inProgress ? 'Now' : event.isAllDay ? 'All day' : startsSoon ? String(minutes) : day,
        description: inProgress ? undefined : startsSoon ? (minutes === 1 ? 'Minute' : 'Minutes')
          : start?.toLocaleDateString(locale, { weekday: 'short', day: 'numeric' }),
        tone: inProgress ? 'success' : startsSoon ? 'accent' : 'neutral'
      },
      actions: [
        ...(join ? [{ text: 'Join', ariaLabel: `Join ${title}`, href: join, iconName: 'Video', isPrimary: true }] : []),
        ...(event.webLink ? [{
          text: 'Details', ariaLabel: `Open ${title} in Outlook`, href: event.webLink, iconName: 'Calendar'
        }] : [])
      ]
    };
  });
}
