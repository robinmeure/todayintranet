import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Event } from '@microsoft/microsoft-graph-types';
import { calendarDateKey, calendarEventTime, calendarItems, calendarTimeFormat, calendarWindow, loadCalendarEvents, parseCalendarDate } from './calendarData';

describe('calendar date and item mapping', () => {
  const now = new Date(2026, 8, 21, 13);
  const event = (title: string, hour: number, extra?: Partial<Event>): Event => ({
    id: title, subject: title,
    start: { dateTime: new Date(2026, 8, 21, hour).toISOString(), timeZone: 'UTC' },
    end: { dateTime: new Date(2026, 8, 21, hour + 1).toISOString(), timeZone: 'UTC' },
    ...extra
  });

  it('defaults the clock to the locale while honoring an explicit selection', () => {
    expect(calendarTimeFormat(undefined, 'en-US')).toBe('12');
    expect(calendarTimeFormat(undefined, 'en-GB')).toBe('24');
    expect(calendarTimeFormat('invalid', 'nl-NL')).toBe('24');
    expect(calendarTimeFormat('24', 'en-US')).toBe('24');
    expect(calendarTimeFormat('12', 'en-GB')).toBe('12');
  });

  it.each(['12', '24'] as const)('formats meeting times with an explicit %s-hour clock', (format) => {
    const item = calendarItems([event('Meeting', 14)], new Date(2026, 8, 21), now, 5, 'en-GB', format)[0];
    expect(item.agendaMeta?.[0]).toBe(format === '12' ? '2:00 pm - 3:00 pm' : '14:00 - 15:00');
    expect(item.meta?.[0]).toContain(item.agendaMeta?.[0]);
  });

  it.each(['2026-02-30', '2026-13-01', '2026-00-01', '2026-09-00', '26-09-21', '', 'not-a-date'])(
    'rejects invalid or normalized dates (%s)', (value) => expect(parseCalendarDate(value)).toBeUndefined()
  );

  it('round trips a leap date', () => {
    const date = parseCalendarDate('2028-02-29');
    expect(date && calendarDateKey(date)).toBe('2028-02-29');
  });

  it.each([new Date(2026, 2, 29, 12), new Date(2026, 9, 25, 12)])(
    'uses calendar days rather than 24-hour offsets across daylight saving boundaries', (date) => {
      const { start, end } = calendarWindow(date, 'today', '', 7);
      expect(start.getHours()).toBe(0);
      expect(end.getHours()).toBe(0);
      const expected = new Date(start);
      expected.setDate(expected.getDate() + 1);
      expect(end).toEqual(expected);
    }
  );

  it('interprets Graph UTC strings without an offset as UTC', () => {
    expect(calendarEventTime({ dateTime: '2026-09-21T10:30:00.0000000', timeZone: 'UTC' })?.toISOString())
      .toBe('2026-09-21T10:30:00.000Z');
    expect(calendarEventTime({ dateTime: '2026-09-21T12:30:00+02:00', timeZone: 'UTC' })?.toISOString())
      .toBe('2026-09-21T10:30:00.000Z');
  });

  it('filters cancelled and completed events before taking the display limit', () => {
    const events = [8, 9, 10, 11, 12].map((hour) => event(`Past ${hour}`, hour));
    events.push(event('Cancelled', 14, { isCancelled: true }), event('Later', 15), event('Now', 13));
    const items = calendarItems(events, calendarWindow(now, 'today', '', 7).start, now, 2, 'en-GB');
    expect(items.map((item) => item.title)).toEqual(['Now', 'Later']);
    expect(items[0].leadingLabel?.text).toBe('Now');
  });

  it('keeps historical appointments when browsing a past date', () => {
    const items = calendarItems([event('History', 8)], new Date(2026, 8, 21), new Date(2026, 8, 22), 5, 'en-GB');
    expect(items[0].title).toBe('History');
  });

  it('does not label all-day events as a live meeting', () => {
    const items = calendarItems([event('All day', 0, {
      isAllDay: true, end: { dateTime: new Date(2026, 8, 22).toISOString(), timeZone: 'UTC' }
    })], new Date(2026, 8, 21), now, 5, 'en-GB');
    expect(items[0].leadingLabel?.text).toBe('All day');
    expect(items[0].badge).toBeUndefined();
  });

  it.each<[number, number, string, string | undefined]>([
    [12, 59, 'Now', undefined],
    [13, 1, '1', 'minute'],
    [13, 15, '15', 'minutes'],
    [13, 59, '59', 'minutes'],
    [14, 0, 'Today', 'Mon 21'],
    [15, 0, 'Today', 'Mon 21']
  ])('uses the reference badge hierarchy for a meeting at %s:%s', (hour, minute, text, description) => {
    const start = new Date(2026, 8, 21, hour, minute);
    const items = calendarItems([event('Meeting', 13, {
      start: { dateTime: start.toISOString(), timeZone: 'UTC' },
      end: { dateTime: new Date(start.getTime() + 30 * 60000).toISOString(), timeZone: 'UTC' }
    })], new Date(2026, 8, 21), now, 5, 'en-GB');
    expect(items[0].leadingLabel).toMatchObject({ text });
    expect(items[0].leadingLabel?.description).toBe(description);
  });

  it.each<[Date, string, string]>([
    [new Date(2026, 8, 22, 10), 'Tue 22', 'Tomorrow'],
    [new Date(2026, 8, 23, 10), 'Wed 23', 'Sept'],
    [new Date(2026, 9, 1, 10), 'Thu 1', 'Oct']
  ])('puts the weekday and date above the relative day or month (%s)', (start, text, description) => {
    const items = calendarItems([event('Future', 14, {
      start: { dateTime: start.toISOString(), timeZone: 'UTC' },
      end: { dateTime: new Date(start.getTime() + 30 * 60000).toISOString(), timeZone: 'UTC' }
    })], new Date(2026, 8, 21), now, 5, 'en-GB');
    expect(items[0].leadingLabel).toMatchObject({ text, description });
  });

  it('omits unavailable actions and labels online actions with the meeting title', () => {
    const items = calendarItems([
      event('Offline', 14),
      event('Online', 15, { onlineMeeting: { joinUrl: 'https://teams.microsoft.com/test' } })
    ], new Date(2026, 8, 21), now, 5, 'en-GB');
    expect(items[0].actions).toEqual([]);
    expect(items[1].actions?.[0]).toMatchObject({ text: 'Join', ariaLabel: 'Join Online' });
  });
});

describe('calendar Graph paging', () => {
  type GraphRequest = ReturnType<MSGraphClientV3['api']>;
  const request = new (jest.fn<GraphRequest, []>())();
  const get = jest.fn<Promise<{ value: Event[]; '@odata.nextLink'?: string }>, []>();
  const header = jest.fn<GraphRequest, [string, string]>(() => request);
  const query = jest.fn<GraphRequest, [Record<string, string>]>(() => request);
  const api = jest.fn<GraphRequest, [string]>(() => request);
  const client = new (jest.fn<MSGraphClientV3, []>())();
  Object.assign(request, {
    version: () => request, header, query, select: () => request, orderby: () => request, top: () => request, get
  });
  Object.assign(client, { api });

  beforeEach(() => {
    jest.clearAllMocks();
    get.mockReset();
  });

  it('follows nextLink and retains later meetings after a full cancelled page', async () => {
    get.mockResolvedValueOnce({ value: [{ id: 'cancelled', isCancelled: true }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/calendarView?skip=1' })
      .mockResolvedValueOnce({ value: [{ id: 'later' }] });
    expect(await loadCalendarEvents(client, 'start', 'end')).toEqual([{ id: 'later' }]);
    expect(api.mock.calls[1][0]).toContain('skip=1');
    expect(header).toHaveBeenNthCalledWith(2, 'Prefer', 'outlook.timezone="UTC"');
    expect(query).toHaveBeenCalledTimes(1);
  });

  it('stops following nextLink once enough relevant meetings are collected', async () => {
    const now = new Date('2026-09-21T10:00:00Z');
    const ended = { id: 'ended', end: { dateTime: '2026-09-21T09:00:00', timeZone: 'UTC' } };
    const upcoming = (id: string): Event => ({ id, end: { dateTime: '2026-09-21T12:00:00', timeZone: 'UTC' } });
    get.mockResolvedValueOnce({ value: [ended, upcoming('a')], '@odata.nextLink': 'page-2' })
      .mockResolvedValueOnce({ value: [upcoming('b'), upcoming('c')], '@odata.nextLink': 'page-3' })
      .mockResolvedValueOnce({ value: [upcoming('d')] });
    const events = await loadCalendarEvents(client, '2026-09-21T00:00:00Z', '2026-09-28T00:00:00Z', 3, now);
    expect(events.map((item) => item.id)).toEqual(['ended', 'a', 'b', 'c']);
    expect(get).toHaveBeenCalledTimes(2);
  });

  it('fails explicitly rather than looping on a repeated page', async () => {
    get.mockResolvedValue({ value: [], '@odata.nextLink': 'repeat' });
    await expect(loadCalendarEvents(client, 'start', 'end')).rejects.toThrow('repeated calendar page');
    expect(get).toHaveBeenCalledTimes(2);
  });
});
