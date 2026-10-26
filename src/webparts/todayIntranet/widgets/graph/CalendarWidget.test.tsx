/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */

jest.mock('../content/WidgetCollections.module.scss', () => ({
  collection: 'collection', agenda: 'agenda', datePrimary: 'datePrimary', dateSecondary: 'dateSecondary',
  actionButton: 'actionButton'
}), { virtual: true });
jest.mock('../content/WidgetContent.module.scss', () => ({ itemCompact: 'compact', cardsGrid: 'gallery' }), { virtual: true });
jest.mock('../content/WidgetAdaptiveCard', () => ({ WidgetAdaptiveCard: jest.fn(() => null) }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { Event } from '@microsoft/microsoft-graph-types';
import { CalendarWidget, CalendarWidgetSettings } from './CalendarWidget';
import { IWidgetContext } from '../IWidget';
import { clearWidgetDataCache } from '../data/WidgetDataCache';
import { WidgetAdaptiveCard } from '../content/WidgetAdaptiveCard';

type GraphRequest = ReturnType<MSGraphClientV3['api']>;
const request = new (jest.fn<GraphRequest, []>())();
const get = jest.fn<Promise<{ value: Event[] }>, []>();
const query = jest.fn<GraphRequest, [Record<string, string>]>(() => request);
Object.assign(request, { version: () => request, header: () => request, select: () => request, orderby: () => request, top: () => request, query, get });
const client = new (jest.fn<MSGraphClientV3, []>())();
Object.assign(client, { api: () => request });
const spContext = new (jest.fn<WebPartContext, []>())();
Object.defineProperties(spContext, {
  pageContext: { value: {
    cultureInfo: { currentUICultureName: 'en-GB' }, aadInfo: { tenantId: 'tenant', userId: 'user' },
    site: { id: 'site' }, web: { id: 'web' }, user: { loginName: 'test@example.test' }
  } },
  msGraphClientFactory: { value: { getClient: async () => client } }
});

function meeting(title: string, hour: number, minute: number = 0): Event {
  return {
    id: title, subject: title,
    start: { dateTime: new Date(2026, 8, 21, hour, minute).toISOString(), timeZone: 'UTC' },
    end: { dateTime: new Date(2026, 8, 21, hour, minute + 30).toISOString(), timeZone: 'UTC' },
    webLink: 'https://outlook.office.com/test', onlineMeeting: { joinUrl: 'https://teams.microsoft.com/test' }
  };
}

describe('calendar interactions', () => {
  let container: HTMLDivElement;
  let context: IWidgetContext;

  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => {
    jest.useFakeTimers();
    jest.setSystemTime(new Date(2026, 8, 21, 13));
    jest.clearAllMocks();
    clearWidgetDataCache();
    get.mockReset().mockResolvedValue({ value: [meeting('Later meeting', 13, 25)] });
    context = { instanceId: 'calendar', spContext, settings: {}, isEditing: false, refreshToken: 0, updateSettings: jest.fn() };
    container = document.createElement('div');
    document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    clearWidgetDataCache();
    jest.clearAllTimers();
    jest.useRealTimers();
  });
  const render = async (): Promise<void> => {
    await act(async () => { ReactDom.render(<CalendarWidget context={context} />, container); });
  };
  const button = (text: string): HTMLButtonElement => {
    const found = Array.from(container.querySelectorAll('button')).find((candidate) => candidate.textContent?.includes(text));
    if (!found) { throw new Error(`Missing button: ${text}`); }
    return found;
  };
  const click = async (text: string): Promise<void> => {
    await act(async () => { Simulate.click(button(text)); });
  };
  const dateInput = (): HTMLInputElement => {
    const input = container.querySelector<HTMLInputElement>('input[type="date"]');
    if (!input) { throw new Error('Missing native date input.'); }
    return input;
  };
  const chooseDate = async (value: string): Promise<void> => {
    await act(async () => { dateInput().value = value; Simulate.change(dateInput()); });
  };

  it('keeps Today, Tomorrow, custom date and the Graph query aligned without persisting navigation', async () => {
    await render();
    expect(dateInput().value).toBe('2026-09-21');
    await click('Tomorrow');
    expect(dateInput().value).toBe('2026-09-22');
    expect(new Date(query.mock.calls[1][0].startDateTime).getDate()).toBe(22);
    await chooseDate('2026-09-24');
    expect(dateInput().value).toBe('2026-09-24');
    await click('Today');
    expect(dateInput().value).toBe('2026-09-21');
    expect(context.updateSettings).not.toHaveBeenCalled();
  });

  it('announces invalid date input instead of querying an invented date', async () => {
    await render();
    await chooseDate('');
    expect(container.querySelector('[role="alert"]')?.textContent).toBe('Choose a valid calendar date.');
    expect(dateInput().getAttribute('aria-invalid')).toBe('true');
    expect(get).toHaveBeenCalledTimes(1);
    await click('Tomorrow');
    expect(container.querySelector('[role="alert"]')).toBeNull();
  });

  it('preserves keyboard focus when changing ranges and custom dates', async () => {
    await render();
    const tomorrow = button('Tomorrow');
    tomorrow.focus();
    await click('Tomorrow');
    expect(document.activeElement).toBe(tomorrow);
    const input = dateInput();
    input.focus();
    await chooseDate('2026-09-24');
    expect(document.activeElement).toBe(input);
    expect(input.value).toBe('2026-09-24');
  });

  it('updates countdown without requerying each minute, then shows Now and removes completed events', async () => {
    await render();
    expect(container.querySelector('.datePrimary')?.textContent).toBe('25');
    await act(async () => { jest.advanceTimersByTime(60000); });
    expect(container.querySelector('.datePrimary')?.textContent).toBe('24');
    expect(get).toHaveBeenCalledTimes(1);
    await act(async () => { jest.advanceTimersByTime(24 * 60000); });
    expect(container.querySelector('.datePrimary')?.textContent).toBe('Now');
    await act(async () => { jest.advanceTimersByTime(30 * 60000); });
    expect(container.textContent).toContain('No meetings scheduled');
    expect(container.textContent).toContain("Looks like you don't have any meetings for today.");
  });

  it('renders the reference badge lines for Now, minutes, Today and Tomorrow in the agenda', async () => {
    const tomorrow = new Date(2026, 8, 22, 10);
    get.mockResolvedValue({ value: [
      meeting('In progress', 13),
      meeting('Starting soon', 13, 15),
      meeting('Later today', 15),
      { ...meeting('Tomorrow meeting', 10),
        start: { dateTime: tomorrow.toISOString(), timeZone: 'UTC' },
        end: { dateTime: new Date(tomorrow.getTime() + 30 * 60000).toISOString(), timeZone: 'UTC' } }
    ] });
    await render();
    await click('Upcoming');
    const rows = Array.from(container.querySelectorAll('ol[aria-label="Calendar events"] > li'));
    expect(rows.map((row) => row.querySelector('.datePrimary')?.textContent))
      .toEqual(['Now', '15', 'Today', 'Tue 22']);
    expect(rows.map((row) => row.querySelector('.dateSecondary')?.textContent))
      .toEqual([undefined, 'minutes', 'Mon 21', 'Tomorrow']);
  });

  it('updates day-dependent queries after midnight', async () => {
    jest.setSystemTime(new Date(2026, 8, 21, 23, 59, 50));
    await render();
    await click('Tomorrow');
    await act(async () => { jest.advanceTimersByTime(15000); });
    expect(dateInput().value).toBe('2026-09-23');
    const latest = query.mock.calls[query.mock.calls.length - 1][0];
    expect(new Date(latest.startDateTime).getDate()).toBe(23);
  });

  it('supplies an accessible Join name independently of the visible caption', async () => {
    await render();
    expect(container.querySelector('a[aria-label="Join Later meeting"]')?.getAttribute('href'))
      .toBe('https://teams.microsoft.com/test');
  });

  it('does not show the previous date under a new date while loading', async () => {
    await render();
    let resolvePending!: (value: { value: Event[] }) => void;
    get.mockReturnValueOnce(new Promise((resolve) => { resolvePending = resolve; }));
    await click('Tomorrow');
    expect(container.textContent).not.toContain('Later meeting');
    expect(container.querySelector('[aria-label="Loading your calendar..."]')).not.toBeNull();
    await act(async () => { resolvePending({ value: [] }); });
    expect(container.textContent).toContain("Looks like you don't have any meetings for tomorrow.");
  });

  it('keeps same-range data during sync and reports transient errors', async () => {
    await render();
    let rejectPending!: (error: unknown) => void;
    get.mockReturnValueOnce(new Promise((_resolve, reject) => { rejectPending = reject; }));
    await click('Sync calendar');
    expect(container.textContent).toContain('Later meeting');
    expect(button('Syncing').disabled).toBe(true);
    await act(async () => { rejectPending({ statusCode: 503 }); });
    expect(container.textContent).toContain('Could not refresh.');
    expect(container.textContent).toContain('Later meeting');
  });

  it('retains stale data when an automatic refresh fails and recovers on the next refresh', async () => {
    await render();
    get.mockRejectedValueOnce({ statusCode: 503 });
    await act(async () => { jest.advanceTimersByTime(5 * 60000); });
    expect(get).toHaveBeenCalledTimes(2);
    expect(container.textContent).toContain('Could not refresh.');
    expect(container.textContent).toContain('Later meeting');
    get.mockResolvedValueOnce({ value: [meeting('Updated meeting', 14)] });
    await act(async () => { jest.advanceTimersByTime(5 * 60000); });
    expect(get).toHaveBeenCalledTimes(3);
    expect(container.textContent).not.toContain('Could not refresh.');
    expect(container.textContent).toContain('Updated meeting');
  });

  it('updates the clock and query on focus after sleeping across midnight', async () => {
    await render();
    jest.setSystemTime(new Date(2026, 8, 22, 8));
    await act(async () => { window.dispatchEvent(new Event('focus')); });
    expect(dateInput().value).toBe('2026-09-22');
    expect(get).toHaveBeenCalledTimes(2);
  });

  it('renders permission errors rather than retaining inaccessible calendar data', async () => {
    await render();
    get.mockRejectedValueOnce({ statusCode: 403 });
    await click('Sync calendar');
    act(() => { jest.advanceTimersByTime(200); });
    expect(container.textContent).toContain('Calendars.ReadBasic');
    expect(container.textContent).not.toContain('Later meeting');
  });

  it.each(['list', 'compact', 'cards', 'gallery', 'adaptive'])('honors a previously saved %s view', async (view) => {
    context.settings = { view };
    await render();
    expect(container.querySelector('.agenda')).toBeNull();
    if (view === 'adaptive') {
      expect(WidgetAdaptiveCard).toHaveBeenCalled();
    } else {
      expect(container.textContent).toContain('Later meeting');
    }
  });

  it.each([[1e200, 30], [-5, 1], [2.9, 2], [undefined, 7]])('bounds upcoming days (%p)', async (days, expected) => {
    context.settings = { days };
    await render();
    await click('Upcoming');
    const latest = query.mock.calls[query.mock.calls.length - 1][0];
    const end = new Date(latest.startDateTime);
    end.setDate(end.getDate() + expected);
    expect(latest.endDateTime).toBe(end.toISOString());
  });

  it.each([[1e200, 6], [-5, 3], [3.8, 3], [undefined, 5]])('migrates the old display limit to bounded page size (%p)', async (maxItems, expected) => {
    context.settings = { maxItems };
    get.mockResolvedValue({ value: Array.from({ length: 25 }, (_, index) => meeting(`Event ${index}`, 14)) });
    await render();
    expect(container.querySelectorAll('ol > li')).toHaveLength(expected);
  });

  it('pages every loaded meeting and shows all of them with pagination disabled', async () => {
    context.settings = { maxItems: 1, itemsPerPage: 3 };
    get.mockResolvedValue({ value: Array.from({ length: 8 }, (_, index) => meeting(`Event ${index}`, 14)) });
    await render();
    expect(container.querySelectorAll('ol > li')).toHaveLength(3);
    const next = container.querySelector<HTMLButtonElement>('[aria-label="Next calendar events page"]');
    if (!next) { throw new Error('Missing next page button'); }
    await act(async () => { Simulate.click(next); });
    expect(container.textContent).toContain('Event 3');
    expect(container.textContent).not.toContain('Event 0');
    context.settings = { ...context.settings, paginationEnabled: false };
    await render();
    expect(container.querySelectorAll('ol > li')).toHaveLength(8);
    expect(container.querySelector('nav')).toBeNull();
    expect(get).toHaveBeenCalledTimes(1);
  });

  it('changes time format without refetching events', async () => {
    context.settings = { timeFormat: '24' };
    await render();
    expect(container.textContent).toContain('13:25');
    context.settings = { timeFormat: '12' };
    await render();
    expect(container.textContent).toContain('1:25 pm');
    expect(get).toHaveBeenCalledTimes(1);
  });

  it('renders a contextual empty notice for the chosen period', async () => {
    get.mockResolvedValue({ value: [] });
    await render();
    expect(container.textContent).toContain("Looks like you don't have any meetings for today.");
    await click('Tomorrow');
    expect(container.textContent).toContain("Looks like you don't have any meetings for tomorrow.");
    await click('Upcoming');
    expect(container.textContent).toContain('the next 7 days');
    await chooseDate('2026-09-24');
    expect(container.textContent).toContain('Thursday, 24 September 2026');
  });

  it('offers a bounded slider and locale-default time format without losing other settings', async () => {
    context.settings = { itemsPerPage: 4, unrelated: 'preserve' };
    await act(async () => { ReactDom.render(<CalendarWidgetSettings context={context} />, container); });
    const slider = container.querySelector<HTMLElement>('[role="slider"]');
    if (!slider) { throw new Error('Missing meeting count slider'); }
    expect(slider.getAttribute('aria-valuemin')).toBe('3');
    expect(slider.getAttribute('aria-valuemax')).toBe('6');
    expect(slider.getAttribute('aria-valuenow')).toBe('4');
    act(() => { Simulate.keyDown(slider, { which: 39, keyCode: 39, key: 'ArrowRight' }); });
    expect(context.updateSettings).toHaveBeenCalledWith({ ...context.settings, itemsPerPage: 5 });
    const radio = container.querySelector<HTMLInputElement>('input[type="radio"][id$="-12"]');
    const selected = container.querySelector<HTMLInputElement>('input[type="radio"][id$="-24"]');
    expect(selected?.checked).toBe(true);
    if (!radio) { throw new Error('Missing AM/PM option'); }
    act(() => { Simulate.change(radio); });
    expect(context.updateSettings).toHaveBeenCalledWith({ ...context.settings, timeFormat: '12' });
  });
});
