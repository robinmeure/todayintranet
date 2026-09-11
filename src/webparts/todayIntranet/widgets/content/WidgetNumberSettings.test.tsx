jest.mock('./WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content', () => ({
  ...jest.requireActual<typeof import('./WidgetSettings')>('./WidgetSettings'),
  WidgetView: () => null
}));
jest.mock('../graph/useGraphData', () => ({ useGraphData: jest.fn() }));
jest.mock('../search/useSearchData', () => ({ useSearchData: jest.fn() }));
jest.mock('../search/SearchQuerySetting', () => ({ SearchQuerySetting: () => null }));
jest.mock('../search/SearchService', () => ({}));

import * as React from 'react';
import type { ISpinButtonProps } from '@fluentui/react/lib/SpinButton';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import { numberSetting, NumberSetting } from './WidgetSettings';
import { CalendarWidget } from '../graph/CalendarWidget';
import { MailWidget } from '../graph/MailWidget';
import { TasksWidget } from '../graph/TasksWidget';
import { NewsWidget } from '../search/NewsWidget';
import { SearchResultsWidget } from '../search/SearchResultsWidget';
import { useGraphData } from '../graph/useGraphData';
import { useSearchData } from '../search/useSearchData';

const spContext = new (jest.fn<WebPartContext, []>())();
Object.defineProperty(spContext, 'pageContext', {
  value: { cultureInfo: { currentUICultureName: 'en-US' } }
});

const contextFor = (settings: Record<string, unknown>): IWidgetContext => ({
  instanceId: 'bounded-widget',
  spContext,
  settings,
  isEditing: false,
  refreshToken: 0,
  updateSettings: jest.fn<void, [Record<string, unknown>]>()
});

describe('numberSetting persisted values', () => {
  it.each([undefined, null, '12', true, {}, [], NaN, Infinity, -Infinity])(
    'preserves the default for invalid numeric settings (%p)',
    (value) => {
      expect(numberSetting(contextFor({ count: value }), 'count', 6, { min: 1, max: 20 })).toBe(6);
    }
  );

  it.each([
    [-1e200, 1],
    [-3, 1],
    [0, 1],
    [1, 1],
    [3.9, 3],
    [20, 20],
    [20.9, 20],
    [1e200, 20]
  ])('clamps %p to an integer within the declared bounds (%p)', (value, expected) => {
    expect(numberSetting(contextFor({ count: value }), 'count', 6, { min: 1, max: 20 })).toBe(expected);
  });

  it('preserves the original API and finite numeric behavior without bounds', () => {
    expect(numberSetting(contextFor({ count: -4.75 }), 'count', 6)).toBe(-4.75);
    expect(numberSetting(contextFor({}), 'count', 6)).toBe(6);
    expect(numberSetting(contextFor({ count: Infinity }), 'count', 6)).toBe(6);
  });

  it('does not mutate or persist the stored settings while reading them', () => {
    const context = contextFor({ count: 1000 });
    expect(numberSetting(context, 'count', 6, { min: 1, max: 20 })).toBe(20);
    expect(context.settings.count).toBe(1000);
    expect(context.updateSettings).not.toHaveBeenCalled();
  });

  it('uses the same bounds in the numeric editor and clamps zero rather than using the default', () => {
    const context = contextFor({ count: 1000 });
    const control = NumberSetting({ context, settingKey: 'count', label: 'Count', fallback: 6, min: 1, max: 20 });
    if (!React.isValidElement<ISpinButtonProps>(control)) {
      throw new Error('Numeric editor did not return a control.');
    }
    expect(control.props.value).toBe('20');
    control.props.onValidate?.('0');
    expect(context.updateSettings).toHaveBeenLastCalledWith({ count: 1 });
    control.props.onValidate?.('invalid');
    expect(context.updateSettings).toHaveBeenLastCalledWith({ count: 6 });
    control.props.onDecrement?.('20');
    expect(context.updateSettings).toHaveBeenLastCalledWith({ count: 19 });
  });
});

type GraphRequest = ReturnType<MSGraphClientV3['api']>;
let fetchGraph: ((client: MSGraphClientV3) => Promise<unknown>) | undefined;

function graphClient(): {
  client: MSGraphClientV3;
  top: jest.Mock<GraphRequest, [number]>;
  query: jest.Mock<GraphRequest, [Record<string, string>]>;
} {
  const request = new (jest.fn<GraphRequest, []>())();
  const top = jest.fn<GraphRequest, [number]>(() => request);
  const query = jest.fn<GraphRequest, [Record<string, string>]>(() => request);
  Object.assign(request, {
    version: jest.fn<GraphRequest, [string]>(() => request),
    header: jest.fn<GraphRequest, [string, string]>(() => request),
    select: jest.fn<GraphRequest, [string]>(() => request),
    filter: jest.fn<GraphRequest, [string]>(() => request),
    orderby: jest.fn<GraphRequest, [string]>(() => request),
    top,
    query,
    get: jest.fn<Promise<{ value: Array<{ id: string }> }>, []>()
      .mockResolvedValueOnce({ value: [{ id: 'default-list' }] })
      .mockResolvedValue({ value: [] })
  });
  const client = new (jest.fn<MSGraphClientV3, []>())();
  Object.assign(client, { api: jest.fn<GraphRequest, [string]>(() => request) });
  return { client, top, query };
}

describe('widget request numeric bounds', () => {
  beforeEach(() => {
    fetchGraph = undefined;
    jest.mocked(useGraphData).mockImplementation((_context, _scope, fetcher) => {
      fetchGraph = fetcher;
      return { status: 'loading', reload: () => undefined };
    });
    jest.mocked(useSearchData).mockReturnValue({ status: 'loading', reload: () => undefined });
  });

  const request = async (): Promise<ReturnType<typeof graphClient>> => {
    if (!fetchGraph) {
      throw new Error('Graph fetcher was not registered.');
    }
    const graph = graphClient();
    await fetchGraph(graph.client);
    return graph;
  };

  describe.each([
    ['Mail', MailWidget, 6, 0],
    ['Calendar', CalendarWidget, 5, 5],
    ['Tasks', TasksWidget, 6, 10]
  ] as const)('%s Graph request', (_name, Widget, fallback, overfetch) => {
    it.each([1e200, -5, 3.8, undefined])('bounds a persisted count before calling Graph (%p)', async (maxItems) => {
      Widget({ context: contextFor({ maxItems }) });
      const graph = await request();
      const count = maxItems === undefined ? fallback : Math.max(1, Math.min(20, Math.floor(maxItems)));
      expect(graph.top).toHaveBeenCalledWith(count + overfetch);
    });
  });

  it.each([
    [1e200, 30],
    [-5, 1],
    [2.9, 2],
    [undefined, 7]
  ])('bounds calendar days before constructing ISO date query parameters (%p)', async (days, expected) => {
    CalendarWidget({ context: contextFor({ days }) });
    const graph = await request();
    const query = graph.query.mock.calls[0][0];
    const duration = Date.parse(query.endDateTime) - Date.parse(query.startDateTime);
    expect(duration).toBe(expected * 24 * 60 * 60 * 1000);
  });

  describe.each([
    ['News', NewsWidget, 4],
    ['Search results', SearchResultsWidget, 6]
  ] as const)('%s search request', (_name, Widget, fallback) => {
    it.each([1e200, -5, 3.8, undefined])('bounds a persisted count before building the search request (%p)', (maxItems) => {
      Widget({ context: contextFor({ maxItems, scope: 'everywhere', terms: 'example' }) });
      const count = maxItems === undefined ? fallback : Math.max(1, Math.min(20, Math.floor(maxItems)));
      expect(jest.mocked(useSearchData).mock.calls[0][1].rowLimit).toBe(count);
    });
  });
});
