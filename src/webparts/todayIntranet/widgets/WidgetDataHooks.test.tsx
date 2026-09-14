/* eslint-disable @rushstack/pair-react-dom-render-unmount -- The shared container is unmounted in afterEach. */

jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from './IWidget';
import { IWidgetDataState } from './content';
import { clearWidgetDataCache } from './data/WidgetDataCache';
import { useGraphData } from './graph/useGraphData';
import type { ISearchRequest, ISearchResults } from './search/SearchService';
import * as SearchService from './search/SearchService';
import { useSearchData } from './search/useSearchData';

function deferred<T>(): {
  promise: Promise<T>;
  resolve(value: T): void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((_resolve) => {
    resolve = _resolve;
  });
  return { promise, resolve };
}

const graphClient = new (jest.fn<MSGraphClientV3, []>())();
const getGraphClient = jest.fn<Promise<MSGraphClientV3>, [string]>().mockResolvedValue(graphClient);
const spContext = new (jest.fn<WebPartContext, []>())();
Object.defineProperties(spContext, {
  pageContext: {
    value: {
      aadInfo: { tenantId: 'tenant-one', userId: 'user-one' },
      site: { id: 'site-one' },
      web: { id: 'web-one', absoluteUrl: 'https://example.test/sites/one' },
      user: { loginName: 'user-one@example.com' }
    }
  },
  msGraphClientFactory: { value: { getClient: getGraphClient } }
});

function contextFor(instanceId: string, refreshToken: number): IWidgetContext {
  return {
    instanceId,
    settings: {},
    spContext,
    isEditing: false,
    refreshToken,
    updateSettings: jest.fn<void, [Record<string, unknown>]>()
  };
}

let graphState: IWidgetDataState<string[]>;
const GraphHost: React.FunctionComponent<{
  instanceId: string;
  refreshToken: number;
  value?: string;
  fetcher?(client: MSGraphClientV3): Promise<string[]>;
}> = ({ instanceId, refreshToken, value = 'message', fetcher }) => {
  graphState = useGraphData(
    contextFor(instanceId, refreshToken),
    'Mail.ReadBasic',
    fetcher ?? (async () => [value]),
    [6, value]
  );
  return <span data-status={graphState.status} />;
};

const searchRequest: ISearchRequest = {
  queryText: 'PromotedState:2',
  selectProperties: ['Title', 'Path'],
  rowLimit: 5
};

let searchState: IWidgetDataState<ISearchResults>;
const SearchHost: React.FunctionComponent<{
  instanceId: string;
  refreshToken: number;
}> = ({ instanceId, refreshToken }) => {
  searchState = useSearchData(
    contextFor(instanceId, refreshToken),
    searchRequest,
    [searchRequest.queryText, searchRequest.rowLimit]
  );
  return <span data-status={searchState.status} />;
};

describe('widget data hook caching', () => {
  let container: HTMLDivElement;
  let runSearchQuery: jest.SpyInstance<Promise<ISearchResults>, Parameters<typeof SearchService.runSearchQuery>>;

  beforeEach(() => {
    clearWidgetDataCache();
    getGraphClient.mockReset().mockResolvedValue(graphClient);
    runSearchQuery = jest.spyOn(SearchService, 'runSearchQuery');
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    clearWidgetDataCache();
    jest.restoreAllMocks();
  });

  it('shares one Graph request between duplicate widgets and bypasses the cache on refresh', async () => {
    const fetcher = jest.fn<Promise<string[]>, [MSGraphClientV3]>().mockResolvedValue(['message']);
    const render = async (refreshToken: number): Promise<void> => {
      await act(async () => {
        ReactDom.render(
          <>
            <GraphHost instanceId="mail-one" refreshToken={refreshToken} fetcher={fetcher} />
            <GraphHost instanceId="mail-two" refreshToken={refreshToken} fetcher={fetcher} />
          </>,
          container
        );
      });
    };

    await render(0);
    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(fetcher).toHaveBeenCalledTimes(1);

    await render(0);
    expect(fetcher).toHaveBeenCalledTimes(1);

    await render(1);
    expect(getGraphClient).toHaveBeenCalledTimes(2);
    expect(fetcher).toHaveBeenCalledTimes(2);
  });

  it('keeps an in-flight Graph fetcher bound to its original dependency key', async () => {
    const client = deferred<MSGraphClientV3>();
    getGraphClient.mockReturnValue(client.promise);

    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} value="all" />, container);
    });
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} value="unread" />, container);
    });
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} value="all" />, container);
    });
    await act(async () => {
      client.resolve(graphClient);
    });

    expect(graphState.status).toBe('ready');
    expect(graphState.data).toEqual(['all']);
    expect(getGraphClient).toHaveBeenCalledTimes(2);
  });

  it('surfaces status-less access denial and invalidates the cached Graph result', async () => {
    const fetcher = jest.fn<Promise<string[]>, [MSGraphClientV3]>()
      .mockResolvedValueOnce(['cached mail'])
      .mockRejectedValue({ code: 'accessDenied' })
      .mockRejectedValue({ code: 'accessDenied' });

    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} fetcher={fetcher} />, container);
    });
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={1} fetcher={fetcher} />, container);
    });

    expect(graphState.status).toBe('error');
    expect(graphState.error?.isActionRequired).toBe(true);

    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} fetcher={fetcher} />, container);
    });

    expect(fetcher).toHaveBeenCalledTimes(3);
    expect(graphState.status).toBe('error');
  });

  it('retains Graph data with timestamp and refresh error after a transient failure', async () => {
    const fetcher = jest.fn<Promise<string[]>, [MSGraphClientV3]>()
      .mockResolvedValueOnce(['cached mail'])
      .mockRejectedValue({ statusCode: 503 });

    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} fetcher={fetcher} />, container);
    });
    const fetchedAt = graphState.lastUpdated;
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={1} fetcher={fetcher} />, container);
    });

    expect(graphState).toMatchObject({
      status: 'ready',
      data: ['cached mail'],
      lastUpdated: fetchedAt,
      refreshError: { message: 'Microsoft 365 is busy right now. Try again in a moment.' }
    });

    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    await act(async () => {
      ReactDom.render(<GraphHost instanceId="mail" refreshToken={0} fetcher={fetcher} />, container);
    });
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(graphState.refreshError?.message).toBe('Microsoft 365 is busy right now. Try again in a moment.');
  });

  it('shares one Search request between duplicate widgets and bypasses the cache on refresh', async () => {
    runSearchQuery.mockResolvedValue({ rows: [], totalRows: 0 });
    const render = async (refreshToken: number): Promise<void> => {
      await act(async () => {
        ReactDom.render(
          <>
            <SearchHost instanceId="news-one" refreshToken={refreshToken} />
            <SearchHost instanceId="news-two" refreshToken={refreshToken} />
          </>,
          container
        );
      });
    };

    await render(0);
    expect(runSearchQuery).toHaveBeenCalledTimes(1);

    await render(0);
    expect(runSearchQuery).toHaveBeenCalledTimes(1);

    await render(1);
    expect(runSearchQuery).toHaveBeenCalledTimes(2);
  });

  it('retains Search data only for recognized transient failures', async () => {
    const cached = { rows: [{ Title: 'Cached result' }], totalRows: 1 };
    runSearchQuery
      .mockResolvedValueOnce(cached)
      .mockRejectedValueOnce(new SearchService.SearchError({ message: 'Busy' }, 503))
      .mockRejectedValueOnce(new SearchService.SearchError({ message: 'Bad query' }, 500));

    await act(async () => {
      ReactDom.render(<SearchHost instanceId="search" refreshToken={0} />, container);
    });
    await act(async () => {
      ReactDom.render(<SearchHost instanceId="search" refreshToken={1} />, container);
    });
    expect(searchState).toMatchObject({
      status: 'ready',
      data: cached,
      refreshError: { message: 'Busy' }
    });

    await act(async () => {
      ReactDom.render(<SearchHost instanceId="search" refreshToken={2} />, container);
    });
    expect(searchState).toMatchObject({
      status: 'error',
      error: { message: 'Bad query' }
    });
  });
});
