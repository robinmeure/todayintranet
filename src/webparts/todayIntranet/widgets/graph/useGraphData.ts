import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { IWidgetContext } from '../IWidget';
import { IWidgetDataState, IWidgetError } from '../content';
import { getCachedWidgetData, widgetDataCacheKey } from '../data/WidgetDataCache';
import {
  classifyGraphDataError,
  isWidgetDataAuthorizationError,
  widgetDataErrorCodes,
  widgetDataErrorStatus
} from '../data/WidgetDataError';

/** What `useGraphData` returns: the shared data state, always with a retry. */
export interface IGraphDataResult<T> extends IWidgetDataState<T> {
  reload(): void;
}

interface IGraphErrorShape {
  code?: string;
  message?: string;
}

const DEFAULT_GRAPH_TTL_MS: number = 5 * 60 * 1000;
const GRAPH_TTL_BY_SCOPE: Record<string, number> = {
  'Mail.ReadBasic': 3 * 60 * 1000,
  'Calendars.ReadBasic': 5 * 60 * 1000,
  'Tasks.Read': 10 * 60 * 1000
};

function toGraphError(error: unknown, scope: string): IWidgetError {
  const shape = (error ?? {}) as IGraphErrorShape;
  const status = widgetDataErrorStatus(error);
  const codes = widgetDataErrorCodes(error);

  if (isWidgetDataAuthorizationError(error)) {
    const isAuthenticationFailure =
      status === 401 ||
      codes.indexOf('invalidauthenticationtoken') >= 0 ||
      codes.indexOf('unauthenticated') >= 0;
    return {
      isActionRequired: !isAuthenticationFailure,
      message: isAuthenticationFailure
        ? 'Your Microsoft 365 session could not be validated. Refresh the page or sign in again.'
        : `This widget needs the Microsoft Graph "${scope}" permission. ` +
          'A tenant administrator has to approve it in the SharePoint admin center under Advanced > API access.'
    };
  }
  if (status === 429 || status === 500 || status === 502 || status === 503 || status === 504) {
    return { message: 'Microsoft 365 is busy right now. Try again in a moment.' };
  }
  return { message: shape.message || 'Could not load data from Microsoft 365.' };
}

/**
 * Runs a Microsoft Graph query for a widget and exposes it as an `IWidgetDataState`,
 * so `WidgetView` can render the loading, error, empty and content cases the same way
 * it does for every other widget. `scope` is only used to explain which permission is
 * missing when consent fails.
 */
export function useGraphData<T>(
  context: IWidgetContext,
  scope: string,
  fetcher: (client: MSGraphClientV3) => Promise<T>,
  deps: unknown[]
): IGraphDataResult<T> {
  const spContext = context.spContext;
  const refreshToken = context.refreshToken;
  const [result, setResult] = React.useState<IWidgetDataState<T>>({ status: 'loading' });
  const [reloadToken, setReloadToken] = React.useState<number>(0);
  const previousReloadToken = React.useRef<number>(reloadToken);
  const previousRefreshToken = React.useRef<number>(refreshToken);

  // Keep the latest fetcher out of the effect dependencies so callers can pass
  // an inline arrow function without re-querying on every render.
  const fetcherRef = React.useRef(fetcher);
  fetcherRef.current = fetcher;

  React.useEffect(() => {
    let cancelled = false;
    const bypassCache =
      previousReloadToken.current !== reloadToken ||
      previousRefreshToken.current !== refreshToken;
    previousReloadToken.current = reloadToken;
    previousRefreshToken.current = refreshToken;
    const requestFetcher = fetcherRef.current;
    const requestKey = widgetDataCacheKey(context, `graph:${scope}`, deps);

    // Data already on screen stays there and is marked stale: re-reading after a
    // refresh or a settings change should not blank a tile the user is reading.
    setResult((previous) =>
      previous.status === 'ready' ? { ...previous, isRefreshing: true } : { status: 'loading' }
    );

    getCachedWidgetData({
      key: requestKey,
      ttlMilliseconds: GRAPH_TTL_BY_SCOPE[scope] ?? DEFAULT_GRAPH_TTL_MS,
      bypassCache,
      classifyError: classifyGraphDataError,
      load: () =>
        spContext.msGraphClientFactory
          .getClient('3')
          .then((client) => requestFetcher(client))
    })
      .then((response) => {
        if (!cancelled) {
          setResult({
            status: 'ready',
            data: response.data,
            lastUpdated: response.fetchedAt,
            refreshError: response.refreshError
              ? toGraphError(response.refreshError, scope)
              : undefined
          });
        }
      })
      .catch((error: unknown) => {
        if (!cancelled) {
          setResult({ status: 'error', error: toGraphError(error, scope) });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [spContext, scope, reloadToken, refreshToken, ...deps]);

  const reload = React.useCallback(() => setReloadToken((token) => token + 1), []);

  return { ...result, reload };
}
