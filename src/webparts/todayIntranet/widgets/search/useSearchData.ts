import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import { IWidgetDataState } from '../content';
import { ISearchRequest, ISearchResults, toWidgetError } from './SearchService';
import { getSearchData, searchDataKey } from './cachedSearchData';

export interface ISearchDataResult extends IWidgetDataState<ISearchResults> {
  reload(): void;
}

/**
 * The search equivalent of `useGraphData`: same `IWidgetDataState` shape, so search
 * backed widgets get the skeleton, refresh, empty and error handling for free.
 *
 * The request is re-run whenever any part of it, or the user/site identity, changes.
 * Callers can build it inline on every render.
 */
export function useSearchData(
  context: IWidgetContext,
  request: ISearchRequest
): ISearchDataResult {
  const spContext = context.spContext;
  const refreshToken = context.refreshToken;
  const requestKey = searchDataKey(context, request);
  const [result, setResult] = React.useState<IWidgetDataState<ISearchResults>>({ status: 'loading' });
  const [reloadToken, setReloadToken] = React.useState<number>(0);
  const previousReloadToken = React.useRef<number>(reloadToken);
  const previousRefreshToken = React.useRef<number>(refreshToken);

  const requestRef = React.useRef(request);
  requestRef.current = request;

  React.useEffect(() => {
    let cancelled = false;
    const bypassCache =
      previousReloadToken.current !== reloadToken ||
      previousRefreshToken.current !== refreshToken;
    previousReloadToken.current = reloadToken;
    previousRefreshToken.current = refreshToken;

    // An unconfigured tile is not an error, it simply has nothing to show yet.
    if (!requestRef.current.queryText.trim()) {
      setResult({ status: 'ready', data: { rows: [], totalRows: 0 } });
      return undefined;
    }

    setResult((previous) =>
      previous.status === 'ready' ? { ...previous, isRefreshing: true } : { status: 'loading' }
    );

    const activeRequest = requestRef.current;
    getSearchData({ spContext }, activeRequest, bypassCache)
      .then((response) => {
        if (!cancelled) {
          setResult({
            status: 'ready',
            data: response.data,
            lastUpdated: response.fetchedAt,
            refreshError: response.refreshError
              ? toWidgetError(response.refreshError)
              : undefined
          });
        }
      })
      .catch((error: unknown) => {
        if (!cancelled) {
          setResult({ status: 'error', error: toWidgetError(error) });
        }
      });

    return () => {
      cancelled = true;
    };
  }, [spContext, requestKey, reloadToken, refreshToken]);

  const reload = React.useCallback(() => setReloadToken((token) => token + 1), []);

  return { ...result, reload };
}
