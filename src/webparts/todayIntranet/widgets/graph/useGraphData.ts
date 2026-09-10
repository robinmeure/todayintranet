import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IWidgetContext } from '../IWidget';
import { IWidgetDataState, IWidgetError } from '../content';

/** What `useGraphData` returns: the shared data state, always with a retry. */
export interface IGraphDataResult<T> extends IWidgetDataState<T> {
  reload(): void;
}

interface IGraphErrorShape {
  statusCode?: number;
  code?: string;
  message?: string;
}

function toGraphError(error: unknown, scope: string): IWidgetError {
  const shape = (error ?? {}) as IGraphErrorShape;
  const status = shape.statusCode;

  if (status === 401 || status === 403 || shape.code === 'accessDenied') {
    return {
      isActionRequired: true,
      message:
        `This widget needs the Microsoft Graph "${scope}" permission. ` +
        'A tenant administrator has to approve it in the SharePoint admin center under Advanced > API access.'
    };
  }
  if (status === 429 || status === 503) {
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

  // Keep the latest fetcher out of the effect dependencies so callers can pass
  // an inline arrow function without re-querying on every render.
  const fetcherRef = React.useRef(fetcher);
  fetcherRef.current = fetcher;

  React.useEffect(() => {
    let cancelled = false;
    // Data already on screen stays there and is marked stale: re-reading after a
    // refresh or a settings change should not blank a tile the user is reading.
    setResult((previous) =>
      previous.status === 'ready' ? { ...previous, isRefreshing: true } : { status: 'loading' }
    );

    spContext.msGraphClientFactory
      .getClient('3')
      .then((client) => fetcherRef.current(client))
      .then((data) => {
        if (!cancelled) {
          setResult({ status: 'ready', data });
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
