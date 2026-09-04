import * as React from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IGraphError {
  message: string;
  /** True when the tenant admin has not approved the required Graph permission. */
  needsConsent: boolean;
}

export interface IGraphDataResult<T> {
  status: 'loading' | 'ready' | 'error';
  data?: T;
  error?: IGraphError;
  reload(): void;
}

interface IGraphErrorShape {
  statusCode?: number;
  code?: string;
  message?: string;
}

function toGraphError(error: unknown, scope: string): IGraphError {
  const shape = (error ?? {}) as IGraphErrorShape;
  const status = shape.statusCode;

  if (status === 401 || status === 403 || shape.code === 'accessDenied') {
    return {
      needsConsent: true,
      message:
        `This widget needs the Microsoft Graph "${scope}" permission. ` +
        'A tenant administrator has to approve it in the SharePoint admin center under Advanced > API access.'
    };
  }
  if (status === 429 || status === 503) {
    return { needsConsent: false, message: 'Microsoft 365 is busy right now. Try again in a moment.' };
  }
  return { needsConsent: false, message: shape.message || 'Could not load data from Microsoft 365.' };
}

/**
 * Runs a Microsoft Graph query for a widget and exposes loading / error state.
 * `scope` is only used to explain which permission is missing when consent fails.
 */
export function useGraphData<T>(
  spContext: WebPartContext,
  scope: string,
  fetcher: (client: MSGraphClientV3) => Promise<T>,
  deps: unknown[]
): IGraphDataResult<T> {
  const [result, setResult] = React.useState<{ status: 'loading' | 'ready' | 'error'; data?: T; error?: IGraphError }>({
    status: 'loading'
  });
  const [reloadToken, setReloadToken] = React.useState<number>(0);

  // Keep the latest fetcher out of the effect dependencies so callers can pass
  // an inline arrow function without re-querying on every render.
  const fetcherRef = React.useRef(fetcher);
  fetcherRef.current = fetcher;

  React.useEffect(() => {
    let cancelled = false;
    setResult({ status: 'loading' });

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
  }, [spContext, scope, reloadToken, ...deps]);

  const reload = React.useCallback(() => setReloadToken((token) => token + 1), []);

  return { status: result.status, data: result.data, error: result.error, reload };
}
