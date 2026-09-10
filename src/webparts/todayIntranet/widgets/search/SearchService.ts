import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetError } from '../content';

/** One search hit, as a flat bag of the managed properties that were selected. */
export interface ISearchRow {
  [property: string]: string | undefined;
}

export interface ISearchResults {
  rows: ISearchRow[];
  /** Total matches, which is usually larger than `rows`. */
  totalRows: number;
}

export interface ISearchRequest {
  queryText: string;
  selectProperties: string[];
  rowLimit: number;
  /** e.g. `Created:descending`. Omit to sort by relevance. */
  sortList?: string;
}

const JSON_HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=nometadata',
  'odata-version': ''
};

/** Search returns collections bare or wrapped, depending on the OData mode it picks. */
function unwrap<T>(value: unknown): T[] {
  if (Array.isArray(value)) {
    return value as T[];
  }
  const wrapped = value as { results?: T[] } | undefined;
  return wrapped && Array.isArray(wrapped.results) ? wrapped.results : [];
}

interface ISearchPayload {
  PrimaryQueryResult?: {
    RelevantResults?: {
      TotalRows?: number;
      Table?: { Rows?: unknown };
    };
  };
}

function toQueryString(request: ISearchRequest): string {
  const quoted = (value: string): string => `'${encodeURIComponent(value.replace(/'/g, "''"))}'`;

  const parameters: string[] = [
    `querytext=${quoted(request.queryText)}`,
    `selectproperties=${quoted(request.selectProperties.join(','))}`,
    `rowlimit=${Math.max(1, Math.min(50, request.rowLimit))}`,
    'trimduplicates=false',
    `clienttype=${quoted('ContentSearchRegular')}`
  ];
  if (request.sortList) {
    parameters.push(`sortlist=${quoted(request.sortList)}`);
  }
  return parameters.join('&');
}

async function toSearchError(response: SPHttpClientResponse): Promise<IWidgetError> {
  if (response.status === 403 || response.status === 401) {
    return {
      isActionRequired: true,
      message: 'You do not have access to search results on this tenant.'
    };
  }
  if (response.status === 429 || response.status === 503) {
    return { message: 'Search is busy right now. Try again in a moment.' };
  }

  // Search reports a bad query as a 500 with the reason in the body, which is the
  // most useful thing to show someone who is still editing that query.
  let detail: string | undefined;
  try {
    const body = (await response.json()) as {
      error?: { message?: string | { value?: string } };
    };
    const message = body.error ? body.error.message : undefined;
    detail = typeof message === 'string' ? message : message ? message.value : undefined;
  } catch {
    detail = undefined;
  }

  return { message: detail || `Search failed (HTTP ${response.status}).` };
}

/** Thrown by `runSearchQuery` so callers can surface the reason unchanged. */
export class SearchError extends Error {
  public readonly widgetError: IWidgetError;

  constructor(widgetError: IWidgetError) {
    super(widgetError.message);
    this.widgetError = widgetError;
    // Required for `instanceof` to survive the ES5 target.
    Object.setPrototypeOf(this, SearchError.prototype);
  }
}

/** Turns whatever a failed query threw into something worth showing a user. */
export function toWidgetError(error: unknown): IWidgetError {
  if (error instanceof SearchError) {
    return error.widgetError;
  }
  const message = error instanceof Error ? error.message : undefined;
  return { message: message || 'Could not reach SharePoint search.' };
}

/**
 * Runs a query against SharePoint search as the signed in user. Nothing here needs
 * admin consent: results are already trimmed to what the user is allowed to see.
 */
export async function runSearchQuery(
  spContext: WebPartContext,
  request: ISearchRequest
): Promise<ISearchResults> {
  const url = `${spContext.pageContext.web.absoluteUrl}/_api/search/query?${toQueryString(request)}`;

  const response = await spContext.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: JSON_HEADERS
  });
  if (!response.ok) {
    throw new SearchError(await toSearchError(response));
  }

  const payload = (await response.json()) as ISearchPayload;
  const relevant = payload.PrimaryQueryResult ? payload.PrimaryQueryResult.RelevantResults : undefined;
  const rawRows = unwrap<{ Cells?: unknown }>(relevant && relevant.Table ? relevant.Table.Rows : undefined);

  const rows: ISearchRow[] = rawRows.map((rawRow) => {
    const row: ISearchRow = {};
    unwrap<{ Key?: string; Value?: string }>(rawRow.Cells).forEach((cell) => {
      if (cell.Key) {
        row[cell.Key] = cell.Value === null ? undefined : cell.Value;
      }
    });
    return row;
  });

  return { rows, totalRows: relevant && relevant.TotalRows ? relevant.TotalRows : rows.length };
}

/** Strips the hit-highlighting markup search puts around matched terms. */
export function plainSummary(value: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  return value
    .replace(/<\/?c0>/g, '')
    .replace(/<ddd\/>/g, '…')
    .replace(/\s+/g, ' ')
    .trim();
}

export function formatSearchDate(value: string | undefined, locale: string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return undefined;
  }
  const today = new Date();
  const isToday = date.toDateString() === today.toDateString();
  return isToday
    ? date.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })
    : date.toLocaleDateString(locale, { day: 'numeric', month: 'short', year: 'numeric' });
}
