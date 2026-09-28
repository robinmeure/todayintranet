import { IWidgetContext } from '../IWidget';
import {
  ICachedWidgetData,
  getCachedWidgetData,
  widgetDataCacheKey
} from '../data/WidgetDataCache';
import {
  ISearchRequest,
  ISearchResults,
  classifySearchDataError,
  runSearchQuery
} from './SearchService';

const SEARCH_CACHE_TTL_MS: number = 15 * 60 * 1000;
const SEARCH_CACHE_SOURCE: string = 'sharepoint-search';

/** Identifies one search request for this user and site; equal keys share cached data. */
export function searchDataKey(context: Pick<IWidgetContext, 'spContext'>, request: ISearchRequest): string {
  return widgetDataCacheKey(context, SEARCH_CACHE_SOURCE, request);
}

export function getSearchData(
  context: Pick<IWidgetContext, 'spContext'>,
  request: ISearchRequest,
  bypassCache: boolean
): Promise<ICachedWidgetData<ISearchResults>> {
  return getCachedWidgetData({
    key: searchDataKey(context, request),
    ttlMilliseconds: SEARCH_CACHE_TTL_MS,
    bypassCache,
    classifyError: classifySearchDataError,
    load: () => runSearchQuery(context.spContext, request)
  });
}
