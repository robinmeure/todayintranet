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

export function getSearchData(
  context: IWidgetContext,
  request: ISearchRequest,
  bypassCache: boolean
): Promise<ICachedWidgetData<ISearchResults>> {
  return getCachedWidgetData({
    key: widgetDataCacheKey(context, 'sharepoint-search', request),
    ttlMilliseconds: SEARCH_CACHE_TTL_MS,
    bypassCache,
    classifyError: classifySearchDataError,
    load: () => runSearchQuery(context.spContext, request)
  });
}
