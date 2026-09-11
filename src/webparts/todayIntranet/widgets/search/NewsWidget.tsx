import * as React from 'react';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import {
  IWidgetListItem,
  NumberSetting,
  SettingsSurface,
  WidgetItems,
  WidgetItemsView,
  WidgetView,
  numberSetting,
  viewSetting
} from '../content';
import { useSearchData } from './useSearchData';
import { ISearchRow, formatSearchDate, plainSummary } from './SearchService';
import { SearchQuerySetting } from './SearchQuerySetting';
import { SearchScope, buildQuery, scopeSetting } from './searchQuery';

/** News posts are site pages promoted as news. */
const NEWS_BASE: string = 'PromotedState:2';
const SELECT_PROPERTIES: string[] = [
  'Title',
  'Path',
  'Created',
  'FirstPublishedDate',
  'Description',
  'SiteTitle',
  'HitHighlightedSummary'
];
const SORT_LIST: string = 'Created:descending';
const DEFAULT_MAX_ITEMS: number = 4;
const MAX_ITEMS_BOUNDS = { min: 1, max: 20 };
const DEFAULT_SCOPE: SearchScope = 'siteCollection';

export const NEWS_DEFAULT_VIEW: WidgetItemsView = 'cards';

/** Site relative, so it has to be resolved against the page rather than hard coded. */
export function newsFooterLink(context: IWidgetContext): IWidgetLink {
  return { text: 'All news', href: `${context.spContext.pageContext.web.absoluteUrl}/_layouts/15/news.aspx` };
}

function toItem(row: ISearchRow, index: number, locale: string | undefined, showSite: boolean): IWidgetListItem {
  const published = formatSearchDate(row.FirstPublishedDate ?? row.Created, locale);
  return {
    key: row.Path ?? String(index),
    title: row.Title || '(Untitled post)',
    meta: [showSite ? row.SiteTitle : undefined, published],
    description: plainSummary(row.Description) ?? plainSummary(row.HitHighlightedSummary),
    hasAccentBar: true,
    tone: 'accent',
    href: row.Path
  };
}

export const NewsWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const scope = scopeSetting(context, DEFAULT_SCOPE);
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, MAX_ITEMS_BOUNDS);
  const view = viewSetting(context, NEWS_DEFAULT_VIEW);
  const terms = typeof context.settings.terms === 'string' ? context.settings.terms : '';
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const queryText = buildQuery({ base: NEWS_BASE, scope, terms, spContext: context.spContext });

  const state = useSearchData(
    context,
    { queryText, selectProperties: SELECT_PROPERTIES, rowLimit: maxItems, sortList: SORT_LIST },
    [queryText, maxItems]
  );

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Loading news…', rows: Math.min(maxItems, 3) }}
      empty={{ iconName: 'News', text: 'No news here yet. Try a wider scope in the tile settings.' }}
      isEmpty={(results) => results.rows.length === 0}
    >
      {(results) => (
        <WidgetItems
          view={view}
          ariaLabel="News"
          items={results.rows.map((row, index) => toItem(row, index, locale, scope !== 'site'))}
        />
      )}
    </WidgetView>
  );
};

export const NewsWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="News posts are found with SharePoint search, so you only ever see what you already have access to.">
    <SearchQuerySetting
      context={context}
      base={NEWS_BASE}
      defaultScope={DEFAULT_SCOPE}
      termsLabel="Narrow it down"
      termsPlaceholder="e.g. Title:release or a keyword"
      termsDescription="Optional. Plain words or KQL, added to the query below."
      selectProperties={SELECT_PROPERTIES}
      sortList={SORT_LIST}
    />
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Posts to show"
      fallback={DEFAULT_MAX_ITEMS}
      {...MAX_ITEMS_BOUNDS}
    />
  </SettingsSurface>
);
