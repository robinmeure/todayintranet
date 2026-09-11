import * as React from 'react';
import { IWidgetContext } from '../IWidget';
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

const SELECT_PROPERTIES: string[] = [
  'Title',
  'Path',
  'SiteTitle',
  'FileType',
  'LastModifiedTime',
  'HitHighlightedSummary'
];
const DEFAULT_MAX_ITEMS: number = 6;
const MAX_ITEMS_BOUNDS = { min: 1, max: 20 };
const DEFAULT_SCOPE: SearchScope = 'siteCollection';

export const SEARCH_DEFAULT_VIEW: WidgetItemsView = 'list';

/** Fluent icons for the file types worth distinguishing at a glance. */
const FILE_TYPE_ICONS: Record<string, string> = {
  docx: 'WordDocument',
  doc: 'WordDocument',
  xlsx: 'ExcelDocument',
  xls: 'ExcelDocument',
  pptx: 'PowerPointDocument',
  ppt: 'PowerPointDocument',
  pdf: 'PDF',
  one: 'OneNoteLogo',
  aspx: 'Page',
  html: 'Page',
  txt: 'TextDocument',
  csv: 'ExcelDocument',
  msg: 'Mail',
  vsdx: 'VisioDocument'
};

function toItem(row: ISearchRow, index: number, locale: string | undefined): IWidgetListItem {
  const fileType = (row.FileType ?? '').toLowerCase();
  return {
    key: row.Path ?? String(index),
    title: row.Title || row.Path || '(Untitled)',
    meta: [row.SiteTitle, formatSearchDate(row.LastModifiedTime, locale)],
    description: plainSummary(row.HitHighlightedSummary),
    iconName: FILE_TYPE_ICONS[fileType] ?? 'TextDocument',
    tone: 'neutral',
    href: row.Path
  };
}

export const SearchResultsWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const scope = scopeSetting(context, DEFAULT_SCOPE);
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, MAX_ITEMS_BOUNDS);
  const view = viewSetting(context, SEARCH_DEFAULT_VIEW);
  const terms = typeof context.settings.terms === 'string' ? context.settings.terms : '';
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  // Without terms this widget has no query of its own, only a scope, which would
  // return the whole site. Better to say so than to show a random sample of it.
  const queryText = terms.trim() ? buildQuery({ scope, terms, spContext: context.spContext }) : '';

  const state = useSearchData(
    context,
    { queryText, selectProperties: SELECT_PROPERTIES, rowLimit: maxItems },
    [queryText, maxItems]
  );

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Searching…', rows: Math.min(maxItems, 4) }}
      empty={{
        iconName: 'Search',
        text: terms.trim()
          ? 'Nothing matched. Try different terms in the tile settings.'
          : 'Add a query in the tile settings to fill this tile.'
      }}
      isEmpty={(results) => results.rows.length === 0}
    >
      {(results) => (
        <WidgetItems
          view={view}
          ariaLabel="Search results"
          items={results.rows.map((row, index) => toItem(row, index, locale))}
        />
      )}
    </WidgetView>
  );
};

export const SearchResultsWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({
  context
}) => (
  <SettingsSurface description="Results are trimmed to what you have access to, exactly like the search box.">
    <SearchQuerySetting
      context={context}
      defaultScope={DEFAULT_SCOPE}
      termsLabel="Query"
      termsPlaceholder="e.g. FileType:docx handbook"
      termsDescription="Plain words or KQL. The scope above is added to it."
      requiresTerms={true}
      selectProperties={SELECT_PROPERTIES}
    />
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Results to show"
      fallback={DEFAULT_MAX_ITEMS}
      {...MAX_ITEMS_BOUNDS}
    />
  </SettingsSurface>
);
