import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';

/** Where a search widget looks. Persisted in widget settings, so never rename one. */
export type SearchScope = 'site' | 'siteCollection' | 'hub' | 'everywhere';

export interface ISearchScopeInfo {
  id: SearchScope;
  name: string;
  description: string;
}

export const SEARCH_SCOPES: ISearchScopeInfo[] = [
  { id: 'site', name: 'This site', description: 'Only content on this site.' },
  { id: 'siteCollection', name: 'This site collection', description: 'This site and everything under it.' },
  { id: 'hub', name: 'This hub', description: 'Every site associated with this hub.' },
  { id: 'everywhere', name: 'Everywhere', description: 'Everything you have access to.' }
];

export const SCOPE_SETTING_KEY: string = 'scope';
export const TERMS_SETTING_KEY: string = 'terms';

export function getSearchScope(id: string | undefined): ISearchScopeInfo | undefined {
  return SEARCH_SCOPES.filter((scope) => scope.id === id)[0];
}

/** The scope this tile is set to, falling back to the widget's own default. */
export function scopeSetting(context: IWidgetContext, fallback: SearchScope): SearchScope {
  const value: unknown = context.settings[SCOPE_SETTING_KEY];
  const known = typeof value === 'string' ? getSearchScope(value) : undefined;
  return known ? known.id : fallback;
}

/** Hub this site belongs to, or undefined when it is not in a hub. */
export function hubSiteId(spContext: WebPartContext): string | undefined {
  const legacy = spContext.pageContext.legacyPageContext as { hubSiteId?: string } | undefined;
  const id = legacy ? legacy.hubSiteId : undefined;
  return id && id !== '00000000-0000-0000-0000-000000000000' ? id : undefined;
}

/** True when the scope cannot be used here, e.g. "This hub" on a site with no hub. */
export function isScopeAvailable(scope: SearchScope, spContext: WebPartContext): boolean {
  return scope !== 'hub' || !!hubSiteId(spContext);
}

/** The KQL fragment that limits a query to a scope. Empty for "everywhere". */
export function scopeFilter(scope: SearchScope, spContext: WebPartContext): string {
  switch (scope) {
    case 'site':
      return `SPWebUrl:"${spContext.pageContext.web.absoluteUrl}"`;
    case 'siteCollection':
      return `SPSiteURL:"${spContext.pageContext.site.absoluteUrl}"`;
    case 'hub': {
      const id = hubSiteId(spContext);
      return id ? `DepartmentId:{${id}}` : '';
    }
    default:
      return '';
  }
}

export interface IQueryParts {
  /** Fixed part the widget always applies, e.g. `PromotedState:2` for news. */
  base?: string;
  scope: SearchScope;
  /** Whatever the user typed: free text, KQL, or nothing. */
  terms?: string;
  spContext: WebPartContext;
}

/**
 * The query that actually runs. Shown to the user in the settings preview, so that
 * building one is a matter of seeing what each choice does rather than guessing.
 */
export function buildQuery(parts: IQueryParts): string {
  const { base, scope, terms, spContext } = parts;
  return [base, scopeFilter(scope, spContext), terms]
    .map((part) => (part ?? '').trim())
    .filter((part) => part.length > 0)
    .join(' ');
}
