import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Link } from '@fluentui/react/lib/Link';
import { IWidgetContext } from '../IWidget';
import { IWidgetError, WidgetErrorMessage, useDraftSetting } from '../content';
import { ISearchResults, plainSummary, runSearchQuery, toWidgetError } from './SearchService';
import {
  SCOPE_SETTING_KEY,
  SEARCH_SCOPES,
  SearchScope,
  TERMS_SETTING_KEY,
  buildQuery,
  isScopeAvailable,
  scopeSetting
} from './searchQuery';
import styles from './SearchQuerySetting.module.scss';

const PREVIEW_ROWS: number = 5;
const PREVIEW_DEBOUNCE_MS: number = 500;

interface IPreviewState {
  status: 'idle' | 'running' | 'done' | 'error';
  results?: ISearchResults;
  error?: IWidgetError;
}

export interface ISearchQuerySettingProps {
  context: IWidgetContext;
  /** Fixed part the widget always applies, e.g. `PromotedState:2` for news. */
  base?: string;
  defaultScope: SearchScope;
  termsLabel: string;
  termsPlaceholder?: string;
  termsDescription?: string;
  /**
   * True for widgets that refuse to run on a scope alone. Keeps the preview honest:
   * it only ever runs what the tile itself would run.
   */
  requiresTerms?: boolean;
  /** Managed properties the preview asks for; the title is enough to show a hit. */
  selectProperties: string[];
  sortList?: string;
}

/**
 * Scope picker, query box, and a preview that runs the real query while you edit it.
 * Building a search query blind is the worst part of configuring this kind of widget,
 * so the effective KQL and its first hits are always on screen.
 */
export const SearchQuerySetting: React.FunctionComponent<ISearchQuerySettingProps> = (props) => {
  const {
    context,
    base,
    defaultScope,
    termsLabel,
    termsPlaceholder,
    termsDescription,
    requiresTerms,
    selectProperties,
    sortList
  } = props;

  const spContext = context.spContext;
  const scope = scopeSetting(context, defaultScope);
  const terms = useDraftSetting(context, TERMS_SETTING_KEY, '');

  const [runToken, setRunToken] = React.useState<number>(0);
  const [preview, setPreview] = React.useState<IPreviewState>({ status: 'idle' });

  const query = buildQuery({ base, scope, terms: terms.value, spContext });
  // Mirrors what the widget will do, so the preview never shows results the tile
  // would refuse to fetch.
  const isRunnable: boolean = !!query.trim() && (!requiresTerms || !!terms.value.trim());

  // The preview follows whatever is on screen, debounced, so a query takes shape as
  // it is typed rather than only once it has been saved.
  React.useEffect(() => {
    if (!isRunnable) {
      setPreview({ status: 'idle' });
      return undefined;
    }

    let cancelled = false;
    setPreview((current) => ({ ...current, status: 'running' }));

    const timer = window.setTimeout(() => {
      runSearchQuery(spContext, {
        queryText: query,
        selectProperties,
        rowLimit: PREVIEW_ROWS,
        sortList
      })
        .then((results) => {
          if (!cancelled) {
            setPreview({ status: 'done', results });
          }
        })
        .catch((error: unknown) => {
          if (!cancelled) {
            setPreview({ status: 'error', error: toWidgetError(error) });
          }
        });
    }, PREVIEW_DEBOUNCE_MS);

    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
    // selectProperties / sortList are constants supplied by the widget.
  }, [spContext, query, isRunnable, runToken]);

  const scopeOptions: IDropdownOption[] = SEARCH_SCOPES.map((option) => ({
    key: option.id,
    text: option.name,
    disabled: !isScopeAvailable(option.id, spContext),
    title: option.description
  }));

  const selectedScope = SEARCH_SCOPES.filter((option) => option.id === scope)[0];
  const rows = preview.results ? preview.results.rows : [];

  return (
    <>
      <div>
        <Dropdown
          label="Look in"
          selectedKey={scope}
          options={scopeOptions}
          onChange={(_, option) =>
            option &&
            context.updateSettings({ ...context.settings, [SCOPE_SETTING_KEY]: option.key as SearchScope })
          }
        />
        <p className={styles.scopeDescription}>
          {selectedScope ? selectedScope.description : ''}
          {!isScopeAvailable('hub', spContext) && ' This site is not associated with a hub.'}
        </p>
      </div>

      <TextField
        label={termsLabel}
        placeholder={termsPlaceholder}
        description={termsDescription}
        value={terms.value}
        onChange={(_, next) => terms.setValue(next ?? '')}
        onBlur={terms.commit}
      />

      <div className={styles.preview}>
        <div className={styles.previewHeader}>
          <span className={styles.previewTitle}>Preview</span>
          <DefaultButton
            iconProps={{ iconName: 'Refresh' }}
            text="Run again"
            disabled={!isRunnable}
            onClick={() => setRunToken((token) => token + 1)}
          />
        </div>

        <code className={`${styles.query} ${isRunnable ? '' : styles.queryEmpty}`}>
          {isRunnable
            ? query
            : requiresTerms
              ? `Add a query above. The scope alone (${query || 'everywhere'}) is not enough.`
              : 'Nothing to run yet — pick a scope or add some terms.'}
        </code>

        {preview.status === 'running' && <Spinner size={SpinnerSize.xSmall} label="Running the query…" />}

        {preview.status === 'error' && preview.error && <WidgetErrorMessage error={preview.error} />}

        {preview.status === 'done' && preview.results && (
          <>
            <div className={styles.count}>
              {preview.results.totalRows === 0
                ? 'No results. Widen the scope or change the terms.'
                : `${preview.results.totalRows} result${preview.results.totalRows === 1 ? '' : 's'}${
                    rows.length < preview.results.totalRows ? `, first ${rows.length} shown` : ''
                  }`}
            </div>
            <ul className={styles.hits}>
              {rows.map((row, index) => (
                <li key={row.Path ?? index} className={styles.hit}>
                  {row.Path ? (
                    <Link href={row.Path} target="_blank" rel="noreferrer">
                      {row.Title || row.Path}
                    </Link>
                  ) : (
                    row.Title || '(Untitled)'
                  )}
                  {row.SiteTitle && <span className={styles.hitMeta}> · {row.SiteTitle}</span>}
                  {!row.SiteTitle && row.HitHighlightedSummary && (
                    <span className={styles.hitMeta}> · {plainSummary(row.HitHighlightedSummary)}</span>
                  )}
                </li>
              ))}
            </ul>
          </>
        )}
      </div>
    </>
  );
};
