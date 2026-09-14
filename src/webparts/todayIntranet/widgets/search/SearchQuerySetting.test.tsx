jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));
jest.mock('./SearchQuerySetting.module.scss', () => ({
  scopeDescription: 'scopeDescription',
  preview: 'preview',
  previewHeader: 'previewHeader',
  previewTitle: 'previewTitle',
  query: 'query',
  queryEmpty: 'queryEmpty',
  count: 'count',
  hits: 'hits',
  hit: 'hit',
  hitMeta: 'hitMeta'
}), { virtual: true });
jest.mock('../content/WidgetContent.module.scss', () => ({
  staleNotice: 'staleNotice'
}), { virtual: true });
jest.mock('./cachedSearchData', () => ({
  getSearchData: jest.fn()
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import { SearchError } from './SearchService';
import { getSearchData } from './cachedSearchData';
import { SearchQuerySetting } from './SearchQuerySetting';

const spContext = new (jest.fn<WebPartContext, []>())();
Object.defineProperty(spContext, 'pageContext', {
  value: {
    aadInfo: { tenantId: 'tenant-one', userId: 'user-one' },
    legacyPageContext: {},
    site: {
      id: 'site-one',
      absoluteUrl: 'https://example.test/sites/one'
    },
    web: {
      id: 'web-one',
      absoluteUrl: 'https://example.test/sites/one'
    },
    user: { loginName: 'user-one@example.com' }
  }
});

const context: IWidgetContext = {
  instanceId: 'search',
  settings: { scope: 'everywhere', terms: 'handbook' },
  spContext,
  isEditing: true,
  refreshToken: 0,
  updateSettings: jest.fn<void, [Record<string, unknown>]>()
};

describe('SearchQuerySetting cached preview', () => {
  let container: HTMLDivElement;

  beforeAll(() => {
    initializeIcons(undefined, { disableWarnings: true });
  });

  beforeEach(() => {
    jest.useFakeTimers();
    jest.mocked(getSearchData).mockReset().mockResolvedValue({
      data: {
        rows: [{ Title: 'Employee handbook', Path: 'https://example.test/handbook' }],
        totalRows: 1
      },
      fetchedAt: Date.parse('2026-09-14T08:20:00Z'),
      refreshError: new SearchError({ message: 'Search is busy right now.' }, 503)
    });
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  const render = async (): Promise<void> => {
    await act(async () => {
      ReactDom.render(
        <SearchQuerySetting
          context={context}
          defaultScope="everywhere"
          termsLabel="Query"
          requiresTerms={true}
          selectProperties={['Title', 'Path']}
        />,
        container
      );
    });
    await act(async () => {
      jest.runOnlyPendingTimers();
    });
  };

  it('shows retained preview results with a quiet stale notice and bypasses on Run again', async () => {
    await render();

    expect(container.textContent).toContain('Could not refresh. Showing data from');
    expect(container.textContent).toContain('Employee handbook');
    expect(getSearchData).toHaveBeenCalledTimes(1);
    expect(jest.mocked(getSearchData).mock.calls[0][2]).toBe(false);

    const runAgainLabel = Array.from(container.querySelectorAll<HTMLElement>('.ms-Button-label'))
      .find((label) => label.textContent === 'Run again');
    const runAgain = runAgainLabel?.closest('button');
    if (!runAgain) {
      throw new Error('Run again button was not rendered.');
    }
    await act(async () => {
      Simulate.click(runAgain);
    });
    await act(async () => {
      jest.runOnlyPendingTimers();
    });

    expect(getSearchData).toHaveBeenCalledTimes(2);
    expect(jest.mocked(getSearchData).mock.calls[1][2]).toBe(true);
  });
});
