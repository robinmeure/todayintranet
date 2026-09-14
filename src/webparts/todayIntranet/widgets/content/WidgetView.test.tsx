/* eslint-disable @rushstack/pair-react-dom-render-unmount -- The shared container is unmounted in afterEach. */

jest.mock('./WidgetContent.module.scss', () => ({
  stale: 'stale',
  staleNotice: 'staleNotice',
  refreshing: 'refreshing',
  refreshingBar: 'refreshingBar'
}), { virtual: true });

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { WidgetView } from './WidgetView';

describe('WidgetView stale fallback', () => {
  let container: HTMLDivElement;

  beforeAll(() => {
    initializeIcons(undefined, { disableWarnings: true });
  });

  beforeEach(() => {
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
  });

  it('keeps content visible with a polite timestamped retry notice', () => {
    const reload = jest.fn<void, []>();
    act(() => {
      ReactDom.render(
        <WidgetView
          state={{
            status: 'ready',
            data: ['retained row'],
            lastUpdated: Date.parse('2026-09-14T08:20:00Z'),
            refreshError: { message: 'Service unavailable' },
            reload
          }}
          empty={{ iconName: 'Mail', text: 'Empty' }}
        >
          {(items) => <div>{items[0]}</div>}
        </WidgetView>,
        container
      );
    });

    expect(container.textContent).toContain('Could not refresh. Showing data from');
    expect(container.textContent).toContain('retained row');
    expect(container.querySelector('[role="status"]')?.getAttribute('aria-live')).toBe('polite');
    const retry = Array.from(container.querySelectorAll<HTMLElement>('a, button'))
      .find((item) => item.textContent === 'Try again');
    if (!retry) {
      throw new Error('Retry link was not rendered.');
    }
    act(() => {
      Simulate.click(retry);
    });
    expect(reload).toHaveBeenCalledTimes(1);
  });

  it('marks a retained empty response as stale instead of presenting it as newly verified', () => {
    act(() => {
      ReactDom.render(
        <WidgetView
          state={{
            status: 'ready',
            data: [],
            lastUpdated: Date.parse('2026-09-14T08:20:00Z'),
            refreshError: { message: 'Service unavailable' }
          }}
          empty={{ iconName: 'Mail', text: 'Your inbox is empty.' }}
        >
          {() => <div>Unexpected content</div>}
        </WidgetView>,
        container
      );
    });

    expect(container.textContent).toContain('Could not refresh. Showing data from');
    expect(container.textContent).toContain('Your inbox is empty.');
  });
});
