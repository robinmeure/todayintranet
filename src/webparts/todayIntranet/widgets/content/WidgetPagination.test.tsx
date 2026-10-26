/* eslint-disable @rushstack/pair-react-dom-render-unmount -- The shared container is unmounted after each test. */
jest.mock('./WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('./WidgetCollections.module.scss', () => ({}), { virtual: true });
jest.mock('./WidgetAdaptiveCard', () => ({
  WidgetAdaptiveCard: ({ card }: { card: { body: unknown[] } }) => <div data-count={card.body.length} />
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import { WidgetItems } from './WidgetItems';
import { WidgetItemsView } from './IWidgetContent';
import { paginationSize, PaginationSettings } from './WidgetPagination';

const items = Array.from({ length: 13 }, (_, index) => ({ key: String(index), title: `Item ${index + 1}` }));

describe('Shared widget pagination', () => {
  let container: HTMLDivElement;
  const render = (size?: number, count = 13, view: WidgetItemsView = 'list'): void => {
    act(() => { ReactDom.render(<WidgetItems items={items.slice(0, count)} pageSize={size} view={view}
      ariaLabel="Example" />, container); });
  };
  const click = (label: string): void => {
    const button = container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
    if (!button) { throw new Error(`Missing button: ${label}`); }
    act(() => Simulate.click(button));
  };
  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => { container = document.createElement('div'); document.body.appendChild(container); });
  afterEach(() => { act(() => { ReactDom.unmountComponentAtNode(container); }); container.remove(); });

  it.each<WidgetItemsView>(['list', 'compact', 'cards', 'gallery', 'agenda', 'links', 'adaptive'])(
    'paginates the %s view with accessible controls', (view) => {
      render(3, 13, view);
      expect(container.querySelector('nav')?.getAttribute('aria-label')).toBe('Example pages');
      expect(container.querySelector('[aria-current="page"]')?.textContent).toBe('1');
      click('Example page 5');
      expect(container.querySelector('[aria-current="page"]')?.textContent).toBe('5');
      expect(container.querySelector<HTMLButtonElement>('[aria-label="Next example page"]')?.disabled).toBe(true);
      if (view === 'adaptive') {
        expect(container.querySelector('[data-count]')?.getAttribute('data-count')).toBe('1');
      } else {
        expect(container.textContent).toContain('Item 13');
        expect(container.textContent).not.toContain('Item 12');
      }
    }
  );

  it('shows all content when disabled and hides navigation for a single page', () => {
    render();
    expect(container.textContent).toContain('Item 13');
    expect(container.querySelector('nav')).toBeNull();
    render(20);
    expect(container.querySelector('nav')).toBeNull();
  });

  it('clamps a removed last page, resets page size changes, and handles empty results', () => {
    render(3);
    click('Example page 5');
    render(3, 4);
    expect(container.textContent).toContain('Page 2 of 2');
    render(3, 13);
    expect(container.textContent).toContain('Page 2 of 5');
    render(6);
    expect(container.textContent).toContain('Page 1 of 3');
    render(6, 0);
    expect(container.querySelector('nav')).toBeNull();
  });

  it('bounds saved sizes and exposes settings without changing unrelated configuration', () => {
    const context: IWidgetContext = {
      instanceId: 'example', spContext: new (jest.fn<WebPartContext, []>())(), isEditing: true,
      refreshToken: 0, settings: { paginationEnabled: true, itemsPerPage: 1e9, unrelated: 'keep' },
      updateSettings: jest.fn()
    };
    expect(paginationSize(context)).toBe(100);
    expect(paginationSize({ ...context, settings: {} })).toBeUndefined();
    act(() => { ReactDom.render(<PaginationSettings context={context} />, container); });
    const toggle = container.querySelector<HTMLButtonElement>('[role="switch"]');
    if (!toggle) { throw new Error('Missing pagination toggle'); }
    act(() => Simulate.click(toggle));
    expect(context.updateSettings).toHaveBeenCalledWith({ ...context.settings, paginationEnabled: false });
  });
});
