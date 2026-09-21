/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each container is unmounted in afterEach. */

jest.mock('../content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetCollections.module.scss', () => ({ linksGrid: 'linksGrid' }), { virtual: true });
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { MyLinksWidget, MyLinksWidgetSettings } from './MyLinksWidget';
import { DEFAULT_LINKS_SETTING, parseLinksSetting } from './linksSettings';
import { IWidgetContext } from '../IWidget';
import { WidgetRegistry } from '../WidgetRegistry';
import { ALL_ITEM_VIEWS } from '../content';
import { WidgetFrame } from '../../components/WidgetFrame';

const links = Array.from({ length: 12 }, (_, index) => ({
  title: `Tool ${index + 1}`, url: `https://example.test/tool/${index + 1}`,
  description: index === 10 ? 'Mandatory training' : 'Work tools',
  badge: index === 8 ? 'VPN' : undefined
}));
const spContext = new (jest.fn<WebPartContext, []>())();
const updateSettings = jest.fn<void, [Record<string, unknown>]>();

describe('My links interactions and widget contract', () => {
  let container: HTMLDivElement;
  let context: IWidgetContext;
  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => {
    jest.useFakeTimers();
    updateSettings.mockClear();
    context = {
      instanceId: 'my-links', spContext, refreshToken: 0, isEditing: false,
      settings: { links: JSON.stringify(links), itemsPerPage: 9, unrelated: 'preserve-me' },
      updateSettings
    };
    container = document.createElement('div');
    document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    jest.clearAllTimers();
    jest.useRealTimers();
  });
  const render = (): void => { act(() => { ReactDom.render(<MyLinksWidget context={context} />, container); }); };
  const button = (label: string): HTMLButtonElement => {
    const found = container.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`);
    if (!found) { throw new Error(`Missing button: ${label}`); }
    return found;
  };
  const search = (query: string): void => {
    const input = container.querySelector<HTMLInputElement>('input[type="search"]');
    if (!input) { throw new Error('Missing search field.'); }
    act(() => { input.value = query; Simulate.change(input); });
  };

  it('shows sample links for unconfigured instances without saving them automatically', () => {
    context.settings = {};
    render();
    expect(container.querySelectorAll('a')).toHaveLength(6);
    expect(Array.from(container.querySelectorAll('a')).map((link) => link.getAttribute('href')))
      .toEqual(parseLinksSetting(DEFAULT_LINKS_SETTING).links?.map((link) => link.url));
    expect(container.textContent).toContain('Sample:');
    expect(container.textContent).not.toContain('VPN');
    expect(updateSettings).not.toHaveBeenCalled();
  });

  it('shows the same editable samples in settings without saving untouched defaults', () => {
    context.settings = { unrelated: 'preserve-me' };
    act(() => { ReactDom.render(<MyLinksWidgetSettings context={context} />, container); });
    const input = container.querySelector('textarea');
    if (!input) { throw new Error('Missing Links JSON editor.'); }
    expect(input.value).toBe(DEFAULT_LINKS_SETTING);
    act(() => { Simulate.blur(input); });
    render();
    expect(updateSettings).not.toHaveBeenCalled();
    expect(container.querySelectorAll('a')).toHaveLength(6);
  });

  it('treats an explicitly empty list as setup rather than a failed search', () => {
    context.settings = { links: '[]' };
    render();
    expect(container.textContent).toContain('No links configured.');
    expect(container.textContent).not.toContain('No links match');
    expect(container.querySelectorAll('a')).toHaveLength(0);
  });

  it('reports configuration errors without making any link clickable', () => {
    context.settings = { links: '{broken' };
    render();
    act(() => { jest.runOnlyPendingTimers(); });
    expect(container.textContent).toContain('not valid JSON');
    expect(container.querySelectorAll('a')).toHaveLength(0);
  });

  it('paginates and resets to page one when searching all pages by description or badge', () => {
    render();
    expect(container.querySelectorAll('a')).toHaveLength(9);
    expect(button('Previous links page').disabled).toBe(true);
    act(() => { Simulate.click(button('Next links page')); });
    expect(container.querySelectorAll('a')).toHaveLength(3);
    expect(button('Links page 2').getAttribute('aria-current')).toBe('page');
    expect(button('Next links page').disabled).toBe(true);
    search(' TRAINING ');
    expect(container.querySelectorAll('a')).toHaveLength(1);
    expect(container.textContent).toContain('Tool 11');
    expect(container.querySelector('nav')).toBeNull();
    search('vpn');
    expect(container.querySelectorAll('a')).toHaveLength(1);
    expect(container.textContent).toContain('Tool 9');
    search('no such link');
    expect(container.textContent).toContain('No links match');
    search('');
    expect(button('Links page 1').getAttribute('aria-current')).toBe('page');
    expect(updateSettings).not.toHaveBeenCalled();
  });

  it('keeps pagination valid after configuration changes', () => {
    render();
    act(() => { Simulate.click(button('Next links page')); });
    context.settings = { ...context.settings, links: JSON.stringify(links.slice(0, 2)) };
    render();
    expect(container.querySelectorAll('a')).toHaveLength(2);
    expect(container.querySelector('nav')).toBeNull();
    expect(container.querySelector('[role="status"]')?.textContent).toContain('Page 1 of 1');
  });

  it('keeps separate instances independently searchable', () => {
    act(() => {
      ReactDom.render(<><MyLinksWidget context={context} /><MyLinksWidget context={{ ...context, instanceId: 'second' }} /></>, container);
    });
    search('Tool 12');
    expect(container.querySelectorAll('a')).toHaveLength(10);
  });

  it('commits JSON through the existing settings callback, preserving unrelated settings', () => {
    act(() => { ReactDom.render(<MyLinksWidgetSettings context={context} />, container); });
    const input = container.querySelector('textarea');
    if (!input) { throw new Error('Missing Links JSON editor.'); }
    const next = JSON.stringify([{ title: 'Our site', url: 'https://example.test' }]);
    act(() => { input.value = next; Simulate.change(input); });
    expect(updateSettings).not.toHaveBeenCalled();
    act(() => { Simulate.blur(input); });
    expect(updateSettings).toHaveBeenLastCalledWith({ ...context.settings, links: next });
    context.settings = updateSettings.mock.calls[0][0];
    render();
    expect(container.textContent).toContain('Our site');
    expect(container.querySelector('a')?.getAttribute('href')).toBe('https://example.test/');
  });

  it('commits final drafts on settings dismissal', () => {
    act(() => { ReactDom.render(<MyLinksWidgetSettings context={context} />, container); });
    const input = container.querySelector('textarea');
    if (!input) { throw new Error('Missing Links JSON editor.'); }
    act(() => { input.value = '[]'; Simulate.change(input); });
    render();
    expect(updateSettings).toHaveBeenLastCalledWith({ ...context.settings, links: '[]' });
  });

  it.each([[1e200, 12], [-5, 3], [4.8, 4], [undefined, 9]])(
    'bounds links per page (%p)', (itemsPerPage, expected) => {
      context.settings = { ...context.settings, itemsPerPage };
      render();
      expect(container.querySelectorAll('a')).toHaveLength(expected);
    }
  );

  it('registers specialized views without advertising them for unrelated widgets', () => {
    expect(WidgetRegistry.get('custom.myLinks')?.supportedViews).toEqual(['links', ...ALL_ITEM_VIEWS]);
    expect(WidgetRegistry.get('m365.calendar')?.supportedViews).toEqual(['agenda', ...ALL_ITEM_VIEWS]);
    expect(WidgetRegistry.get('m365.mail')?.supportedViews).toEqual(ALL_ITEM_VIEWS);
    expect(WidgetRegistry.get('custom.myLinks')?.requiredPermission).toBeUndefined();
    expect(WidgetRegistry.get('custom.myLinks')?.isRefreshable).toBeUndefined();
    expect(WidgetRegistry.get('custom.myLinks')?.contentSurface).toBe('none');
  });

  it('uses the host title and settings chrome and exposes the shared view picker', () => {
    act(() => {
      ReactDom.render(<WidgetFrame
        instance={{ id: 'my-links', type: 'custom.myLinks', title: 'Work tools', x: 0, y: 0, w: 12, h: 9 }}
        widgetContext={context} isEditing={true}
        onRemove={jest.fn()} onNudge={jest.fn()} onUpdateTitle={jest.fn()}
      />, container);
    });
    expect(container.querySelector('section')?.getAttribute('aria-label')).toBe('Work tools');
    act(() => { Simulate.click(button('Work tools settings')); jest.runOnlyPendingTimers(); });
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain('Link tiles');
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain('Links (JSON)');
    expect(container.querySelector('button[aria-label="Refresh Work tools"]')).toBeNull();
  });
});
