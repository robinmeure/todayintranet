/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each container is unmounted in afterEach. */

jest.mock('../content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetCollections.module.scss', () => ({ linksGrid: 'linksGrid' }), { virtual: true });
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('../graph/useTaskOrganization', () => ({ useTaskOrganization: jest.fn() }));
jest.mock('../graph/TasksWidget.module.scss', () => ({}), { virtual: true });
jest.mock('./MyLinksWidgetSettings.module.scss', () => ({ row: 'authoring-row' }), { virtual: true });

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
    updateSettings.mockReset();
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
  const pointer = (row: HTMLLIElement, type: string, clientY: number): void => {
    const event = new MouseEvent(type, { bubbles: true, button: 0, clientX: 10, clientY });
    Object.defineProperties(event, { isPrimary: { value: true }, pointerId: { value: 1 } });
    row.dispatchEvent(event);
  };
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

  it('shows the no-links-added state for unconfigured instances without saving defaults', () => {
    context.settings = {};
    render();
    expect(container.querySelectorAll('a')).toHaveLength(0);
    expect(Array.from(container.querySelectorAll('a')).map((link) => link.getAttribute('href')))
      .toEqual(parseLinksSetting(DEFAULT_LINKS_SETTING).links?.map((link) => link.url));
    expect(container.textContent).toContain('No links added');
    expect(container.textContent).not.toContain('VPN');
    expect(updateSettings).not.toHaveBeenCalled();
  });

  it('starts settings empty without saving untouched defaults', () => {
    context.settings = { unrelated: 'preserve-me' };
    act(() => { ReactDom.render(<MyLinksWidgetSettings context={context} />, container); });
    const input = container.querySelector('textarea');
    if (!input) { throw new Error('Missing Links JSON editor.'); }
    expect(input.value).toBe(DEFAULT_LINKS_SETTING);
    act(() => { Simulate.blur(input); });
    render();
    expect(updateSettings).not.toHaveBeenCalled();
    expect(container.querySelectorAll('a')).toHaveLength(0);
  });

  it('shows the designed no-links-added state without search for an empty list', () => {
    context.settings = { links: '[]' };
    render();
    expect(container.textContent).toContain('No links added');
    expect(container.textContent).not.toContain('No links match');
    expect(container.querySelectorAll('a')).toHaveLength(0);
    expect(container.querySelector('input[type="search"]')).toBeNull();
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
    expect(container.textContent).toContain('No links found');
    search('');
    expect(button('Links page 1').getAttribute('aria-current')).toBe('page');
    expect(updateSettings).not.toHaveBeenCalled();
  });

  it('searches URLs and titles case-insensitively across pages without changing settings', () => {
    render();
    act(() => { Simulate.click(button('Next links page')); });
    search('EXAMPLE.TEST/TOOL/12');
    expect(container.querySelectorAll('a')).toHaveLength(1);
    expect(container.querySelector('a')?.textContent).toContain('Tool 12');
    search(' tOOL 2 ');
    expect(container.querySelectorAll('a')).toHaveLength(1);
    expect(container.querySelector('a')?.textContent).toContain('Tool 2');
    search('');
    expect(button('Links page 1').getAttribute('aria-current')).toBe('page');
    expect(updateSettings).not.toHaveBeenCalled();
  });

  it('paginates filtered results rather than filtering only the current page', () => {
    context.settings = { ...context.settings, itemsPerPage: 3 };
    render();
    act(() => { Simulate.click(button('Links page 4')); });
    search('EXAMPLE.TEST');
    expect(button('Links page 1').getAttribute('aria-current')).toBe('page');
    expect(container.querySelectorAll('a')).toHaveLength(3);
    act(() => { Simulate.click(button('Links page 4')); });
    expect(container.querySelectorAll('a')).toHaveLength(3);
    expect(container.textContent).toContain('Tool 12');
  });

  it('displays VPN labels and opens links in a new tab', () => {
    render();
    expect(container.textContent).toContain('VPN');
    container.querySelectorAll('a').forEach((link) => {
      expect(link.target).toBe('_blank');
      expect(link.rel).toContain('noreferrer');
    });
  });

  it.each(ALL_ITEM_VIEWS.filter((view) => view !== 'adaptive'))('opens links in a new tab in the %s view', (view) => {
    context.settings = { ...context.settings, view };
    render();
    expect(container.querySelectorAll('a')).toHaveLength(9);
    container.querySelectorAll('a').forEach((link) => expect(link.target).toBe('_blank'));
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

  it.each(['{broken', '[{"title":"Unsafe","url":"http://example.test"}]'])(
    'never saves invalid JSON drafts on blur or dismissal (%s)', (value) => {
      act(() => { ReactDom.render(<MyLinksWidgetSettings context={context} />, container); });
      const input = container.querySelector('textarea');
      if (!input) { throw new Error('Missing Links JSON editor.'); }
      act(() => { input.value = value; Simulate.change(input); Simulate.blur(input); });
      expect(container.querySelector('[role="alert"]')).not.toBeNull();
      render();
      expect(updateSettings).not.toHaveBeenCalled();
    }
  );

  describe('link management', () => {
    const settings = (): void => {
      act(() => {
        ReactDom.render(<MyLinksWidgetSettings context={context} />, container);
      });
      act(() => { jest.runOnlyPendingTimers(); });
    };
    const click = (text: string): void => {
      const found = Array.from(document.querySelectorAll('button'))
        .find((item) => item.querySelector('.ms-Button-label')?.textContent === text);
      if (!found) { throw new Error(`Missing button: ${text}`); }
      act(() => {
        if (found.type === 'submit' && found.form) { Simulate.submit(found.form); }
        else { Simulate.click(found); }
      });
      act(() => { jest.runOnlyPendingTimers(); });
    };
    const field = (label: string, value: string): void => {
      const id = Array.from(container.querySelectorAll('label')).find((item) => item.textContent?.startsWith(label))?.htmlFor;
      const input = id ? document.getElementById(id) : undefined;
      if (!(input instanceof HTMLInputElement) && !(input instanceof HTMLTextAreaElement)) {
        throw new Error(`Missing field: ${label}`);
      }
      act(() => { input.value = value; Simulate.change(input); });
    };
    const saved = (): NonNullable<ReturnType<typeof parseLinksSetting>['links']> => {
      const last = updateSettings.mock.calls[updateSettings.mock.calls.length - 1]?.[0];
      if (!last || typeof last.links !== 'string') { throw new Error('Links were not saved.'); }
      const parsed = parseLinksSetting(last.links);
      if (!parsed.links) { throw new Error(parsed.error); }
      expect(last.unrelated).toBe('preserve-me');
      return parsed.links;
    };

    it('adds a validated HTTPS link with a VPN tag and renders the saved result', () => {
      settings();
      click('Add a new link');
      expect(container.querySelector('ol')).toBeNull();
      field('Name', '  VPN tool  ');
      field('Link', 'https://vpn.example.test/resource');
      field('Description', '  Private tool  ');
      const toggle = container.querySelector<HTMLButtonElement>('button[role="switch"]');
      if (!toggle) { throw new Error('Missing VPN toggle.'); }
      act(() => { Simulate.click(toggle); });
      click('Add link');
      expect(saved()[12]).toMatchObject({
        title: 'VPN tool', url: 'https://vpn.example.test/resource', description: 'Private tool', badge: 'VPN'
      });
      context.settings = updateSettings.mock.calls[0][0];
      render();
      search('VPN.EXAMPLE.TEST');
      expect(container.querySelector('a')?.textContent).toContain('VPN');
    });

    it.each([
      ['', 'https://example.test', 'needs a title'],
      ['Unsafe', 'http://example.test', 'HTTPS URL'],
      ['Unsafe', '/relative', 'HTTPS URL'],
      ['Unsafe', 'https://user:password@example.test', 'without credentials']
    ])('rejects invalid additions (%s, %s)', (title, url, message) => {
      settings();
      click('Add a new link');
      field('Name', title);
      field('Link', url);
      click('Add link');
      expect(container.querySelector('[role="alert"]')?.textContent).toContain(message);
      expect(updateSettings).not.toHaveBeenCalled();
    });

    it('updates an existing link and preserves optional presentation fields', () => {
      context.settings = { ...context.settings, links: JSON.stringify([{
        title: 'Old tool', url: 'https://old.example.test', badge: 'VPN', iconName: 'Education', tone: 'warning'
      }]) };
      settings();
      act(() => { Simulate.click(button('Edit Old tool')); });
      expect(container.querySelector('h2')?.textContent).toBe('Edit link');
      field('Name', 'Updated tool');
      field('Link', 'https://new.example.test');
      click('Update');
      expect(saved()).toEqual([{
        title: 'Updated tool', url: 'https://new.example.test/', badge: 'VPN', iconName: 'Education', tone: 'warning'
      }]);
    });

    it('cancels or dismisses an unfinished form without persisting it', () => {
      settings();
      click('Add a new link');
      field('Name', 'Not saved');
      field('Link', 'https://example.test');
      click('Cancel');
      expect(updateSettings).not.toHaveBeenCalled();
      click('Add a new link');
      field('Name', 'Still not saved');
      render();
      expect(updateSettings).not.toHaveBeenCalled();
    });

    it('requires confirmation before deleting and supports cancellation', () => {
      settings();
      act(() => { Simulate.click(button('Delete Tool 1')); });
      expect(updateSettings).not.toHaveBeenCalled();
      click('Cancel delete');
      expect(updateSettings).not.toHaveBeenCalled();
      act(() => { Simulate.click(button('Delete Tool 1')); });
      click('Confirm delete');
      expect(saved().map((link) => link.title)).toEqual(links.slice(1).map((link) => link.title));
    });

    it('deletes the last link without restoring defaults', () => {
      context.settings = { ...context.settings, links: JSON.stringify(links.slice(0, 1)) };
      settings();
      act(() => { Simulate.click(button('Delete Tool 1')); });
      click('Confirm delete');
      expect(saved()).toEqual([]);
      context.settings = updateSettings.mock.calls[0][0];
      render();
      expect(container.textContent).toContain('No links added');
    });

    it('persists accessible keyboard reordering and ignores moves past the boundaries', () => {
      settings();
      const row = container.querySelector<HTMLLIElement>('.authoring-row');
      if (!row) { throw new Error('Missing authoring row.'); }
      act(() => { Simulate.keyDown(row, { altKey: true, key: 'ArrowUp' }); });
      expect(updateSettings).not.toHaveBeenCalled();
      act(() => { Simulate.keyDown(row, { altKey: true, key: 'ArrowDown' }); });
      expect(saved().slice(0, 2).map((link) => link.title)).toEqual(['Tool 2', 'Tool 1']);
      context.settings = updateSettings.mock.calls[0][0];
      settings();
      const moved = container.querySelectorAll<HTMLLIElement>('.authoring-row')[1];
      act(() => { Simulate.keyDown(moved, { altKey: true, key: 'ArrowUp' }); });
      expect(saved().map((link) => link.title)).toEqual(links.map((link) => link.title));
    });

    it('persists case-insensitive title sorting in both directions', () => {
      context.settings = { ...context.settings, links: JSON.stringify([
        { title: 'zebra', url: 'https://z.example.test' },
        { title: 'Alpha', url: 'https://a.example.test' },
        { title: 'middle', url: 'https://m.example.test' }
      ]) };
      settings();
      click('Sort A-Z');
      expect(saved().map((link) => link.title)).toEqual(['Alpha', 'middle', 'zebra']);
      context.settings = updateSettings.mock.calls[0][0];
      settings();
      click('Sort Z-A');
      expect(saved().map((link) => link.title)).toEqual(['zebra', 'middle', 'Alpha']);
    });

    it('pointer-drags filtered rows in the complete saved list and ignores external drops', () => {
      settings();
      const input = container.querySelector<HTMLInputElement>('input[aria-label="Find your links"]');
      if (!input) { throw new Error('Missing authoring search.'); }
      act(() => { input.value = 'TOOL 1'; Simulate.change(input); });
      const rows = container.querySelectorAll<HTMLLIElement>('.authoring-row');
      expect(rows).toHaveLength(4);
      const original = document.elementFromPoint;
      document.elementFromPoint = jest.fn(() => rows[3]);
      Object.assign(rows[0], { setPointerCapture: jest.fn() });
      try {
        act(() => { pointer(rows[1], 'pointerup', 100); });
        expect(updateSettings).not.toHaveBeenCalled();
        act(() => { pointer(rows[0], 'pointerdown', 10); });
        act(() => { pointer(rows[0], 'pointermove', 100); pointer(rows[0], 'pointerup', 100); });
        expect(saved().map((link) => link.title)).toEqual([
          ...links.slice(1).map((link) => link.title), 'Tool 1'
        ]);
        expect(rows[0].setPointerCapture).toHaveBeenCalledWith(1);
      } finally {
        document.elementFromPoint = original;
      }
    });

    it('does not reorder on a click, cancelled pointer drag or a drop outside this list', () => {
      settings();
      const row = container.querySelector<HTMLLIElement>('.authoring-row');
      if (!row) { throw new Error('Missing row.'); }
      const original = document.elementFromPoint;
      document.elementFromPoint = jest.fn(() => null);
      Object.assign(row, { setPointerCapture: jest.fn() });
      try {
        act(() => { pointer(row, 'pointerdown', 10); pointer(row, 'pointerup', 10); });
        expect(updateSettings).not.toHaveBeenCalled();
        act(() => { pointer(row, 'pointerdown', 10); pointer(row, 'pointermove', 100); pointer(row, 'pointercancel', 100); });
        act(() => { pointer(row, 'pointerup', 100); });
        expect(updateSettings).not.toHaveBeenCalled();
        act(() => { pointer(row, 'pointerdown', 10); pointer(row, 'pointermove', 100); pointer(row, 'pointerup', 100); });
        expect(updateSettings).not.toHaveBeenCalled();
      } finally {
        document.elementFromPoint = original;
      }
    });
    it('returns to the list on back and dismisses settings without saving a draft', () => {
      settings();
      click('Add a new link');
      field('Name', 'Unsaved');
      act(() => { Simulate.click(button('Back to My links')); });
      expect(container.querySelector('h2')).toBeNull();
      expect(container.querySelector('form')).toBeNull();
      expect(updateSettings).not.toHaveBeenCalled();
      click('Add a new link');
      field('Name', 'Another unsaved link');
      render();
      expect(updateSettings).not.toHaveBeenCalled();
    });

    it('offers deletion on the separate edit screen with confirmation', () => {
      settings();
      act(() => { Simulate.click(button('Edit Tool 1')); });
      click('Delete');
      expect(updateSettings).not.toHaveBeenCalled();
      click('Confirm delete');
      expect(saved().map((link) => link.title)).toEqual(links.slice(1).map((link) => link.title));
      expect(container.querySelector('form')).toBeNull();
    });

    it('blocks additions at the limit without preventing deletion', () => {
      context.settings = { ...context.settings, links: JSON.stringify(Array.from({ length: 100 }, () => links[0])) };
      settings();
      const add = Array.from(container.querySelectorAll('button'))
        .find((item) => item.querySelector('.ms-Button-label')?.textContent === 'Add a new link');
      expect(add?.disabled).toBe(true);
      expect(button('Delete Tool 1').disabled).toBe(false);
      expect(container.textContent).toContain('at most 100');
    });

    it('opens add authoring from an empty widget in viewing mode and persists the first link', () => {
      const Host: React.FunctionComponent = () => {
        const [editing, setEditing] = React.useState(false);
        const [settings, setSettings] = React.useState<Record<string, unknown>>({ links: '[]', unrelated: 'preserve-me' });
        return <WidgetFrame
          instance={{ id: 'my-links', type: 'custom.myLinks', x: 0, y: 0, w: 12, h: 9 }}
          widgetContext={{ ...context, isEditing: editing, settings, updateSettings: (next) => {
            updateSettings(next); setSettings(next);
          } }}
          isEditing={editing} canEdit onRequestEdit={() => setEditing(true)}
          onRemove={jest.fn()} onNudge={jest.fn()} onUpdateTitle={jest.fn()} />;
      };
      act(() => { ReactDom.render(<Host />, container); });
      const add = container.querySelector('button');
      if (!add) { throw new Error('Missing empty-state add action.'); }
      act(() => { Simulate.click(add); jest.runOnlyPendingTimers(); });
      const panel = document.querySelector('[role="dialog"]');
      expect(panel?.querySelector('h2')?.textContent).toBe('Add a new link');
      expect(panel?.getAttribute('aria-label')).toBe('My links settings');
      expect(document.querySelector('.ms-Callout')).not.toBeNull();
      expect(document.querySelector('.ms-Panel')).toBeNull();
      expect(panel?.textContent).toContain('Title');
      expect(panel?.textContent).toContain('Link tiles');
      expect(panel?.querySelector('ol')).toBeNull();
      expect(panel?.textContent).not.toContain('Widget options');
      ['Name', 'Link'].forEach((label, index) => {
        const id = Array.from(panel?.querySelectorAll('label') ?? [])
          .find((item) => item.textContent?.startsWith(label))?.htmlFor;
        const input = id ? document.getElementById(id) : undefined;
        if (!(input instanceof HTMLInputElement)) { throw new Error(`Missing field: ${label}`); }
        act(() => { input.value = index === 0 ? 'First link' : 'https://first.example.test'; Simulate.change(input); });
      });
      const form = panel?.querySelector('form');
      if (!form) { throw new Error('Missing add form.'); }
      act(() => { Simulate.submit(form); });
      expect(updateSettings).toHaveBeenCalledWith({
        links: '[{"title":"First link","url":"https://first.example.test/"}]', unrelated: 'preserve-me'
      });
      expect(panel?.querySelector('h2')).toBeNull();
      expect(panel?.querySelector('ol')).not.toBeNull();
      expect(container.querySelector('a')?.textContent).toContain('First link');
    });

    it('disables empty-state authoring while persistence blocks editing', () => {
      context.settings = { links: '[]' };
      act(() => { ReactDom.render(<WidgetFrame
        instance={{ id: 'my-links', type: 'custom.myLinks', x: 0, y: 0, w: 12, h: 9 }}
        widgetContext={context} isEditing={false} canEdit={false}
        onRemove={jest.fn()} onNudge={jest.fn()} onUpdateTitle={jest.fn()} />, container); });
      expect(container.querySelector('button')?.disabled).toBe(true);
    });

    it('reports malformed stored configuration and permits repair through the JSON editor', () => {
      context.settings = { ...context.settings, links: '{broken' };
      settings();
      const add = Array.from(container.querySelectorAll('button'))
        .find((item) => item.querySelector('.ms-Button-label')?.textContent === 'Add a new link');
      expect(add?.disabled).toBe(true);
      expect(container.textContent).toContain('not valid JSON');
      const input = container.querySelector('textarea');
      if (!input) { throw new Error('Missing Links JSON editor.'); }
      act(() => { input.value = '[]'; Simulate.change(input); Simulate.blur(input); });
      expect(saved()).toEqual([]);
    });
  });

  it.each([[1e200, 12], [-5, 3], [4.8, 4], [undefined, 9]])(
    'bounds links per page (%p)', (itemsPerPage, expected) => {
      context.settings = { ...context.settings, itemsPerPage };
      render();
      expect(container.querySelectorAll('a')).toHaveLength(expected);
    }
  );

  it('allows My Links to shrink in both dimensions while keeping its initial size', () => {
    const definition = WidgetRegistry.get('custom.myLinks');
    expect(definition?.defaultSize).toEqual({ w: 12, h: 9 });
    expect(definition?.minSize).toEqual({ w: 3, h: 4 });
  });

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
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain('The list of your favorite links');
    expect(document.querySelector('[role="dialog"]')?.textContent).toContain('Link tiles');
    expect(document.querySelector('.ms-Callout')).not.toBeNull();
    expect(document.querySelector('.ms-Panel')).toBeNull();
    expect(container.querySelector('button[aria-label="Refresh Work tools"]')).toBeNull();
  });
});
