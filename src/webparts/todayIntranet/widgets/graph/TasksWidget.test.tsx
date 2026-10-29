/* eslint-disable @rushstack/pair-react-dom-render-unmount -- The shared test container is unmounted in afterEach. */

jest.mock('./TasksWidget.module.scss', () => ({ tasks: 'tasks', relaxed: 'relaxed', tags: 'tags', due: 'due' }), { virtual: true });
jest.mock('../content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetCollections.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetAdaptiveCard', () => ({ WidgetAdaptiveCard: () => null }));
jest.mock('./useTaskOrganization', () => ({ useTaskOrganization: jest.fn() }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import { clearWidgetDataCache } from '../data/WidgetDataCache';
import { TasksWidget, TASKS_DEFAULT_VIEW, TASKS_VIEWS } from './TasksWidget';
import { ViewSetting } from '../content/WidgetSettings';
import { useTaskOrganization } from './useTaskOrganization';
import { TaskOrganizationStore } from '../../services/TaskOrganizationStore';
import { emptyTaskDocument, TaskDocument } from './taskOrganization';

describe('task table and organization interactions', () => {
  type Request = ReturnType<MSGraphClientV3['api']>;
  let container: HTMLDivElement;
  let context: IWidgetContext;
  let documentState: TaskDocument;
  const get = jest.fn<Promise<unknown>, [string]>();
  const api = jest.fn<Request, [string]>((path) => {
    const request = new (jest.fn<Request, []>())();
    Object.assign(request, { version: () => request, select: () => request, header: () => request,
      filter: () => request, top: () => request, get: () => get(path) });
    return request;
  });
  const client = new (jest.fn<MSGraphClientV3, []>())();
  Object.assign(client, { api });
  const spContext = new (jest.fn<WebPartContext, []>())();
  Object.defineProperties(spContext, {
    pageContext: { value: { cultureInfo: { currentUICultureName: 'en-GB' }, aadInfo: { userId: 'user', tenantId: 'tenant' },
      user: { loginName: 'test@example.test' }, site: { id: 'site' }, web: { id: 'web' } } },
    msGraphClientFactory: { value: { getClient: async () => client } }
  });
  const store = new (jest.fn<TaskOrganizationStore, []>())();
  const edit = jest.fn<void, [TaskDocument]>((value) => { documentState = value; });
  Object.assign(store, { edit, revision: () => '1790588000001:test' });
  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => {
    jest.clearAllMocks();
    clearWidgetDataCache();
    documentState = emptyTaskDocument();
    jest.mocked(useTaskOrganization).mockImplementation(() => ({
      store, state: { document: documentState, status: 'saved', message: 'Saved to SharePoint.', canEdit: true }
    }));
    context = { instanceId: 'tasks', spContext, settings: {}, isEditing: false, refreshToken: 0, updateSettings: jest.fn() };
    get.mockImplementation(async (path) => path === '/me/todo/lists' ? { value: [{ id: 'list' }] }
      : path === '/me/planner/tasks' ? { value: [{ id: 'p', title: 'Planner work', percentComplete: 0 }] }
        : path === '/me/messages' ? { value: [{ id: 'mail', subject: 'Flagged email', flag: { flagStatus: 'flagged' } }] }
          : { value: [
            { id: 'undated', title: 'Undated task', status: 'notStarted' },
            { id: 'later', title: 'Later task', status: 'notStarted', dueDateTime: { dateTime: '2026-09-30T00:00:00', timeZone: 'UTC' } },
            { id: 'early', title: 'Early task', status: 'notStarted', dueDateTime: { dateTime: '2026-09-01T00:00:00', timeZone: 'UTC' } }
          ] });
    container = window.document.createElement('div');
    window.document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    clearWidgetDataCache();
  });
  const render = async (): Promise<void> => {
    await act(async () => { ReactDom.render(<TasksWidget context={context} />, container); });
  };
  const button = (label: string): HTMLButtonElement => {
    const found = Array.from(container.querySelectorAll('button')).find((element) =>
      element.getAttribute('aria-label') === label || element.textContent?.trim() === label);
    if (!found) { throw new Error(`Button not found: ${label}`); }
    return found;
  };
  const click = async (label: string): Promise<void> => {
    await act(async () => { Simulate.click(button(label)); });
  };
  const titles = (): string[] => Array.from(container.querySelectorAll('td:nth-child(2) span[title]')).map((element) => element.getAttribute('title') ?? '');
  const pageNumbers = (): number[] => Array.from(container.querySelectorAll('nav button[aria-label^="Tasks page "]'))
    .map((element) => Number(element.textContent));
  const currentPage = (): number => Number(container.querySelector('[aria-current="page"]')?.textContent);
  const goTo = async (page: number): Promise<void> => {
    for (let step = 0; step < 30 && currentPage() !== page; step++) {
      await click(pageNumbers().indexOf(page) >= 0 ? `Tasks page ${page}`
        : currentPage() < page ? 'Next tasks page' : 'Previous tasks page');
    }
    expect(currentPage()).toBe(page);
  };

  it('defaults old tiles to all To Do lists only and renders a read-only compact table', async () => {
    await render();
    expect(titles()).toEqual(['Early task', 'Later task', 'Undated task']);
    expect(api).not.toHaveBeenCalledWith('/me/messages');
    expect(api).not.toHaveBeenCalledWith('/me/planner/tasks');
    expect(container.querySelector('input[type="checkbox"]')).toBeNull();
    expect(container.querySelector('th:nth-child(3)')?.textContent).toBe('Quick look');
    expect(container.querySelector('a')?.textContent).toContain('To Do app');
  });
  it('shows no organization status, explanation or management controls when saved', async () => {
    await render();
    expect(container.querySelector('details')).toBeNull();
    expect(container.textContent).not.toContain('Saved to SharePoint');
    expect(container.textContent).not.toContain('Site owners');
    expect(container.textContent).not.toContain('Export recovery data');
    expect(titles()).toHaveLength(3);
  });
  it('keeps organization management available in dashboard edit mode', async () => {
    context.isEditing = true;
    await render();
    expect(container.querySelector('summary')?.textContent).toBe('Manage personal organization');
    expect(button('Export recovery data')).toBeDefined();
    expect(button('Clear organization')).toBeDefined();
    expect(container.textContent).not.toContain('Saved to SharePoint');
  });
  it.each(['loading', 'pending', 'local', 'error', 'conflict'] as const)(
    'keeps the %s organization notice visible and removes it once saved', async (status) => {
      jest.mocked(useTaskOrganization).mockReturnValue({
        store, state: { document: documentState, status, message: `${status} notice`, canEdit: false }
      });
      await render();
      expect(container.querySelector('summary [role="status"]')?.textContent).toBe(`${status} notice`);
      if (status === 'conflict') {
        expect(button('Keep this browser organization')).toBeDefined();
        expect(button('Use SharePoint organization')).toBeDefined();
      }
      jest.mocked(useTaskOrganization).mockReturnValue({
        store, state: { document: documentState, status: 'saved', message: 'Saved to SharePoint.', canEdit: true }
      });
      await render();
      expect(container.querySelector('details')).toBeNull();
      expect(container.textContent).not.toContain('Saved to SharePoint');
    }
  );
  it.each([[1e200, 3], [-10, 1], [1.9, 1], [undefined, 3]])('bounds display count %p after due sorting', async (maxItems, count) => {
    context.settings = { maxItems };
    await render();
    expect(titles()).toHaveLength(count);
    expect(titles()[0]).toBe('Early task');
  });
  it('filters before the display limit and does not refetch when maxItems changes', async () => {
    context.settings = { dueOnly: true, maxItems: 1 };
    await render();
    expect(titles()).toEqual(['Early task']);
    context.settings = { dueOnly: true, maxItems: 2 };
    await render();
    expect(titles()).toEqual(['Early task', 'Later task']);
    expect(get).toHaveBeenCalledTimes(2);
  });
  it('pages all loaded tasks and resets pagination after filtering without refetching', async () => {
    context.settings = { maxItems: 1, paginationEnabled: true, itemsPerPage: 2 };
    await render();
    expect(titles()).toEqual(['Early task', 'Later task']);
    await click('Next tasks page');
    expect(titles()).toEqual(['Undated task']);
    context.settings = { ...context.settings, dueOnly: true };
    await render();
    expect(titles()).toEqual(['Early task', 'Later task']);
    expect(container.querySelector('nav')).toBeNull();
    expect(get).toHaveBeenCalledTimes(2);
  });
  it('toggles sources accessibly without persisting temporary navigation', async () => {
    await render();
    button('Planner').focus();
    await click('Planner');
    expect(window.document.activeElement).toBe(button('Planner'));
    expect(button('Planner').getAttribute('aria-pressed')).toBe('true');
    expect(titles()).toContain('Planner work');
    expect(context.updateSettings).not.toHaveBeenCalled();
    await click('All');
    expect(titles()).toContain('Flagged email');
  });
  const pagedTasks = (start: number, count: number): { id: string; title: string; status: string }[] =>
    Array.from({ length: count }, (_, index) => ({
      id: `00${start + index}`.slice(-3), title: `Task ${`00${start + index}`.slice(-3)}`, status: 'notStarted'
    }));
  const mockTaskPages = (): void => {
    get.mockImplementation(async (path) => {
      if (path === '/me/todo/lists') { return { value: [{ id: 'list' }] }; }
      if (path === '/me/todo/lists/list/tasks') { return { value: pagedTasks(0, 50), '@odata.nextLink': '/tasks-50' }; }
      if (path === '/tasks-50') { return { value: pagedTasks(50, 50), '@odata.nextLink': '/tasks-100' }; }
      if (path === '/tasks-100') { return { value: pagedTasks(100, 3) }; }
      throw new Error(`Unexpected request: ${path}`);
    });
  };
  it('shows four page numbers, advertises five extra pages and only loads a batch for an unloaded page', async () => {
    context.settings = { maxItems: 10 };
    mockTaskPages();
    await render();
    expect(get).toHaveBeenCalledTimes(2);
    expect(container.textContent).toContain('Page 1 of 10');
    expect(pageNumbers()).toEqual([1, 2, 3, 4]);
    await goTo(5);
    expect(titles()).toEqual(pagedTasks(40, 10).map((task) => task.title));
    expect(pageNumbers()).toEqual([4, 5, 6, 7]);
    expect(get).toHaveBeenCalledTimes(2);
    await click('Tasks page 6');
    expect(titles()).toEqual(pagedTasks(50, 10).map((task) => task.title));
    expect(get).toHaveBeenCalledTimes(3);
    expect(container.textContent).toContain('Page 6 of 15');
    await goTo(11);
    expect(titles()).toEqual(pagedTasks(100, 3).map((task) => task.title));
    expect(container.textContent).toContain('Page 11 of 11');
    expect(pageNumbers()).toEqual([8, 9, 10, 11]);
    expect(button('Next tasks page').disabled).toBe(true);
    await goTo(5);
    expect(get).toHaveBeenCalledTimes(4);
  });
  it('loads no continuation when pagination is off', async () => {
    context.settings = { maxItems: 10, paginationEnabled: false };
    mockTaskPages();
    await render();
    expect(titles()).toHaveLength(10);
    expect(container.querySelector('nav')).toBeNull();
    expect(get).toHaveBeenCalledTimes(2);
  });
  it('loads multiple batches for a saved large page size without repeating a continuation', async () => {
    context.settings = { maxItems: 10, paginationEnabled: true, itemsPerPage: 100 };
    mockTaskPages();
    await render();
    expect(titles()).toHaveLength(50);
    expect(get).toHaveBeenCalledTimes(2);
    await click('Tasks page 2');
    expect(titles()).toEqual(pagedTasks(100, 3).map((task) => task.title));
    expect(get.mock.calls.map(([path]) => path)).toEqual([
      '/me/todo/lists', '/me/todo/lists/list/tasks', '/tasks-50', '/tasks-100'
    ]);
  });
  it('starts again at the first batch after refresh instead of following an old cursor', async () => {
    context.settings = { maxItems: 10 };
    mockTaskPages();
    await render();
    await goTo(5);
    await click('Tasks page 6');
    context.refreshToken++;
    await render();
    expect(get.mock.calls.map(([path]) => path)).toEqual([
      '/me/todo/lists', '/me/todo/lists/list/tasks', '/tasks-50', '/me/todo/lists', '/me/todo/lists/list/tasks'
    ]);
  });
  it('clamps a provisional page when the final response contains fewer tasks', async () => {
    context.settings = { maxItems: 10 };
    mockTaskPages();
    get.mockImplementation(async (path) => path === '/me/todo/lists' ? { value: [{ id: 'list' }] }
      : path === '/tasks-50' ? { value: [] }
        : { value: pagedTasks(0, 50), '@odata.nextLink': '/tasks-50' });
    await render();
    await goTo(5);
    await click('Tasks page 7');
    expect(container.textContent).toContain('Page 5 of 5');
    expect(titles()).toEqual(pagedTasks(40, 10).map((task) => task.title));
  });
  it('keeps loaded tasks and the current page visible after a continuation failure and allows retry', async () => {
    context.settings = { maxItems: 10 };
    mockTaskPages();
    await render();
    await goTo(5);
    get.mockRejectedValueOnce(new Error('Throttled'));
    await click('Tasks page 6');
    expect(container.textContent).toContain('Throttled');
    expect(container.textContent).toContain('Page 5 of 10');
    expect(titles()).toHaveLength(10);
    await click('Tasks page 6');
    expect(container.textContent).not.toContain('Throttled');
    expect(titles()).toEqual(pagedTasks(50, 10).map((task) => task.title));
  });
  it('keeps healthy tasks visible when another source is denied', async () => {
    await render();
    get.mockImplementation(async () => { throw Object.assign(new Error('Denied'), { statusCode: 403 }); });
    await click('Planner');
    expect(titles()).toContain('Early task');
    expect(container.textContent).toContain('Tasks.Read');
    expect(container.textContent).toContain('Retry Planner');
  });
  it('edits priority and adds, renames and removes tags inline without layout saves or Graph writes', async () => {
    await render();
    const select = container.querySelector<HTMLSelectElement>('select[aria-label="Priority for Early task"]')!;
    await act(async () => { select.value = 'high'; Simulate.change(select); });
    await render();
    await click('Add tag for Early task');
    const input = container.querySelector<HTMLInputElement>('input')!;
    expect(window.document.activeElement).toBe(input);
    expect(container.querySelector('form')).toBeNull();
    await act(async () => {
      input.value = ' Follow  up '; Simulate.change(input);
    });
    await act(async () => { Simulate.keyDown(input, { key: 'Enter' }); });
    expect(edit).toHaveBeenCalledWith(expect.objectContaining({ tasks: expect.objectContaining({
      'todo:list:early': expect.objectContaining({ priority: 'high', tags: ['Follow up'] })
    }) }));
    await render();
    await click('Edit tag Follow up for Early task');
    const renameInput = container.querySelector<HTMLInputElement>('input')!;
    await act(async () => { renameInput.value = 'This week'; Simulate.change(renameInput); });
    await act(async () => { Simulate.blur(renameInput); });
    await render();
    expect(documentState.tasks['todo:list:early'].tags).toEqual(['This week']);
    await click('Remove tag This week from Early task');
    await render();
    expect(documentState.tasks['todo:list:early'].tags).toEqual([]);
    expect(context.updateSettings).not.toHaveBeenCalled();
    expect(get).toHaveBeenCalledTimes(2);
  });
  it('cancels inline tag editing with Escape and reports invalid tags without discarding the input', async () => {
    await render();
    await click('Add tag for Early task');
    const input = container.querySelector<HTMLInputElement>('input')!;
    await act(async () => { input.value = 'x'.repeat(33); Simulate.change(input); });
    await act(async () => { Simulate.keyDown(input, { key: 'Enter' }); });
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('32');
    expect(input.value).toHaveLength(33);
    expect(edit).not.toHaveBeenCalled();
    await act(async () => { Simulate.keyDown(input, { key: 'Escape' }); Simulate.blur(input); });
    expect(container.querySelector('input')).toBeNull();
    expect(edit).not.toHaveBeenCalled();
    await click('Add tag for Early task');
    const emptyInput = container.querySelector<HTMLInputElement>('input')!;
    await act(async () => { Simulate.blur(emptyInput); });
    expect(edit).not.toHaveBeenCalled();
  });
  it('keeps tag limits and priority/order intact while renaming and removing inline', async () => {
    documentState.tasks['todo:list:early'] = {
      priority: 'high', order: 3, tags: ['One', 'Two', 'Three', 'Four', 'Five'], modified: '1790588000000:test'
    };
    await render();
    expect(button('Add tag for Early task').disabled).toBe(true);
    await click('Edit tag Two for Early task');
    const input = container.querySelector<HTMLInputElement>('input')!;
    await act(async () => { input.value = ' ONE '; Simulate.change(input); });
    await act(async () => { Simulate.keyDown(input, { key: 'Enter' }); });
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('already a tag');
    expect(documentState.tasks['todo:list:early'].tags).toEqual(['One', 'Two', 'Three', 'Four', 'Five']);
    expect(container.querySelector('input')).toBe(input);
    await act(async () => { input.value = 'Six'; Simulate.change(input); });
    await act(async () => { Simulate.keyDown(input, { key: 'Enter' }); });
    await render();
    expect(documentState.tasks['todo:list:early']).toMatchObject({
      priority: 'high', order: 3, tags: ['One', 'Six', 'Three', 'Four', 'Five']
    });
    // Focus returns to the renamed tag rather than the first chip in the cell.
    expect(window.document.activeElement).toBe(button('Edit tag Six for Early task'));
    await click('Remove tag One from Early task');
    await render();
    expect(documentState.tasks['todo:list:early'].tags).toEqual(['Six', 'Three', 'Four', 'Five']);
    expect(button('Add tag for Early task').disabled).toBe(false);
  });
  it('restores the add button when the tag being renamed disappears from another tile', async () => {
    documentState.tasks['todo:list:early'] = { priority: 'normal', tags: ['Shared'], modified: '1790588000000:test' };
    await render();
    await click('Edit tag Shared for Early task');
    expect(container.querySelector('input')).not.toBeNull();
    documentState = { version: 1, tasks: {
      'todo:list:early': { priority: 'normal', tags: [], modified: '1790588000002:other' }
    } };
    await render();
    expect(container.querySelector('input')).toBeNull();
    expect(button('Add tag for Early task').disabled).toBe(false);
    await click('Add tag for Early task');
    expect(container.querySelector('input')).not.toBeNull();
  });
  it('offers keyboard ordering on drag handles and adopts the destination priority', async () => {
    documentState.tasks['todo:list:early'] = { priority: 'high', tags: [], modified: '1790588000000:test' };
    await render();
    button('Move Later task').focus();
    await act(async () => { Simulate.keyDown(button('Move Later task'), { key: 'ArrowUp' }); });
    await render();
    expect(titles()[0]).toBe('Later task');
    expect(documentState.tasks['todo:list:later'].priority).toBe('high');
    expect(window.document.activeElement).toBe(button('Move Later task'));
    expect(container.textContent).toContain('Moved Later task before Early task; high priority');
  });
  it('drags to a precise insertion point and cancels without writing', async () => {
    await render();
    const handle = button('Move Undated task');
    const row = button('Move Early task').closest('tr')!;
    Object.defineProperty(handle, 'setPointerCapture', { value: jest.fn(), configurable: true });
    const originalHitTest = window.document.elementFromPoint;
    Object.defineProperty(window.document, 'elementFromPoint', { value: jest.fn(() => row), configurable: true });
    const pointer = (type: string, clientY: number): void => {
      const event = new MouseEvent(type, { bubbles: true, button: 0, clientX: 10, clientY });
      Object.defineProperties(event, { isPrimary: { value: true }, pointerId: { value: 1 } });
      handle.dispatchEvent(event);
    };
    try {
      await act(async () => { pointer('pointerdown', 100); });
      await act(async () => { pointer('pointermove', -10); });
      await act(async () => { pointer('pointercancel', -10); pointer('pointerup', -10); });
      expect(edit).not.toHaveBeenCalled();
      await act(async () => { pointer('pointerdown', 100); });
      await act(async () => { pointer('pointermove', -10); });
      await act(async () => { pointer('pointerup', -10); });
    } finally {
      Object.defineProperty(window.document, 'elementFromPoint', { value: originalHitTest, configurable: true });
    }
    await render();
    expect(titles()).toEqual(['Undated task', 'Early task', 'Later task']);
    expect(edit).toHaveBeenCalledTimes(1);
  });
  it('shows organization errors explicitly and disables edits instead of overwriting defaults', async () => {
    jest.mocked(useTaskOrganization).mockReturnValue({ state: {
      document: emptyTaskDocument(), status: 'error', canEdit: false, message: 'Saved organization could not be loaded.'
    } });
    await render();
    expect(container.textContent).toContain('could not be loaded');
    expect(container.querySelector<HTMLSelectElement>('select[aria-label="Priority for Early task"]')?.disabled).toBe(true);
    expect(button('Add tag for Early task').disabled).toBe(true);
    expect(button('Move Early task').disabled).toBe(true);
    expect(titles()).toHaveLength(3);
  });
  it('renders legacy unsupported views as the compact table without mutating settings', async () => {
    context.settings = { view: 'adaptive' };
    await render();
    expect(container.querySelector('table')).not.toBeNull();
    expect(context.settings.view).toBe('adaptive');
    expect(context.updateSettings).not.toHaveBeenCalled();
  });
  it('shows the compact fallback selected in settings for a legacy unsupported view', async () => {
    context.settings = { view: 'adaptive' };
    await act(async () => { ReactDom.render(
      <ViewSetting context={context} views={TASKS_VIEWS} fallback={TASKS_DEFAULT_VIEW} />, container
    ); });
    expect(container.querySelectorAll('button')[0].getAttribute('aria-pressed')).toBe('true');
    expect(context.updateSettings).not.toHaveBeenCalled();
  });
});
