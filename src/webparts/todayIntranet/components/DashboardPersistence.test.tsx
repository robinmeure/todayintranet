/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */
/* eslint-disable @typescript-eslint/no-use-before-define -- Mock factories reference imports only when React renders them. */

import type { Layout, Layouts } from 'react-grid-layout';
import type { IWidgetFrameProps } from './WidgetFrame';

interface ITestGridProps {
  children?: React.ReactNode;
  layouts: Layouts;
  isDraggable: boolean;
  isResizable: boolean;
  onLayoutChange(layout: Layout[]): void;
  onBreakpointChange(breakpoint: string): void;
}

const mockFrames = new Map<string, IWidgetFrameProps>();
let mockGrid: ITestGridProps;

jest.mock('react-grid-layout', () => ({
  Responsive: (props: ITestGridProps) => {
    mockGrid = props;
    return props.children;
  },
  WidthProvider: (component: React.ComponentType) => component
}));
jest.mock('./Dashboard.module.scss', () => ({}), { virtual: true });
jest.mock('../widgets/content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../widgets/WidgetRegistry', () => ({ WidgetRegistry: { get: () => undefined } }));
jest.mock('./AddWidgetPanel', () => ({ AddWidgetPanel: () => null }));
jest.mock('./LayoutPresetPanel', () => ({ LayoutPresetPanel: () => null }));
jest.mock('./WidgetFrame', () => ({
  WidgetFrame: (props: IWidgetFrameProps) => {
    mockFrames.set(props.instance.id, props);
    return props.isEditing ? (
      <TextSetting
        context={{ ...props.widgetContext, refreshToken: 0 }}
        settingKey="draft"
        label="Draft setting"
        fallback=""
        commitOn="blur"
      />
    ) : null;
  }
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { Dashboard } from './Dashboard';
import { TextSetting } from '../widgets/content/WidgetSettings';
import { IDashboardLayout, cloneLayout, CURRENT_LAYOUT_VERSION } from '../model/IDashboardLayout';
import { ILayoutStore, ILayoutStoreStatus, LayoutStoreAction } from '../services/ILayoutStore';

const spContext = new (jest.fn<WebPartContext, []>())();
const starter: IDashboardLayout = {
  version: CURRENT_LAYOUT_VERSION,
  presetId: 'two',
  rowSizeId: 'medium',
  widgets: [{ id: 'one', type: 'test.clock', x: 0, y: 0, w: 6, h: 6, settings: { view: 'list' } }]
};
const recovered: IDashboardLayout = {
  version: CURRENT_LAYOUT_VERSION,
  presetId: 'three',
  rowSizeId: 'short',
  widgets: [{ id: 'recovered', type: 'test.clock', title: 'Recovered', x: 0, y: 0, w: 4, h: 4 }]
};

class TestStore implements ILayoutStore {
  public status: ILayoutStoreStatus = {
    mode: 'local', state: 'saved', message: 'Saved in this browser.', canEdit: true, actions: []
  };
  public disposed = false;
  public events: string[] = [];
  public subscriptions: Array<() => void> = [];
  private listeners = new Set<() => void>();
  public load = jest.fn<Promise<IDashboardLayout | undefined>, []>(() => Promise.resolve(cloneLayout(starter)));
  public save = jest.fn<Promise<void>, [IDashboardLayout]>(() => {
    if (this.disposed) {
      throw new Error('Save after disposal');
    }
    this.events.push('save');
    return Promise.resolve();
  });
  public publish = jest.fn<Promise<void>, []>(() => {
    this.events.push('publish');
    return Promise.resolve();
  });
  public resolve = jest.fn<Promise<IDashboardLayout | undefined>, [LayoutStoreAction, IDashboardLayout?]>(
    () => Promise.resolve(cloneLayout(recovered))
  );
  public exportRecovery = jest.fn<string, []>(() => '{"recovery":true}');
  public dispose = jest.fn<void, []>(() => {
    this.events.push('dispose');
    this.disposed = true;
  });
  public getStatus(): ILayoutStoreStatus {
    return this.status;
  }
  public subscribe(listener: () => void): () => void {
    this.subscriptions.push(listener);
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
  public emitStatus(status: Partial<ILayoutStoreStatus>): void {
    this.status = { ...this.status, ...status };
    this.listeners.forEach((listener) => listener());
  }
}

function deferred<T>(): { promise: Promise<T>; resolve(value: T): void; reject(reason: Error): void } {
  let resolve!: (value: T) => void;
  let reject!: (reason: Error) => void;
  const promise = new Promise<T>((_resolve, _reject) => { resolve = _resolve; reject = _reject; });
  return { promise, resolve, reject };
}

describe('Dashboard persistence and recovery', () => {
  let container: HTMLDivElement;
  let store: TestStore;

  const frame = (id = 'one'): IWidgetFrameProps => {
    const result = mockFrames.get(id);
    if (!result) {
      throw new Error(`Frame not rendered: ${id}`);
    }
    return result;
  };
  const button = (text: string): HTMLButtonElement => {
    const result = Array.from(document.querySelectorAll('button'))
      .find((item) => item.querySelector('.ms-Button-label')?.textContent === text);
    if (!result) {
      throw new Error(`Button not rendered: ${text}`);
    }
    return result;
  };
  const render = async (nextStore: TestStore = store): Promise<void> => {
    await act(async () => {
      ReactDom.render(<Dashboard title="Dashboard" spContext={spContext} store={nextStore} starterLayout={starter} />, container);
    });
    act(() => { jest.runOnlyPendingTimers(); });
  };
  const click = async (text: string): Promise<void> => {
    await act(async () => { Simulate.click(button(text)); });
    act(() => { jest.runOnlyPendingTimers(); });
  };
  const changeDraft = (value: string): void => {
    const input = container.querySelector('input');
    if (!input) {
      throw new Error('Draft input is missing.');
    }
    act(() => {
      input.value = value;
      Simulate.change(input);
    });
  };

  beforeAll(() => { initializeIcons(undefined, { disableWarnings: true }); });
  beforeEach(() => {
    jest.useFakeTimers();
    mockFrames.clear();
    store = new TestStore();
    container = document.createElement('div');
    document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    jest.restoreAllMocks();
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  it('checkpoints each commit immediately without advancing a debounce timer', async () => {
    await render();
    await act(async () => { frame().onUpdateTitle('one', 'Immediate title'); });
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(store.save.mock.calls[0][0].widgets[0].title).toBe('Immediate title');
  });

  it('hides routine browser-only status', async () => {
    await render();
    expect(container.querySelector('[role="status"]')).toBeNull();
    expect(container.querySelector('[role="alert"]')).toBeNull();
  });

  it.each(['reject', 'throw'])('shows a visible error when save fails (%s)', async (failure) => {
    store.save.mockImplementation(() => {
      if (failure === 'throw') {
        throw new Error('Storage failed');
      }
      return Promise.reject(new Error('Storage failed'));
    });
    await render();
    await act(async () => { frame().onUpdateTitle('one', 'Unsaved'); });
    act(() => { jest.runOnlyPendingTimers(); });
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('could not be saved');
    expect(frame().instance.title).toBe('Unsaved');
  });

  it('hides pending status changes without replacing the current React layout', async () => {
    await render();
    await act(async () => { frame().onUpdateTitle('one', 'Local edit'); });
    act(() => {
      store.emitStatus({ state: 'pending', message: 'Network unavailable.', actions: ['retry', 'export'] });
    });
    act(() => { jest.runOnlyPendingTimers(); });
    expect(container.querySelector('[role="alert"]')).toBeNull();
    expect(container.textContent).not.toContain('Network unavailable.');
    expect(frame().instance.title).toBe('Local edit');
    expect(store.resolve).not.toHaveBeenCalled();
  });

  it('blocks edit controls and automatic grid writes when loading invalid data fails', async () => {
    store.status = { ...store.status, state: 'error', reason: 'invalid-data', canEdit: false, actions: ['reset', 'export'] };
    store.load.mockRejectedValue(new Error('Invalid data'));
    await render();
    expect(button('Edit dashboard').disabled).toBe(true);
    expect(mockGrid.isDraggable).toBe(false);
    expect(mockGrid.isResizable).toBe(false);
    act(() => {
      mockGrid.onLayoutChange([{ i: 'one', x: 2, y: 2, w: 4, h: 4 }]);
      frame().onUpdateTitle('one', 'Must not save');
      frame().onNudge('one', { dx: 1 });
    });
    expect(store.save).not.toHaveBeenCalled();
    expect(frame().instance.title).toBeUndefined();
    expect(button('Export recovery data').disabled).toBe(false);
  });

  it('ignores an old scope load and status callback after the store changes', async () => {
    const oldLoad = deferred<IDashboardLayout | undefined>();
    store.load.mockReturnValue(oldLoad.promise);
    await render();
    const next = new TestStore();
    next.load.mockResolvedValue(cloneLayout(recovered));
    await render(next);
    await act(async () => {
      oldLoad.resolve(cloneLayout(starter));
      store.status = { ...store.status, state: 'error', message: 'Stale failure' };
      store.subscriptions[0]();
    });
    expect(frame('recovered').instance.title).toBe('Recovered');
    expect(mockGrid.layouts.lg[0].i).toBe('recovered');
    expect(container.textContent).not.toContain('Stale failure');
    expect(store.dispose).toHaveBeenCalledTimes(1);
    expect(next.dispose).not.toHaveBeenCalled();
  });

  it('ignores an old scope save rejection after the store changes', async () => {
    const save = deferred<void>();
    store.save.mockReturnValue(save.promise);
    await render();
    act(() => { frame().onUpdateTitle('one', 'Old scope'); });
    const next = new TestStore();
    await render(next);
    await act(async () => { save.reject(new Error('Old failure')); });
    expect(container.textContent).not.toContain('could not be saved');
    expect(frame().instance.title).toBeUndefined();
  });

  it('does not let an older failed save replace the status of a newer successful commit', async () => {
    const firstSave = deferred<void>();
    store.save.mockReturnValueOnce(firstSave.promise);
    await render();
    await act(async () => {
      frame().onUpdateTitle('one', 'Older edit');
      frame().onUpdateTitle('one', 'Newer edit');
    });
    await act(async () => { firstSave.reject(new Error('Old write failed')); });
    act(() => { jest.runOnlyPendingTimers(); });
    expect(container.textContent).not.toContain('could not be saved');
    expect(frame().instance.title).toBe('Newer edit');
  });

  it('ignores a recovery completion from the previous dashboard scope', async () => {
    const resolution = deferred<IDashboardLayout | undefined>();
    store.status = { ...store.status, state: 'error', actions: ['retry'] };
    store.resolve.mockReturnValue(resolution.promise);
    await render();
    await click('Retry');
    expect(button('Edit dashboard').disabled).toBe(true);
    const next = new TestStore();
    await render(next);
    await act(async () => { resolution.resolve(cloneLayout(recovered)); });
    expect(mockGrid.layouts.lg[0].i).toBe('one');
    expect(button('Edit dashboard').disabled).toBe(false);
    expect(next.dispose).not.toHaveBeenCalled();
  });

  it.each(['use-remote', 'keep-local'] as const)('offers explicit conflict choice %s and applies all layout metadata', async (action) => {
    store.status = { ...store.status, mode: 'sharepoint', state: 'conflict', canEdit: false, actions: ['use-remote', 'keep-local'] };
    await render();
    store.resolve.mockImplementation(async () => {
      store.emitStatus({ state: 'saved', canEdit: true, actions: [] });
      return cloneLayout(recovered);
    });
    await click(action === 'use-remote' ? 'Use saved version' : 'Keep my version');
    expect(store.resolve).toHaveBeenCalledWith(action, action === 'use-remote' ? starter : undefined);
    await act(async () => { frame('recovered').onUpdateTitle('recovered', 'Next edit'); });
    expect(store.save.mock.calls[0][0]).toEqual({
      ...recovered, widgets: [{ ...recovered.widgets[0], title: 'Next edit' }]
    });
  });

  it('provides the starter when the selected saved version is absent', async () => {
    store.status = { ...store.status, mode: 'sharepoint', state: 'conflict', canEdit: false, actions: ['use-remote'] };
    store.load.mockResolvedValue(cloneLayout(recovered));
    store.resolve.mockImplementation(async (_action, starterLayout) => {
      store.emitStatus({ state: 'saved', canEdit: true, actions: [] });
      return starterLayout;
    });
    await render();
    await click('Use saved version');
    expect(store.resolve).toHaveBeenCalledWith('use-remote', starter);
    expect(mockGrid.layouts.lg[0].i).toBe('one');
    await act(async () => { frame().onUpdateTitle('one', 'Starter title'); });
    expect(store.save.mock.calls[0][0]).toEqual({
      ...starter, widgets: [{ ...starter.widgets[0], title: 'Starter title' }]
    });
  });

  it('requires confirmation for recovery reset and supports cancellation', async () => {
    store.status = { ...store.status, state: 'error', canEdit: false, actions: ['reset'] };
    await render();
    await click('Reset layout');
    expect(store.resolve).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain('may discard recoverable changes');
    await click('Cancel');
    expect(store.resolve).not.toHaveBeenCalled();
    await click('Reset layout');
    await click('Reset this dashboard');
    expect(store.resolve).toHaveBeenCalledWith('reset', starter);
  });

  it('confirms the unknown legacy source site and possible SharePoint sync before importing', async () => {
    store.status = { ...store.status, mode: 'sharepoint', state: 'error', actions: ['import-legacy'] };
    await render();
    await click('Import legacy layout');
    expect(store.resolve).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain('unknown source site');
    expect(document.body.textContent).toContain('may sync to this dashboard');
    await click('Import into this dashboard');
    expect(store.resolve).toHaveBeenCalledWith('import-legacy', undefined);
  });

  it('shows failed recovery actions without replacing the displayed layout', async () => {
    store.status = { ...store.status, state: 'error', actions: ['retry', 'export'] };
    store.resolve.mockRejectedValue(new Error('Still offline'));
    await render();
    await click('Retry');
    expect(container.querySelector('[role="alert"]')?.textContent).toContain('could not be completed');
    expect(mockGrid.layouts.lg[0].i).toBe('one');
    expect(button('Export recovery data').disabled).toBe(false);
  });

  it('exports recovery JSON as a download and revokes the Blob URL', async () => {
    store.status = { ...store.status, state: 'error', actions: ['export'] };
    const create = jest.fn<string, [Blob]>(() => 'blob:recovery');
    const revoke = jest.fn<void, [string]>();
    const oldCreate = Object.getOwnPropertyDescriptor(URL, 'createObjectURL');
    const oldRevoke = Object.getOwnPropertyDescriptor(URL, 'revokeObjectURL');
    Object.defineProperty(URL, 'createObjectURL', { configurable: true, value: create });
    Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: revoke });
    const anchorClick = jest.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
    try {
      await render();
      await click('Export recovery data');
      expect(store.exportRecovery).toHaveBeenCalledTimes(1);
      expect(create.mock.calls[0][0].type).toBe('application/json');
      expect(anchorClick).toHaveBeenCalledTimes(1);
      expect(revoke).toHaveBeenCalledWith('blob:recovery');
      expect(document.querySelector('a[download]')).toBeNull();
      expect(store.resolve).not.toHaveBeenCalled();
    } finally {
      if (oldCreate) { Object.defineProperty(URL, 'createObjectURL', oldCreate); }
      else { delete (URL as Partial<typeof URL>).createObjectURL; }
      if (oldRevoke) { Object.defineProperty(URL, 'revokeObjectURL', oldRevoke); }
      else { delete (URL as Partial<typeof URL>).revokeObjectURL; }
    }
  });

  it('publishes on Done and keeps healthy Reset as a local immediate commit', async () => {
    await render();
    await click('Edit dashboard');
    await click('Done');
    expect(store.save).not.toHaveBeenCalled();
    expect(store.publish).toHaveBeenCalledTimes(1);
    expect(store.resolve).not.toHaveBeenCalled();
    await click('Edit dashboard');
    await click('Reset');
    expect(store.save).not.toHaveBeenCalled();
    await click('Reset this dashboard');
    expect(store.save).toHaveBeenCalledWith(starter);
    expect(store.resolve).not.toHaveBeenCalled();
  });

  it('does not commit an unchanged draft when its settings close', async () => {
    await render();
    await click('Edit dashboard');
    await click('Done');
    expect(store.save).not.toHaveBeenCalled();
    expect(store.publish).toHaveBeenCalledTimes(1);
  });

  it('commits a changed draft on dismissal, without waiting for blur or a timer', async () => {
    await render();
    await click('Edit dashboard');
    changeDraft('Ready to save');
    expect(store.save).not.toHaveBeenCalled();
    await click('Done');
    expect(store.save).toHaveBeenCalledTimes(1);
    expect(store.save.mock.calls[0][0].widgets[0].settings?.draft).toBe('Ready to save');
    expect(store.events).toEqual(['save', 'publish']);
  });

  it('checkpoints a pending draft before disposing the store on dashboard unmount', async () => {
    await render();
    await click('Edit dashboard');
    changeDraft('Last draft');
    act(() => { ReactDom.unmountComponentAtNode(container); });
    expect(store.events).toEqual(['save', 'dispose']);
    expect(store.save.mock.calls[0][0].widgets[0].settings?.draft).toBe('Last draft');
  });

  it('checkpoints a final change even when the field unmounts in the same event batch', async () => {
    await render();
    await click('Edit dashboard');
    const input = container.querySelector('input');
    if (!input) {
      throw new Error('Draft input is missing.');
    }
    act(() => {
      input.value = 'Final event';
      Simulate.change(input);
      ReactDom.unmountComponentAtNode(container);
    });
    expect(store.events).toEqual(['save', 'dispose']);
    expect(store.save.mock.calls[0][0].widgets[0].settings?.draft).toBe('Final event');
  });

  it('does not save twice when blur is immediately followed by unmount', async () => {
    await render();
    await click('Edit dashboard');
    changeDraft('Blurred draft');
    const input = container.querySelector('input');
    if (!input) {
      throw new Error('Draft input is missing.');
    }
    act(() => {
      Simulate.blur(input);
      ReactDom.unmountComponentAtNode(container);
    });
    expect(store.events).toEqual(['save', 'dispose']);
  });

  it('keeps the newest title when a settings draft is committed during teardown', async () => {
    await render();
    await click('Edit dashboard');
    changeDraft('Draft alongside title');
    act(() => {
      frame().onUpdateTitle('one', 'Newest title');
      ReactDom.unmountComponentAtNode(container);
    });
    expect(store.save.mock.calls[1][0].widgets[0]).toMatchObject({
      title: 'Newest title', settings: { draft: 'Draft alongside title', view: 'list' }
    });
    expect(store.events).toEqual(['save', 'save', 'dispose']);
  });

  it('rejects obsolete widget callbacks after the store has been disposed', async () => {
    await render();
    const oldFrame = frame();
    act(() => { ReactDom.unmountComponentAtNode(container); });
    oldFrame.onUpdateTitle('one', 'Too late');
    oldFrame.widgetContext.updateSettings({ draft: 'Too late' });
    expect(store.events).toEqual(['dispose']);
  });

  it('checkpoints an outgoing scope draft into the old store before disposal', async () => {
    await render();
    await click('Edit dashboard');
    changeDraft('Old dashboard draft');
    const next = new TestStore();
    next.load.mockResolvedValue(cloneLayout(recovered));
    await render(next);
    expect(store.events).toEqual(['save', 'dispose']);
    expect(next.save).not.toHaveBeenCalled();
    expect(mockGrid.layouts.lg[0].i).toBe('recovered');
  });

  it('does not save derived narrow-breakpoint layouts', async () => {
    await render();
    act(() => {
      mockGrid.onBreakpointChange('xs');
      mockGrid.onLayoutChange([{ i: 'one', x: 0, y: 0, w: 4, h: 6 }]);
    });
    expect(store.save).not.toHaveBeenCalled();
  });
});
