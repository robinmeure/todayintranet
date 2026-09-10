/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */

import type { IWidgetFrameProps } from './WidgetFrame';

const mockFrames = new Map<string, IWidgetFrameProps>();

jest.mock('react-grid-layout', () => ({
  Responsive: ({ children }: { children?: React.ReactNode }) => children,
  WidthProvider: (component: React.ComponentType) => component
}));
jest.mock('./Dashboard.module.scss', () => ({}), { virtual: true });
jest.mock('../widgets/WidgetRegistry', () => ({ WidgetRegistry: { get: () => undefined } }));
jest.mock('./AddWidgetPanel', () => ({ AddWidgetPanel: () => null }));
jest.mock('./LayoutPresetPanel', () => ({ LayoutPresetPanel: () => null }));
jest.mock('./WidgetFrame', () => ({
  WidgetFrame: (props: IWidgetFrameProps) => {
    mockFrames.set(props.instance.id, props);
    return null;
  }
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { Dashboard } from './Dashboard';
import { IDashboardLayout, cloneLayout, CURRENT_LAYOUT_VERSION } from '../model/IDashboardLayout';
import { applyPreset, getPreset, getRowSize } from '../model/LayoutPresets';
import { LocalStorageLayoutStore } from '../services/LocalStorageLayoutStore';

const spContext = new (jest.fn<WebPartContext, []>())();
const scope = 'widget-title-tests';

const initialLayout: IDashboardLayout = {
  version: CURRENT_LAYOUT_VERSION,
  presetId: 'two',
  rowSizeId: 'medium',
  widgets: [
    { id: 'clock-1', type: 'test.clock', x: 0, y: 0, w: 6, h: 6, settings: { view: 'list' } },
    { id: 'clock-2', type: 'test.clock', title: 'Second clock', x: 6, y: 0, w: 6, h: 6 }
  ]
};

describe('Dashboard widget title persistence', () => {
  let container: HTMLDivElement;
  let store: LocalStorageLayoutStore;

  const frame = (id: string): IWidgetFrameProps => {
    const props = mockFrames.get(id);
    if (!props) {
      throw new Error(`Widget frame not rendered: ${id}`);
    }
    return props;
  };

  const renderDashboard = async (): Promise<void> => {
    await act(async () => {
      ReactDom.render(
        <Dashboard title="Test dashboard" spContext={spContext} store={store} starterLayout={initialLayout} />,
        container
      );
    });
  };

  const flushSave = async (): Promise<void> => {
    await act(async () => {
      jest.advanceTimersByTime(800);
    });
  };

  beforeAll(() => {
    initializeIcons(undefined, { disableWarnings: true });
  });

  beforeEach(async () => {
    jest.useFakeTimers();
    mockFrames.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
    store = new LocalStorageLayoutStore(scope);
    await store.save(cloneLayout(initialLayout));
    await renderDashboard();
  });

  afterEach(() => {
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    container.remove();
    window.localStorage.removeItem(`todayIntranet.layout.${scope}`);
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  it.each(['title-first', 'settings-first'])('preserves simultaneous edits (%s)', async (order) => {
    const original = frame('clock-1');
    const settings = { ...original.widgetContext.settings, caption: 'New content setting' };

    act(() => {
      if (order === 'title-first') {
        original.onUpdateTitle('clock-1', 'Office time');
        original.widgetContext.updateSettings(settings);
      } else {
        original.widgetContext.updateSettings(settings);
        original.onUpdateTitle('clock-1', 'Office time');
      }
    });
    await flushSave();

    const saved = await store.load();
    expect(saved).toEqual({
      ...initialLayout,
      widgets: [
        { ...initialLayout.widgets[0], title: 'Office time', settings },
        initialLayout.widgets[1]
      ]
    });
  });

  it('retains titles after loading the saved layout again', async () => {
    act(() => {
      frame('clock-1').onUpdateTitle('clock-1', 'Office time');
    });
    await flushSave();
    act(() => {
      ReactDom.unmountComponentAtNode(container);
    });
    mockFrames.clear();

    await renderDashboard();

    expect(frame('clock-1').instance.title).toBe('Office time');
    expect(frame('clock-2').instance.title).toBe('Second clock');
    expect(frame('clock-1').widgetContext.settings).toEqual({ view: 'list' });
  });

  it('persists clearing a title without replacing the instance or its settings', async () => {
    act(() => {
      frame('clock-2').onUpdateTitle('clock-2', '');
    });
    await flushSave();

    const saved = await store.load();
    expect(saved?.widgets[1]).toEqual({ ...initialLayout.widgets[1], title: '' });
    expect(saved?.widgets[0]).toEqual(initialLayout.widgets[0]);
  });

  it('keeps titles when layouts are cloned or reflowed into another preset', () => {
    const copy = cloneLayout(initialLayout);
    const preset = getPreset('three');
    if (!preset) {
      throw new Error('Three-column preset is missing.');
    }
    const reflowed = applyPreset(copy.widgets, preset, getRowSize('short'));

    expect(reflowed.map((widget) => widget.title)).toEqual([undefined, 'Second clock']);
    expect(copy.widgets[1]).not.toBe(initialLayout.widgets[1]);
    expect(initialLayout.widgets[1].title).toBe('Second clock');
    expect(initialLayout.widgets[1].w).toBe(6);
  });
});
