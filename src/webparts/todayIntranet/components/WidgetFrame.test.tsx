/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */

jest.mock('../widgets/WidgetRegistry', () => ({ WidgetRegistry: { get: jest.fn() } }));
jest.mock('./WidgetFrame.module.scss', () => ({
  frame: 'frame',
  title: 'widget-title',
  header: 'header',
  body: 'body',
  actions: 'actions',
  settingsCallout: 'settings-callout'
}), { virtual: true });
jest.mock('../widgets/content/WidgetContent.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { WidgetFrame } from './WidgetFrame';
import { IWidgetInstance } from '../model/IDashboardLayout';
import { IWidgetDefinition, IWidgetHostContext } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

const clock: IWidgetDefinition = {
  type: 'test.clock',
  displayName: 'Clock',
  description: 'Example clock',
  category: 'demo',
  iconName: 'Clock',
  defaultSize: { w: 2, h: 2 },
  render: () => <span>09:41</span>
};

const onUpdateTitle = jest.fn<void, [string, string]>();
const onUpdateSettings = jest.fn<void, [Record<string, unknown>]>();
const onRemove = jest.fn<void, [string]>();
const spContext = new (jest.fn<WebPartContext, []>())();

interface ITestWidgetProps {
  id?: string;
  type?: string;
  title?: string;
  isEditing?: boolean;
}

const TestWidget: React.FunctionComponent<ITestWidgetProps> = ({
  id = 'clock-1',
  type = clock.type,
  title,
  isEditing = true
}) => {
  const [instance, setInstance] = React.useState<IWidgetInstance>({ id, type, title, x: 0, y: 0, w: 2, h: 2 });
  const context: IWidgetHostContext = {
    instanceId: id,
    settings: { title: 'Content-owned title' },
    spContext,
    isEditing,
    updateSettings: onUpdateSettings
  };

  return (
    <WidgetFrame
      instance={instance}
      widgetContext={context}
      isEditing={isEditing}
      onRemove={onRemove}
      onNudge={() => undefined}
      onUpdateTitle={(instanceId, next) => {
        onUpdateTitle(instanceId, next);
        setInstance((current) => ({ ...current, title: next }));
      }}
    />
  );
};

describe('WidgetFrame titles', () => {
  let container: HTMLDivElement;

  const button = (label: string): HTMLButtonElement => {
    const found = Array.from(document.querySelectorAll('button'))
      .find((candidate) => candidate.getAttribute('aria-label') === label);
    if (!found) {
      throw new Error(`Button not found: ${label}`);
    }
    return found;
  };

  const openSettings = (label: string, placeholder: string): HTMLInputElement => {
    act(() => {
      Simulate.click(button(label));
    });
    act(() => {
      jest.runOnlyPendingTimers();
    });
    const input = Array.from(document.querySelectorAll('input'))
      .find((candidate) => candidate.placeholder === placeholder);
    if (!input) {
      throw new Error(`Title input not found: ${placeholder}`);
    }
    return input;
  };

  const changeTitle = (input: HTMLInputElement, value: string): void => {
    act(() => {
      input.value = value;
      Simulate.change(input);
    });
  };

  beforeAll(() => {
    initializeIcons(undefined, { disableWarnings: true });
  });

  beforeEach(() => {
    jest.useFakeTimers();
    jest.clearAllMocks();
    jest.mocked(WidgetRegistry.get).mockImplementation((type) => type === clock.type ? clock : undefined);
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

  it('offers title settings even when the widget has no settings or alternate views', () => {
    act(() => {
      ReactDom.render(<TestWidget />, container);
    });

    const input = openSettings('Clock settings', 'Clock');
    expect(input.value).toBe('');
    expect(document.body.textContent).toContain('Leave blank to use "Clock".');
    expect(onUpdateTitle).not.toHaveBeenCalled();
  });

  it('updates the title, tooltip, region and action labels without changing content settings', () => {
    jest.mocked(WidgetRegistry.get).mockReturnValue({ ...clock, isRefreshable: true });
    act(() => {
      ReactDom.render(<TestWidget />, container);
    });
    const input = openSettings('Clock settings', 'Clock');

    changeTitle(input, '  Office time  ');

    expect(onUpdateTitle).toHaveBeenLastCalledWith('clock-1', '  Office time  ');
    expect(container.querySelector('section')?.getAttribute('aria-label')).toBe('Office time');
    expect(container.querySelector('.widget-title')?.textContent).toBe('Office time');
    expect(container.querySelector('.widget-title')?.getAttribute('title')).toBe('Office time');
    expect(button('Office time settings')).toBeDefined();
    expect(button('Refresh Office time')).toBeDefined();
    expect(button('Remove Office time')).toBeDefined();
    expect(document.querySelector('[role="dialog"]')?.getAttribute('aria-label')).toBe('Office time settings');
    expect(onUpdateSettings).not.toHaveBeenCalled();
    expect(clock.displayName).toBe('Clock');
  });

  it.each(['', '   '])('restores the default name for a cleared title (%j)', (value) => {
    act(() => {
      ReactDom.render(<TestWidget title="Office time" />, container);
    });
    const input = openSettings('Office time settings', 'Clock');

    changeTitle(input, value);

    expect(container.querySelector('.widget-title')?.textContent).toBe('Clock');
    expect(container.querySelector('section')?.getAttribute('aria-label')).toBe('Clock');
    expect(button('Clock settings')).toBeDefined();
  });

  it('renders a saved title in read mode without exposing edit controls', () => {
    act(() => {
      ReactDom.render(<TestWidget title="Office time" isEditing={false} />, container);
    });

    expect(container.querySelector('.widget-title')?.textContent).toBe('Office time');
    expect(container.querySelectorAll('button')).toHaveLength(0);
  });

  it('renames only the selected instance of a widget type', () => {
    act(() => {
      ReactDom.render(<><TestWidget /><TestWidget id="clock-2" title="Second clock" /></>, container);
    });
    const input = openSettings('Clock settings', 'Clock');

    changeTitle(input, 'Office time');

    expect(Array.from(container.querySelectorAll('.widget-title')).map((node) => node.textContent))
      .toEqual(['Office time', 'Second clock']);
    expect(onUpdateTitle).toHaveBeenCalledTimes(1);
    expect(onUpdateTitle).toHaveBeenCalledWith('clock-1', 'Office time');
  });

  it('also provides title settings for an unavailable widget', () => {
    act(() => {
      ReactDom.render(<TestWidget type="retired.widget" />, container);
    });
    const input = openSettings('retired.widget settings', 'retired.widget');

    changeTitle(input, 'Archived widget');

    expect(container.querySelector('section')?.getAttribute('aria-label')).toBe('Archived widget');
    expect(container.textContent).toContain('This widget (retired.widget) is no longer available.');
  });

  it('renders titles as plain text rather than markup', () => {
    const title = '<b>Office</b> & "Team"';
    act(() => {
      ReactDom.render(<TestWidget title={title} isEditing={false} />, container);
    });

    expect(container.querySelector('.widget-title')?.textContent).toBe(title);
    expect(container.querySelector('.widget-title')?.getAttribute('title')).toBe(title);
    expect(container.querySelector('b')).toBeNull();
  });

  it('keeps the shared view picker and widget-specific settings alongside the title', () => {
    jest.mocked(WidgetRegistry.get).mockReturnValue({
      ...clock,
      supportedViews: ['list', 'compact'],
      defaultView: 'list',
      renderSettings: () => <p>Widget-specific settings</p>
    });
    act(() => {
      ReactDom.render(<TestWidget />, container);
    });

    openSettings('Clock settings', 'Clock');

    expect(document.querySelector('[role="group"][aria-labelledby]')?.textContent).toContain('Compact');
    expect(document.body.textContent).toContain('Widget-specific settings');
  });
});
