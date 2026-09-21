/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */

jest.mock('../widgets/WidgetRegistry', () => ({ WidgetRegistry: { get: jest.fn() } }));
jest.mock('./WidgetFrame.module.scss', () => ({
  frame: 'frame',
  title: 'widget-title',
  header: 'header',
  body: 'body',
  bodyBare: 'body-bare',
  experienceLink: 'experience-link',
  actions: 'actions',
  settingsCallout: 'settings-callout'
}), { virtual: true });
jest.mock('../widgets/content/WidgetContent.module.scss', () => ({}), { virtual: true });

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { INudge, WidgetFrame } from './WidgetFrame';
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
const onNudge = jest.fn<void, [string, INudge]>();
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
      onNudge={onNudge}
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

  it('renders the title as a heading above the content surface, without a decorative title icon', () => {
    act(() => { ReactDom.render(<TestWidget isEditing={false} />, container); });
    expect(container.querySelector('.header h3')?.textContent).toBe('Clock');
    expect(container.querySelector('.body h3')).toBeNull();
    expect(container.querySelector('.header [data-icon-name="Clock"]')).toBeNull();
    expect(container.querySelector('.body')?.classList.contains('body-bare')).toBe(false);
  });

  it('places an existing footerLink beside the heading, without duplicating it below the content', () => {
    jest.mocked(WidgetRegistry.get).mockReturnValue({
      ...clock, footerLink: { text: 'Open clock', href: 'https://example.test/clock' }
    });
    act(() => { ReactDom.render(<TestWidget isEditing={false} />, container); });
    const link = container.querySelector('.header a');
    expect(link?.textContent).toContain('Open clock');
    expect(link?.getAttribute('href')).toBe('https://example.test/clock');
    expect(link?.getAttribute('target')).toBe('_blank');
    expect(link?.getAttribute('rel')).toBe('noreferrer');
    expect(container.querySelectorAll('a[href="https://example.test/clock"]')).toHaveLength(1);
    expect(container.querySelector('.body a')).toBeNull();
  });

  it('resolves dynamic links using the current context and keeps refresh working', () => {
    jest.mocked(WidgetRegistry.get).mockReturnValue({
      ...clock,
      isRefreshable: true,
      footerLink: (context) => context.refreshToken
        ? { text: 'Updated destination', href: `https://example.test/${context.instanceId}` } : undefined,
      render: (context) => <span>Refresh {context.refreshToken}</span>
    });
    act(() => { ReactDom.render(<TestWidget isEditing={false} />, container); });
    expect(container.querySelector('.experience-link')).toBeNull();
    expect(container.querySelector('.body')?.textContent).toBe('Refresh 0');
    act(() => { Simulate.click(button('Refresh Clock')); });
    expect(container.querySelector('.body')?.textContent).toBe('Refresh 1');
    expect(container.querySelector('.experience-link')?.getAttribute('href')).toBe('https://example.test/clock-1');
    expect(onUpdateSettings).not.toHaveBeenCalled();
  });

  it('allows a definition to supply its own content surfaces', () => {
    jest.mocked(WidgetRegistry.get).mockReturnValue({ ...clock, contentSurface: 'none' });
    act(() => { ReactDom.render(<TestWidget isEditing={false} />, container); });
    expect(container.querySelector('.body-bare')?.textContent).toBe('09:41');
    expect(container.querySelector('.header h3')?.textContent).toBe('Clock');
  });

  it('preserves keyboard move/resize and removal without treating action keys as movement', () => {
    act(() => { ReactDom.render(<TestWidget />, container); });
    const header = container.querySelector<HTMLElement>('.header');
    if (!header) { throw new Error('Missing widget header.'); }
    expect(header.classList.contains('widget-drag-handle')).toBe(true);
    expect(header.tabIndex).toBe(0);
    act(() => {
      Simulate.keyDown(header, { key: 'ArrowRight' });
      Simulate.keyDown(header, { key: 'ArrowDown', shiftKey: true });
      Simulate.keyDown(button('Clock settings'), { key: 'ArrowRight' });
    });
    expect(onNudge.mock.calls).toEqual([['clock-1', { dx: 1 }], ['clock-1', { dh: 1 }]]);
    act(() => { Simulate.click(button('Remove Clock')); });
    expect(onRemove).toHaveBeenCalledWith('clock-1');
  });

  it('closes settings and removes edit affordances when leaving edit mode', () => {
    act(() => { ReactDom.render(<TestWidget />, container); });
    openSettings('Clock settings', 'Clock');
    expect(document.querySelector('[role="dialog"]')).not.toBeNull();
    act(() => { ReactDom.render(<TestWidget isEditing={false} />, container); });
    expect(document.querySelector('[role="dialog"]')).toBeNull();
    expect(container.querySelector('.widget-drag-handle')).toBeNull();
    expect(container.querySelectorAll('button')).toHaveLength(0);
    expect(container.querySelector('.header')?.getAttribute('tabindex')).toBeNull();
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
