/* eslint-disable @rushstack/pair-react-dom-render-unmount -- Each test unmounts its container in afterEach. */

jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@microsoft/sp-page-context', () => ({ SPPermission: {} }));
jest.mock('../../components/WidgetFrame.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetCollections.module.scss', () => ({}), { virtual: true });
jest.mock('../graph/TasksWidget.module.scss', () => ({}), { virtual: true });
jest.mock('../graph/VivaEngageWidget.module.scss', () => ({}), { virtual: true });
jest.mock('../search/SearchQuerySetting.module.scss', () => ({}), { virtual: true });
jest.mock('../links/MyLinksWidgetSettings.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetAdaptiveCard', () => ({
  WidgetAdaptiveCard: ({ card }: { card: Record<string, unknown> }) => <div>{JSON.stringify(card)}</div>
}));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import { WidgetRegistry } from '../WidgetRegistry';

describe('mock content across the widget catalogue', () => {
  let container: HTMLDivElement;
  const graphClient = jest.fn();
  const http = jest.fn();
  const spContext = new (jest.fn<WebPartContext, []>())();
  Object.defineProperties(spContext, {
    pageContext: { value: {
      cultureInfo: { currentUICultureName: 'en-GB' },
      aadInfo: { userId: 'user', tenantId: 'tenant' },
      user: { loginName: 'test@example.test', displayName: 'Test User' },
      site: { id: 'site', absoluteUrl: 'https://example.test' },
      web: { id: 'web', absoluteUrl: 'https://example.test' }, legacyPageContext: {}
    } },
    msGraphClientFactory: { value: { getClient: graphClient } },
    spHttpClient: { value: { get: http, post: http } }
  });

  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => {
    jest.useFakeTimers();
    jest.clearAllMocks();
    container = document.createElement('div');
    document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  it.each(WidgetRegistry.list())('renders $type from its fixture without settings or API calls', async (definition) => {
    const context: IWidgetContext = {
      instanceId: definition.type, settings: {}, spContext,
      refreshToken: 0, isEditing: false, useMockData: true, updateSettings: jest.fn()
    };
    await act(async () => { ReactDom.render(definition.render(context), container); });
    expect(container.textContent?.toLowerCase()).toMatch(/lorem|ipsum/);
    expect(container.textContent).not.toMatch(/Loading|No community chosen|Add a query/);
    context.refreshToken++;
    await act(async () => { ReactDom.render(definition.render(context), container); });
    expect(graphClient).not.toHaveBeenCalled();
    expect(http).not.toHaveBeenCalled();
    expect(context.updateSettings).not.toHaveBeenCalled();
  });
});
