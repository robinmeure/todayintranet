jest.mock('./VivaEngageWidget.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetCollections.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetContent.module.scss', () => ({}), { virtual: true });
jest.mock('../content/WidgetAdaptiveCard', () => ({ WidgetAdaptiveCard: jest.fn(() => null) }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act, Simulate } from 'react-dom/test-utils';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { VivaEngageWidget, vivaEngageFooterLink } from './VivaEngageWidget';
import { IWidgetContext } from '../IWidget';
import { clearWidgetDataCache } from '../data/WidgetDataCache';

const respond = jest.fn<Promise<unknown>, [string]>();
const client = new (jest.fn<MSGraphClientV3, []>())();
Object.assign(client, {
  api: (url: string) => {
    const request = { version: () => request, select: () => request, header: () => request, get: () => respond(url) };
    return request;
  }
});
const spContext = new (jest.fn<WebPartContext, []>())();
Object.defineProperties(spContext, {
  pageContext: { value: {
    cultureInfo: { currentUICultureName: 'en-GB' }, aadInfo: { tenantId: 'tenant', userId: 'user' },
    site: { id: 'site' }, web: { id: 'web' }, user: { loginName: 'test@example.test' }
  } },
  msGraphClientFactory: { value: { getClient: async () => client } }
});

describe('VivaEngageWidget', () => {
  let container: HTMLDivElement;

  beforeAll(() => initializeIcons(undefined, { disableWarnings: true }));
  beforeEach(() => {
    clearWidgetDataCache();
    respond.mockReset();
    container = document.createElement('div');
    document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove();
  });

  const contextWith = (settings: Record<string, unknown>, openSettings?: () => void): IWidgetContext => ({
    instanceId: 'engage', spContext, settings, isEditing: false, refreshToken: 0, updateSettings: jest.fn(), openSettings
  });
  const render = async (context: IWidgetContext): Promise<void> => {
    await act(async () => { ReactDom.render(<VivaEngageWidget context={context} />, container); });
  };

  it('asks for a community instead of querying Graph when none is chosen', async () => {
    const openSettings = jest.fn();
    await render(contextWith({}, openSettings));
    expect(container.textContent).toContain('No community chosen');
    expect(respond).not.toHaveBeenCalled();
    const button = Array.from(container.querySelectorAll('button')).find((item) => item.textContent?.includes('Choose community'));
    await act(async () => { Simulate.click(button!); });
    expect(openSettings).toHaveBeenCalled();
  });

  it('renders community details, member count and recent files', async () => {
    respond.mockImplementation(async (url) => {
      if (url.startsWith('/employeeExperience')) {
        return { id: 'c1', displayName: 'Engineering', description: 'Builders', privacy: 'private', groupId: 'g1' };
      }
      if (url.includes('/members')) { return { '@odata.count': 1234 }; }
      return { value: [{ id: 'f1', name: 'Roadmap.pptx', file: {}, webUrl: 'https://contoso.sharepoint.com/f1',
        lastModifiedDateTime: '2026-05-01T10:00:00Z' }] };
    });
    await render(contextWith({ communityId: 'c1' }));

    const name = container.querySelector<HTMLAnchorElement>('a[href="https://engage.cloud.microsoft/main/groups/c1"]');
    expect(name?.textContent).toBe('Engineering');
    expect(container.textContent).toContain('Private');
    expect(container.textContent).toContain('1,234 members');
    expect(container.textContent).toContain('Builders');
    expect(container.textContent).toContain('Roadmap.pptx');
  });

  it('explains a missing Community.Read.All consent', async () => {
    respond.mockRejectedValue({ statusCode: 403, code: 'Authorization_RequestDenied' });
    await render(contextWith({ communityId: 'c1' }));
    // MessageBar renders its text through Fluent's DelayedRender.
    await act(async () => { await new Promise((resolve) => setTimeout(resolve, 0)); });
    expect(container.textContent).toContain('Community.Read.All');
  });

  it('links the tile header to the chosen community or Engage home', () => {
    expect(vivaEngageFooterLink(contextWith({ communityId: 'c1' })).href).toBe('https://engage.cloud.microsoft/main/groups/c1');
    expect(vivaEngageFooterLink(contextWith({})).href).toBe('https://engage.cloud.microsoft/main/feed');
  });
});
