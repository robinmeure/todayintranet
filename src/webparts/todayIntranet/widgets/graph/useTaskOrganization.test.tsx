jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@microsoft/sp-page-context', () => ({ SPPermission: {} }));

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { act } from 'react-dom/test-utils';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointDocumentClient } from '../../services/SharePointLayoutClient';
import { IWidgetContext } from '../IWidget';
import { useTaskOrganization } from './useTaskOrganization';

describe('shared Tasks organization subscription', () => {
  let container: HTMLDivElement;
  const contexts: IWidgetContext[] = [];
  const values: Array<ReturnType<typeof useTaskOrganization>> = [];
  const Host: React.FunctionComponent<{ index: number }> = ({ index }) => {
    values[index] = useTaskOrganization(contexts[index]);
    return <span>{values[index].state.message}</span>;
  };
  beforeEach(() => {
    localStorage.clear();
    jest.spyOn(SharePointDocumentClient.prototype, 'initialize').mockResolvedValue(true);
    jest.spyOn(SharePointDocumentClient.prototype, 'read').mockResolvedValue(undefined);
    const spContext = new (jest.fn<WebPartContext, []>())();
    Object.defineProperty(spContext, 'pageContext', { value: {
      aadInfo: { tenantId: 'tenant', userId: 'user' }, site: { id: 'site' },
      web: { id: 'web', absoluteUrl: 'https://example.test/site' }, user: { loginName: 'user' }
    } });
    contexts[0] = { instanceId: 'one', settings: {}, spContext, refreshToken: 0, isEditing: false, updateSettings: jest.fn() };
    contexts[1] = { ...contexts[0], instanceId: 'two' };
    container = document.createElement('div'); document.body.appendChild(container);
  });
  afterEach(() => {
    act(() => { ReactDom.unmountComponentAtNode(container); });
    container.remove(); localStorage.clear(); jest.restoreAllMocks();
  });
  it('shares one store and updates duplicate tiles immediately without needing Done', async () => {
    await act(async () => { ReactDom.render(<><Host index={0} /><Host index={1} /></>, container); });
    expect(values[0].state.canEdit).toBe(true);
    expect(values[0].store).toBe(values[1].store);
    expect(SharePointDocumentClient.prototype.initialize).toHaveBeenCalledTimes(1);
    const store = values[0].store!;
    act(() => store.edit({ version: 1, tasks: {
      'todo:a': { priority: 'high', tags: ['Shared'], modified: store.revision() }
    } }));
    expect(values[1].state.document.tasks['todo:a'].tags).toEqual(['Shared']);
    expect(contexts[0].updateSettings).not.toHaveBeenCalled();
    expect(contexts[1].updateSettings).not.toHaveBeenCalled();
  });
});
