import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import { widgetDataCacheKey } from '../data/WidgetDataCache';
import { SharePointDocumentClient } from '../../services/SharePointLayoutClient';
import { TaskOrganizationState, TaskOrganizationStore } from '../../services/TaskOrganizationStore';
import { parseTaskDocument, TaskDocument } from './taskOrganization';

interface SharedStore { store: TaskOrganizationStore; users: number; cleanup(): void }
const stores = new Map<string, SharedStore>();

function acquire(context: IWidgetContext, scope: string): SharedStore {
  const existing = stores.get(scope);
  if (existing) { existing.users++; return existing; }
  const page = context.spContext.pageContext;
  const user = String(page.aadInfo?.userId ?? page.user.loginName);
  const store = new TaskOrganizationStore(() => new SharePointDocumentClient<TaskDocument>({
    spHttpClient: context.spContext.spHttpClient, webAbsoluteUrl: page.web.absoluteUrl,
    // Layout keys always contain "|"; this encoded key cannot collide with them.
    itemKey: `tasks:v1:${encodeURIComponent(user)}`
  }, 'TaskOrganizationJson', parseTaskDocument), window.localStorage, `today-tasks:${scope}:`, crypto.randomUUID());
  const storageChanged = (event: StorageEvent): void => {
    if (event.key === null || event.key.startsWith(store.prefix)) { store.run(() => store.sync()); }
  };
  const focus = (): void => { store.run(() => store.sync()); };
  window.addEventListener('storage', storageChanged);
  window.addEventListener('focus', focus);
  const entry = { store, users: 1, cleanup: () => {
    window.removeEventListener('storage', storageChanged);
    window.removeEventListener('focus', focus);
    store.dispose();
  } };
  stores.set(scope, entry);
  return entry;
}

export function useTaskOrganization(context: IWidgetContext): {
  state: TaskOrganizationState; store?: TaskOrganizationStore
} {
  const scope = widgetDataCacheKey(context, 'task-organization', []);
  const [value, setValue] = React.useState<{ scope: string; state: TaskOrganizationState; store?: TaskOrganizationStore }>();
  React.useEffect(() => {
    let entry: SharedStore;
    try {
      entry = acquire(context, scope);
    } catch (error) {
      setValue({ scope, state: { document: { version: 1, tasks: {} }, status: 'error', canEdit: false,
        message: `Personal organization is unavailable: ${error instanceof Error ? error.message : 'browser storage could not be opened'}` } });
      return;
    }
    const update = (): void => setValue({ scope, state: entry.store.getState(), store: entry.store });
    const unsubscribe = entry.store.subscribe(update);
    update();
    entry.store.run(() => entry.store.load());
    return () => {
      unsubscribe();
      if (--entry.users === 0) { entry.cleanup(); stores.delete(scope); }
    };
  }, [scope]);
  React.useEffect(() => {
    if (context.refreshToken && value?.scope === scope && value.store) { value.store.run(() => value.store!.sync()); }
  }, [context.refreshToken, scope]);
  return value?.scope === scope ? value : {
    state: { document: { version: 1, tasks: {} }, status: 'loading', message: 'Loading personal organization...', canEdit: false }
  };
}
