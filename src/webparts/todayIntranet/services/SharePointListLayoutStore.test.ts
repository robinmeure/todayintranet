/* eslint-disable require-atomic-updates -- Tests deliberately change server state between awaited requests. */

jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@microsoft/sp-core-library', () => ({ _SPKillSwitch: { isActivated: () => false } }));
jest.mock('@microsoft/sp-page-context', () => ({
  SPPermission: jest.requireActual<{ default: typeof SPPermission }>('@microsoft/sp-page-context/lib-commonjs/SPPermission').default
}));

import type { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { IDashboardLayout } from '../model/IDashboardLayout';
import { LayoutCache } from './LayoutCache';
import { layoutScopeKey } from './ILayoutStore';
import { SharePointLayoutClient } from './SharePointLayoutClient';
import {
  CONFIGURATION_CACHE_MAX_AGE_MS,
  LAYOUT_CACHE_MAX_AGE_MS,
  SharePointListLayoutStore
} from './SharePointListLayoutStore';

const starter: IDashboardLayout = {
  version: 2, widgets: [{ id: 'mail', type: 'm365.mail', x: 0, y: 0, w: 4, h: 6, settings: { maxItems: 6 } }]
};
const changed: IDashboardLayout = { ...starter, widgets: [{ ...starter.widgets[0], title: 'Personal inbox' }] };
const newest: IDashboardLayout = { ...starter, widgets: [{ ...starter.widgets[0], title: 'Latest title' }] };
const options = { siteId: 'site-a', webId: 'web-a', dashboardKey: 'default', userKey: 'test-user',
  webAbsoluteUrl: 'https://example.test/sites/a' };

function response(body: unknown, status: number = 200, headers: Record<string, string> = {}): SPHttpClientResponse {
  const result = new (jest.fn<SPHttpClientResponse, []>())();
  const responseHeaders = new (jest.fn<Headers, []>())();
  responseHeaders.get = (key) => headers[key] ?? null;
  return Object.assign(result, {
    ok: status >= 200 && status < 300, status, headers: responseHeaders,
    json: async () => body
  });
}

interface IItem {
  Id: number;
  Title: string;
  LayoutJson: string;
  version: number;
}

class TestServer {
  public items: IItem[] = [];
  public exists: boolean = true;
  public unique: boolean = false;
  public indexed: boolean = true;
  public field: boolean = true;
  public richText: boolean = false;
  public writable: boolean = true;
  public owner: boolean = false;
  public hidden: boolean = true;
  public readSecurity: number = 2;
  public failGet: number | undefined;
  public failPost: number | undefined;
  public retryAfter: string = '1';
  public beforeWrite: (() => void) | undefined;
  public delayWrite: Promise<void> | undefined;
  public loseCreateResponse: boolean = false;
  public collectionEtags: boolean = true;
  public readonly client = new (jest.fn<SPHttpClient, []>())();
  public readonly get = jest.fn<ReturnType<SPHttpClient['get']>, Parameters<SPHttpClient['get']>>();
  public readonly post = jest.fn<ReturnType<SPHttpClient['post']>, Parameters<SPHttpClient['post']>>();
  private _nextId: number = 1;

  public constructor() {
    this.client.get = this.get;
    this.client.post = this.post;
    this.get.mockImplementation(async (url) => {
      if (this.failGet) {
        return response({}, this.failGet, { 'Retry-After': this.retryAfter });
      }
      if (url.indexOf('/_api/web?$select=EffectiveBasePermissions') >= 0) {
        return response({ EffectiveBasePermissions: this._permissions() });
      }
      if (!this.exists) {
        return response({}, 404);
      }
      if (url.indexOf('/fields?') >= 0) {
        return response({ value: [
          { InternalName: 'Title', TypeAsString: 'Text', Indexed: this.indexed, EnforceUniqueValues: this.unique },
          ...(this.field ? [{ InternalName: 'LayoutJson', TypeAsString: 'Note',
            SchemaXml: `<Field RichText='${this.richText ? 'TRUE' : 'FALSE'}' />` }] : [])
        ] });
      }
      if (url.indexOf('/items?') >= 0) {
        return response({ value: this.items.slice(0, 2).map((item) => ({
          Id: item.Id,
          LayoutJson: item.LayoutJson,
          ...(this.collectionEtags ? { '@odata.etag': `"${item.version}"` } : {})
        })) });
      }
      const itemMatch = /\/items\((\d+)\)/.exec(url);
      if (itemMatch) {
        const item = this.items.filter((candidate) => candidate.Id === Number(itemMatch[1]))[0];
        return item ? response(item, 200, { ETag: `"${item.version}"` }) : response({}, 404);
      }
      return response({ Id: 'list-id', ItemCount: this.items.length, Hidden: this.hidden, OnQuickLaunch: false,
        ReadSecurity: this.readSecurity, WriteSecurity: 2, EnableAttachments: false, EffectiveBasePermissions: this._permissions() });
    });
    this.post.mockImplementation(async (url, _configuration, request) => {
      if (this.failPost) {
        return response({}, this.failPost, { 'Retry-After': this.retryAfter });
      }
      const body: Record<string, unknown> = JSON.parse(String(request?.body));
      if (url.endsWith('/_api/web/lists')) {
        if (this.exists) {
          return response({}, 409);
        }
        this.exists = true;
        this.field = false;
        return response({ Id: 'list-id' }, 201);
      }
      if (url.indexOf('/fields/createfieldasxml') >= 0) {
        this.field = true;
        return response({});
      }
      if (url.indexOf('/fields/getbyinternalnameortitle') >= 0) {
        this.indexed = body.Indexed === true || this.indexed;
        this.unique = body.EnforceUniqueValues === true || this.unique;
        return response({}, 204);
      }
      if (url.indexOf('/items') < 0) {
        this.hidden = body.Hidden === true;
        this.readSecurity = Number(body.ReadSecurity);
        return response({}, 204);
      }
      this.beforeWrite?.();
      if (this.delayWrite) {
        await this.delayWrite;
      }
      const itemMatch = /\/items\((\d+)\)/.exec(url);
      if (itemMatch) {
        const item = this.items.filter((candidate) => candidate.Id === Number(itemMatch[1]))[0];
        if (!item) {
          return response({}, 404);
        }
        const headers = request?.headers;
        const etag = headers && 'IF-MATCH' in headers ? headers['IF-MATCH'] : undefined;
        if (etag !== `"${item.version}"`) {
          return response({}, 412);
        }
        item.LayoutJson = String(body.LayoutJson);
        item.version++;
        return response({}, 204);
      }
      if (this.unique && this.items.some((item) => item.Title === body.Title)) {
        return response({}, 400);
      }
      const item: IItem = { Id: this._nextId++, Title: String(body.Title), LayoutJson: String(body.LayoutJson), version: 1 };
      this.items.push(item);
      if (this.loseCreateResponse) {
        this.loseCreateResponse = false;
        throw new Error('Connection closed after the server accepted the create.');
      }
      return response({ Id: item.Id }, 201);
    });
  }

  private _permissions(): { Low: number; High: number } {
    if (this.owner) {
      return SPPermission.fullMask.value;
    }
    return { Low: SPPermission.viewListItems.value.Low |
      (this.writable ? SPPermission.addListItems.value.Low | SPPermission.editListItems.value.Low : 0), High: 0 };
  }

  public seed(layout: IDashboardLayout = starter): void {
    this.items.push({ Id: this._nextId++, Title: 'default|test-user', LayoutJson: JSON.stringify(layout), version: 1 });
  }

  public itemWrites(): number {
    return this.post.mock.calls.filter(([url]) => url.indexOf('/items') >= 0).length;
  }
}

async function settle(): Promise<void> {
  for (let i = 0; i < 50; i++) {
    await Promise.resolve();
  }
}

async function publish(store: SharePointListLayoutStore): Promise<void> {
  const publishing = store.publish();
  await settle();
  await publishing;
  await settle();
}

describe('SharePoint layout persistence', () => {
  let server: TestServer;
  let store: SharePointListLayoutStore;
  const stores: SharePointListLayoutStore[] = [];
  const createStore = (): SharePointListLayoutStore => {
    const next = new SharePointListLayoutStore({ ...options, spHttpClient: server.client });
    stores.push(next);
    return next;
  };

  beforeEach(() => {
    jest.useFakeTimers();
    window.localStorage.clear();
    jest.spyOn(console, 'warn').mockImplementation(() => undefined);
    server = new TestServer();
    store = createStore();
  });

  afterEach(() => {
    stores.splice(0).forEach((entry) => entry.dispose());
    jest.clearAllTimers();
    jest.useRealTimers();
    jest.restoreAllMocks();
    window.localStorage.clear();
  });

  it('checkpoints every edit locally and publishes only when Done requests it', async () => {
    await store.load();
    await store.save(starter);
    await store.save(changed);
    expect(new LayoutCache(layoutScopeKey(options)).read().record?.layout).toEqual(changed);
    jest.advanceTimersByTime(10000);
    await settle();
    expect(server.itemWrites()).toBe(0);
    await publish(store);
    expect(server.itemWrites()).toBe(1);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(changed);
    expect(store.getStatus().state).toBe('saved');
  });

  it('loads old records and updates with a real ETag', async () => {
    server.seed();
    expect(await store.load()).toEqual(starter);
    await store.save(changed);
    await publish(store);
    expect(server.items[0].version).toBe(2);
    expect(store.getStatus()).toMatchObject({ state: 'saved', mode: 'sharepoint' });
  });

  it('uses three GETs for a healthy cold load and two requests for a normal save', async () => {
    server.seed();
    await store.load();
    expect(server.get).toHaveBeenCalledTimes(3);
    expect(server.post).not.toHaveBeenCalled();
    server.get.mockClear();
    await store.save(changed);
    await publish(store);
    expect(server.post).toHaveBeenCalledTimes(1);
    expect(server.get).toHaveBeenCalledTimes(1);
    expect(server.get.mock.calls[0][0]).toContain('/items(1)');
  });

  it('loads a fresh cloud-validated layout entirely from the scoped local cache', async () => {
    server.seed();
    expect(await store.load()).toEqual(starter);
    store.dispose();
    server.get.mockClear();
    server.post.mockClear();

    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(starter);
    expect(server.get).not.toHaveBeenCalled();
    expect(server.post).not.toHaveBeenCalled();
    expect(reloaded.getStatus()).toMatchObject({ mode: 'sharepoint', state: 'saved' });
  });

  it('revalidates only the layout while cached list configuration is fresh', async () => {
    server.seed();
    await store.load();
    store.dispose();
    server.items[0].LayoutJson = JSON.stringify(newest);
    server.items[0].version++;
    jest.advanceTimersByTime(LAYOUT_CACHE_MAX_AGE_MS + 1);
    server.get.mockClear();

    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(newest);
    expect(server.get).toHaveBeenCalledTimes(1);
    expect(server.get.mock.calls[0][0]).toContain('/items?');
    expect(new LayoutCache(layoutScopeKey(options)).read().record?.layout).toEqual(newest);
  });

  it('revalidates list configuration after its bounded cache lifetime', async () => {
    server.seed();
    await store.load();
    store.dispose();
    jest.advanceTimersByTime(CONFIGURATION_CACHE_MAX_AGE_MS + 1);
    server.get.mockClear();

    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(starter);
    expect(server.get).toHaveBeenCalledTimes(3);
  });

  it('bypasses fresh local validation metadata when Retry is requested', async () => {
    server.seed();
    await store.load();
    store.dispose();
    server.get.mockClear();
    const reloaded = createStore();
    await reloaded.load();
    expect(server.get).not.toHaveBeenCalled();

    await reloaded.resolve('retry');
    expect(server.get).toHaveBeenCalledTimes(3);
  });

  it('falls back to a direct item read when a collection omits its ETag', async () => {
    server.seed();
    server.collectionEtags = false;
    expect(await store.load()).toEqual(starter);
    expect(server.get).toHaveBeenCalledTimes(4);
    expect(server.get.mock.calls[3][0]).toContain('/items(1)');
  });

  it('uses local-first persistence for visitors without attempting cloud writes', async () => {
    server.writable = false;
    server.seed();
    await store.load();
    await store.save(changed);
    store.dispose();
    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(changed);
    expect(server.get).toHaveBeenCalledTimes(3);
    jest.advanceTimersByTime(10000);
    await settle();
    expect(server.post).not.toHaveBeenCalled();
    expect(reloaded.getStatus()).toMatchObject({ mode: 'local', state: 'saved', reason: 'read-only' });
    expect(reloaded.getStatus().message).toContain('do not follow you');
  });

  it('uses a readable existing layout without requiring incompatible Title uniqueness', async () => {
    server.seed(changed);
    server.unique = false;
    expect(await store.load()).toEqual(changed);
    expect(store.getStatus()).toMatchObject({ mode: 'sharepoint', state: 'saved', canEdit: true });
    await store.save(newest);
    jest.advanceTimersByTime(10000);
    await settle();
    expect(server.itemWrites()).toBe(0);
    await publish(store);
    expect(server.itemWrites()).toBe(1);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(newest);
  });

  it('detects a competing server edit instead of overwriting it', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    server.items[0].LayoutJson = JSON.stringify(newest);
    server.items[0].version++;
    await publish(store);
    expect(server.itemWrites()).toBe(1);
    expect(store.getStatus().state).toBe('conflict');
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(newest);
    expect(new LayoutCache(layoutScopeKey(options)).read().record?.layout).toEqual(changed);
    expect(await store.resolve('use-remote', starter)).toEqual(newest);
    expect(store.getStatus().state).toBe('saved');
    expect(store.exportRecovery()).toContain('Personal inbox');
  });

  it('uses conditional writes even after a user chooses to keep the local draft', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    server.items[0].version++;
    await publish(store);
    await store.resolve('keep-local');
    await settle();
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(starter);
    await publish(store);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(changed);
    expect(store.getStatus().state).toBe('saved');
  });

  it('handles a 412 race between the version check and the write', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    server.beforeWrite = () => { server.items[0].version++; };
    await publish(store);
    expect(store.getStatus().state).toBe('conflict');
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(starter);
  });

  it('detects simultaneous first-save duplicates without acknowledging either as canonical', async () => {
    await store.load();
    await store.save(changed);
    server.beforeWrite = () => {
      server.beforeWrite = undefined;
      server.seed(newest);
    };
    await publish(store);
    expect(server.items).toHaveLength(2);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(newest);
    expect(store.getStatus()).toMatchObject({ mode: 'local', reason: 'configuration' });
    expect(store.getStatus().message).toContain('Duplicate layouts');
  });

  it('reconciles a successful create with a lost response before retrying creation', async () => {
    await store.load();
    await store.save(changed);
    server.loseCreateResponse = true;
    await publish(store);
    expect(server.items).toHaveLength(1);
    expect(server.itemWrites()).toBe(1);
    expect(store.getStatus().state).toBe('saved');
  });

  it('serializes slow saves and does not acknowledge a newer edit as saved', async () => {
    server.seed();
    await store.load();
    let release: (() => void) | undefined;
    server.delayWrite = new Promise<void>((resolve) => { release = resolve; });
    await store.save(changed);
    const firstPublish = store.publish();
    await settle();
    await store.save(newest);
    await settle();
    expect(server.itemWrites()).toBe(1);
    server.delayWrite = undefined;
    release?.();
    await firstPublish;
    await settle();
    expect(new LayoutCache(layoutScopeKey(options)).read().record?.pendingSync).toBe(true);
    jest.advanceTimersByTime(1);
    await settle();
    expect(server.itemWrites()).toBe(2);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(newest);
    expect(store.getStatus().state).toBe('saved');
  });

  it('preserves an offline draft across reload and detects a changed cloud base', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    store.dispose();
    server.items[0].LayoutJson = JSON.stringify(newest);
    server.items[0].version++;
    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(changed);
    expect(reloaded.getStatus().state).toBe('conflict');
    expect(server.itemWrites()).toBe(0);
  });

  it('resumes a pending draft when its server base has not changed', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    store.dispose();
    const reloaded = createStore();
    expect(await reloaded.load()).toEqual(changed);
    jest.advanceTimersByTime(10000);
    await settle();
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(starter);
    await publish(reloaded);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(changed);
  });

  it('respects Retry-After and stops after three automatic retries', async () => {
    server.failGet = 429;
    server.retryAfter = '10';
    await store.load();
    const firstCount = server.get.mock.calls.length;
    jest.advanceTimersByTime(9999);
    await settle();
    expect(server.get).toHaveBeenCalledTimes(firstCount);
    await store.resolve('retry');
    expect(server.get).toHaveBeenCalledTimes(firstCount);
    jest.advanceTimersByTime(1);
    await settle();
    expect(server.get).toHaveBeenCalledTimes(firstCount + 1);
    for (let i = 0; i < 2; i++) {
      jest.advanceTimersByTime(10000);
      await settle();
    }
    const finalCount = server.get.mock.calls.length;
    jest.advanceTimersByTime(60000);
    await settle();
    expect(server.get).toHaveBeenCalledTimes(finalCount);
    expect(store.getStatus().message).toContain('Automatic retries stopped');
    server.failGet = undefined;
    await store.resolve('retry');
    expect(store.getStatus().state).toBe('saved');
  });

  it('can save to SharePoint when browser storage is blocked, without claiming a local copy', async () => {
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('quota'); });
    await store.load();
    await store.save(changed);
    expect(store.getStatus().message).not.toContain('Saved in this browser');
    await publish(store);
    expect(store.getStatus()).toMatchObject({ mode: 'sharepoint', state: 'saved', reason: 'local-storage' });
    expect(store.getStatus().message).toContain('recovery is unavailable');
  });

  it('does not claim successful persistence when both stores fail', async () => {
    await store.load();
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('quota'); });
    server.failPost = 503;
    await store.save(changed);
    await publish(store);
    expect(store.getStatus().state).toBe('error');
    expect(store.getStatus().message).toContain('have not been saved');
    expect(store.exportRecovery()).toContain('Personal inbox');
  });

  it('can explicitly export in-memory edits even when browser reads are blocked', async () => {
    await store.load();
    const read = jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('policy'); });
    await store.save(changed);
    const exported = store.exportRecovery();
    expect(exported).toContain('Personal inbox');
    expect(exported).toContain('"partial": true');
    expect(store.getStatus().message).toContain('Browser backups could not be read');
    read.mockRestore();
  });

  it('preserves corrupt server data until an explicit backed-up reset', async () => {
    server.seed();
    server.items[0].LayoutJson = '{"version":999,"widgets":[]}';
    await store.load();
    expect(store.getStatus()).toMatchObject({ state: 'error', reason: 'invalid-data', canEdit: false });
    await expect(store.save(changed)).rejects.toThrow();
    await store.resolve('reset', starter);
    expect(server.items[0].LayoutJson).toContain('999');
    await publish(store);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(starter);
    expect(store.exportRecovery()).toContain('999');
  });

  it('keeps invalid server data protected if reset cannot archive its original', async () => {
    server.seed();
    server.items[0].LayoutJson = '{"version":999,"widgets":[]}';
    await store.load();
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('quota'); });
    await expect(store.resolve('reset', starter)).rejects.toThrow();
    expect(store.getStatus().canEdit).toBe(false);
    expect(server.items[0].LayoutJson).toContain('999');
    expect(server.itemWrites()).toBe(0);
  });

  it('does not auto-import the old site-less local key', async () => {
    server.writable = false;
    window.localStorage.setItem('todayIntranet.layout.default|test-user', JSON.stringify(changed));
    expect(await store.load()).toBeUndefined();
    expect(store.getStatus().actions).toContain('import-legacy');
    expect(await store.resolve('import-legacy')).toEqual(changed);
    expect(window.localStorage.getItem('todayIntranet.layout.default|test-user')).not.toBeNull();
    expect(server.itemWrites()).toBe(0);
  });

  it('can choose the starter when reconciling an imported draft with an absent cloud version', async () => {
    window.localStorage.setItem('todayIntranet.layout.default|test-user', JSON.stringify(changed));
    await store.load();
    await store.resolve('import-legacy');
    expect(store.getStatus().state).toBe('conflict');
    expect(await store.resolve('use-remote', starter)).toEqual(starter);
    expect(server.items).toHaveLength(0);
    await publish(store);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(starter);
    expect(store.exportRecovery()).toContain('Personal inbox');
  });

  it('preserves competing browser drafts rather than uploading one silently', async () => {
    server.seed();
    await store.load();
    const other = createStore();
    await other.load();
    await store.save(changed);
    await other.save(newest);
    expect(other.getStatus().state).toBe('conflict');
    jest.advanceTimersByTime(10000);
    await settle();
    expect(server.itemWrites()).toBe(0);
    expect(store.getStatus().state).toBe('pending');
    await publish(store);
    expect(store.getStatus().state).toBe('conflict');
    expect(server.itemWrites()).toBe(0);
  });

  it('keeps unpublished changes local when disposed', async () => {
    await store.load();
    await store.save(changed);
    store.dispose();
    jest.advanceTimersByTime(10000);
    await settle();
    expect(server.itemWrites()).toBe(0);
    expect(new LayoutCache(layoutScopeKey(options)).read().record?.layout).toEqual(changed);
  });

  it('recreates a deleted item only after confirming its absence', async () => {
    server.seed();
    await store.load();
    await store.save(changed);
    server.items = [];
    await publish(store);
    expect(server.items).toHaveLength(1);
    expect(JSON.parse(server.items[0].LayoutJson)).toEqual(changed);
  });
});

describe('SharePoint list configuration', () => {
  let server: TestServer;
  let client: SharePointLayoutClient;
  beforeEach(() => {
    server = new TestServer();
    client = new SharePointLayoutClient({ spHttpClient: server.client, webAbsoluteUrl: options.webAbsoluteUrl, itemKey: 'default|test-user' });
  });
  afterEach(() => client.dispose());

  it('uses effective list permissions to identify read-only users', async () => {
    server.writable = false;
    expect(await client.initialize()).toBe(false);
    expect(server.post).not.toHaveBeenCalled();
  });

  it('does not repeat metadata reads for an already healthy list', async () => {
    expect(await client.initialize()).toBe(true);
    expect(server.get).toHaveBeenCalledTimes(2);
    expect(server.get.mock.calls[0][0]).not.toContain('/fields?');
    expect(server.get.mock.calls[1][0]).toContain('/fields?');
  });

  it('does not provision a missing list for a visitor', async () => {
    server.exists = false;
    server.writable = false;
    await expect(client.initialize()).rejects.toMatchObject({ reason: 'configuration' });
    expect(server.post).not.toHaveBeenCalled();
  });

  it('provisions and verifies an empty list for an owner exactly once', async () => {
    server.exists = false;
    server.owner = true;
    server.indexed = false;
    const results = await Promise.all([client.initialize(), client.initialize()]);
    expect(results).toEqual([true, true]);
    expect(server.field && server.indexed).toBe(true);
    expect(server.unique).toBe(false);
    expect(server.post.mock.calls.filter(([url]) => url.endsWith('/_api/web/lists'))).toHaveLength(1);
  });

  it('does not attempt incompatible Title uniqueness on populated own-item lists', async () => {
    server.seed();
    server.unique = false;
    server.owner = true;
    expect(await client.initialize()).toBe(true);
    expect(server.unique).toBe(false);
    expect(server.itemWrites()).toBe(0);
    expect(server.post.mock.calls.some(([, , request]) => String(request?.body).indexOf('EnforceUniqueValues') >= 0)).toBe(false);
  });

  it('checks the result of item-access configuration repair', async () => {
    server.owner = true;
    server.readSecurity = 1;
    server.failPost = 403;
    await expect(client.initialize()).rejects.toThrow('HTTP 403');
  });

  it('does not read personal data from an unverified access configuration', async () => {
    server.readSecurity = 1;
    await expect(client.initialize()).rejects.toMatchObject({ reason: 'configuration' });
    expect(server.get.mock.calls.some(([url]) => url.indexOf('/items') >= 0)).toBe(false);
  });

  it('rejects incompatible rich-text layout fields without replacing them', async () => {
    server.richText = true;
    server.owner = true;
    await expect(client.initialize()).rejects.toMatchObject({ reason: 'configuration' });
    expect(server.post).not.toHaveBeenCalled();
  });

  it('detects duplicate layout records instead of choosing the first', async () => {
    server.seed();
    server.seed(changed);
    await client.initialize();
    await expect(client.read()).rejects.toThrow('Duplicate layouts');
  });

  it('rejects overlong keys instead of truncating two identities into one', async () => {
    const longKey = new SharePointLayoutClient({
      spHttpClient: server.client, webAbsoluteUrl: options.webAbsoluteUrl, itemKey: 'x'.repeat(256)
    });
    await expect(longKey.initialize()).rejects.toMatchObject({ reason: 'configuration' });
    expect(server.get).not.toHaveBeenCalled();
  });
});
