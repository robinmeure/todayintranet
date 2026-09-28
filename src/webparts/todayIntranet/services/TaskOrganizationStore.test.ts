jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@microsoft/sp-page-context', () => ({ SPPermission: {} }));

import { TaskDocumentClient, TaskOrganizationStore } from './TaskOrganizationStore';
import { IRemoteLayout, LayoutStorageError } from './SharePointLayoutClient';
import { emptyTaskDocument, TaskDocument, TaskMetadata } from '../widgets/graph/taskOrganization';

const metadata = (extra?: Partial<TaskMetadata>): TaskMetadata => ({
  priority: 'normal', tags: [], modified: '1790588000000:device', ...extra
});
const document = (tasks: Record<string, TaskMetadata>): TaskDocument => ({ version: 1, tasks });

describe('independent personal task organization store', () => {
  let cloud: IRemoteLayout<TaskDocument> | undefined;
  let client: TaskDocumentClient;
  let write: jest.MockedFunction<TaskDocumentClient['write']>;
  let read: jest.MockedFunction<TaskDocumentClient['read']>;
  let update: jest.MockedFunction<TaskDocumentClient['update']>;
  let initialize: jest.MockedFunction<TaskDocumentClient['initialize']>;
  let store: TaskOrganizationStore;
  beforeEach(() => {
    jest.useFakeTimers();
    localStorage.clear();
    cloud = undefined;
    initialize = jest.fn().mockResolvedValue(true);
    read = jest.fn(async () => cloud);
    write = jest.fn(async (layout, version) => {
      if (cloud && cloud.etag !== version?.etag) { throw new LayoutStorageError('Conflict', 'unavailable', 412); }
      cloud = { layout, itemId: 1, etag: String(Number(cloud?.etag ?? 0) + 1) };
      return cloud;
    });
    update = jest.fn(async (layout, version) => write(layout, version));
    client = { initialize, read, write, update, dispose: jest.fn(), recoveryRaw: undefined };
    store = new TaskOrganizationStore(() => client, localStorage, 'tasks:scope:', 'writer');
  });
  afterEach(() => { store.dispose(); jest.clearAllTimers(); jest.useRealTimers(); localStorage.clear(); });

  it('loads without creating or updating metadata on a read-only refresh', async () => {
    await store.load();
    await store.sync();
    expect(write).not.toHaveBeenCalled();
    expect(store.getState()).toMatchObject({ status: 'saved', canEdit: true });
  });
  it('checkpoints immediately, then publishes without dashboard edit mode or Done', async () => {
    await store.load();
    store.edit(document({ 'todo:a': metadata({ tags: ['follow up'] }) }));
    expect(store.getState().status).toBe('pending');
    expect(localStorage.length).toBe(1);
    expect(write).not.toHaveBeenCalled();
    await store.sync();
    expect(cloud?.layout.tasks['todo:a'].tags).toEqual(['follow up']);
    expect(store.getState().status).toBe('saved');
    expect(localStorage.length).toBe(0);
  });
  it('automatically updates with cached ETags and no reads before or after acknowledged saves', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    store.edit(document({ 'todo:a': metadata() }));
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledTimes(1);
    expect(update.mock.calls[0][1].etag).toBe('1');
    expect(store.getState().status).toBe('saved');
    expect(localStorage.length).toBe(0);
    store.edit(document({ 'todo:a': metadata({ tags: ['new'], modified: store.revision() }) }));
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(1);
    expect(update).toHaveBeenCalledTimes(2);
    expect(update.mock.calls[1][1].etag).toBe('2');
  });
  it('defers reading a missing response ETag until another edit needs to be saved', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    update.mockImplementation(async (layout, version) => { await write(layout, version); return undefined; });
    store.edit(document({ 'todo:a': metadata() }));
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(1);
    expect(store.getState().status).toBe('saved');
    expect(localStorage.length).toBe(0);
    await jest.advanceTimersByTimeAsync(60000);
    expect(read).toHaveBeenCalledTimes(1);
    store.edit(document({ 'todo:a': metadata({ tags: ['later'], modified: store.revision() }) }));
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(2);
    expect(update).toHaveBeenCalledTimes(2);
    expect(update.mock.calls[1][1].etag).toBe('2');
    expect(cloud?.layout.tasks['todo:a'].tags).toEqual(['later']);
  });
  it('reads and merges after a cached ETag conflicts without losing another device edits', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    store.edit(document({ 'todo:ours': metadata() }));
    cloud = { itemId: 1, etag: '2', layout: document({ 'todo:theirs': metadata() }) };
    await jest.advanceTimersByTimeAsync(600);
    expect(update.mock.calls[0][1].etag).toBe('1');
    expect(read).toHaveBeenCalledTimes(1);
    expect(store.getState().status).toBe('pending');
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(2);
    expect(update.mock.calls[1][1].etag).toBe('2');
    expect(Object.keys(cloud.layout.tasks).sort()).toEqual(['todo:ours', 'todo:theirs']);
    expect(store.getState().status).toBe('saved');
  });
  it('does not repeat an update after its acknowledgement is lost', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    update.mockImplementationOnce(async (layout, version) => {
      await write(layout, version);
      throw new LayoutStorageError('Connection lost', 'unavailable');
    });
    store.edit(document({ 'todo:a': metadata() }));
    await jest.advanceTimersByTimeAsync(600);
    expect(store.getState().status).toBe('error');
    expect(localStorage.length).toBe(1);
    await store.sync(false);
    expect(read).toHaveBeenCalledTimes(2);
    expect(update).toHaveBeenCalledTimes(1);
    expect(store.getState().status).toBe('saved');
    expect(localStorage.length).toBe(0);
  });
  it('retains edits and obeys Retry-After on an acknowledged-update request that is throttled', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    update.mockRejectedValueOnce(new LayoutStorageError('Throttled', 'unavailable', 429, 60000));
    store.edit(document({ 'todo:a': metadata() }));
    await jest.advanceTimersByTimeAsync(600);
    expect(store.getState().status).toBe('error');
    expect(localStorage.length).toBe(1);
    await store.sync(false);
    await store.sync();
    expect(update).toHaveBeenCalledTimes(1);
    expect(read).toHaveBeenCalledTimes(1);
    await jest.advanceTimersByTimeAsync(60000);
    await store.sync(false);
    expect(read).toHaveBeenCalledTimes(2);
    expect(update).toHaveBeenCalledTimes(2);
    expect(store.getState().status).toBe('saved');
  });
  it('does not checkpoint or schedule saves for an unchanged document', async () => {
    await store.load();
    store.edit(emptyTaskDocument());
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(1);
    expect(write).not.toHaveBeenCalled();
    expect(localStorage.length).toBe(0);
  });
  it('retains newer edits during an acknowledged update with no response ETag', async () => {
    cloud = { itemId: 1, etag: '1', layout: emptyTaskDocument() };
    await store.load();
    let finish!: () => void;
    update.mockImplementationOnce(async (layout, version) => {
      await new Promise<void>((resolve) => { finish = resolve; });
      await write(layout, version);
      return undefined;
    });
    const first = document({ 'todo:a': metadata() });
    store.edit(first);
    const publishing = store.sync(false);
    const newer = document({ 'todo:a': metadata({ tags: ['newer'], modified: store.revision() }) });
    store.edit(newer);
    finish();
    await publishing;
    expect(store.getState()).toMatchObject({ status: 'pending', document: newer });
    expect(read).toHaveBeenCalledTimes(1);
    expect(localStorage.length).toBe(1);
    await jest.advanceTimersByTimeAsync(600);
    expect(read).toHaveBeenCalledTimes(2);
    expect(cloud?.layout).toEqual(newer);
    expect(store.getState().status).toBe('saved');
    expect(localStorage.length).toBe(0);
  });
  it('retains explicit browser-only edits when provisioning is unavailable', async () => {
    initialize.mockRejectedValue(new LayoutStorageError('An owner must provision the field.', 'configuration'));
    await store.load();
    expect(store.getState().status).toBe('local');
    store.edit(document({ 'todo:a': metadata() }));
    expect(store.getState().message).toContain('browser only');
    expect(localStorage.length).toBe(1);
    expect(write).not.toHaveBeenCalled();
  });
  it('does not claim roaming when a visitor has read-only list permissions', async () => {
    initialize.mockResolvedValue(false);
    cloud = { itemId: 1, etag: '1', layout: document({ 'todo:remote': metadata() }) };
    await store.load();
    store.edit(document({ ...store.getState().document.tasks, 'todo:local': metadata() }));
    await store.sync();
    expect(store.getState().status).toBe('local');
    expect(write).not.toHaveBeenCalled();
    expect(store.getState().document.tasks['todo:remote']).toBeDefined();
  });
  it('merges different task edits from another device without losing either', async () => {
    await store.load();
    store.edit(document({ 'todo:ours': metadata() }));
    cloud = { itemId: 1, etag: '1', layout: document({ 'todo:theirs': metadata() }) };
    await store.sync();
    expect(Object.keys(cloud.layout.tasks).sort()).toEqual(['todo:ours', 'todo:theirs']);
    expect(write.mock.calls[0][1]?.etag).toBe('1');
  });
  it('restores offline checkpoints and deletions against the correct saved baseline', async () => {
    const original = document({ 'todo:a': metadata(), 'todo:b': metadata() });
    cloud = { itemId: 1, etag: '1', layout: original };
    localStorage.setItem('tasks:scope:closed-tab', JSON.stringify({
      version: 1, base: original, document: document({ 'todo:b': metadata() })
    }));
    await store.load();
    expect(cloud.layout.tasks['todo:a']).toBeUndefined();
    expect(cloud.layout.tasks['todo:b']).toBeDefined();
  });
  it('merges separate tab outboxes and leaves other user scopes untouched', async () => {
    localStorage.setItem('tasks:scope:tab-a', JSON.stringify({ version: 1, base: emptyTaskDocument(), document: document({ 'todo:a': metadata() }) }));
    localStorage.setItem('tasks:scope:tab-b', JSON.stringify({ version: 1, base: emptyTaskDocument(), document: document({ 'todo:b': metadata() }) }));
    localStorage.setItem('tasks:other:tab', 'unrelated');
    await store.load();
    expect(Object.keys(cloud?.layout.tasks ?? {}).sort()).toEqual(['todo:a', 'todo:b']);
    expect(localStorage.getItem('tasks:other:tab')).toBe('unrelated');
  });
  it('does not overwrite edits made while a publish is in flight', async () => {
    await store.load();
    let finish!: (value: IRemoteLayout<TaskDocument>) => void;
    write.mockImplementationOnce(() => new Promise((resolve) => { finish = resolve; }));
    const first = document({ 'todo:a': metadata() });
    store.edit(first);
    const publishing = store.sync();
    await Promise.resolve(); await Promise.resolve();
    const newer = document({ 'todo:a': metadata({ tags: ['newer'], modified: '1790588000001:device' }) });
    store.edit(newer);
    cloud = { layout: first, itemId: 1, etag: '1' };
    finish(cloud);
    await publishing;
    expect(store.getState().document).toEqual(newer);
    expect(store.getState().status).toBe('pending');
    await store.sync();
    expect(cloud.layout).toEqual(newer);
  });
  it('requires an explicit choice for conflicting order groups', async () => {
    const base = document({ 'todo:a': metadata({ order: 0 }), 'todo:b': metadata({ order: 1 }) });
    cloud = { itemId: 1, etag: '1', layout: base };
    await store.load();
    store.edit(document({ 'todo:a': metadata({ order: 1 }), 'todo:b': metadata({ order: 0 }) }));
    cloud = { itemId: 1, etag: '2', layout: document({ 'todo:a': metadata({ order: 2 }), 'todo:b': metadata({ order: 1 }) }) };
    await store.sync();
    expect(store.getState()).toMatchObject({ status: 'conflict', canEdit: false });
    expect(write).not.toHaveBeenCalled();
    await store.resolveConflict('remote');
    expect(store.getState().document).toEqual(cloud.layout);
    expect(store.getState().status).toBe('saved');
  });
  it('does not claim success after a failed local checkpoint or publish', async () => {
    await store.load();
    const set = jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('Quota'); });
    expect(() => store.edit(document({ 'todo:a': metadata() }))).toThrow('Quota');
    expect(store.getState()).toMatchObject({ status: 'error' });
    expect(store.getState().message).toContain('Not saved');
    set.mockRestore();
    write.mockRejectedValue(new LayoutStorageError('Unavailable', 'unavailable', 503));
    await store.sync();
    expect(store.getState().status).toBe('error');
    expect(localStorage.length).toBe(1);
  });
  it('blocks default organization writes when remote organization fails to load', async () => {
    read.mockRejectedValue(new LayoutStorageError('Unavailable', 'unavailable', 503));
    await store.load();
    expect(store.getState()).toMatchObject({ status: 'error', canEdit: false });
    expect(write).not.toHaveBeenCalled();
  });
  it('switches to browser-only when an existing organization loses write permission', async () => {
    cloud = { itemId: 1, etag: '1', layout: document({ 'todo:a': metadata() }) };
    await store.load();
    store.edit(document({ 'todo:a': metadata({ priority: 'high', modified: '1790588000001:device' }) }));
    client.recoveryRaw = JSON.stringify(cloud.layout);
    write.mockRejectedValueOnce(new LayoutStorageError('Denied', 'read-only', 403));
    await store.sync();
    expect(store.getState()).toMatchObject({ status: 'local', canEdit: true });
    expect(localStorage.length).toBe(1);
    expect(cloud.layout.tasks['todo:a'].priority).toBe('normal');
  });
  it('honors Retry-After without repeatedly calling a throttled server', async () => {
    await store.load();
    store.edit(document({ 'todo:a': metadata() }));
    write.mockRejectedValueOnce(new LayoutStorageError('Throttled', 'unavailable', 429, 60000));
    await store.sync();
    const reads = read.mock.calls.length;
    await store.sync();
    expect(read).toHaveBeenCalledTimes(reads);
    expect(store.getState().message).toContain('Retry after');
    expect(write).toHaveBeenCalledTimes(1);
  });
  it('bounds automatic reconciliation when the server repeatedly rejects ETags', async () => {
    await store.load();
    store.edit(document({ 'todo:a': metadata() }));
    write.mockRejectedValue(new LayoutStorageError('Conflict', 'unavailable', 412));
    await store.sync(); await store.sync(); await store.sync();
    expect(write).toHaveBeenCalledTimes(3);
    expect(store.getState().status).toBe('error');
  });
  it('exports corrupt checkpoints and clears only this user scope after explicit recovery', async () => {
    localStorage.setItem('tasks:scope:bad', '{corrupt');
    localStorage.setItem('tasks:other:keep', 'keep');
    await store.load();
    expect(store.getState()).toMatchObject({ status: 'error', canEdit: false });
    expect(store.exportRecovery()).toContain('{corrupt');
    await store.clearBrowserRecovery();
    expect(store.getState().status).toBe('saved');
    expect(localStorage.getItem('tasks:other:keep')).toBe('keep');
  });
});
