import { IDashboardLayout } from '../model/IDashboardLayout';
import { LayoutCache, LayoutCacheError } from './LayoutCache';

const layout: IDashboardLayout = { version: 2, widgets: [
  { id: 'one', type: 'unknown', x: 0, y: 0, w: 4, h: 6, settings: { nested: { value: 1 } } }
] };
const changed: IDashboardLayout = { ...layout, widgets: [{ ...layout.widgets[0], title: 'Changed' }] };
const other: IDashboardLayout = { ...layout, widgets: [{ ...layout.widgets[0], title: 'Other' }] };
const metadata = { pendingSync: true, baseKnown: true, server: { itemId: 7, etag: '"3"' } };

function keys(): string[] {
  const result: string[] = [];
  for (let index = 0; index < window.localStorage.length; index++) {
    result.push(window.localStorage.key(index)!);
  }
  return result.sort();
}

describe('immutable browser layout cache', () => {
  const caches: LayoutCache[] = [];
  const cache = (scope: string = 'scope', legacy?: string): LayoutCache => {
    const result = new LayoutCache(scope, legacy);
    caches.push(result);
    return result;
  };
  beforeEach(() => window.localStorage.clear());
  afterEach(() => {
    caches.forEach((entry) => entry.dispose());
    caches.length = 0;
    jest.restoreAllMocks();
    window.localStorage.clear();
  });

  it('isolates complete encoded scopes, including namespace-like and Unicode scopes', () => {
    const a = cache('site|web|dashboard|user');
    const b = cache('site|web|dashboard|user.snapshot.child');
    const c = cache('site|web|dashboard|user/日本');
    a.write(layout, metadata);
    b.write(changed, metadata);
    c.write(other, metadata);
    expect(a.read().record?.layout).toEqual(layout);
    expect(b.read().record?.layout).toEqual(changed);
    expect(c.read().record?.layout).toEqual(other);
    expect(keys().every((key) => key.indexOf('todayIntranet.layout.v3.') === 0)).toBe(true);
    const otherKeys = keys().filter((key) => key.indexOf('snapshot%2Echild') >= 0 || key.indexOf('%E6%97%A5') >= 0);
    a.reset(other, metadata);
    otherKeys.forEach((key) => expect(window.localStorage.getItem(key)).not.toBeNull());
  });

  it('returns missing data without silently importing or deleting legacy data', () => {
    const legacy = 'todayIntranet.layout.dashboard|user';
    window.localStorage.setItem(legacy, JSON.stringify(layout));
    const store = cache('scope', legacy);
    expect(store.read()).toEqual({ record: undefined, conflicts: [], legacyAvailable: true });
    expect(keys()).toEqual([legacy]);
    expect(store.readLegacy()).toEqual(layout);
    store.resolve(store.readLegacy()!, metadata);
    expect(store.read().record?.layout).toEqual(layout);
    expect(store.read().legacyAvailable).toBe(false);
    expect(window.localStorage.getItem(legacy)).toBe(JSON.stringify(layout));
  });

  it('does not offer legacy import when any scoped snapshot already exists', () => {
    const legacy = 'todayIntranet.layout.dashboard|user';
    window.localStorage.setItem(legacy, JSON.stringify(layout));
    const a = cache('scope', legacy);
    const b = cache('scope', legacy);
    a.write(layout, metadata);
    expect(a.read().legacyAvailable).toBe(false);
    b.write(changed, metadata);
    expect(a.read().conflicts).toHaveLength(2);
    expect(a.read().legacyAvailable).toBe(false);
    expect(a.readLegacy()).toEqual(layout);
  });

  it('persists server acknowledgement metadata and prunes only superseded snapshots', () => {
    const writer = cache();
    const first = writer.write(layout, metadata);
    const second = writer.write(changed, { ...metadata, pendingSync: false });
    expect(first.revision).not.toBe(second.revision);
    expect(keys()).toHaveLength(1);
    expect(cache().read()).toEqual({ record: second, conflicts: [], legacyAvailable: false });
    expect(second.server).toEqual(metadata.server);
    expect(second.pendingSync).toBe(false);
  });

  it('invalidates cached capability metadata from the incompatible uniqueness build', () => {
    const writer = cache();
    const legacyMetadata = {
      ...metadata,
      cloud: {
        writable: false,
        configurationCheckedAt: Date.now(),
        writeBlock: 'enable unique values'
      },
      remoteCheckedAt: Date.now()
    };
    const record = writer.write(layout, legacyMetadata);
    expect(record.cloud).toBeUndefined();
    expect(record.remoteCheckedAt).toBe(legacyMetadata.remoteCheckedAt);
    expect(cache().read().record?.cloud).toBeUndefined();
  });

  it('projects server metadata to itemId and etag without persisting extra remote fields', () => {
    const a = cache();
    const remote = { ...metadata.server, layout: changed, transport: () => undefined };
    const supplied = { ...metadata, server: remote, extra: 'not metadata' };
    const record = a.write(layout, supplied);
    expect(record.server).toEqual({ itemId: metadata.server.itemId, etag: metadata.server.etag });
    expect(Object.keys(record.server!).sort()).toEqual(['etag', 'itemId']);
    expect(record).not.toHaveProperty('extra');
    const persisted = JSON.parse(window.localStorage.getItem(keys()[0])!);
    expect(persisted.server).toEqual(metadata.server);
    expect(persisted.layout).toEqual(layout);
  });

  it('accepts a complete record as resolve metadata without overwriting the new revision or selected layout', () => {
    const a = cache();
    const previous = a.write(layout, metadata);
    const resolved = a.resolve(changed, previous);
    expect(resolved.layout).toEqual(changed);
    expect(resolved.revision).not.toBe(previous.revision);
    expect(resolved.scope).toBe(previous.scope);
    Object.keys(previous.clock).forEach((writer) =>
      expect(resolved.clock[writer]).toBeGreaterThan(previous.clock[writer]));
    expect(resolved.server).toEqual(metadata.server);
    expect(resolved.pendingSync).toBe(metadata.pendingSync);
    expect(resolved.baseKnown).toBe(metadata.baseKnown);
    expect(a.read().record).toEqual(resolved);
    expect(a.read().conflicts).toEqual([]);
  });

  it('keeps two immutable incomparable heads even after the stale writer reads the other edit', () => {
    const a = cache();
    const b = cache();
    const original = a.write(layout, metadata);
    b.adopt(original);
    const local = a.write(changed, metadata);
    expect(b.read().record?.revision).toBe(local.revision);
    const remote = b.write(other, metadata);
    const result = a.read();
    expect(keys()).toHaveLength(2);
    expect(result.record?.revision).toBe(local.revision);
    expect(b.read().record?.revision).toBe(remote.revision);
    expect(result.conflicts.map((record) => record.revision).sort()).toEqual([local.revision, remote.revision].sort());
    expect(cache().read().record?.revision).toBe(cache().read().record?.revision);
  });

  it('does not adopt any storage head unless explicitly asked to adopt it', () => {
    const a = cache();
    const b = cache();
    const original = a.write(layout, metadata);
    b.read();
    b.write(changed, metadata);
    expect(a.read().conflicts).toHaveLength(2);
    const c = cache('clean');
    const d = cache('clean');
    d.adopt(c.write(layout, metadata));
    d.write(changed, metadata);
    expect(c.read().conflicts).toEqual([]);
    expect(original.clock).toBeDefined();
  });

  it('explicit resolve merges clocks and archives exact discarded versions before pruning', () => {
    const a = cache();
    const b = cache();
    b.adopt(a.write(layout, metadata));
    const local = a.write(changed, metadata);
    const remote = b.write(other, metadata);
    const raw = keys().map((key) => window.localStorage.getItem(key));
    const resolved = a.resolve(changed, metadata);
    expect(a.read().conflicts).toEqual([]);
    expect(a.read().record).toEqual(resolved);
    [local, remote].forEach((record) => Object.keys(record.clock).forEach((writer) =>
      expect(resolved.clock[writer]).toBeGreaterThanOrEqual(record.clock[writer])));
    const exported = JSON.parse(a.exportRecovery());
    const archives = exported.entries.filter((entry: { key: string }) => entry.key.indexOf('.recovery.') >= 0);
    expect(archives).toHaveLength(1);
    expect(JSON.parse(archives[0].raw).entries.map((entry: { raw: string }) => entry.raw).sort()).toEqual(raw.sort());
  });

  it('archives exact server payloads in distinct scoped immutable keys included in recovery exports', () => {
    const a = cache('scope-a');
    const b = cache('scope-b');
    const payloads = ['', '{malformed server JSON \r\n 日本', '{"version":99,"widgets":[]}'];
    payloads.forEach((raw) => a.archive(raw));
    a.archive(payloads[1]);
    b.archive('other scope backup');
    const exported = JSON.parse(a.exportRecovery());
    expect(exported.entries).toHaveLength(4);
    expect(new Set(exported.entries.map((entry: { key: string }) => entry.key)).size).toBe(4);
    expect(exported.entries.every((entry: { key: string }) => entry.key.indexOf('.recovery.') >= 0)).toBe(true);
    expect(exported.entries.map((entry: { raw: string }) => JSON.parse(entry.raw).raw).sort())
      .toEqual(payloads.concat(payloads[1]).sort());
    expect(a.read().record).toBeUndefined();
    expect(JSON.parse(b.exportRecovery()).entries).toHaveLength(1);
    a.write(layout, metadata);
    a.reset(changed, metadata);
    const after = JSON.parse(a.exportRecovery());
    exported.entries.forEach((entry: { key: string; raw: string }) => expect(after.entries).toContainEqual(entry));
  });

  it('can archive invalid server content even when existing local snapshots are unreadable', () => {
    const a = cache();
    a.write(layout, metadata);
    const key = keys()[0];
    window.localStorage.setItem(key, '{corrupt local snapshot');
    const serverRaw = '{corrupt server snapshot';
    a.archive(serverRaw);
    expect(() => a.read()).toThrow(expect.objectContaining({ reason: 'invalid-data' }));
    const exported = JSON.parse(a.exportRecovery());
    expect(exported.entries).toContainEqual({ key, raw: '{corrupt local snapshot' });
    const archive = exported.entries.filter((entry: { key: string }) => entry.key.indexOf('.recovery.') >= 0)[0];
    expect(JSON.parse(archive.raw).raw).toBe(serverRaw);
  });

  it('rejects archive quota failures or silent storage rejection without claiming a durable backup', () => {
    const a = cache();
    a.write(layout, metadata);
    a.archive('previous server backup');
    const before = a.exportRecovery();
    const write = jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('private quota details'); });
    expect(() => a.archive('{private server content')).toThrow(expect.objectContaining({ reason: 'local-storage' }));
    expect(() => a.archive('{private server content')).not.toThrow('private');
    write.mockImplementation(() => undefined);
    expect(() => a.archive('{private server content')).toThrow('were not retained');
    write.mockRestore();
    expect(a.exportRecovery()).toBe(before);
    expect(a.read().record?.layout).toEqual(layout);
  });

  it('rejects non-string archive input as invalid-data', () => {
    const a = cache();
    expect(() => a.archive(42 as unknown as string)).toThrow(expect.objectContaining({ reason: 'invalid-data' }));
    expect(keys()).toEqual([]);
  });

  it('keeps a concurrently created key even when it appears during pruning', () => {
    const a = cache();
    const b = cache();
    b.adopt(a.write(layout, metadata));
    const nativeSet = Storage.prototype.setItem;
    let injected = false;
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(function (this: Storage, key: string, value: string): void {
      nativeSet.call(this, key, value);
      if (!injected) {
        injected = true;
        b.write(other, metadata);
      }
    });
    a.write(changed, metadata);
    expect(a.read().conflicts).toHaveLength(2);
    expect(keys()).toHaveLength(2);
  });

  it.each(['{malformed private data', '{"schema":99}', '{"schema":1,"layout":{"version":9}}'])
  ('blocks malformed/future snapshots until explicit reset and preserves exact raw data', (raw) => {
    const a = cache();
    a.write(layout, metadata);
    const key = keys()[0];
    window.localStorage.setItem(key, raw);
    expect(() => a.read()).toThrow(LayoutCacheError);
    expect(() => a.read()).toThrow(expect.objectContaining({ reason: 'invalid-data' }));
    expect(() => a.write(changed, metadata)).toThrow(LayoutCacheError);
    expect(window.localStorage.getItem(key)).toBe(raw);
    a.reset(changed, metadata);
    expect(a.read().record?.layout).toEqual(changed);
    const exported = JSON.parse(a.exportRecovery());
    const recovery = exported.entries.filter((entry: { key: string }) => entry.key.indexOf('.recovery.') >= 0)[0];
    expect(JSON.parse(recovery.raw).entries).toContainEqual({ key, raw });
  });

  it('does not prune existing settings when a quota error prevents saving or archiving', () => {
    const a = cache();
    a.write(layout, metadata);
    const before = a.exportRecovery();
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('secret quota details'); });
    expect(() => a.write(changed, metadata)).toThrow('Storage may be full or blocked');
    expect(() => a.write(changed, metadata)).toThrow(expect.objectContaining({ reason: 'local-storage' }));
    expect(() => a.resolve(changed, metadata)).toThrow(LayoutCacheError);
    expect(() => a.reset(changed, metadata)).toThrow(LayoutCacheError);
    expect(a.exportRecovery()).toBe(before);
  });

  it('reports blocked reads and silent storage rejection with safe errors', () => {
    const a = cache();
    a.write(layout, metadata);
    const read = jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('private'); });
    expect(() => a.read()).toThrow(expect.objectContaining({ reason: 'local-storage' }));
    expect(() => a.read()).not.toThrow('private');
    read.mockRestore();
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => undefined);
    expect(() => a.write(layout, metadata)).toThrow('were not retained');
  });

  it('validates metadata, clocks, layout and snapshot identity', () => {
    const a = cache();
    const record = a.write(layout, metadata);
    expect(() => a.write({ ...layout, version: 9 }, metadata)).toThrow(LayoutCacheError);
    expect(() => a.write(layout, { ...metadata, baseKnown: false })).toThrow(LayoutCacheError);
    const key = keys()[0];
    window.localStorage.setItem(key, JSON.stringify({ ...record, clock: { writer: -1 } }));
    expect(() => a.read()).toThrow(LayoutCacheError);
    window.localStorage.setItem(key, JSON.stringify({ ...record, scope: 'another' }));
    expect(() => a.read()).toThrow(LayoutCacheError);
  });

  it('classifies invalid scopes, versions and legacy JSON as invalid-data', () => {
    ['', '\uD800'].forEach((scope) =>
      expect(() => cache(scope)).toThrow(expect.objectContaining({ reason: 'invalid-data' })));
    const a = cache('scope', 'todayIntranet.layout.dashboard|user');
    window.localStorage.setItem('todayIntranet.layout.dashboard|user', '{private malformed data');
    expect(() => a.readLegacy()).toThrow(expect.objectContaining({ reason: 'invalid-data' }));
    expect(() => a.write({ ...layout, version: 99 }, metadata))
      .toThrow(expect.objectContaining({ reason: 'invalid-data' }));
  });

  it('recovers a scan interrupted by another writer removing a key', () => {
    const a = cache();
    a.write(layout, metadata);
    const key = keys()[0];
    const nativeGet = Storage.prototype.getItem;
    let removed = false;
    jest.spyOn(Storage.prototype, 'getItem').mockImplementation(function (this: Storage, item: string): string | null {
      if (item === key && !removed) {
        removed = true;
        this.removeItem(key);
      }
      return nativeGet.call(this, item);
    });
    expect(a.read().record).toBeUndefined();
  });

  it('filters storage notifications by scope and storage area and unsubscribes on disposal', () => {
    const a = cache();
    a.write(layout, metadata);
    const key = keys()[0];
    const listener = jest.fn();
    const unsubscribe = a.subscribe(listener);
    window.dispatchEvent(new StorageEvent('storage', { key: 'unrelated', storageArea: window.localStorage }));
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.sessionStorage }));
    expect(listener).not.toHaveBeenCalled();
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.localStorage }));
    window.dispatchEvent(new StorageEvent('storage', { key: null, storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(2);
    unsubscribe();
    a.dispose();
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(2);
  });

  it('resubscribes and reads after disposal without restoring old listeners or adopting unseen heads', () => {
    const a = cache();
    const before = a.write(layout, metadata);
    const oldListener = jest.fn();
    a.subscribe(oldListener);
    a.dispose();
    const b = cache();
    b.adopt(before);
    b.write(changed, metadata);
    const listener = jest.fn();
    a.subscribe(listener);
    expect(a.read().record?.layout).toEqual(changed);
    const own = a.write(other, metadata);
    expect(a.read().conflicts).toHaveLength(2);
    const key = keys().filter((entry) => entry.indexOf(own.revision) >= 0)[0];
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(1);
    expect(oldListener).not.toHaveBeenCalled();
  });
});
