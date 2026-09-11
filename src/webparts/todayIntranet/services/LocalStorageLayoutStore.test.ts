import { IDashboardLayout } from '../model/IDashboardLayout';
import { LayoutCache } from './LayoutCache';
import { LocalStorageLayoutStore } from './LocalStorageLayoutStore';

const layout: IDashboardLayout = { version: 2, widgets: [{ id: 'one', type: 'unknown', x: 0, y: 0, w: 4, h: 6 }] };
const changed: IDashboardLayout = { ...layout, widgets: [{ ...layout.widgets[0], title: 'Changed' }] };
const attempted: IDashboardLayout = { ...layout, widgets: [{ ...layout.widgets[0], title: 'Just typed locally' }] };
const metadata = { pendingSync: false, baseKnown: false };

describe('LocalStorageLayoutStore', () => {
  const stores: LocalStorageLayoutStore[] = [];
  const caches: LayoutCache[] = [];
  const store = (scope: string = 'dashboard|user'): LocalStorageLayoutStore => {
    const result = new LocalStorageLayoutStore(scope);
    stores.push(result);
    return result;
  };
  const cache = (): LayoutCache => {
    const result = new LayoutCache('dashboard|user');
    caches.push(result);
    return result;
  };
  beforeEach(() => window.localStorage.clear());
  afterEach(() => {
    stores.splice(0).forEach((entry) => entry.dispose());
    caches.splice(0).forEach((entry) => entry.dispose());
    jest.restoreAllMocks();
    window.localStorage.clear();
  });

  it('allows a starter only for missing settings and checkpoints synchronously', async () => {
    const a = store();
    expect(await a.load()).toBeUndefined();
    expect(a.getStatus()).toMatchObject({ mode: 'local', state: 'saved', canEdit: true });
    const saved = a.save(layout);
    expect(cache().read().record?.layout).toEqual(layout);
    await saved;
    expect(await store().load()).toEqual(layout);
  });

  it('supports the existing constructor and initial save before load', async () => {
    const a = store();
    await a.save(layout);
    expect(await a.load()).toEqual(layout);
  });

  it('does not auto-import legacy settings; explicit import preserves the old key', async () => {
    const legacyKey = 'todayIntranet.layout.dashboard|user';
    window.localStorage.setItem(legacyKey, JSON.stringify(layout));
    const a = store(JSON.stringify(['site', 'web', 'dashboard', 'user']));
    expect(await a.load()).toBeUndefined();
    expect(a.getStatus().actions).toContain('import-legacy');
    expect(await a.resolve('import-legacy')).toEqual(layout);
    expect(a.getStatus().actions).not.toContain('import-legacy');
    expect(window.localStorage.getItem(legacyKey)).toBe(JSON.stringify(layout));
    expect(await a.load()).toEqual(layout);
    expect(JSON.parse(a.exportRecovery()).browser.legacy.raw).toBe(JSON.stringify(layout));
  });

  it.each(['keep-local', 'use-remote'] as const)('blocks conflicting edits until explicit %s', async (action) => {
    const a = store();
    await a.save(layout);
    const other = cache();
    other.adopt(other.read().record);
    const remote = other.write(changed, metadata);
    await expect(a.save(attempted)).rejects.toThrow('Another browser tab');
    expect(a.getStatus()).toMatchObject({ state: 'conflict', canEdit: false });
    expect(a.getStatus().actions).toEqual(['keep-local', 'use-remote', 'export']);
    expect(other.read().record).toEqual(remote);
    expect(other.read().conflicts).toEqual([]);
    expect(JSON.parse(a.exportRecovery())).toMatchObject({ current: layout, unsaved: attempted });
    const selected = action === 'keep-local' ? attempted : changed;
    expect(await a.resolve(action)).toEqual(selected);
    expect(a.getStatus()).toMatchObject({ state: 'saved', canEdit: true });
    expect(other.read().record?.layout).toEqual(selected);
    expect(JSON.parse(a.exportRecovery()).browser.entries.some((entry: { key: string }) => entry.key.indexOf('.recovery.') >= 0)).toBe(true);
  });

  it('surfaces competing immutable heads on load without discarding either', async () => {
    const a = cache();
    const b = cache();
    a.write(layout, metadata);
    b.write(changed, metadata);
    const ui = store();
    await ui.load();
    expect(ui.getStatus()).toMatchObject({ state: 'conflict', canEdit: false });
    await expect(ui.save(layout)).rejects.toThrow();
    expect(a.read().conflicts).toHaveLength(2);
  });

  it('blocks corrupt settings, exports original raw data, and requires explicit reset', async () => {
    const a = cache();
    a.write(layout, metadata);
    const key = window.localStorage.key(0)!;
    const raw = '{invalid private settings';
    window.localStorage.setItem(key, raw);
    const ui = store();
    await expect(ui.load()).rejects.toThrow();
    expect(ui.getStatus()).toMatchObject({ state: 'error', reason: 'invalid-data', canEdit: false });
    expect(ui.getStatus().actions).toEqual(['retry', 'export', 'reset']);
    await expect(ui.save(layout)).rejects.toThrow();
    expect(JSON.parse(ui.exportRecovery()).browser.entries).toContainEqual({ key, raw });
    expect(await ui.resolve('reset', changed)).toEqual(changed);
    expect(ui.getStatus()).toMatchObject({ state: 'saved', canEdit: true });
    expect(ui.exportRecovery()).toContain('invalid private settings');
  });

  it('rejects quota failure rather than reporting success and retries the unsaved draft', async () => {
    const a = store();
    await a.save(layout);
    const set = jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('private'); });
    await expect(a.save(changed)).rejects.toThrow('Storage may be full or blocked');
    expect(a.getStatus()).toMatchObject({ state: 'error', reason: 'local-storage', canEdit: false });
    expect(JSON.parse(a.exportRecovery()).unsaved).toEqual(changed);
    window.dispatchEvent(new StorageEvent('storage', { key: window.localStorage.key(0), storageArea: window.localStorage }));
    expect(a.getStatus()).toMatchObject({ state: 'error', reason: 'local-storage', canEdit: false });
    set.mockRestore();
    expect(await a.resolve('retry')).toEqual(changed);
    expect(a.getStatus().state).toBe('saved');
    expect(cache().read().record?.layout).toEqual(changed);
  });

  it.each(['getItem', 'length', 'localStorage'] as const)
  ('exports current and unsaved memory with an explicit partial warning when %s is blocked', async (blocked) => {
    const a = store();
    await a.save(layout);
    const write = jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('quota'); });
    await expect(a.save(attempted)).rejects.toThrow();
    write.mockRestore();
    const failRead = (): never => { throw new Error('private storage policy detail'); };
    const read = blocked === 'getItem' ? jest.spyOn(Storage.prototype, 'getItem').mockImplementation(failRead)
      : blocked === 'length' ? jest.spyOn(Storage.prototype, 'length', 'get').mockImplementation(failRead)
        : jest.spyOn(window, 'localStorage', 'get').mockImplementation(failRead);
    const exported = JSON.parse(a.exportRecovery());
    expect(exported).toMatchObject({ partial: true, current: layout, unsaved: attempted });
    expect(exported).not.toHaveProperty('browser');
    expect(exported.browserReadError).toContain('Browser backups could not be read');
    expect(exported.browserReadError).not.toContain('private');
    expect(a.getStatus()).toMatchObject({ state: 'error', reason: 'local-storage', canEdit: false,
      message: exported.browserReadError, actions: ['retry', 'export'] });
    read.mockRestore();
    expect(await a.resolve('retry')).toEqual(attempted);
    const complete = JSON.parse(a.exportRecovery());
    expect(complete.partial).not.toBe(true);
    expect(complete.browser.entries).toBeDefined();
  });

  it('explains a partial export when blocked browser access leaves no in-memory settings', () => {
    const a = store();
    jest.spyOn(Storage.prototype, 'length', 'get').mockImplementation(() => { throw new Error('private'); });
    const exported = JSON.parse(a.exportRecovery());
    expect(exported.partial).toBe(true);
    expect(exported.browserReadError).toContain('no in-memory settings');
    expect(exported).not.toHaveProperty('current');
    expect(exported).not.toHaveProperty('unsaved');
    expect(a.getStatus()).toMatchObject({ state: 'error', reason: 'local-storage', canEdit: false,
      message: exported.browserReadError });
  });

  it('preserves an attempted draft before a browser read failure and exports it from memory', async () => {
    const a = store();
    await a.save(layout);
    jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => { throw new Error('private'); });
    await expect(a.save(attempted)).rejects.toThrow('Storage may be full or blocked');
    expect(JSON.parse(a.exportRecovery())).toMatchObject({ partial: true, current: layout, unsaved: attempted });
  });

  it('keeps recovery export usable when a caller supplies cyclic invalid settings', async () => {
    const a = store();
    await a.save(layout);
    const settings: Record<string, unknown> = {};
    settings.self = settings;
    await expect(a.save({ ...layout, widgets: [{ ...layout.widgets[0], settings }] })).rejects.toThrow('invalid');
    expect(a.getStatus()).toMatchObject({ state: 'error', reason: 'invalid-data', canEdit: false });
    expect(JSON.parse(a.exportRecovery()).current).toEqual(layout);
    expect(cache().read().record?.layout).toEqual(layout);
  });

  it('notifies subscribers of external changes and stops after disposal', async () => {
    const a = store();
    await a.save(layout);
    const listener = jest.fn();
    a.subscribe(listener);
    const external = cache();
    external.adopt(external.read().record);
    external.write(changed, metadata);
    const key = window.localStorage.key(0)!;
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(1);
    expect(a.getStatus()).toMatchObject({ state: 'conflict', canEdit: false });
    a.dispose();
    window.dispatchEvent(new StorageEvent('storage', { key, storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it('loads and receives external changes when the same store is reused after disposal', async () => {
    const a = store();
    await a.save(layout);
    a.dispose();
    expect(await a.load()).toEqual(layout);
    expect(a.getStatus()).toMatchObject({ state: 'saved', canEdit: true });
    const listener = jest.fn();
    a.subscribe(listener);
    const external = cache();
    external.adopt(external.read().record);
    external.write(changed, metadata);
    window.dispatchEvent(new StorageEvent('storage', { key: window.localStorage.key(0), storageArea: window.localStorage }));
    expect(listener).toHaveBeenCalledTimes(1);
    expect(a.getStatus()).toMatchObject({ state: 'conflict', canEdit: false });
  });
});
