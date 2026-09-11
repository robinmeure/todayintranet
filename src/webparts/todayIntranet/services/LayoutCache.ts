import { IDashboardLayout } from '../model/IDashboardLayout';
import { decodeLayout, parseLayout } from '../model/layoutDecoder';
import { ILayoutCacheRead, ILayoutRecord, ILayoutSyncMetadata } from './ILayoutCache';

export class LayoutCacheError extends Error {
  public constructor(message: string, public readonly reason: 'invalid-data' | 'local-storage') {
    super(message);
    this.name = 'LayoutCacheError';
    Object.setPrototypeOf(this, LayoutCacheError.prototype);
  }
}

interface IRawEntry { key: string; raw: string; }
interface ISnapshot extends IRawEntry { record: ILayoutRecord; }

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

function uuid(): string {
  try {
    const bytes = window.crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 15) | 64;
    bytes[8] = (bytes[8] & 63) | 128;
    const hex = Array.prototype.map.call(bytes, (byte: number) => (`0${byte.toString(16)}`).slice(-2)).join('');
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  } catch {
    throw new LayoutCacheError('Browser settings storage is unavailable.', 'local-storage');
  }
}

function invalid(): never {
  throw new LayoutCacheError('Browser settings are invalid or use an unsupported version. Export a backup before resetting.', 'invalid-data');
}

function object(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}

function dominates(left: Record<string, number>, right: Record<string, number>): boolean {
  return Object.keys(right).every((writer) => (left[writer] || 0) >= right[writer]) &&
    Object.keys(left).some((writer) => left[writer] > (right[writer] || 0));
}

/** Immutable, scoped browser checkpoints. Reading never changes a writer's ancestry. */
export class LayoutCache {
  private readonly _prefix: string;
  private readonly _writer: string;
  private _clock: Record<string, number> = {};
  private _revision: string | undefined;
  private _counter: number = 0;
  private readonly _listeners: Set<() => void> = new Set();
  private _disposed: boolean = false;

  public constructor(private readonly _scope: string, private readonly _legacyKey?: string) {
    try {
      if (typeof _scope !== 'string' || !_scope) {
        invalid();
      }
      // encodeURIComponent leaves dots intact; escape them too so no scope can
      // masquerade as a child key in another scope's namespace.
      const encoded = encodeURIComponent(_scope).replace(/[.!'()*]/g, (character) =>
        `%${character.charCodeAt(0).toString(16).toUpperCase()}`);
      this._prefix = `todayIntranet.layout.v3.${encoded}.`;
    } catch {
      throw new LayoutCacheError('The browser settings scope is invalid.', 'invalid-data');
    }
    this._writer = uuid();
    this._listen();
  }

  private _listen(): void {
    try {
      window.addEventListener('storage', this._onStorage);
      this._disposed = false;
    } catch {
      throw new LayoutCacheError('Browser settings notifications are unavailable.', 'local-storage');
    }
  }

  private _storage<T>(operation: (storage: Storage) => T): T {
    if (this._disposed) {
      throw new LayoutCacheError('Browser settings storage has been closed.', 'local-storage');
    }
    try {
      return operation(window.localStorage);
    } catch (error) {
      if (error instanceof LayoutCacheError) {
        throw error;
      }
      throw new LayoutCacheError('Browser settings could not be read or saved. Storage may be full or blocked. Keep this page open and export a backup.', 'local-storage');
    }
  }

  private _scan(): IRawEntry[] {
    return this._storage((storage) => {
      const keys = (): string[] => {
        const found: string[] = [];
        for (let index = 0; index < storage.length; index++) {
          const key = storage.key(index);
          if (key !== null && key.indexOf(this._prefix) === 0 && found.indexOf(key) < 0) {
            found.push(key);
          }
        }
        return found.sort();
      };
      for (let attempt = 0; attempt < 3; attempt++) {
        const before = keys();
        const entries: IRawEntry[] = [];
        before.forEach((key) => {
          const raw = storage.getItem(key);
          if (raw !== null) {
            entries.push({ key, raw });
          }
        });
        if (entries.length === before.length && JSON.stringify(before) === JSON.stringify(keys()) &&
          entries.every((entry) => storage.getItem(entry.key) === entry.raw)) {
          return entries;
        }
      }
      throw new LayoutCacheError('Browser settings changed while being read. Retry before editing.', 'local-storage');
    });
  }

  private _active(entries: IRawEntry[]): IRawEntry[] {
    return entries.filter((entry) => entry.key.indexOf(`${this._prefix}recovery.`) !== 0);
  }

  private _validate(value: unknown, key?: string): ILayoutRecord {
    if (!object(value) || value.schema !== 1 || value.scope !== this._scope ||
      typeof value.revision !== 'string' || !UUID.test(value.revision) ||
      key !== undefined && key !== `${this._prefix}snapshot.${value.revision}` ||
      !object(value.clock) || !Object.keys(value.clock).length) {
      return invalid();
    }
    const clock: Record<string, number> = {};
    for (const writer of Object.keys(value.clock)) {
      const counter = value.clock[writer];
      if (!UUID.test(writer) || typeof counter !== 'number' || !Number.isSafeInteger(counter) || counter < 1) {
        return invalid();
      }
      clock[writer] = counter;
    }
    const metadata = this._metadata(value);
    let layout: IDashboardLayout;
    try {
      layout = decodeLayout(value.layout);
    } catch {
      return invalid();
    }
    return { schema: 1, scope: this._scope, revision: value.revision, clock, layout, ...metadata };
  }

  private _metadata(value: ILayoutSyncMetadata | Record<string, unknown>): ILayoutSyncMetadata {
    if (!object(value) || typeof value.pendingSync !== 'boolean' || typeof value.baseKnown !== 'boolean') {
      return invalid();
    }
    const metadata: ILayoutSyncMetadata = { pendingSync: value.pendingSync, baseKnown: value.baseKnown };
    if (value.remoteCheckedAt !== undefined) {
      if (typeof value.remoteCheckedAt !== 'number' || !Number.isSafeInteger(value.remoteCheckedAt) ||
          value.remoteCheckedAt < 0) {
        return invalid();
      }
      metadata.remoteCheckedAt = value.remoteCheckedAt;
    }
    if (value.cloud !== undefined) {
      if (!object(value.cloud) || typeof value.cloud.writable !== 'boolean' ||
          typeof value.cloud.configurationCheckedAt !== 'number' ||
          !Number.isSafeInteger(value.cloud.configurationCheckedAt) ||
          value.cloud.configurationCheckedAt < 0) {
        return invalid();
      }
      // Builds that incorrectly required Title uniqueness stored a writeBlock.
      // SharePoint forbids that setting with own-item visibility, so discard this
      // legacy capability result and force one fresh configuration check.
      if (value.cloud.writeBlock === undefined) {
        metadata.cloud = {
          writable: value.cloud.writable,
          configurationCheckedAt: value.cloud.configurationCheckedAt
        };
      }
    }
    const server = value.server;
    if (server === undefined) {
      return metadata;
    }
    if (!object(server)) {
      return invalid();
    }
    const { itemId, etag } = server;
    if (typeof itemId !== 'number' || !Number.isSafeInteger(itemId) || itemId < 1 ||
      typeof etag !== 'string' || !etag.trim() || !metadata.baseKnown) {
      return invalid();
    }
    return { ...metadata, server: { itemId, etag } };
  }

  private _snapshots(entries: IRawEntry[]): ISnapshot[] {
    return this._active(entries).map((entry) => {
      let value: unknown;
      try {
        value = JSON.parse(entry.raw);
      } catch {
        return invalid();
      }
      return { ...entry, record: this._validate(value, entry.key) };
    });
  }

  private _heads(snapshots: ISnapshot[]): ILayoutRecord[] {
    return snapshots.filter((candidate) => !snapshots.some((other) =>
      dominates(other.record.clock, candidate.record.clock))).map((entry) => entry.record)
      .sort((left, right) => left.revision < right.revision ? -1 : left.revision === right.revision ? 0 : 1);
  }

  public read(): ILayoutCacheRead {
    const heads = this._heads(this._snapshots(this._scan()));
    const record = heads.filter((head) => head.revision === this._revision)[0] ?? heads[0];
    const legacyAvailable = heads.length === 0 && this._legacyKey !== undefined &&
      this._storage((storage) => storage.getItem(this._legacyKey!) !== null);
    return { record, conflicts: heads.length > 1 ? heads : [], legacyAvailable };
  }

  public adopt(record?: ILayoutRecord): void {
    const valid = record === undefined ? undefined : this._validate(record);
    this._clock = { ...valid?.clock };
    this._revision = valid?.revision;
  }

  private _next(layout: IDashboardLayout, metadata: ILayoutSyncMetadata, clock: Record<string, number>): ILayoutRecord {
    let valid: IDashboardLayout;
    try {
      valid = decodeLayout(layout);
    } catch {
      return invalid();
    }
    const counter = Math.max(this._counter, clock[this._writer] || 0) + 1;
    if (!Number.isSafeInteger(counter)) {
      return invalid();
    }
    return { schema: 1, scope: this._scope, revision: uuid(), clock: { ...clock, [this._writer]: counter },
      layout: valid, ...this._metadata(metadata) };
  }

  private _put(key: string, raw: string): void {
    this._storage((storage) => {
      if (storage.getItem(key) !== null) {
        throw new LayoutCacheError('A browser settings revision already exists. Retry the save.', 'local-storage');
      }
      storage.setItem(key, raw);
      if (storage.getItem(key) !== raw) {
        throw new LayoutCacheError('Browser settings were not retained. Keep this page open and export a backup.', 'local-storage');
      }
    });
  }

  private _persist(record: ILayoutRecord): void {
    this._put(`${this._prefix}snapshot.${record.revision}`, JSON.stringify(record));
    this._counter = record.clock[this._writer];
    this.adopt(record);
  }

  private _remove(entries: IRawEntry[]): void {
    this._storage((storage) => {
      entries.forEach((entry) => {
        // Never remove a key that changed after the scan, including during reset.
        if (storage.getItem(entry.key) === entry.raw) {
          storage.removeItem(entry.key);
          if (storage.getItem(entry.key) === entry.raw) {
            throw new LayoutCacheError('Old browser settings could not be removed. Export a backup and retry.', 'local-storage');
          }
        }
      });
    });
  }

  public write(layout: IDashboardLayout, metadata: ILayoutSyncMetadata): ILayoutRecord {
    const snapshots = this._snapshots(this._scan());
    const record = this._next(layout, metadata, this._clock);
    this._persist(record);
    this._remove(snapshots.filter((entry) => dominates(record.clock, entry.record.clock)));
    return record;
  }

  private _merge(snapshots: ISnapshot[]): Record<string, number> {
    const clock = { ...this._clock };
    snapshots.forEach(({ record }) => {
      Object.keys(record.clock).forEach((writer) => {
        clock[writer] = Math.max(clock[writer] || 0, record.clock[writer]);
      });
    });
    return clock;
  }

  private _archiveEntries(entries: IRawEntry[]): void {
    if (entries.length) {
      this._put(`${this._prefix}recovery.${uuid()}`, JSON.stringify({ schema: 1, scope: this._scope, entries }));
    }
  }

  /**
   * Durably preserve exact server content in an immutable, scoped backup.
   * Callers must not overwrite server data if this verified write throws.
   */
  public archive(raw: string): void {
    if (typeof raw !== 'string') {
      invalid();
    }
    this._put(`${this._prefix}recovery.${uuid()}`, JSON.stringify({ schema: 1, scope: this._scope, raw }));
  }

  public resolve(layout: IDashboardLayout, metadata: ILayoutSyncMetadata): ILayoutRecord {
    const snapshots = this._snapshots(this._scan());
    const record = this._next(layout, metadata, this._merge(snapshots));
    const discarded = snapshots.filter((entry) => dominates(record.clock, entry.record.clock));
    this._archiveEntries(discarded.map(({ key, raw }) => ({ key, raw })));
    this._persist(record);
    this._remove(discarded);
    return record;
  }

  public reset(layout: IDashboardLayout, metadata: ILayoutSyncMetadata): ILayoutRecord {
    const entries = this._active(this._scan());
    const valid: ISnapshot[] = [];
    entries.forEach((entry) => {
      try {
        valid.push(...this._snapshots([entry]));
      } catch (error) {
        // Explicit reset is the only operation allowed to replace invalid data.
        if (!(error instanceof LayoutCacheError) || error.reason !== 'invalid-data') {
          throw error;
        }
      }
    });
    const record = this._next(layout, metadata, this._merge(valid));
    this._archiveEntries(entries);
    this._persist(record);
    this._remove(entries);
    return record;
  }

  public readLegacy(): IDashboardLayout | undefined {
    if (!this._legacyKey) {
      return undefined;
    }
    const raw = this._storage((storage) => storage.getItem(this._legacyKey!));
    if (raw === null) {
      return undefined;
    }
    try {
      return parseLayout(raw);
    } catch {
      return invalid();
    }
  }

  public exportRecovery(): string {
    const entries = this._scan();
    const legacy = this._legacyKey === undefined ? undefined
      : { key: this._legacyKey, raw: this._storage((storage) => storage.getItem(this._legacyKey!)) };
    return JSON.stringify({ schema: 1, scope: this._scope, entries, legacy }, undefined, 2);
  }

  private readonly _onStorage = (event: StorageEvent): void => {
    if (this._disposed || event.key !== null && event.key !== this._legacyKey &&
      (event.key.indexOf(this._prefix) !== 0 || event.key.indexOf(`${this._prefix}recovery.`) === 0)) {
      return;
    }
    this._storage((storage) => {
      if (event.storageArea === null || event.storageArea === storage) {
        this._listeners.forEach((listener) => listener());
      }
    });
  };

  public subscribe(listener: () => void): () => void {
    if (this._disposed) {
      this._listen();
    }
    this._listeners.add(listener);
    return () => { this._listeners.delete(listener); };
  }

  public dispose(): void {
    try {
      window.removeEventListener('storage', this._onStorage);
    } catch {
      throw new LayoutCacheError('Browser settings notifications could not be closed.', 'local-storage');
    } finally {
      this._listeners.clear();
      this._disposed = true;
    }
  }
}
