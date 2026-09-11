import { IDashboardLayout } from '../model/IDashboardLayout';
import { decodeLayout } from '../model/layoutDecoder';
import { ILayoutRecord } from './ILayoutCache';
import { ILayoutStore, ILayoutStoreStatus, LayoutStoreAction } from './ILayoutStore';
import { LayoutCache, LayoutCacheError } from './LayoutCache';

const LOCAL_METADATA = { pendingSync: false, baseKnown: false };

/** Browser-only persistence with explicit conflict and recovery actions. */
export class LocalStorageLayoutStore implements ILayoutStore {
  private readonly _cache: LayoutCache;
  private readonly _listeners: Set<() => void> = new Set();
  private _unsubscribe: (() => void) | undefined;
  private _record: ILayoutRecord | undefined;
  private _unsaved: IDashboardLayout | undefined;
  private _legacyAvailable: boolean = false;
  private _status: ILayoutStoreStatus = {
    mode: 'local', state: 'loading', message: 'Loading browser settings...', canEdit: false, actions: []
  };

  public constructor(scopeKey: string, private readonly _message?: string) {
    let legacyScope = scopeKey;
    try {
      const scope: unknown = JSON.parse(scopeKey);
      if (Array.isArray(scope) && scope.length === 4 && scope.every((part) => typeof part === 'string')) {
        legacyScope = `${scope[2]}|${scope[3]}`;
      }
    } catch {
      // The original constructor also accepts the legacy dashboard|user scope.
    }
    this._cache = new LayoutCache(scopeKey, `todayIntranet.layout.${legacyScope}`);
    this._unsubscribe = this._cache.subscribe(this._onExternalChange);
  }

  private readonly _onExternalChange = (): void => {
    try {
      this._refresh(false);
    } catch (error) {
      this._error(error);
    }
  };

  private _set(status: ILayoutStoreStatus): void {
    this._status = status;
    this._listeners.forEach((listener) => listener());
  }

  private _refresh(adopt: boolean, checkingSave: boolean = false): IDashboardLayout | undefined {
    const result = this._cache.read();
    this._legacyAvailable = result.legacyAvailable;
    if (adopt) {
      this._record = result.record;
      this._cache.adopt(result.record);
    }
    if (result.conflicts.length || !adopt && result.record?.revision !== this._record?.revision) {
      this._set({ mode: 'local', state: 'conflict', message: 'Another browser tab changed these settings. Export a backup, then keep your version or use the other tab’s version.',
        canEdit: false, actions: ['keep-local', 'use-remote', 'export'] });
    } else if (this._unsaved && !checkingSave) {
      this._error(new LayoutCacheError('The latest browser settings have not been saved.', 'local-storage'));
    } else if (!checkingSave) {
      this._set({ mode: 'local', state: 'saved',
        message: this._message ? `${this._message} Settings are saved in this browser only.` : 'Settings are saved in this browser only.',
        canEdit: true, actions: this._legacyAvailable ? ['import-legacy', 'export'] : ['export'] });
    }
    return this._record?.layout;
  }

  private _error(error: unknown): void {
    const invalid = error instanceof LayoutCacheError && error.reason === 'invalid-data';
    this._set({ mode: 'local', state: 'error', reason: invalid ? 'invalid-data' : 'local-storage',
      message: invalid ? 'Browser settings are invalid or use an unsupported version. Editing is blocked; export a backup before resetting.'
        : 'Browser settings could not be read or saved. Keep this page open, export a backup, and retry.',
      canEdit: false, actions: invalid ? ['retry', 'export', 'reset'] : ['retry', 'export'] });
  }

  public async load(): Promise<IDashboardLayout | undefined> {
    try {
      this._unsubscribe?.();
      this._unsubscribe = this._cache.subscribe(this._onExternalChange);
      return this._refresh(true);
    } catch (error) {
      this._error(error);
      throw error;
    }
  }

  public async save(layout: IDashboardLayout): Promise<void> {
    if (this._status.state === 'loading') {
      try {
        this._refresh(true);
      } catch (error) {
        this._error(error);
        throw error;
      }
    }
    try {
      let valid: IDashboardLayout;
      try {
        valid = decodeLayout(layout);
      } catch {
        throw new LayoutCacheError('The edited browser settings are invalid. Export a backup before resetting.', 'invalid-data');
      }
      this._unsaved = valid;
      if (!this._status.canEdit) {
        throw new LayoutCacheError('Resolve the browser settings problem before editing.', this._status.reason === 'invalid-data' ? 'invalid-data' : 'local-storage');
      }
      this._refresh(false, true);
      if (!this._status.canEdit) {
        throw new LayoutCacheError('Another browser tab changed these settings. Choose which version to keep before editing.', 'local-storage');
      }
      this._record = this._cache.write(valid, LOCAL_METADATA);
      this._unsaved = undefined;
      this._refresh(false);
    } catch (error) {
      if (this._status.state !== 'conflict') {
        this._error(error);
      }
      throw error;
    }
  }

  public getStatus(): ILayoutStoreStatus {
    return { ...this._status, actions: this._status.actions.slice() };
  }

  public async publish(): Promise<void> {
    // Browser-only stores are already durable after save().
  }

  public subscribe(listener: () => void): () => void {
    this._listeners.add(listener);
    return () => { this._listeners.delete(listener); };
  }

  public async resolve(action: LayoutStoreAction, starterLayout?: IDashboardLayout): Promise<IDashboardLayout | undefined> {
    try {
      if (action === 'export') {
        return this._unsaved ?? this._record?.layout;
      }
      if (action === 'retry') {
        if (this._unsaved && this._status.reason !== 'invalid-data') {
          const current = this._cache.read();
          if (current.conflicts.length || current.record?.revision !== this._record?.revision) {
            this._refresh(false);
            return this._unsaved;
          }
          this._record = this._cache.write(this._unsaved, LOCAL_METADATA);
          this._unsaved = undefined;
        }
        return this._refresh(true);
      }
      let layout: IDashboardLayout | undefined;
      if (action === 'reset') {
        if (!starterLayout) {
          throw new LayoutCacheError('A starter layout is required to reset browser settings.', 'invalid-data');
        }
        this._record = this._cache.reset(starterLayout, LOCAL_METADATA);
      } else {
        if (action === 'import-legacy') {
          layout = this._cache.readLegacy();
        } else if (action === 'keep-local') {
          layout = this._unsaved ?? this._record?.layout;
        } else if (action === 'use-remote') {
          const current = this._cache.read();
          layout = (current.conflicts.filter((record) => record.revision !== this._record?.revision)[0] ?? current.record)?.layout;
        }
        if (!layout) {
          throw new LayoutCacheError('The selected browser settings version is unavailable. Retry or export a backup.', 'local-storage');
        }
        this._record = this._cache.resolve(layout, LOCAL_METADATA);
      }
      this._unsaved = undefined;
      return this._refresh(false);
    } catch (error) {
      this._error(error);
      throw error;
    }
  }

  public exportRecovery(): string {
    let browser: unknown;
    try {
      browser = JSON.parse(this._cache.exportRecovery());
    } catch {
      const browserReadError = 'Browser backups could not be read. ' +
        (this._record || this._unsaved ? 'This partial export includes only settings available in this page.'
          : 'This partial export contains no in-memory settings.');
      this._set({ mode: 'local', state: 'error', reason: 'local-storage', message: browserReadError,
        canEdit: false, actions: ['retry', 'export'] });
      return JSON.stringify({ partial: true, browserReadError, current: this._record?.layout, unsaved: this._unsaved }, undefined, 2);
    }
    return JSON.stringify({ browser, current: this._record?.layout, unsaved: this._unsaved }, undefined, 2);
  }

  public dispose(): void {
    this._unsubscribe?.();
    this._unsubscribe = undefined;
    this._cache.dispose();
    this._listeners.clear();
  }
}
