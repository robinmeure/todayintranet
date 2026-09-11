import { SPHttpClient } from '@microsoft/sp-http';
import { IDashboardLayout } from '../model/IDashboardLayout';
import { decodeLayout } from '../model/layoutDecoder';
import { ILayoutStore, ILayoutStoreStatus, ILayoutScope, LayoutStoreAction, layoutScopeKey } from './ILayoutStore';
import { ILayoutRecord, ILayoutSyncMetadata, IServerLayoutVersion } from './ILayoutCache';
import { LayoutCache, LayoutCacheError } from './LayoutCache';
import { IRemoteLayout, LayoutStorageError, SharePointLayoutClient } from './SharePointLayoutClient';

export const LAYOUT_CACHE_MAX_AGE_MS: number = 5 * 60 * 1000;
export const CONFIGURATION_CACHE_MAX_AGE_MS: number = 60 * 60 * 1000;
const MAX_RETRIES: number = 3;
const LOCAL_MESSAGE: string = 'Saved in this browser only. These settings do not follow you to other devices.';

export interface ISharePointLayoutStoreOptions extends ILayoutScope {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
}

function sameVersion(a: IServerLayoutVersion | undefined, b: IServerLayoutVersion | undefined): boolean {
  return a?.itemId === b?.itemId && a?.etag === b?.etag;
}

function sameLayout(a: IDashboardLayout, b: IDashboardLayout): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

export class SharePointListLayoutStore implements ILayoutStore {
  private readonly _cache: LayoutCache;
  private _client: SharePointLayoutClient;
  private _status: ILayoutStoreStatus = {
    mode: 'local', state: 'loading', message: 'Loading your saved settings...', canEdit: false, actions: []
  };
  private readonly _listeners = new Set<() => void>();
  private _unsubscribeCache: (() => void) | undefined;
  private _record: ILayoutRecord | undefined;
  private _remote: IRemoteLayout | undefined;
  private _remoteInvalid: boolean = false;
  private _base: IServerLayoutVersion | undefined;
  private _baseKnown: boolean = false;
  private _configurationCheckedAt: number | undefined;
  private _remoteCheckedAt: number | undefined;
  private _resetVersion: IServerLayoutVersion | undefined;
  private _writable: boolean | undefined;
  private _pending: IDashboardLayout | undefined;
  private _localFailed: boolean = false;
  private _localInvalid: boolean = false;
  private _localConflicts: ILayoutRecord[] = [];
  private _legacyAvailable: boolean = false;
  private _timer: number | undefined;
  private _inFlight: Promise<void> | undefined;
  private _epoch: number = 0;
  private _sequence: number = 0;
  private _dueAt: number = 0;
  private _retryAt: number = 0;
  private _retries: number = 0;
  private _syncRequested: boolean = false;
  private _disposed: boolean = false;

  public constructor(private readonly _options: ISharePointLayoutStoreOptions) {
    const legacyKey = `todayIntranet.layout.${`${_options.dashboardKey}|${_options.userKey}`.substring(0, 255)}`;
    this._cache = new LayoutCache(layoutScopeKey(_options), legacyKey);
    this._client = this._newClient();
  }

  private _newClient(): SharePointLayoutClient {
    return new SharePointLayoutClient({
      spHttpClient: this._options.spHttpClient,
      webAbsoluteUrl: this._options.webAbsoluteUrl,
      itemKey: `${this._options.dashboardKey}|${this._options.userKey}`
    });
  }

  public getStatus(): ILayoutStoreStatus {
    return { ...this._status, actions: [...this._status.actions] };
  }

  public subscribe(listener: () => void): () => void {
    this._listeners.add(listener);
    return () => this._listeners.delete(listener);
  }

  private _setStatus(status: ILayoutStoreStatus): void {
    if (!this._disposed) {
      this._status = status;
      this._listeners.forEach((listener) => listener());
    }
  }

  private _mode(): 'sharepoint' | 'local' {
    return this._writable ? 'sharepoint' : 'local';
  }

  private _withLegacy(actions: LayoutStoreAction[] = []): LayoutStoreAction[] {
    return this._legacyAvailable ? [...actions, 'import-legacy'] : actions;
  }

  public async load(): Promise<IDashboardLayout | undefined> {
    this._disposed = false;
    const epoch = ++this._epoch;
    this._clearTimer();
    this._client.dispose();
    this._client = this._newClient();
    this._writable = undefined;
    this._pending = undefined;
    this._inFlight = undefined;
    this._retryAt = 0;
    this._retries = 0;
    this._syncRequested = false;
    this._unsubscribeCache?.();
    this._unsubscribeCache = this._cache.subscribe(() => this._onExternalChange());
    window.removeEventListener('online', this._onOnline);
    window.addEventListener('online', this._onOnline);
    this._setStatus({ mode: 'local', state: 'loading', message: 'Loading your saved settings...', canEdit: false, actions: [] });
    try {
      const local = this._cache.read();
      this._record = local.record;
      this._cache.adopt(local.record);
      this._base = local.record?.server;
      this._baseKnown = local.record?.baseKnown ?? false;
      this._configurationCheckedAt = local.record?.cloud?.configurationCheckedAt;
      this._remoteCheckedAt = local.record?.remoteCheckedAt;
      this._legacyAvailable = local.legacyAvailable;
      this._localConflicts = local.conflicts;
      this._localFailed = false;
      this._localInvalid = false;
      if (local.conflicts.length) {
        this._conflict('This dashboard has competing changes from another browser tab. Export a backup, then choose which version to keep.');
        return local.record?.layout;
      }
    } catch (error) {
      this._record = undefined;
      this._base = undefined;
      this._baseKnown = false;
      this._localFailed = true;
      if (error instanceof LayoutCacheError && error.reason === 'invalid-data') {
        this._localError(error);
        throw error;
      }
      this._localError(error);
    }
    return this._connect(epoch, false);
  }

  private async _connect(epoch: number, force: boolean): Promise<IDashboardLayout | undefined> {
    const client = this._client;
    try {
      const now = Date.now();
      const cachedCloud = this._record?.cloud;
      const configurationFresh = !force && cachedCloud !== undefined &&
        now - cachedCloud.configurationCheckedAt < CONFIGURATION_CACHE_MAX_AGE_MS;
      const layoutFresh = !force && this._record !== undefined && !this._record.pendingSync &&
        this._record.remoteCheckedAt !== undefined &&
        now - this._record.remoteCheckedAt < LAYOUT_CACHE_MAX_AGE_MS;
      if (configurationFresh) {
        client.prime(cachedCloud.writable);
        this._writable = cachedCloud.writable;
        this._configurationCheckedAt = cachedCloud.configurationCheckedAt;
      }
      if (configurationFresh && layoutFresh && this._record) {
        if (cachedCloud.writable) {
          this._setCloudSaved();
        } else {
          this._setLocalStatus();
        }
        return this._record.layout;
      }
      const writable = await client.initialize();
      if (!this._active(epoch)) {
        return undefined;
      }
      this._writable = writable;
      if (!configurationFresh) {
        this._configurationCheckedAt = Date.now();
      }
      if (!writable) {
        this._remoteInvalid = false;
      }
      if (!writable && (this._pending || this._record)) {
        const layout = this._pending ?? this._record?.layout;
        if (this._pending) {
          this._checkpoint(this._pending, this._metadata(false, false));
        } else if (this._record) {
          this._checkpoint(this._record.layout, this._metadata(false, false));
        }
        if (!this._localFailed) {
          this._pending = undefined;
        }
        this._setLocalStatus();
        return layout;
      }
      const remote = await client.read();
      if (!this._active(epoch)) {
        return undefined;
      }
      this._remote = remote;
      this._remoteInvalid = false;
      this._remoteCheckedAt = Date.now();
      const local = this._pending ?? this._record?.layout;
      if (!writable) {
        if (remote) {
          this._checkpoint(remote.layout, this._metadata(false, false));
        }
        this._setLocalStatus();
        return remote?.layout;
      }
      const hasDraft = !!this._pending || !!this._record?.pendingSync || (!!this._record && !this._record.baseKnown);
      if (local && hasDraft) {
        if (remote && sameLayout(local, remote.layout)) {
          this._base = remote;
          this._baseKnown = true;
          this._pending = undefined;
          this._checkpoint(local, this._metadata(false, true, remote));
          this._setCloudSaved();
        } else if (this._baseKnown && sameVersion(this._base, remote)) {
          this._pending = local;
          this._queue();
        } else {
          this._pending = local;
          this._conflict('Your browser has changes based on a different or unknown SharePoint version. Choose which version to keep; neither will be overwritten automatically.');
        }
        return local;
      }
      this._base = remote;
      this._baseKnown = true;
      if (remote && (!this._record || !sameLayout(this._record.layout, remote.layout) ||
          !sameVersion(this._record.server, remote) || this._record.pendingSync)) {
        this._checkpoint(remote.layout, this._metadata(false, true, remote));
      }
      this._setCloudSaved();
      return remote?.layout;
    } catch (error) {
      if (this._active(epoch)) {
        this._remoteError(error);
      }
      return this._pending ?? this._record?.layout;
    }
  }

  public async save(layout: IDashboardLayout): Promise<void> {
    if (this._disposed || !this._status.canEdit) {
      throw new Error('Settings cannot be edited until storage recovery is complete.');
    }
    let valid: IDashboardLayout;
    try {
      valid = decodeLayout(layout);
    } catch (error) {
      this._setStatus({ mode: this._mode(), state: 'error', reason: 'invalid-data',
        message: 'These settings are invalid and could not be saved. Export a backup before resetting.',
        canEdit: false, actions: ['export', 'reset'] });
      throw error;
    }
    this._sequence++;
    this._pending = valid;
    this._legacyAvailable = false;
    this._checkpoint(valid, this._metadata(true));
    if (this._localInvalid) {
      return;
    }
    if (!this._checkLocalConflict()) {
      return;
    }
    if (this._writable === false) {
      if (!this._localFailed) {
        this._pending = undefined;
      }
      this._setLocalStatus();
      return;
    }
    this._setPending('Saved in this browser. Click Done to save these changes to SharePoint.');
  }

  public async publish(): Promise<void> {
    if (this._disposed || !this._pending || this._writable === false) {
      return;
    }
    this._syncRequested = true;
    this._dueAt = Date.now();
    this._clearTimer();
    this._startSync(this._epoch);
    const current = this._inFlight;
    if (current) {
      await current;
    }
  }

  private _metadata(
    pending: boolean,
    baseKnown: boolean = this._baseKnown,
    server: IServerLayoutVersion | undefined = this._base
  ): ILayoutSyncMetadata {
    return {
      pendingSync: pending && this._writable !== false,
      baseKnown: this._writable !== false && baseKnown,
      server: this._writable !== false ? server : undefined,
      ...(this._configurationCheckedAt !== undefined && this._writable !== undefined ? {
        cloud: {
          writable: this._writable,
          configurationCheckedAt: this._configurationCheckedAt
        }
      } : {}),
      ...(this._remoteCheckedAt !== undefined ? { remoteCheckedAt: this._remoteCheckedAt } : {})
    };
  }

  private _checkpoint(layout: IDashboardLayout, metadata: ILayoutSyncMetadata): void {
    try {
      this._record = this._cache.write(layout, metadata);
      this._localFailed = false;
      this._localInvalid = false;
      this._checkLocalConflict();
    } catch (error) {
      this._localFailed = true;
      this._localError(error);
    }
  }

  private _checkLocalConflict(): boolean {
    if (this._localFailed) {
      return true;
    }
    try {
      const local = this._cache.read();
      if (local.conflicts.length) {
        this._localConflicts = local.conflicts;
        this._clearTimer();
        this._conflict('Another tab has different settings for this dashboard. Both drafts are preserved. Choose which version to keep.');
        return false;
      }
      if (local.record && this._record && local.record.revision !== this._record.revision) {
        this._localConflicts = [this._record, local.record];
        this._clearTimer();
        this._conflict('Another tab updated this dashboard. Choose whether to keep your displayed settings or load the saved version.');
        return false;
      }
      return true;
    } catch (error) {
      this._localError(error);
      return false;
    }
  }

  private _queue(): void {
    this._setPending(this._syncRequested
      ? 'Saved in this browser. Saving to SharePoint was requested by Done.'
      : 'Saved in this browser. Click Done to save these changes to SharePoint.');
    if (!this._syncRequested) {
      return;
    }
    this._schedule(Math.max(this._dueAt, this._retryAt, Date.now()) - Date.now());
  }

  private _schedule(delay: number): void {
    this._clearTimer();
    const epoch = this._epoch;
    this._timer = window.setTimeout(() => {
      this._timer = undefined;
      if (this._active(epoch)) {
        this._startSync(epoch);
      }
    }, Math.min(Math.max(0, delay), 2147483647));
  }

  private _startSync(epoch: number): void {
    if (this._inFlight || !this._active(epoch) || !this._status.canEdit) {
      return;
    }
    if (Date.now() < this._retryAt || Date.now() < this._dueAt) {
      this._schedule(Math.max(this._retryAt, this._dueAt) - Date.now());
      return;
    }
    const finish = (): void => {
      if (this._active(epoch)) {
        this._inFlight = undefined;
        if (this._pending && this._writable && this._syncRequested &&
            this._status.canEdit && this._timer === undefined) {
          this._queue();
        }
      }
    };
    this._inFlight = this._sync(epoch).then(finish, (error: unknown) => {
      if (this._active(epoch)) {
        this._remoteError(error);
      }
      finish();
    });
  }

  private async _sync(epoch: number): Promise<void> {
    if (!this._checkLocalConflict()) {
      return;
    }
    if (this._writable === undefined) {
      this._client.dispose();
      this._client = this._newClient();
      await this._connect(epoch, false);
      return;
    }
    const snapshot = this._pending;
    if (!snapshot || !this._writable) {
      return;
    }
    const sequence = this._sequence;
    const client = this._client;
    this._setStatus({ mode: 'sharepoint', state: 'saving', message: 'Saving to SharePoint...', canEdit: true, actions: [] });
    try {
      let remote: IRemoteLayout | undefined;
      let resettingInvalid = false;
      // A known item version (or known absence) can be written conditionally
      // without a speculative read. Unknown-base recovery still reconciles first.
      if (!this._baseKnown) {
        try {
          remote = await client.read();
        } catch (error) {
          if (error instanceof LayoutStorageError && error.reason === 'invalid-data' &&
              this._resetVersion && sameVersion(this._resetVersion, client.recoveryVersion)) {
            resettingInvalid = true;
          } else {
            throw error;
          }
        }
        if (!this._active(epoch)) {
          return;
        }
        this._remote = remote;
        this._remoteCheckedAt = Date.now();
        if (!resettingInvalid && remote && sameLayout(remote.layout, snapshot)) {
          this._acknowledge(remote, sequence);
          return;
        }
        if (!resettingInvalid) {
          this._conflict('The SharePoint layout changed. Your browser draft is preserved; choose which version to keep.');
          return;
        }
      }
      const saved = await client.write(snapshot, this._base);
      if (this._active(epoch)) {
        this._acknowledge(saved, sequence);
      }
    } catch (error) {
      if (!this._active(epoch)) {
        return;
      }
      if (error instanceof LayoutStorageError && error.status === 404 && this._baseKnown && this._base) {
        try {
          const found = await client.read();
          if (!this._active(epoch)) {
            return;
          }
          if (!found) {
            this._base = undefined;
            const recreated = await client.write(snapshot);
            if (this._active(epoch)) {
              this._acknowledge(recreated, sequence);
            }
          } else {
            this._remote = found;
            this._conflict('The SharePoint layout was replaced while saving. Your browser draft is preserved.');
          }
        } catch (readError) {
          if (this._active(epoch)) {
            this._remoteError(readError);
          }
        }
      } else if (error instanceof LayoutStorageError && error.status === 412) {
        this._conflict('Another client changed the SharePoint layout while saving. Your changes are preserved in this browser.');
      } else if (error instanceof LayoutStorageError && (error.status === 400 || error.status === 409)) {
        // Unique-value violations may be returned as 400, not 409.
        try {
          const competing = await client.read();
          if (!this._active(epoch)) {
            return;
          }
          if (competing) {
            this._remote = competing;
            if (sameLayout(competing.layout, snapshot)) {
              this._acknowledge(competing, sequence);
            } else {
              this._conflict('Another client created this dashboard layout first. Choose which version to keep.');
            }
          } else {
            this._remoteError(error);
          }
        } catch (readError) {
          if (this._active(epoch)) {
            this._remoteError(readError);
          }
        }
      } else {
        this._remoteError(error);
      }
    }
  }

  private _acknowledge(remote: IRemoteLayout, sequence: number): void {
    this._remote = remote;
    this._remoteCheckedAt = Date.now();
    this._base = remote;
    this._baseKnown = true;
    this._retries = 0;
    this._retryAt = 0;
    this._resetVersion = undefined;
    if (!this._checkLocalConflict()) {
      return;
    }
    if (sequence === this._sequence) {
      this._pending = undefined;
      this._syncRequested = false;
      this._checkpoint(remote.layout, this._metadata(false));
      this._setCloudSaved();
    } else if (this._pending) {
      this._checkpoint(this._pending, this._metadata(true));
      this._setPending('Saved in this browser. Newer changes are waiting to sync.');
    }
  }

  private _setPending(message: string): void {
    if (this._localInvalid || this._localConflicts.length) {
      return;
    }
    this._setStatus({ mode: this._mode(), state: this._localFailed ? 'error' : 'pending',
      reason: this._localFailed ? 'local-storage' : undefined,
      message: this._localFailed ? 'Your changes are not saved locally. A SharePoint save will be attempted; keep this page open.' : message,
      canEdit: true, actions: ['retry', 'export'] });
  }

  private _setCloudSaved(): void {
    if (this._localInvalid || this._localConflicts.length) {
      return;
    }
    this._setStatus({ mode: 'sharepoint', state: 'saved', reason: this._localFailed ? 'local-storage' : undefined,
      message: this._localFailed ? 'Saved to SharePoint. Browser recovery is unavailable because local storage failed.' : 'Saved to SharePoint.',
      canEdit: true, actions: this._withLegacy(this._localFailed ? ['retry', 'export'] : []) });
  }

  private _setLocalStatus(): void {
    if (this._localInvalid || this._localConflicts.length) {
      return;
    }
    this._setStatus({ mode: 'local', state: this._localFailed ? 'error' : 'saved',
      reason: this._localFailed ? 'local-storage' : 'read-only',
      message: this._localFailed ? 'Your changes could not be saved in this browser, and SharePoint storage is unavailable for this dashboard. Keep this page open and export a backup.'
        : `${LOCAL_MESSAGE} You have read-only access to the layouts list.`,
      canEdit: true, actions: this._withLegacy(['retry', 'export']) });
  }

  private _conflict(message: string): void {
    this._clearTimer();
    this._syncRequested = false;
    this._setStatus({ mode: this._mode(), state: 'conflict', message, canEdit: false,
      actions: ['keep-local', 'use-remote', 'export'] });
  }

  private _localError(error: unknown): void {
    const invalid = error instanceof LayoutCacheError && error.reason === 'invalid-data';
    this._localFailed = true;
    this._localInvalid = invalid;
    console.warn('[TodayIntranet] Browser settings storage failed.', invalid ? 'invalid-data' : 'local-storage');
    this._setStatus({ mode: this._mode(), state: 'error', reason: invalid ? 'invalid-data' : 'local-storage',
      message: invalid ? 'The browser settings are invalid or use an unsupported version. Export a backup before resetting.'
        : 'Browser settings could not be read or saved. Keep this page open and export a backup.',
      canEdit: !invalid && !this._remoteInvalid && !this._localConflicts.length,
      actions: invalid || this._remoteInvalid ? ['retry', 'export', 'reset'] : ['retry', 'export'] });
  }

  private _remoteError(error: unknown): void {
    const failure = error instanceof LayoutStorageError ? error
      : new LayoutStorageError('SharePoint returned an unreadable response.', 'configuration');
    console.warn('[TodayIntranet] SharePoint settings storage failed.', failure.reason, failure.status);
    if (failure.reason === 'invalid-data') {
      this._remoteInvalid = true;
      this._clearTimer();
      this._setStatus({ mode: this._mode(), state: 'error', reason: failure.reason, message: failure.message,
        canEdit: false, actions: ['retry', 'export', 'reset'] });
      return;
    }
    if (failure.reason === 'read-only') {
      this._writable = false;
      this._syncRequested = false;
      this._clearTimer();
      if (this._pending) {
        this._checkpoint(this._pending, this._metadata(false, false));
        if (!this._localFailed) {
          this._pending = undefined;
        }
      }
      this._setLocalStatus();
      return;
    }
    const transient = failure.reason === 'unavailable';
    this._writable = undefined;
    this._setStatus({ mode: 'local', state: this._localFailed ? 'error' : transient ? 'pending' : 'saved',
      reason: this._localFailed ? 'local-storage' : failure.reason,
      message: this._localFailed ? `${failure.message} Your changes have not been saved. Keep this page open and export a backup.`
        : `${failure.message} ${transient ? 'Settings are kept in this browser while synchronization is pending.' : LOCAL_MESSAGE}`,
      canEdit: true, actions: this._withLegacy(['retry', 'export']) });
    if (transient && (this._syncRequested || !this._pending) && this._retries < MAX_RETRIES) {
      const backoff = 1000 * Math.pow(2, this._retries++) + Math.floor(Math.random() * 250);
      this._retryAt = Date.now() + Math.max(backoff, failure.retryAfterMs ?? 0);
      this._schedule(this._retryAt - Date.now());
    } else {
      this._clearTimer();
      // Prevent the queue's completion handler from immediately retrying forever.
      this._writable = undefined;
      this._syncRequested = false;
      if (transient) {
        this._setStatus({ ...this._status, message: `${this._status.message} Automatic retries stopped; use Retry to try again.` });
      }
    }
  }

  public async resolve(action: LayoutStoreAction, starterLayout?: IDashboardLayout): Promise<IDashboardLayout | undefined> {
    const epoch = this._epoch;
    this._clearTimer();
    this._syncRequested = false;
    if (action === 'export') {
      return this._pending ?? this._record?.layout;
    }
    try {
      if (action === 'retry') {
        if (this._inFlight) {
          await this._inFlight;
        }
        if (!this._active(epoch)) {
          return undefined;
        }
        if (Date.now() < this._retryAt) {
          this._schedule(this._retryAt - Date.now());
          return this._pending ?? this._record?.layout;
        }
        this._retries = 0;
        this._retryAt = 0;
        this._client.dispose();
        this._client = this._newClient();
        this._writable = undefined;
        if (!this._pending) {
          const local = this._cache.read();
          this._record = local.record;
          this._cache.adopt(local.record);
          this._base = local.record?.server;
          this._baseKnown = local.record?.baseKnown ?? false;
          this._localConflicts = local.conflicts;
          this._localFailed = false;
          this._localInvalid = false;
          if (local.conflicts.length) {
            this._conflict('Competing browser drafts need to be resolved before retrying.');
            return local.record?.layout;
          }
        }
        return await this._connect(epoch, true);
      }
      if (action === 'import-legacy') {
        const legacy = this._cache.readLegacy();
        if (!legacy) {
          throw new Error('The legacy browser layout is no longer available.');
        }
        this._record = this._cache.resolve(legacy, { pendingSync: this._writable !== false, baseKnown: false });
        this._localFailed = false;
        this._localInvalid = false;
        this._legacyAvailable = false;
        this._base = undefined;
        this._baseKnown = false;
        this._pending = this._writable === false ? undefined : legacy;
        if (this._writable === false) {
          this._setLocalStatus();
          return legacy;
        }
        return await this._connect(epoch, false);
      }
      if (action === 'reset') {
        if (!starterLayout) {
          throw new Error('A starter layout is required to reset settings.');
        }
        if (this._client.recoveryRaw !== undefined) {
          this._cache.archive(this._client.recoveryRaw);
        }
        this._base = this._client.recoveryVersion ?? this._base;
        this._resetVersion = this._client.recoveryVersion;
        this._baseKnown = !!this._base || this._baseKnown;
        const resetLayout = decodeLayout(starterLayout);
        this._record = this._cache.reset(resetLayout, this._metadata(true));
        this._remoteInvalid = false;
        this._pending = this._writable === false ? undefined : resetLayout;
        this._localFailed = false;
        this._localInvalid = false;
        this._localConflicts = [];
        this._sequence++;
        if (this._writable === false) {
          this._setLocalStatus();
        } else if (this._writable) {
          this._queue();
        } else {
          await this._connect(epoch, false);
        }
        return resetLayout;
      }
      if (this._localConflicts.length) {
        const selected = action === 'keep-local' ? this._record
          : this._localConflicts.filter((record) => record.revision !== this._record?.revision)[0];
        if (!selected) {
          throw new Error('The selected browser draft is unavailable.');
        }
        this._record = this._cache.resolve(selected.layout, selected);
        this._localFailed = false;
        this._localInvalid = false;
        this._base = selected.server;
        this._baseKnown = selected.baseKnown;
        this._localConflicts = [];
        this._pending = this._writable === false ? undefined : selected.layout;
        if (this._writable === false) {
          this._setLocalStatus();
          return selected.layout;
        }
        return await this._connect(epoch, false);
      }
      const remote = await this._client.read();
      if (!this._active(epoch)) {
        return undefined;
      }
      const selected = action === 'keep-local' ? this._pending ?? this._record?.layout : remote?.layout ?? starterLayout;
      if (!selected) {
        throw new Error('The selected layout is unavailable.');
      }
      if (this._client.recoveryRaw !== undefined) {
        this._cache.archive(this._client.recoveryRaw);
      }
      this._base = remote;
      this._baseKnown = true;
      this._remote = remote;
      this._remoteCheckedAt = Date.now();
      const pending = action === 'keep-local' || !remote;
      this._record = this._cache.resolve(selected, { pendingSync: pending, baseKnown: true, server: remote });
      this._localFailed = false;
      this._localInvalid = false;
      this._pending = pending ? selected : undefined;
      this._sequence++;
      if (pending) {
        this._queue();
      } else {
        this._setCloudSaved();
      }
      return selected;
    } catch (error) {
      if (this._active(epoch)) {
        if (error instanceof LayoutStorageError) {
          this._remoteError(error);
        } else {
          this._localError(error);
        }
      }
      throw error;
    }
  }

  public exportRecovery(): string {
    let browser: unknown;
    let browserReadError: string | undefined;
    try {
      browser = JSON.parse(this._cache.exportRecovery());
    } catch (error) {
      browserReadError = 'Browser backups could not be read. This export contains only available in-memory and SharePoint recovery data.';
      this._localError(error);
      this._setStatus({ ...this._status, message: browserReadError });
    }
    return JSON.stringify({
      partial: browserReadError !== undefined,
      browserReadError,
      browser,
      sharepoint: this._client.recoveryRaw,
      current: this._record?.layout,
      unsaved: this._pending
    }, undefined, 2);
  }

  private _onExternalChange(): void {
    if (!this._disposed) {
      this._checkLocalConflict();
    }
  }

  private readonly _onOnline = (): void => {
    if (!this._disposed && this._syncRequested &&
        this._status.reason === 'unavailable' && this._writable !== false) {
      this._retries = 0;
      this._schedule(Math.max(this._retryAt - Date.now(), 0));
    }
  };

  private _active(epoch: number): boolean {
    return !this._disposed && epoch === this._epoch;
  }

  private _clearTimer(): void {
    if (this._timer !== undefined) {
      window.clearTimeout(this._timer);
      this._timer = undefined;
    }
  }

  public dispose(): void {
    this._disposed = true;
    this._epoch++;
    this._clearTimer();
    this._client.dispose();
    this._unsubscribeCache?.();
    this._unsubscribeCache = undefined;
    this._cache.dispose();
    this._listeners.clear();
    window.removeEventListener('online', this._onOnline);
  }
}
