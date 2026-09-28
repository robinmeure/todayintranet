import { IRemoteLayout, LayoutStorageError, SharePointDocumentClient } from './SharePointLayoutClient';
import {
  emptyTaskDocument, mergeTaskDocuments, parseTaskDocument, TaskDocument, TaskOrderConflict
} from '../widgets/graph/taskOrganization';

export interface TaskOrganizationState {
  document: TaskDocument;
  status: 'loading' | 'saved' | 'pending' | 'local' | 'conflict' | 'error';
  message: string;
  canEdit: boolean;
}

interface Checkpoint { version: 1; base: TaskDocument; document: TaskDocument }
interface PendingCheckpoint { key: string; raw: string; value: Checkpoint }
export type TaskDocumentClient = Pick<SharePointDocumentClient<TaskDocument>,
  'initialize' | 'read' | 'write' | 'update' | 'dispose' | 'recoveryRaw'>;

function message(error: unknown): string {
  return error instanceof Error ? error.message : 'Task organization could not be saved.';
}
function same(a: TaskDocument, b: TaskDocument): boolean { return JSON.stringify(a) === JSON.stringify(b); }

/** Separate per-writer outboxes prevent two tabs from overwriting each other's unsent edits. */
export class TaskOrganizationStore {
  private _state: TaskOrganizationState = {
    document: emptyTaskDocument(), status: 'loading', message: 'Loading personal organization...', canEdit: false
  };
  private _listeners = new Set<() => void>();
  private _client: TaskDocumentClient;
  private _base: TaskDocument = emptyTaskDocument();
  private _remote: IRemoteLayout<TaskDocument> | undefined;
  private _remoteKnown = false;
  private _loaded: Promise<void> | undefined;
  private _busy = false;
  private _writable = false;
  private _connected = false;
  private _disposed = false;
  private _timer: ReturnType<typeof setTimeout> | undefined;
  private _seen: PendingCheckpoint[] = [];
  private _conflictRemote: TaskDocument | undefined;
  private _badCheckpoint = false;
  private _lastClock = 0;
  private _conflictRetries = 0;
  private _retryAt = 0;
  private _recovery: string[] = [];
  private readonly _key: string;

  public constructor(
    private readonly _factory: () => TaskDocumentClient,
    private readonly _storage: Storage,
    public readonly prefix: string,
    private readonly _writer: string
  ) {
    this._key = `${prefix}${_writer}`;
    this._client = _factory();
  }

  public getState = (): TaskOrganizationState => this._state;
  public subscribe = (listener: () => void): (() => void) => {
    this._listeners.add(listener);
    return () => { this._listeners.delete(listener); };
  };
  public run(operation: () => Promise<void>): void {
    operation().catch((error: unknown) => {
      this._set({ status: 'error', message: message(error), canEdit: false });
    });
  }
  private _set(values: Partial<TaskOrganizationState>): void {
    if (this._disposed) { return; }
    this._state = { ...this._state, ...values };
    this._listeners.forEach((listener) => listener());
  }

  public revision(): string {
    const tasks = this._state.document.tasks;
    const largest = Object.keys(tasks).reduce((max, key) => Math.max(max, Number(tasks[key].modified.split(':')[0])), 0);
    this._lastClock = Math.max(Date.now(), this._lastClock + 1, largest + 1);
    return `${this._lastClock}:${this._writer}`;
  }

  private _checkpoints(): PendingCheckpoint[] {
    const result: PendingCheckpoint[] = [];
    this._recovery = [];
    for (let index = 0; index < this._storage.length; index++) {
      const key = this._storage.key(index);
      if (!key?.startsWith(this.prefix)) { continue; }
      const raw = this._storage.getItem(key);
      if (!raw) { continue; }
      this._recovery.push(raw);
      const value: unknown = JSON.parse(raw);
      if (!value || typeof value !== 'object' || Array.isArray(value)) { throw new Error('Invalid browser organization checkpoint.'); }
      const record = value as Record<string, unknown>;
      if (record.version !== 1) { throw new Error('Unsupported browser organization checkpoint.'); }
      result.push({ key, raw, value: { version: 1, base: parseTaskDocument(JSON.stringify(record.base)),
        document: parseTaskDocument(JSON.stringify(record.document)) } });
    }
    return result.sort((a, b) => a.key.localeCompare(b.key));
  }

  public load(): Promise<void> {
    if (!this._loaded) { this._loaded = this._load(); }
    return this._loaded;
  }

  private async _load(): Promise<void> {
    try {
      this._seen = this._checkpoints();
      let document = this._seen[0]?.value.document ?? emptyTaskDocument();
      this._base = this._seen[0]?.value.base ?? emptyTaskDocument();
      this._set({ document });
      for (const pending of this._seen) {
        if (pending === this._seen[0]) { continue; }
        document = mergeTaskDocuments(pending.value.base, pending.value.document, document);
      }
      this._set({ document });
    } catch (error) {
      if (!(error instanceof TaskOrderConflict)) {
        this._badCheckpoint = true;
        this._set({ status: 'error', message: `${message(error)} Export recovery data before clearing browser checkpoints.`, canEdit: false });
        return;
      }
    }
    await this.sync();
  }

  private _checkpoint(document: TaskDocument): void {
    const value: Checkpoint = { version: 1, base: this._base, document };
    this._storage.setItem(this._key, JSON.stringify(value));
  }

  public edit(document: TaskDocument): void {
    if (!this._state.canEdit) { throw new Error('Resolve the organization loading or recovery state before editing.'); }
    const validated = parseTaskDocument(JSON.stringify(document));
    if (this._state.status !== 'error' && same(validated, this._state.document)) { return; }
    this._conflictRetries = 0;
    try {
      this._checkpoint(validated);
    } catch (error) {
      this._set({ document: validated, status: 'error', message: `Not saved in this browser: ${message(error)} Export your edits or retry.` });
      throw error;
    }
    this._set({ document: validated, status: this._connected && this._writable ? 'pending' : 'local',
      message: this._connected && this._writable ? 'Pending SharePoint save...' : 'Saved in this browser only. Organization will not follow you to other devices.' });
    if (this._connected && this._writable) { this._schedule(); }
  }

  private _schedule(): void {
    if (this._timer) { clearTimeout(this._timer); }
    this._timer = setTimeout(() => { this._timer = undefined; this.run(() => this.sync(false)); }, 600);
  }

  public async sync(refreshRemote: boolean = true): Promise<void> {
    if (this._busy || this._disposed || this._badCheckpoint || this._state.status === 'conflict') { return; }
    if (this._timer) { clearTimeout(this._timer); this._timer = undefined; }
    if (Date.now() < this._retryAt) {
      this._set({ message: `SharePoint is busy. Retry after ${new Date(this._retryAt).toLocaleTimeString()}. Your edits are retained.` });
      return;
    }
    this._busy = true;
    let again = false;
    try {
      if (!this._connected) {
        this._remoteKnown = false;
        this._client.dispose();
        this._client = this._factory();
        this._writable = await this._client.initialize();
      }
      const remote = refreshRemote || !this._remoteKnown ? await this._client.read() : this._remote;
      if (this._disposed) { return; }
      this._remote = remote;
      this._remoteKnown = true;
      this._connected = true;
      const remoteDocument = remote?.layout ?? emptyTaskDocument();
      this._conflictRemote = remoteDocument;
      this._seen = this._checkpoints();
      let merged = mergeTaskDocuments(this._base, this._state.document, remoteDocument);
      for (const checkpoint of this._seen) {
        // Our current state includes edits made during the remote read.
        if (checkpoint.key !== this._key) {
          merged = mergeTaskDocuments(checkpoint.value.base, checkpoint.value.document, merged);
        }
      }
      this._base = remoteDocument;
      this._set({ document: merged, canEdit: true });
      if (same(merged, remoteDocument)) {
        this._clearSeen();
        this._set({ status: 'saved', message: 'Organization is up to date in SharePoint.' });
        return;
      }
      this._checkpoint(merged);
      if (!this._writable) {
        this._set({ status: 'local', message: 'Saved in this browser only. You do not have permission to save organization to SharePoint.' });
        return;
      }
      this._set({ status: 'pending', message: 'Pending SharePoint save...' });
      const snapshot = merged;
      const checkpointRaw = this._storage.getItem(this._key);
      try {
        if (remote) {
          const version = await this._client.update(snapshot, remote);
          this._remote = version ? { ...version, layout: snapshot } : undefined;
          this._remoteKnown = version !== undefined;
        } else {
          this._remote = await this._client.write(snapshot);
          this._remoteKnown = true;
        }
      } catch (error) {
        this._remoteKnown = false;
        if (error instanceof LayoutStorageError && error.status === 412 && this._conflictRetries++ < 2) {
          this._set({ status: 'pending', message: 'Another device saved changes. Reconciling organization...' });
          again = true;
          return;
        }
        throw error;
      }
      if (this._disposed) { return; }
      this._conflictRetries = 0;
      this._base = snapshot;
      if (same(this._state.document, snapshot)) {
        this._clearSeen();
        if (this._storage.getItem(this._key) === checkpointRaw) { this._storage.removeItem(this._key); }
        this._set({ status: 'saved', message: 'Saved to SharePoint.' });
      } else {
        this._checkpoint(this._state.document);
        again = true;
      }
    } catch (error) {
      this._remoteKnown = false;
      if (error instanceof LayoutStorageError && (error.status === 429 || error.status === 503)) {
        this._retryAt = Date.now() + (error.retryAfterMs ?? 30000);
      }
      if (error instanceof TaskOrderConflict) {
        this._set({ status: 'conflict', message: error.message, canEdit: false });
      } else if (error instanceof LayoutStorageError &&
          (error.reason === 'read-only' ||
            (error.reason === 'configuration' && !this._connected && !this._client.recoveryRaw))) {
        this._connected = false;
        this._set({ status: 'local', message: `${message(error)} Organization is browser-only until access is restored.`, canEdit: true });
      } else {
        this._set({ status: 'error', message: `${message(error)} Saved organization could not be synchronized; existing edits are retained.`,
          canEdit: this._connected && !(error instanceof LayoutStorageError && error.reason === 'invalid-data') });
      }
    } finally {
      this._busy = false;
      if (again && !this._disposed) { this._schedule(); }
    }
  }

  private _clearSeen(): void {
    for (const checkpoint of this._seen) {
      if (this._storage.getItem(checkpoint.key) === checkpoint.raw) { this._storage.removeItem(checkpoint.key); }
    }
    this._seen = [];
  }

  public async resolveConflict(choice: 'local' | 'remote'): Promise<void> {
    if (this._state.status !== 'conflict' || !this._conflictRemote) { return; }
    const document = choice === 'remote' ? this._conflictRemote : this._state.document;
    this._base = this._conflictRemote;
    try {
      this._checkpoint(document);
      this._clearSeen();
      this._checkpoint(document);
      this._set({ document, status: 'pending', canEdit: true, message: 'Saving your conflict choice...' });
      await this.sync();
    } catch (error) {
      this._set({ status: 'error', message: `Could not checkpoint the conflict choice: ${message(error)}`, canEdit: false });
    }
  }

  public async clearBrowserRecovery(): Promise<void> {
    try {
      const keys: string[] = [];
      for (let index = 0; index < this._storage.length; index++) {
        const key = this._storage.key(index);
        if (key?.startsWith(this.prefix)) { keys.push(key); }
      }
      keys.forEach((key) => this._storage.removeItem(key));
      this._badCheckpoint = false;
      this._base = emptyTaskDocument();
      this._set({ document: emptyTaskDocument(), status: 'loading', canEdit: false, message: 'Reloading SharePoint organization...' });
      await this.sync();
    } catch (error) {
      this._set({ status: 'error', message: `Could not clear browser checkpoints: ${message(error)}`, canEdit: false });
    }
  }

  public exportRecovery(): string {
    return JSON.stringify({ document: this._state.document, browserCheckpoints: this._recovery,
      sharePointRecovery: this._client.recoveryRaw }, undefined, 2);
  }

  public dispose(): void {
    this._disposed = true;
    if (this._timer) { clearTimeout(this._timer); }
    this._client.dispose();
    this._listeners.clear();
  }
}
