import { IDashboardLayout } from '../model/IDashboardLayout';

export type LayoutStoreMode = 'sharepoint' | 'local';
export type LayoutSaveState = 'loading' | 'saved' | 'saving' | 'pending' | 'conflict' | 'error';
export type LayoutStoreReason = 'read-only' | 'unavailable' | 'configuration' | 'invalid-data' | 'local-storage';
export type LayoutStoreAction = 'retry' | 'use-remote' | 'keep-local' | 'import-legacy' | 'reset' | 'export';

export interface ILayoutStoreStatus {
  mode: LayoutStoreMode;
  state: LayoutSaveState;
  reason?: LayoutStoreReason;
  message: string;
  canEdit: boolean;
  actions: LayoutStoreAction[];
}

export interface ILayoutStore {
  load(): Promise<IDashboardLayout | undefined>;
  /** Checkpoint locally immediately without publishing to SharePoint. */
  save(layout: IDashboardLayout): Promise<void>;
  /** Publish the latest local checkpoint. Called when the user finishes editing. */
  publish(): Promise<void>;
  getStatus(): ILayoutStoreStatus;
  subscribe(listener: () => void): () => void;
  resolve(action: LayoutStoreAction, starterLayout?: IDashboardLayout): Promise<IDashboardLayout | undefined>;
  exportRecovery(): string;
  dispose(): void;
}

export interface ILayoutScope {
  siteId: string;
  webId: string;
  dashboardKey: string;
  userKey: string;
}

export function layoutScopeKey(scope: ILayoutScope): string {
  return JSON.stringify([scope.siteId.toLowerCase(), scope.webId.toLowerCase(), scope.dashboardKey, scope.userKey]);
}
