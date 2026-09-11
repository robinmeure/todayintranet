import { IDashboardLayout } from '../model/IDashboardLayout';

export interface IServerLayoutVersion {
  itemId: number;
  etag: string;
}

export interface ILayoutSyncMetadata {
  pendingSync: boolean;
  /** True also represents a confirmed missing server item when server is absent. */
  baseKnown: boolean;
  server?: IServerLayoutVersion;
  /** Last verified capability/schema state for this scoped user and SharePoint web. */
  cloud?: {
    writable: boolean;
    configurationCheckedAt: number;
  };
  /** Time the corresponding server layout/version was last observed. */
  remoteCheckedAt?: number;
}

export interface ILayoutRecord extends ILayoutSyncMetadata {
  schema: 1;
  scope: string;
  revision: string;
  clock: Record<string, number>;
  layout: IDashboardLayout;
}

export interface ILayoutCacheRead {
  record?: ILayoutRecord;
  /** Incomparable revisions; empty when there is only one head. */
  conflicts: ILayoutRecord[];
  legacyAvailable: boolean;
}
