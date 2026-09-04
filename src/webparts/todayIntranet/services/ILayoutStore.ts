import { IDashboardLayout } from '../model/IDashboardLayout';

export type LayoutStoreMode = 'sharepoint' | 'local';

export interface ILayoutStoreStatus {
  /** Where the layout is actually being persisted right now. */
  mode: LayoutStoreMode;
  /** Set when the SharePoint list could not be used, so the UI can explain itself. */
  message?: string;
}

/**
 * Persists one user's dashboard layout. Implementations must be safe to call
 * repeatedly; `save` is invoked whenever the user stops dragging or resizing.
 */
export interface ILayoutStore {
  load(): Promise<IDashboardLayout | undefined>;
  save(layout: IDashboardLayout): Promise<void>;
  getStatus(): ILayoutStoreStatus;
}
