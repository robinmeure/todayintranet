import { IDashboardLayout } from '../model/IDashboardLayout';
import { ILayoutStore, ILayoutStoreStatus } from './ILayoutStore';

/**
 * Fallback store used when the SharePoint list is unavailable (for example when
 * the user only has read access to the site). Layouts saved here do not roam.
 */
export class LocalStorageLayoutStore implements ILayoutStore {
  private readonly _key: string;
  private _message: string | undefined;

  constructor(scopeKey: string, message?: string) {
    this._key = `todayIntranet.layout.${scopeKey}`;
    this._message = message;
  }

  public async load(): Promise<IDashboardLayout | undefined> {
    try {
      const raw = window.localStorage.getItem(this._key);
      return raw ? (JSON.parse(raw) as IDashboardLayout) : undefined;
    } catch {
      return undefined;
    }
  }

  public async save(layout: IDashboardLayout): Promise<void> {
    try {
      window.localStorage.setItem(this._key, JSON.stringify(layout));
    } catch {
      // Private browsing or a full quota - nothing sensible to do beyond dropping it.
    }
  }

  public getStatus(): ILayoutStoreStatus {
    return { mode: 'local', message: this._message };
  }
}
