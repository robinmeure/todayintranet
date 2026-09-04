import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IDashboardLayout } from '../model/IDashboardLayout';
import { ILayoutStore, ILayoutStoreStatus } from './ILayoutStore';
import { LocalStorageLayoutStore } from './LocalStorageLayoutStore';

const LIST_TITLE: string = 'TodayIntranetLayouts';
const LAYOUT_FIELD: string = 'LayoutJson';

const JSON_HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=nometadata',
  'Content-type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export interface ISharePointLayoutStoreOptions {
  spHttpClient: SPHttpClient;
  /** Absolute url of the web hosting the dashboard. */
  webAbsoluteUrl: string;
  /** Login name of the current user; scopes the stored layout. */
  userKey: string;
  /** Distinguishes multiple dashboards on the same site. */
  dashboardKey: string;
}

/**
 * Stores one layout item per user in a hidden list on the hosting site.
 *
 * The list is provisioned on first use, which requires Manage Lists rights, so
 * in practice a site owner creates it the first time they open the page. Any
 * failure (missing list, read-only user, throttling) degrades to localStorage
 * rather than losing the user's arrangement.
 */
export class SharePointListLayoutStore implements ILayoutStore {
  private readonly _options: ISharePointLayoutStoreOptions;
  private readonly _fallback: LocalStorageLayoutStore;
  private readonly _itemKey: string;

  private _itemId: number | undefined;
  private _listReady: boolean = false;
  private _degradedMessage: string | undefined;

  constructor(options: ISharePointLayoutStoreOptions) {
    this._options = options;
    this._itemKey = `${options.dashboardKey}|${options.userKey}`.substring(0, 255);
    this._fallback = new LocalStorageLayoutStore(this._itemKey);
  }

  public getStatus(): ILayoutStoreStatus {
    return this._degradedMessage
      ? { mode: 'local', message: this._degradedMessage }
      : { mode: 'sharepoint' };
  }

  public async load(): Promise<IDashboardLayout | undefined> {
    try {
      await this._ensureList();

      const filter = encodeURIComponent(`Title eq '${this._itemKey.replace(/'/g, "''")}'`);
      const url =
        `${this._listUrl()}/items?$select=Id,${LAYOUT_FIELD}&$filter=${filter}&$top=1`;

      const response = await this._options.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: JSON_HEADERS
      });
      if (!response.ok) {
        throw new Error(`Reading the saved layout failed (HTTP ${response.status}).`);
      }

      const payload = (await response.json()) as { value: { Id: number; LayoutJson: string }[] };
      const item = payload.value && payload.value[0];
      if (!item) {
        return undefined;
      }

      this._itemId = item.Id;
      return item.LayoutJson ? (JSON.parse(item.LayoutJson) as IDashboardLayout) : undefined;
    } catch (error) {
      this._degrade(error);
      return this._fallback.load();
    }
  }

  public async save(layout: IDashboardLayout): Promise<void> {
    // Always keep the local copy in sync so a later outage still has something to show.
    await this._fallback.save(layout);

    if (this._degradedMessage) {
      return;
    }

    try {
      await this._ensureList();

      const body = { Title: this._itemKey, [LAYOUT_FIELD]: JSON.stringify(layout) };

      if (this._itemId === undefined) {
        const response = await this._post(`${this._listUrl()}/items`, body);
        if (!response.ok) {
          throw new Error(`Saving the layout failed (HTTP ${response.status}).`);
        }
        const created = (await response.json()) as { Id: number };
        this._itemId = created.Id;
        return;
      }

      const response = await this._post(`${this._listUrl()}/items(${this._itemId})`, body, {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      });
      if (response.status === 404) {
        // The item was removed behind our back - recreate it on the next save.
        this._itemId = undefined;
        throw new Error('The saved layout item no longer exists.');
      }
      if (!response.ok) {
        throw new Error(`Saving the layout failed (HTTP ${response.status}).`);
      }
    } catch (error) {
      this._degrade(error);
    }
  }

  private _listUrl(): string {
    return `${this._options.webAbsoluteUrl}/_api/web/lists/getByTitle('${LIST_TITLE}')`;
  }

  private async _post(
    url: string,
    body: unknown,
    extraHeaders?: Record<string, string>
  ): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: { ...JSON_HEADERS, ...extraHeaders },
      body: JSON.stringify(body)
    };
    return this._options.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }

  private async _ensureList(): Promise<void> {
    if (this._listReady) {
      return;
    }

    const probe = await this._options.spHttpClient.get(
      `${this._listUrl()}?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: JSON_HEADERS }
    );

    if (probe.ok) {
      this._listReady = true;
      return;
    }
    if (probe.status !== 404) {
      throw new Error(`Could not open the '${LIST_TITLE}' list (HTTP ${probe.status}).`);
    }

    await this._provisionList();
    this._listReady = true;
  }

  private async _provisionList(): Promise<void> {
    const createList = await this._post(`${this._options.webAbsoluteUrl}/_api/web/lists`, {
      Title: LIST_TITLE,
      BaseTemplate: 100,
      Description: 'Stores each user\'s personal Today dashboard layout. Managed by the Today intranet web part.'
    });
    if (!createList.ok) {
      throw new Error(
        `The '${LIST_TITLE}' list does not exist and could not be created (HTTP ${createList.status}). ` +
          'A site owner needs to open this page once to provision it.'
      );
    }

    const schemaXml =
      `<Field Type='Note' DisplayName='${LAYOUT_FIELD}' Name='${LAYOUT_FIELD}' ` +
      `StaticName='${LAYOUT_FIELD}' NumLines='6' RichText='FALSE' RestrictedMode='TRUE' />`;
    const addField = await this._post(`${this._listUrl()}/fields/createfieldasxml`, {
      parameters: { SchemaXml: schemaXml, Options: 9 }
    });
    if (!addField.ok) {
      throw new Error(`Could not add the '${LAYOUT_FIELD}' column (HTTP ${addField.status}).`);
    }

    // Hide the list and restrict item level access so users only ever see their own layout.
    await this._post(
      this._listUrl(),
      { Hidden: true, OnQuickLaunch: false, ReadSecurity: 2, WriteSecurity: 2, EnableAttachments: false },
      { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }
    );
  }

  private _degrade(error: unknown): void {
    if (this._degradedMessage) {
      return;
    }
    const detail = error instanceof Error ? error.message : String(error);
    this._degradedMessage = `${detail} Your layout is being kept in this browser only.`;
    console.warn('[TodayIntranet] Falling back to local storage:', error);
  }
}
