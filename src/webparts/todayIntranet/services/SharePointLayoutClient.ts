import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { IDashboardLayout } from '../model/IDashboardLayout';
import { parseLayout } from '../model/layoutDecoder';
import { IServerLayoutVersion } from './ILayoutCache';
import { LayoutStoreReason } from './ILayoutStore';

export const LAYOUT_LIST_TITLE: string = 'TodayIntranetLayouts';
const FIELD: string = 'LayoutJson';
const HEADERS: Record<string, string> = {
  Accept: 'application/json;odata=minimalmetadata',
  'Content-type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export class LayoutStorageError extends Error {
  public constructor(
    message: string,
    public readonly reason: LayoutStoreReason,
    public readonly status?: number,
    public readonly retryAfterMs?: number
  ) {
    super(message);
    this.name = 'LayoutStorageError';
    Object.setPrototypeOf(this, LayoutStorageError.prototype);
  }
}

export interface IRemoteLayout extends IServerLayoutVersion {
  layout: IDashboardLayout;
}

export interface ISharePointLayoutClientOptions {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  itemKey: string;
}

function object(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new LayoutStorageError('SharePoint returned an unexpected response.', 'configuration');
  }
  return value as Record<string, unknown>;
}

function permissions(value: unknown): SPPermission {
  const mask = object(value);
  const low = typeof mask.Low === 'string' && /^\d+$/.test(mask.Low) ? Number(mask.Low) : mask.Low;
  const high = typeof mask.High === 'string' && /^\d+$/.test(mask.High) ? Number(mask.High) : mask.High;
  if (typeof low !== 'number' || typeof high !== 'number' || !Number.isInteger(low) || !Number.isInteger(high)) {
    throw new LayoutStorageError('SharePoint did not return effective list permissions.', 'configuration');
  }
  return new SPPermission({ Low: low, High: high });
}

/** All requests stay in the hosting web and run with the signed-in user's permissions. */
export class SharePointLayoutClient {
  private _initializing: Promise<boolean> | undefined;
  private _disposed: boolean = false;
  public recoveryRaw: string | undefined;
  public recoveryVersion: IServerLayoutVersion | undefined;

  public constructor(private readonly _options: ISharePointLayoutClientOptions) {}

  public prime(writable: boolean): void {
    this._initializing = Promise.resolve(writable);
  }

  public dispose(): void {
    this._disposed = true;
  }

  public get listUrl(): string {
    return `${this._options.webAbsoluteUrl}/_api/web/lists/getByTitle('${LAYOUT_LIST_TITLE}')`;
  }

  public initialize(): Promise<boolean> {
    if (!this._initializing) {
      this._initializing = this._initialize();
    }
    return this._initializing;
  }

  private async _initialize(): Promise<boolean> {
    if (this._options.itemKey.length > 255) {
      throw new LayoutStorageError('The dashboard/user key exceeds the list field limit. An owner must configure a shorter dashboard ID.', 'configuration');
    }
    const list = await this._getOrCreateList();
    const access = permissions(list.EffectiveBasePermissions);
    const owner = access.hasPermission(SPPermission.manageLists);
    const writable = access.hasAllPermissions(SPPermission.addListItems, SPPermission.editListItems);
    let fields = await this._fields();
    let configurationChanged = false;
    let layoutField = fields.filter((field) => field.InternalName === FIELD)[0];
    const titleField = fields.filter((field) => field.InternalName === 'Title')[0];
    if (!titleField || titleField.TypeAsString !== 'Text') {
      throw new LayoutStorageError('The layouts list needs a compatible Title text column.', 'configuration');
    }
    if (!layoutField && owner) {
      try {
        await this._request(`${this.listUrl}/fields/createfieldasxml`, {
          parameters: {
            SchemaXml: `<Field Type='Note' DisplayName='${FIELD}' Name='${FIELD}' StaticName='${FIELD}' RichText='FALSE' />`,
            Options: 9
          }
        });
        configurationChanged = true;
      } catch (error) {
        fields = await this._fields();
        layoutField = fields.filter((field) => field.InternalName === FIELD)[0];
        if (!layoutField) {
          throw error;
        }
      }
    } else if (layoutField && !this._validLayoutField(layoutField)) {
      throw new LayoutStorageError('The LayoutJson column is incompatible. An owner must preserve its data and repair the schema.', 'configuration');
    }
    const expected: Record<string, boolean | number> = { Hidden: true, OnQuickLaunch: false, ReadSecurity: 2, WriteSecurity: 2, EnableAttachments: false };
    if (Object.keys(expected).some((key) => list[key] !== expected[key]) && owner) {
      await this._request(this.listUrl, expected, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
      configurationChanged = true;
    }
    if (titleField.Indexed !== true && owner) {
      await this._request(`${this.listUrl}/fields/getbyinternalnameortitle('Title')`, { Indexed: true },
        { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
      configurationChanged = true;
    }
    // Healthy lists are already described by the responses above. Only repeat
    // metadata reads after this client actually changed the configuration.
    const verified = configurationChanged ? await this._probe() : list;
    const verifiedFields = configurationChanged ? await this._fields() : fields;
    layoutField = verifiedFields.filter((field) => field.InternalName === FIELD)[0];
    const verifiedTitle = verifiedFields.filter((field) => field.InternalName === 'Title')[0];
    if (!verified || Object.keys(expected).some((key) => verified[key] !== expected[key]) ||
        !layoutField || !this._validLayoutField(layoutField) || verifiedTitle?.Indexed !== true) {
      throw new LayoutStorageError('The layouts list configuration is incomplete. A site owner must repair its field, index, and own-item access settings.', 'configuration');
    }
    return writable;
  }

  private async _getOrCreateList(): Promise<Record<string, unknown>> {
    const existing = await this._probe();
    if (existing) {
      return existing;
    }
    const web = object(await (await this._request(`${this._options.webAbsoluteUrl}/_api/web?$select=EffectiveBasePermissions`)).json());
    if (!permissions(web.EffectiveBasePermissions).hasPermission(SPPermission.manageLists)) {
      throw new LayoutStorageError('The layouts list is missing. A site owner must open this dashboard to provision it.', 'configuration');
    }
    try {
      await this._request(`${this._options.webAbsoluteUrl}/_api/web/lists`, {
        Title: LAYOUT_LIST_TITLE, BaseTemplate: 100,
        Description: 'Personal Today dashboard layouts.',
        Hidden: true, OnQuickLaunch: false, ReadSecurity: 2, WriteSecurity: 2, EnableAttachments: false
      });
    } catch (error) {
      // A competing owner may have created the list, including after a lost response.
      const competing = await this._probe();
      if (competing) {
        return competing;
      }
      throw error;
    }
    const created = await this._probe();
    if (!created) {
      throw new LayoutStorageError('The layouts list could not be verified after creation.', 'configuration');
    }
    return created;
  }

  private _validLayoutField(field: Record<string, unknown>): boolean {
    return field.TypeAsString === 'Note' && typeof field.SchemaXml === 'string' &&
      !/\bRichText\s*=\s*["']TRUE["']/i.test(field.SchemaXml);
  }

  private async _probe(): Promise<Record<string, unknown> | undefined> {
    try {
      const response = await this._request(`${this.listUrl}?$select=Id,ItemCount,Hidden,OnQuickLaunch,ReadSecurity,WriteSecurity,EnableAttachments,EffectiveBasePermissions`);
      return object(await response.json());
    } catch (error) {
      if (error instanceof LayoutStorageError && error.status === 404) {
        return undefined;
      }
      throw error;
    }
  }

  private async _fields(): Promise<Record<string, unknown>[]> {
    const response = await this._request(`${this.listUrl}/fields?$select=InternalName,TypeAsString,Indexed,SchemaXml`);
    const data = object(await response.json());
    if (!Array.isArray(data.value)) {
      throw new LayoutStorageError('The layouts list fields could not be verified.', 'configuration');
    }
    return data.value.map(object);
  }

  public async read(): Promise<IRemoteLayout | undefined> {
    const filter = encodeURIComponent(`Title eq '${this._options.itemKey.replace(/'/g, "''")}'`);
    const response = await this._request(`${this.listUrl}/items?$select=Id,${FIELD}&$filter=${filter}&$top=2`);
    const data = object(await response.json());
    if (!Array.isArray(data.value)) {
      throw new LayoutStorageError('SharePoint returned an invalid list response.', 'configuration');
    }
    if (data.value.length > 1) {
      throw new LayoutStorageError('Duplicate layouts exist for this dashboard. An owner must back up and reconcile them before saving.', 'configuration');
    }
    if (data.value.length === 0) {
      return undefined;
    }
    const item = object(data.value[0]);
    const id = item.Id;
    if (typeof id !== 'number' || !Number.isInteger(id) || id < 1) {
      throw new LayoutStorageError('SharePoint returned an invalid layout item ID.', 'configuration');
    }
    const etag = item['odata.etag'] ?? item['@odata.etag'];
    // Minimal-metadata collection responses normally carry each item's ETag.
    // Fall back to a direct read for tenants that omit it.
    if (typeof etag !== 'string' || !etag || etag === '*') {
      return this._readById(id);
    }
    return this._toRemote(item, id, etag);
  }

  private async _readById(id: number): Promise<IRemoteLayout> {
    const response = await this._request(`${this.listUrl}/items(${id})?$select=Id,${FIELD}`);
    const item = object(await response.json());
    const etag = response.headers.get('ETag') ?? item['odata.etag'] ?? item['@odata.etag'];
    if (typeof etag !== 'string' || !etag || etag === '*') {
      throw new LayoutStorageError('SharePoint did not return the layout version required for safe updates.', 'configuration');
    }
    return this._toRemote(item, id, etag);
  }

  private _toRemote(item: Record<string, unknown>, id: number, etag: string): IRemoteLayout {
    this.recoveryVersion = { itemId: id, etag };
    this.recoveryRaw = typeof item[FIELD] === 'string' ? item[FIELD] : JSON.stringify(item);
    if (typeof item[FIELD] !== 'string') {
      throw new LayoutStorageError('The saved SharePoint layout is empty or invalid. Export it before resetting.', 'invalid-data');
    }
    try {
      return { itemId: id, etag, layout: parseLayout(item[FIELD]) };
    } catch {
      throw new LayoutStorageError('The saved SharePoint layout is invalid or uses an unsupported version. Export it before resetting.', 'invalid-data');
    }
  }

  public async write(layout: IDashboardLayout, version?: IServerLayoutVersion): Promise<IRemoteLayout> {
    const body = { Title: this._options.itemKey, [FIELD]: JSON.stringify(layout) };
    let response: SPHttpClientResponse;
    try {
      response = await this._request(
        version ? `${this.listUrl}/items(${version.itemId})` : `${this.listUrl}/items`,
        body, version ? { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': version.etag } : undefined
      );
    } catch (error) {
      if (!version && error instanceof LayoutStorageError && error.reason === 'unavailable') {
        // The server may have accepted a create whose response was lost.
        const found = await this.read();
        if (found && JSON.stringify(found.layout) === JSON.stringify(layout)) {
          return found;
        }
        if (found) {
          throw new LayoutStorageError('Another client created this dashboard layout first.', 'unavailable', 412);
        }
      }
      throw error;
    }
    let itemId = version?.itemId;
    if (itemId === undefined) {
      const created = object(await response.json());
      if (typeof created.Id !== 'number' || !Number.isInteger(created.Id) || created.Id < 1) {
        // Some SharePoint response modes omit the created entity.
        const found = await this.read();
        if (!found) {
          throw new LayoutStorageError('SharePoint did not return the created layout item.', 'unavailable');
        }
        itemId = found.itemId;
      } else {
        itemId = created.Id;
      }
      // Own-item visibility cannot be combined with a unique Title field.
      // Requery by logical key after creation so a cross-device create race is
      // detected instead of silently selecting either new item.
      const verifiedCreated = await this.read();
      if (!verifiedCreated) {
        throw new LayoutStorageError('The created SharePoint layout could not be found.', 'unavailable');
      }
      if (JSON.stringify(verifiedCreated.layout) !== JSON.stringify(layout)) {
        throw new LayoutStorageError('Another client created this dashboard layout first.', 'unavailable', 412);
      }
      return verifiedCreated;
    }
    // MERGE usually has no response ETag. Verify the exact saved entity directly,
    // without repeating its logical-key lookup.
    const stored = await this._readById(itemId);
    if (!stored || JSON.stringify(stored.layout) !== JSON.stringify(layout)) {
      throw new LayoutStorageError('The SharePoint layout changed while saving.', 'unavailable', 412);
    }
    return stored;
  }

  private async _request(url: string, body?: unknown, extraHeaders?: Record<string, string>): Promise<SPHttpClientResponse> {
    if (this._disposed) {
      throw new LayoutStorageError('The dashboard scope is no longer active.', 'unavailable');
    }
    let response: SPHttpClientResponse;
    try {
      response = body === undefined
        ? await this._options.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: HEADERS })
        : await this._options.spHttpClient.post(url, SPHttpClient.configurations.v1,
          { headers: { ...HEADERS, ...extraHeaders }, body: JSON.stringify(body) });
    } catch {
      throw new LayoutStorageError('SharePoint could not be reached.', 'unavailable');
    }
    if (this._disposed) {
      throw new LayoutStorageError('The dashboard scope is no longer active.', 'unavailable');
    }
    if (!response.ok) {
      const retryAfter = response.headers.get('Retry-After');
      const retryAfterMs = retryAfter
        ? (/^\d+$/.test(retryAfter) ? Number(retryAfter) * 1000 : Math.max(0, Date.parse(retryAfter) - Date.now()))
        : undefined;
      const status = response.status;
      const reason: LayoutStoreReason = status === 401 || status === 403 ? 'read-only'
        : status === 429 || status === 503 || status === 502 || status === 504 || status === 412 || status === 404
          ? 'unavailable' : 'configuration';
      throw new LayoutStorageError(`SharePoint request failed (HTTP ${status}).`, reason, status,
        retryAfterMs !== undefined && Number.isFinite(retryAfterMs) ? retryAfterMs : undefined);
    }
    return response;
  }
}
