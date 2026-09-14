import { IWidgetContext } from '../IWidget';
import { WidgetDataErrorDisposition } from './WidgetDataError';

interface ICacheEntry {
  data: unknown;
  expiresAt: number;
  storedAt: number;
  refreshError?: unknown;
}

export interface ICachedWidgetDataOptions<T> {
  key: string;
  ttlMilliseconds: number;
  load(): Promise<T>;
  bypassCache?: boolean;
  classifyError(error: unknown): WidgetDataErrorDisposition;
}

export interface ICachedWidgetData<T> {
  data: T;
  fetchedAt: number;
  refreshError?: unknown;
}

const MAX_CACHE_ENTRIES: number = 100;
const STALE_RETENTION_MS: number = 30 * 60 * 1000;
const cache = new Map<string, ICacheEntry>();
const inFlight = new Map<string, Promise<unknown>>();

function identityPart(value: unknown): string {
  if (value === undefined || value === null) {
    return '';
  }
  return String(value);
}

function normalize(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(normalize);
  }
  if (value && typeof value === 'object') {
    return Object.keys(value as Record<string, unknown>)
      .sort()
      .reduce<Record<string, unknown>>((result, key) => {
        const next = (value as Record<string, unknown>)[key];
        if (next !== undefined) {
          result[key] = normalize(next);
        }
        return result;
      }, {});
  }
  return value;
}

function pruneExpiredEntries(now: number): void {
  cache.forEach((entry, key) => {
    if (entry.expiresAt + STALE_RETENTION_MS <= now) {
      cache.delete(key);
    }
  });
}

function enforceCapacity(): void {
  while (cache.size > MAX_CACHE_ENTRIES) {
    let oldestKey: string | undefined;
    let oldestStoredAt = Number.MAX_SAFE_INTEGER;
    cache.forEach((entry, key) => {
      if (entry.storedAt < oldestStoredAt) {
        oldestKey = key;
        oldestStoredAt = entry.storedAt;
      }
    });
    if (!oldestKey) {
      break;
    }
    cache.delete(oldestKey);
  }
}

/**
 * Builds a cache key that cannot cross tenant, user, site, web, data source or
 * normalized request parameters.
 */
export function widgetDataCacheKey(
  context: IWidgetContext,
  source: string,
  parameters: unknown
): string {
  const pageContext = context.spContext.pageContext;
  const aadInfo = pageContext.aadInfo;

  return JSON.stringify({
    version: 1,
    tenant: identityPart(aadInfo?.tenantId ?? pageContext.site.id),
    user: identityPart(aadInfo?.userId ?? pageContext.user.loginName),
    site: identityPart(pageContext.site.id),
    web: identityPart(pageContext.web.id),
    source,
    parameters: normalize(parameters)
  });
}

/**
 * Reads through a process-local cache and shares an active request with every
 * consumer of the same key. Expired data remains available for transient-error
 * fallback, but is never returned as a normal cache hit.
 */
export function getCachedWidgetData<T>(options: ICachedWidgetDataOptions<T>): Promise<ICachedWidgetData<T>> {
  const now = Date.now();
  pruneExpiredEntries(now);

  const active = inFlight.get(options.key);
  if (active) {
    return active as Promise<ICachedWidgetData<T>>;
  }

  const existing = cache.get(options.key);
  if (!options.bypassCache && existing && existing.expiresAt > now) {
    return Promise.resolve({
      data: existing.data as T,
      fetchedAt: existing.storedAt,
      refreshError: existing.refreshError
    });
  }

  const loaded = options
    .load()
    .then((data) => {
      const storedAt = Date.now();
      cache.set(options.key, {
        data,
        storedAt,
        expiresAt: storedAt + options.ttlMilliseconds
      });
      enforceCapacity();
      return { data, fetchedAt: storedAt };
    })
    .catch((error: unknown) => {
      const failedAt = Date.now();
      if (
        existing &&
        existing.expiresAt + STALE_RETENTION_MS > failedAt &&
        options.classifyError(error) === 'stale'
      ) {
        existing.refreshError = error;
        return {
          data: existing.data as T,
          fetchedAt: existing.storedAt,
          refreshError: error
        };
      }
      cache.delete(options.key);
      throw error;
    });
  const request: Promise<ICachedWidgetData<T>> = loaded.then(
    (data) => {
      if (inFlight.get(options.key) === request) {
        inFlight.delete(options.key);
      }
      return data;
    },
    (error: unknown) => {
      if (inFlight.get(options.key) === request) {
        inFlight.delete(options.key);
      }
      throw error;
    }
  );

  inFlight.set(options.key, request);
  return request;
}

/** Test-only reset; production entries expire and are bounded by MAX_CACHE_ENTRIES. */
export function clearWidgetDataCache(): void {
  cache.clear();
  inFlight.clear();
}

/** Exposes bounded cache size for focused diagnostics and regression tests. */
export function widgetDataCacheEntryCount(): number {
  return cache.size;
}
