import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { IWidgetContext } from '../IWidget';
import {
  clearWidgetDataCache,
  getCachedWidgetData,
  widgetDataCacheEntryCount,
  widgetDataCacheKey
} from './WidgetDataCache';
import { WidgetDataErrorDisposition } from './WidgetDataError';

const stale = (): WidgetDataErrorDisposition => 'stale';
const invalidate = (): WidgetDataErrorDisposition => 'invalidate';

function deferred<T>(): {
  promise: Promise<T>;
  resolve(value: T): void;
  reject(reason: unknown): void;
} {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((_resolve, _reject) => {
    resolve = _resolve;
    reject = _reject;
  });
  return { promise, resolve, reject };
}

function contextFor(userId: string): IWidgetContext {
  const spContext = new (jest.fn<WebPartContext, []>())();
  Object.defineProperty(spContext, 'pageContext', {
    value: {
      aadInfo: { tenantId: 'tenant-one', userId },
      site: { id: 'site-one' },
      web: { id: 'web-one' },
      user: { loginName: `${userId}@example.com` }
    }
  });
  return {
    instanceId: `widget-${userId}`,
    settings: {},
    spContext,
    isEditing: false,
    refreshToken: 0,
    updateSettings: jest.fn<void, [Record<string, unknown>]>()
  };
}

describe('WidgetDataCache', () => {
  let now: number;

  beforeEach(() => {
    now = 10_000;
    jest.spyOn(Date, 'now').mockImplementation(() => now);
    clearWidgetDataCache();
  });

  afterEach(() => {
    clearWidgetDataCache();
    jest.restoreAllMocks();
  });

  it('deduplicates concurrent requests for the same key', async () => {
    const pending = deferred<string[]>();
    const load = jest.fn<Promise<string[]>, []>(() => pending.promise);
    const options = { key: 'mail', ttlMilliseconds: 1_000, load, classifyError: invalidate };

    const first = getCachedWidgetData(options);
    const second = getCachedWidgetData(options);

    expect(load).toHaveBeenCalledTimes(1);
    pending.resolve(['message']);
    await expect(Promise.all([first, second])).resolves.toEqual([
      { data: ['message'], fetchedAt: now },
      { data: ['message'], fetchedAt: now }
    ]);
  });

  it('caches empty responses and reuses them within the TTL', async () => {
    const load = jest.fn<Promise<never[]>, []>().mockResolvedValue([]);

    await expect(getCachedWidgetData({
      key: 'empty',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    })).resolves.toEqual({ data: [], fetchedAt: now });
    now += 999;
    await expect(getCachedWidgetData({
      key: 'empty',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    })).resolves.toEqual({ data: [], fetchedAt: 10_000 });

    expect(load).toHaveBeenCalledTimes(1);
  });

  it('refetches expired entries and allows a manual bypass before expiry', async () => {
    const load = jest.fn<Promise<number>, []>()
      .mockResolvedValueOnce(1)
      .mockResolvedValueOnce(2)
      .mockResolvedValueOnce(3);

    await expect(getCachedWidgetData({
      key: 'tasks',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    })).resolves.toEqual({ data: 1, fetchedAt: now });
    now += 500;
    await expect(getCachedWidgetData({
      key: 'tasks',
      ttlMilliseconds: 1_000,
      load,
      bypassCache: true,
      classifyError: invalidate
    })).resolves.toEqual({ data: 2, fetchedAt: now });
    now += 1_001;
    await expect(getCachedWidgetData({
      key: 'tasks',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    })).resolves.toEqual({ data: 3, fetchedAt: now });

    expect(load).toHaveBeenCalledTimes(3);
  });

  it('keeps tenant, user, source and normalized parameters isolated', () => {
    const firstUser = contextFor('user-one');
    const secondUser = contextFor('user-two');
    const first = widgetDataCacheKey(firstUser, 'graph:mail', { unread: true, count: 6 });
    const reordered = widgetDataCacheKey(firstUser, 'graph:mail', { count: 6, unread: true });

    expect(reordered).toBe(first);
    expect(widgetDataCacheKey(secondUser, 'graph:mail', { count: 6, unread: true })).not.toBe(first);
    expect(widgetDataCacheKey(firstUser, 'graph:calendar', { count: 6, unread: true })).not.toBe(first);
    expect(widgetDataCacheKey(firstUser, 'graph:mail', { count: 7, unread: true })).not.toBe(first);
  });

  it('returns expired data after a transient refresh failure', async () => {
    const load = jest.fn<Promise<string>, []>()
      .mockResolvedValueOnce('stale data')
      .mockRejectedValueOnce(new Error('temporarily unavailable'));

    await getCachedWidgetData({
      key: 'calendar',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    });

    now += 1_001;

    await expect(getCachedWidgetData({
      key: 'calendar',
      ttlMilliseconds: 1_000,
      load,
      classifyError: stale
    })).resolves.toEqual({
      data: 'stale data',
      fetchedAt: 10_000,
      refreshError: expect.any(Error)
    });
  });

  it('checks stale eligibility when a failed request settles', async () => {
    const failure = new Error('service unavailable');
    const pending = deferred<string>();
    const initial = jest.fn<Promise<string>, []>().mockResolvedValue('stale data');

    await getCachedWidgetData({
      key: 'calendar',
      ttlMilliseconds: 1_000,
      load: initial,
      classifyError: invalidate
    });
    now += 1_001;
    const refresh = getCachedWidgetData({
      key: 'calendar',
      ttlMilliseconds: 1_000,
      load: () => pending.promise,
      classifyError: stale
    });
    now = 10_000 + 1_000 + 30 * 60 * 1_000;
    pending.reject(failure);

    await expect(refresh).rejects.toBe(failure);
  });

  it('retains refresh-failure metadata on a fresh entry until a successful fetch', async () => {
    const failure = new Error('service unavailable');
    const load = jest.fn<Promise<string>, []>()
      .mockResolvedValueOnce('cached data')
      .mockRejectedValueOnce(failure)
      .mockResolvedValueOnce('updated data');
    const base = {
      key: 'mail',
      ttlMilliseconds: 1_000,
      load,
      classifyError: stale
    };

    await getCachedWidgetData(base);
    await getCachedWidgetData({ ...base, bypassCache: true });
    await expect(getCachedWidgetData(base)).resolves.toEqual({
      data: 'cached data',
      fetchedAt: now,
      refreshError: failure
    });

    now += 1_001;
    await expect(getCachedWidgetData(base)).resolves.toEqual({
      data: 'updated data',
      fetchedAt: now
    });
  });

  it('surfaces non-transient refresh failures instead of hiding them', async () => {
    const failure = new Error('access denied');
    const load = jest.fn<Promise<string>, []>()
      .mockResolvedValueOnce('old data')
      .mockRejectedValueOnce(failure);

    await getCachedWidgetData({
      key: 'mail',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    });
    now += 1_001;

    await expect(getCachedWidgetData({
      key: 'mail',
      ttlMilliseconds: 1_000,
      load,
      classifyError: invalidate
    })).rejects.toBe(failure);

    const replacement = jest.fn<Promise<string>, []>().mockResolvedValue('new data');
    await expect(getCachedWidgetData({
      key: 'mail',
      ttlMilliseconds: 1_000,
      load: replacement,
      classifyError: invalidate
    })).resolves.toEqual({ data: 'new data', fetchedAt: now });
    expect(replacement).toHaveBeenCalledTimes(1);
  });

  it('joins an active forced refresh before considering a fresh cache hit', async () => {
    const pending = deferred<string>();
    const load = jest.fn<Promise<string>, []>()
      .mockResolvedValueOnce('old data')
      .mockImplementationOnce(() => pending.promise);
    const base = { key: 'mail', ttlMilliseconds: 1_000, load, classifyError: invalidate };

    await getCachedWidgetData(base);
    const refresh = getCachedWidgetData({ ...base, bypassCache: true });
    const consumer = getCachedWidgetData(base);
    pending.resolve('new data');

    await expect(Promise.all([refresh, consumer])).resolves.toEqual([
      { data: 'new data', fetchedAt: now },
      { data: 'new data', fetchedAt: now }
    ]);
    expect(load).toHaveBeenCalledTimes(2);
  });

  it('enforces capacity on concurrent completion without evicting on hits or replacements', async () => {
    const pending = Array.from({ length: 105 }, () => deferred<number>());
    const requests = pending.map((item, index) =>
      getCachedWidgetData({
        key: `key-${index}`,
        ttlMilliseconds: 1_000,
        load: () => item.promise,
        classifyError: invalidate
      })
    );

    pending.forEach((item, index) => {
      now += 1;
      item.resolve(index);
    });
    await Promise.all(requests);
    expect(widgetDataCacheEntryCount()).toBe(100);

    const hit = jest.fn<Promise<number>, []>().mockResolvedValue(999);
    await getCachedWidgetData({
      key: 'key-104',
      ttlMilliseconds: 1_000,
      load: hit,
      classifyError: invalidate
    });
    expect(hit).not.toHaveBeenCalled();
    expect(widgetDataCacheEntryCount()).toBe(100);

    now += 1;
    await getCachedWidgetData({
      key: 'key-104',
      ttlMilliseconds: 1_000,
      load: hit,
      bypassCache: true,
      classifyError: invalidate
    });
    expect(hit).toHaveBeenCalledTimes(1);
    expect(widgetDataCacheEntryCount()).toBe(100);
  });
});
