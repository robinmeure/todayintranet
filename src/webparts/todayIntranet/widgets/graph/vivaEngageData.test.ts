import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { engageCommunityUrl, listCommunities, loadCommunityCard } from './vivaEngageData';

type Responder = (url: string, headers: Record<string, string>) => Promise<unknown>;

function graphClient(respond: Responder): { client: MSGraphClientV3; urls: string[] } {
  const urls: string[] = [];
  const client = new (jest.fn<MSGraphClientV3, []>())();
  Object.assign(client, {
    api: (url: string) => {
      const headers: Record<string, string> = {};
      let fullUrl = url;
      const request = {
        version: () => request,
        select: (fields: string) => { fullUrl += `|select=${fields}`; return request; },
        header: (name: string, value: string) => { headers[name] = value; return request; },
        get: () => { urls.push(fullUrl); return respond(fullUrl, headers); }
      };
      return request;
    }
  });
  return { client, urls };
}

const notFound = (): Promise<never> => Promise.reject({ statusCode: 404, code: 'itemNotFound' });
const community = { id: 'eyJfdHlwZSI6Ikdyb3VwIn0=', displayName: 'Ops', description: 'Ops talk', privacy: 'public', groupId: 'g1' };

describe('vivaEngageData', () => {
  it('builds the documented community deep link', () => {
    expect(engageCommunityUrl('abc=')).toBe('https://engage.cloud.microsoft/main/groups/abc%3D');
  });

  it('lists communities across pages', async () => {
    const { client, urls } = graphClient(async (url) => url.startsWith('/employeeExperience')
      ? { value: [{ id: 'a', displayName: 'A' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' }
      : { value: [{ id: 'b', displayName: 'B', privacy: 'private' }] });
    const result = await listCommunities(client);
    expect(result.map((item) => item.id)).toEqual(['a', 'b']);
    expect(urls[0]).toContain('$orderby=displayName');
  });

  it('loads the community, member count and newest files only', async () => {
    const { client, urls } = graphClient(async (url, headers) => {
      if (url.startsWith('/employeeExperience')) { return community; }
      if (url.includes('/members')) {
        expect(headers.ConsistencyLevel).toBe('eventual');
        return { '@odata.count': 42, value: [] };
      }
      if (url.includes('Apps/Viva%20Engage')) {
        return { value: [
          { id: 'old', name: 'old.docx', file: {}, lastModifiedDateTime: '2026-01-01T00:00:00Z' },
          { id: 'folder', name: 'Folder', folder: {}, lastModifiedDateTime: '2026-06-01T00:00:00Z' },
          { id: 'new', name: 'new.pptx', file: {}, lastModifiedDateTime: '2026-05-01T00:00:00Z',
            lastModifiedBy: { user: { displayName: 'Ann' } } }
        ] };
      }
      throw new Error(`unexpected ${url}`);
    });

    const card = await loadCommunityCard(client, community.id, 1);
    expect(urls[0]).toContain(encodeURIComponent(community.id));
    expect(card.community.displayName).toBe('Ops');
    expect(card.memberCount).toBe(42);
    expect(card.files).toEqual([expect.objectContaining({ id: 'new', lastModifiedBy: 'Ann' })]);
    expect(card.filesUnavailable).toBe(false);
  });

  it('falls back to the legacy Yammer folder and treats missing folders as no files', async () => {
    const { client, urls } = graphClient(async (url) => {
      if (url.startsWith('/employeeExperience')) { return community; }
      if (url.includes('/members')) { return { '@odata.count': 3 }; }
      return notFound();
    });
    const card = await loadCommunityCard(client, community.id, 5);
    expect(urls.some((url) => url.includes('Apps/Yammer'))).toBe(true);
    expect(card.files).toEqual([]);
    expect(card.filesUnavailable).toBe(false);
  });

  it('degrades member count and files when their permissions are missing', async () => {
    const { client } = graphClient(async (url) => url.startsWith('/employeeExperience')
      ? community
      : Promise.reject({ statusCode: 403, code: 'Authorization_RequestDenied' }));
    const card = await loadCommunityCard(client, community.id, 5);
    expect(card.memberCount).toBeUndefined();
    expect(card.filesUnavailable).toBe(true);
  });

  it('propagates a failure to read the community itself', async () => {
    const { client } = graphClient(() => Promise.reject({ statusCode: 403 }));
    await expect(loadCommunityCard(client, 'x', 5)).rejects.toEqual({ statusCode: 403 });
  });
});
