import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { widgetDataErrorStatus } from '../data/WidgetDataError';

/**
 * Viva Engage through Microsoft Graph.
 *
 * Graph only exposes communities (v1.0 `employeeExperience/communities`) and the
 * Microsoft 365 group behind each one. Posts — discussions, praise, polls, replies,
 * reactions — are not available, so this module reads community details, the group's
 * member count and the files people shared into the community.
 */

export type EngageCommunityPrivacy = 'public' | 'private' | 'unknownFutureValue';

export interface IEngageCommunity {
  id: string;
  displayName: string;
  description?: string;
  privacy?: EngageCommunityPrivacy;
  groupId?: string;
}

export interface IEngageFile {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: string;
}

export interface IEngageCommunityCard {
  community: IEngageCommunity;
  /** Undefined when the member count could not be read (for example, no GroupMember.Read.All). */
  memberCount?: number;
  files: IEngageFile[];
  /** True when the community's file folder could not be read, as opposed to being empty. */
  filesUnavailable: boolean;
}

interface IGraphCollection<T> {
  value?: T[];
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
}

interface IDriveItemShape {
  id?: string;
  name?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  file?: unknown;
}

const ENGAGE_WEB_URL: string = 'https://engage.cloud.microsoft';
/** Files people post into a community land in one of these folders of the group's library. */
const COMMUNITY_FILE_FOLDERS: string[] = ['Apps/Viva Engage', 'Apps/Yammer'];
const FILE_PAGE_SIZE: number = 200;
const COMMUNITY_PAGE_SIZE: number = 100;
const MAX_COMMUNITY_PAGES: number = 10;

export const ENGAGE_HOME_URL: string = `${ENGAGE_WEB_URL}/main/feed`;

/** Documented deep link format: https://learn.microsoft.com/viva/engage/deeplinks */
export function engageCommunityUrl(communityId: string): string {
  return `${ENGAGE_WEB_URL}/main/groups/${encodeURIComponent(communityId)}`;
}

function toCommunity(raw: IEngageCommunity): IEngageCommunity {
  return {
    id: raw.id,
    displayName: raw.displayName || 'Untitled community',
    description: raw.description || undefined,
    privacy: raw.privacy,
    groupId: raw.groupId || undefined
  };
}

/** Communities the signed-in user can see, alphabetically. Bounded so a huge network cannot stall the picker. */
export async function listCommunities(client: MSGraphClientV3): Promise<IEngageCommunity[]> {
  const communities: IEngageCommunity[] = [];
  let url: string | undefined =
    `/employeeExperience/communities?$select=id,displayName,privacy&$orderby=displayName&$top=${COMMUNITY_PAGE_SIZE}`;

  for (let page = 0; url && page < MAX_COMMUNITY_PAGES; page++) {
    const response: IGraphCollection<IEngageCommunity> = await client.api(url).version('v1.0').get();
    (response.value ?? []).forEach((community) => {
      if (community.id) {
        communities.push(toCommunity(community));
      }
    });
    url = response['@odata.nextLink'];
  }
  return communities;
}

async function readMemberCount(client: MSGraphClientV3, groupId: string): Promise<number | undefined> {
  try {
    const response: IGraphCollection<unknown> = await client
      .api(`/groups/${encodeURIComponent(groupId)}/members?$count=true&$top=1&$select=id`)
      .version('v1.0')
      .header('ConsistencyLevel', 'eventual')
      .get();
    const count = response['@odata.count'];
    return typeof count === 'number' ? count : undefined;
  } catch {
    return undefined;
  }
}

function toFile(item: IDriveItemShape, index: number): IEngageFile {
  return {
    id: item.id ?? String(index),
    name: item.name || 'Untitled file',
    webUrl: item.webUrl,
    lastModifiedDateTime: item.lastModifiedDateTime,
    lastModifiedBy: item.lastModifiedBy?.user?.displayName
  };
}

async function readRecentFiles(
  client: MSGraphClientV3,
  groupId: string,
  maxFiles: number
): Promise<{ files: IEngageFile[]; filesUnavailable: boolean }> {
  for (const folder of COMMUNITY_FILE_FOLDERS) {
    try {
      const response: IGraphCollection<IDriveItemShape> = await client
        .api(
          `/groups/${encodeURIComponent(groupId)}/drive/root:/${encodeURI(folder)}:/children` +
          `?$select=id,name,webUrl,lastModifiedDateTime,lastModifiedBy,file&$top=${FILE_PAGE_SIZE}`
        )
        .version('v1.0')
        .get();
      // Sorted here: SharePoint-backed drives do not reliably honour $orderby on children.
      const files = (response.value ?? [])
        .filter((item) => !!item.file)
        .sort((a, b) => (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? ''))
        .slice(0, maxFiles)
        .map(toFile);
      return { files, filesUnavailable: false };
    } catch (error) {
      if (widgetDataErrorStatus(error) !== 404) {
        return { files: [], filesUnavailable: true };
      }
    }
  }
  // Neither folder exists: nobody has shared a file into this community yet.
  return { files: [], filesUnavailable: false };
}

/**
 * The community, plus its member count and recent files. The community read is what
 * the tile depends on, so its failure propagates; the other two degrade on their own.
 */
export async function loadCommunityCard(
  client: MSGraphClientV3,
  communityId: string,
  maxFiles: number
): Promise<IEngageCommunityCard> {
  const raw: IEngageCommunity = await client
    .api(`/employeeExperience/communities/${encodeURIComponent(communityId)}`)
    .version('v1.0')
    .select('id,displayName,description,privacy,groupId')
    .get();
  const community = toCommunity({ ...raw, id: raw.id || communityId });

  if (!community.groupId) {
    return { community, files: [], filesUnavailable: true };
  }

  const [memberCount, fileResult] = await Promise.all([
    readMemberCount(client, community.groupId),
    readRecentFiles(client, community.groupId, maxFiles)
  ]);
  return { community, memberCount, ...fileResult };
}
