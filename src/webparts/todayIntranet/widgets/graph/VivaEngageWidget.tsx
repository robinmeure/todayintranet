import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  IWidgetListItem,
  NumberSetting,
  SettingsSurface,
  WidgetErrorMessage,
  WidgetItems,
  WidgetItemsView,
  WidgetNotice,
  WidgetView,
  numberSetting,
  textSetting,
  viewSetting
} from '../content';
import {
  ENGAGE_HOME_URL,
  IEngageCommunity,
  IEngageCommunityCard,
  IEngageFile,
  engageCommunityUrl,
  listCommunities,
  loadCommunityCard
} from './vivaEngageData';
import styles from './VivaEngageWidget.module.scss';

const SCOPE: string = 'Community.Read.All';
const COMMUNITY_KEY: string = 'communityId';
const DEFAULT_MAX_FILES: number = 5;
const MAX_FILES_BOUNDS = { min: 1, max: 20 };

export const VIVA_ENGAGE_DEFAULT_VIEW: WidgetItemsView = 'compact';
export const VIVA_ENGAGE_PERMISSIONS: string[] = [
  'Community.Read.All',
  'GroupMember.Read.All (member count)',
  'Files.Read.All (community files)'
];

/** Opens the chosen community, or Viva Engage itself until one is chosen. */
export function vivaEngageFooterLink(context: IWidgetContext): IWidgetLink {
  const communityId = textSetting(context, COMMUNITY_KEY, '');
  return {
    text: 'Open in Viva Engage',
    href: communityId ? engageCommunityUrl(communityId) : ENGAGE_HOME_URL
  };
}

function formatModified(modified: string | undefined, locale: string | undefined): string {
  return modified
    ? new Date(modified).toLocaleDateString(locale, { day: 'numeric', month: 'short', year: 'numeric' })
    : '';
}

function toItem(file: IEngageFile, locale: string | undefined): IWidgetListItem {
  return {
    key: file.id,
    title: file.name,
    meta: [file.lastModifiedBy ?? '', formatModified(file.lastModifiedDateTime, locale)].filter(Boolean),
    iconName: 'Page',
    href: file.webUrl
  };
}

function privacyLabel(privacy: IEngageCommunity['privacy']): string | undefined {
  if (privacy === 'public') {
    return 'Public';
  }
  return privacy === 'private' ? 'Private' : undefined;
}

const CommunityCard: React.FunctionComponent<{
  card: IEngageCommunityCard; view: WidgetItemsView; locale: string | undefined;
}> = ({ card, view, locale }) => {
  const { community, memberCount, files, filesUnavailable } = card;
  const privacy = privacyLabel(community.privacy);
  return (
    <div className={styles.engage}>
      <div className={styles.header}>
        <a className={styles.name} href={engageCommunityUrl(community.id)} target="_blank" rel="noreferrer">
          {community.displayName}
        </a>
        {(privacy || memberCount !== undefined) && (
          <div className={styles.meta}>
            {privacy && <span className={styles.badge}>{privacy}</span>}
            {memberCount !== undefined && (
              <span>{memberCount === 1 ? '1 member' : `${memberCount.toLocaleString(locale)} members`}</span>
            )}
          </div>
        )}
        {community.description && <p className={styles.description}>{community.description}</p>}
      </div>
      <section aria-label="Recent community files">
        <h3 className={styles.sectionTitle}>Recent files</h3>
        {files.length > 0 ? (
          <WidgetItems view={view} ariaLabel="Recent community files" items={files.map((file) => toItem(file, locale))} />
        ) : (
          <p className={styles.sectionNote}>
            {filesUnavailable ? 'Community files could not be loaded.' : 'No files have been shared in this community yet.'}
          </p>
        )}
      </section>
    </div>
  );
};

export const VivaEngageWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const communityId = textSetting(context, COMMUNITY_KEY, '');
  const maxFiles = numberSetting(context, 'maxFiles', DEFAULT_MAX_FILES, MAX_FILES_BOUNDS);
  const view = viewSetting(context, VIVA_ENGAGE_DEFAULT_VIEW);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const state = useGraphData<IEngageCommunityCard>(
    context,
    SCOPE,
    (client) => loadCommunityCard(client, communityId, maxFiles),
    ['card', communityId, maxFiles],
    !!communityId
  );

  if (!communityId) {
    return (
      <WidgetNotice
        title="No community chosen"
        description="Pick the Viva Engage community this tile should follow."
        actionText="Choose community"
        onAction={context.openSettings ? () => context.openSettings?.() : undefined}
      />
    );
  }

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Loading community…', rows: 3 }}
      empty={{ iconName: 'Group', text: 'This community could not be found.' }}
    >
      {(card) => <CommunityCard card={card} view={view} locale={locale} />}
    </WidgetView>
  );
};

export const VivaEngageWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const communityId = textSetting(context, COMMUNITY_KEY, '');
  const communities = useGraphData<IEngageCommunity[]>(context, SCOPE, listCommunities, ['list']);

  const options: IDropdownOption[] = (communities.data ?? []).map((community) => ({
    key: community.id,
    text: community.privacy === 'private' ? `${community.displayName} (private)` : community.displayName
  }));
  if (communityId && !options.some((option) => option.key === communityId)) {
    // Keep the saved choice visible even if it is no longer in the list.
    options.unshift({ key: communityId, text: textSetting(context, 'communityName', 'Saved community') });
  }

  const choose = (option: IDropdownOption | undefined): void => {
    if (option) {
      context.updateSettings({
        ...context.settings,
        [COMMUNITY_KEY]: String(option.key),
        communityName: option.text.replace(/ \(private\)$/, '')
      });
    }
  };

  return (
    <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
      {communities.status === 'error' && communities.error ? (
        <WidgetErrorMessage error={communities.error} onRetry={communities.reload} />
      ) : (
        <Dropdown
          label="Community"
          placeholder={communities.status === 'loading' ? 'Loading communities…' : 'Choose a community'}
          disabled={communities.status === 'loading'}
          options={options}
          selectedKey={communityId || null}
          onChange={(_event, option) => choose(option)}
        />
      )}
      <NumberSetting
        context={context}
        settingKey="maxFiles"
        label="Files to show"
        fallback={DEFAULT_MAX_FILES}
        {...MAX_FILES_BOUNDS}
      />
    </SettingsSurface>
  );
};
