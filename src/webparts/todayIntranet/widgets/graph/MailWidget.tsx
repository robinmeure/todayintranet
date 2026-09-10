import * as React from 'react';
import { Message } from '@microsoft/microsoft-graph-types';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  IWidgetListItem,
  NumberSetting,
  SettingsSurface,
  ToggleSetting,
  WidgetItems,
  WidgetItemsView,
  WidgetView,
  booleanSetting,
  numberSetting,
  viewSetting
} from '../content';

const SCOPE: string = 'Mail.ReadBasic';
const DEFAULT_MAX_ITEMS: number = 6;
const MAIL_URL: string = 'https://outlook.office.com/mail/';

export const MAIL_LINK: IWidgetLink = { text: 'Open Outlook', href: MAIL_URL };
export const MAIL_DEFAULT_VIEW: WidgetItemsView = 'compact';

function formatReceived(received: string | undefined, locale: string | undefined): string {
  if (!received) {
    return '';
  }
  const date = new Date(received);
  const isToday = new Date().toDateString() === date.toDateString();
  return isToday
    ? date.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })
    : date.toLocaleDateString(locale, { day: 'numeric', month: 'short' });
}

/** One message, described in the shared list vocabulary. */
function toItem(message: Message, index: number, locale: string | undefined): IWidgetListItem {
  const isUnread = !message.isRead;
  return {
    key: message.id ?? String(index),
    title: message.subject || '(No subject)',
    meta: [
      message.from?.emailAddress?.name ?? 'Unknown sender',
      formatReceived(message.receivedDateTime ?? undefined, locale)
    ],
    // The accent bar carries "unread" instead of a second icon: it keeps every row
    // aligned, leaves the full width for the subject, and survives the view switch.
    hasAccentBar: true,
    tone: isUnread ? 'accent' : 'neutral',
    isEmphasized: isUnread,
    href: message.webLink ?? undefined
  };
}

export const MailWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS);
  const unreadOnly = booleanSetting(context, 'unreadOnly', false);
  const view = viewSetting(context, MAIL_DEFAULT_VIEW);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const state = useGraphData<Message[]>(
    context,
    SCOPE,
    async (client) => {
      let request = client
        .api('/me/mailFolders/inbox/messages')
        .version('v1.0')
        .select('subject,from,receivedDateTime,isRead,webLink')
        .orderby('receivedDateTime desc')
        .top(maxItems);

      if (unreadOnly) {
        request = request.filter('isRead eq false');
      }

      const response = await request.get();
      return (response.value ?? []) as Message[];
    },
    [maxItems, unreadOnly]
  );

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Loading your inbox…', rows: Math.min(maxItems, 4) }}
      empty={{
        iconName: 'Mail',
        text: unreadOnly ? 'No unread mail. Nicely done.' : 'Your inbox is empty.'
      }}
    >
      {(messages) => (
        <WidgetItems
          view={view}
          ariaLabel={unreadOnly ? 'Unread messages' : 'Recent messages'}
          items={messages.map((message, index) => toItem(message, index, locale))}
        />
      )}
    </WidgetView>
  );
};

export const MailWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Messages to show"
      fallback={DEFAULT_MAX_ITEMS}
      min={1}
      max={20}
    />
    <ToggleSetting context={context} settingKey="unreadOnly" label="Unread only" fallback={false} />
  </SettingsSurface>
);
