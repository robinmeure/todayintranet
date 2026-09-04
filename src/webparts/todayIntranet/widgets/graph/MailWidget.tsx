import * as React from 'react';
import { Message } from '@microsoft/microsoft-graph-types';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetContext } from '../IWidget';
import { useGraphData } from './useGraphData';
import { WidgetLoading, WidgetEmpty, WidgetError } from '../WidgetMessage';
import { NumberSetting, ToggleSetting, SettingsSurface, numberSetting, booleanSetting } from './settings';
import styles from '../WidgetMessage.module.scss';

const SCOPE: string = 'Mail.ReadBasic';
const DEFAULT_MAX_ITEMS: number = 6;

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

export const MailWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS);
  const unreadOnly = booleanSetting(context, 'unreadOnly', false);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const result = useGraphData<Message[]>(
    context.spContext,
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

  if (result.status === 'loading') {
    return <WidgetLoading label="Loading your inbox…" />;
  }
  if (result.status === 'error' && result.error) {
    return <WidgetError error={result.error} onRetry={result.reload} />;
  }
  if (!result.data || result.data.length === 0) {
    return <WidgetEmpty iconName="Mail" text={unreadOnly ? 'No unread mail. Nicely done.' : 'Your inbox is empty.'} />;
  }

  return (
    <ul className={styles.list}>
      {result.data.map((message) => (
        <li key={message.id}>
          <a className={styles.item} href={message.webLink ?? '#'} target="_blank" rel="noreferrer">
            <Icon
              iconName={message.isRead ? 'Mail' : 'MailFill'}
              className={styles.itemIcon}
              aria-label={message.isRead ? 'Read' : 'Unread'}
            />
            <span className={styles.itemText}>
              <span className={`${styles.itemTitle} ${message.isRead ? '' : styles.itemTitleUnread}`}>
                {message.subject || '(No subject)'}
              </span>
              <span className={styles.itemMeta}>
                {message.from?.emailAddress?.name ?? 'Unknown sender'} ·{' '}
                {formatReceived(message.receivedDateTime ?? undefined, locale)}
              </span>
            </span>
          </a>
        </li>
      ))}
    </ul>
  );
};

export const MailWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface>
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
