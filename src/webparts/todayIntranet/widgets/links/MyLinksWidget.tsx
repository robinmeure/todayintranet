import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import {
  WidgetErrorMessage,
  WidgetFilteredItems,
  WidgetItemsView,
  WidgetNotice,
  numberSetting,
  textSetting,
  viewSetting
} from '../content';
import { DEFAULT_ITEMS_PER_PAGE, DEFAULT_LINKS_SETTING, ITEMS_PER_PAGE_BOUNDS, parseLinksSetting } from './linksSettings';
export { MyLinksWidgetSettings } from './MyLinksWidgetSettings';

export { parseLinksSetting } from './linksSettings';

export const MY_LINKS_DEFAULT_VIEW: WidgetItemsView = 'links';

export const MyLinksWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const parsed = parseLinksSetting(textSetting(context, 'links', DEFAULT_LINKS_SETTING));
  const itemsPerPage = numberSetting(context, 'itemsPerPage', DEFAULT_ITEMS_PER_PAGE, ITEMS_PER_PAGE_BOUNDS);
  if (parsed.error) {
    return <WidgetErrorMessage error={{
      message: `${parsed.error} Edit the dashboard and correct Links (JSON) in My links settings.`,
      isActionRequired: true
    }} />;
  }
  if (!parsed.links?.length) {
    return <WidgetNotice title="No links added"
      description={"Looks like you haven't added any links yet, or you've removed them all. " +
        'Feel free to customize this section with the links and resources you use most often.'}
      actionText="Add link" onAction={context.openSettings ? () => context.openSettings?.('add') : undefined} />;
  }
  return (
    <WidgetFilteredItems
      ariaLabel="My links"
      view={viewSetting(context, MY_LINKS_DEFAULT_VIEW)}
      itemsPerPage={itemsPerPage}
      emptyState={<WidgetNotice title="No links found" />}
      items={(parsed.links ?? []).map((link, index) => ({
        key: String(index),
        title: link.title,
        href: link.url,
        description: link.description,
        iconName: link.iconName || 'Link',
        tone: link.tone ?? 'accent',
        badge: link.badge ? { text: link.badge, iconName: 'Lock', tone: 'neutral' } : undefined
      }))}
    />
  );
};
