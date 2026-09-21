import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import {
  NumberSetting,
  SettingsSurface,
  TextSetting,
  WidgetEmpty,
  WidgetErrorMessage,
  WidgetFilteredItems,
  WidgetItemsView,
  numberSetting,
  textSetting,
  viewSetting
} from '../content';
import { DEFAULT_LINKS_SETTING, parseLinksSetting } from './linksSettings';

export { parseLinksSetting } from './linksSettings';

const DEFAULT_ITEMS_PER_PAGE: number = 9;
const ITEMS_PER_PAGE_BOUNDS = { min: 3, max: 12 };
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
    return <WidgetEmpty iconName="Link"
      text="No links configured. Edit the dashboard and open My links settings to add your organization's HTTPS links." />;
  }
  return (
    <WidgetFilteredItems
      ariaLabel="My links"
      view={viewSetting(context, MY_LINKS_DEFAULT_VIEW)}
      itemsPerPage={itemsPerPage}
      items={parsed.links.map((link, index) => ({
        key: String(index),
        title: link.title,
        href: link.url,
        description: link.description,
        iconName: link.iconName || 'Link',
        tone: link.tone ?? 'accent',
        badge: link.badge ? { text: link.badge, iconName: 'Lock', tone: 'accent' } : undefined
      }))}
    />
  );
};

export const MyLinksWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description={
    'Links are personal to this tile. The defaults are sample links to public resources; replace them or enter [] to start empty. ' +
    'Add a JSON array of objects with title and an absolute HTTPS url. ' +
    'Optional fields: description, iconName (Fluent UI), badge (for example VPN), and tone ' +
    '(neutral, accent, success, warning, danger). Only use destinations and VPN labels verified by your organization.'
  }>
    <NumberSetting context={context} settingKey="itemsPerPage" label="Links per page"
      fallback={DEFAULT_ITEMS_PER_PAGE} {...ITEMS_PER_PAGE_BOUNDS} />
    <TextSetting context={context} settingKey="links" label="Links (JSON)" fallback={DEFAULT_LINKS_SETTING}
      multiline={true} rows={16} commitOn="blur" />
  </SettingsSurface>
);
