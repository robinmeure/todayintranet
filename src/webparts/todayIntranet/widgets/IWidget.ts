import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
// Imported straight from the module rather than the barrel: the barrel pulls in the
// settings controls, which import this file back.
import { WidgetItemsView } from './content/IWidgetContent';

/** Everything a widget is given at render time. */
export interface IWidgetContext {
  /** Id of this placed widget, unique within the dashboard. */
  instanceId: string;
  /** Widget owned configuration. Empty object when the widget has never been configured. */
  settings: Record<string, unknown>;
  /** SPFx context, for Graph / SharePoint calls and page context. */
  spContext: WebPartContext;
  /** True while the user is rearranging the dashboard. */
  isEditing: boolean;
  /**
   * Bumped every time the user refreshes this tile. Data hooks list it among their
   * dependencies; widgets with nothing to re-read simply ignore it.
   */
  refreshToken: number;
  /** Persist new settings for this widget instance. */
  updateSettings(settings: Record<string, unknown>): void;
}

/**
 * What the dashboard itself knows about a placed widget. The tile chrome owns the
 * refresh state, so it is what completes this into an `IWidgetContext`.
 */
export type IWidgetHostContext = Omit<IWidgetContext, 'refreshToken'>;

export interface IWidgetSize {
  w: number;
  h: number;
}

/** Groups widgets in the catalogue. */
export type WidgetCategory = 'microsoft365' | 'content' | 'demo';

export interface IWidgetCategoryInfo {
  id: WidgetCategory;
  name: string;
}

/** Catalogue order. A category with no widgets in it is skipped. */
export const WIDGET_CATEGORIES: IWidgetCategoryInfo[] = [
  { id: 'microsoft365', name: 'Microsoft 365' },
  { id: 'content', name: 'Content' },
  { id: 'demo', name: 'Samples' }
];

/** "Open in Outlook" style link, rendered in the tile footer by the frame. */
export interface IWidgetLink {
  text: string;
  href: string;
}

/**
 * A fixed link, or one resolved at render time — a link into the current site can
 * only be built once the page context is known.
 */
export type WidgetFooterLink = IWidgetLink | ((context: IWidgetContext) => IWidgetLink | undefined);

/**
 * A widget is a self contained tile. Adding one to the catalogue is a matter of
 * implementing this interface and registering it in WidgetRegistry.
 */
export interface IWidgetDefinition {
  /** Stable key persisted in the layout. Never rename an existing one. */
  type: string;
  displayName: string;
  description: string;
  /** Fluent UI icon name shown in the catalogue and tile header. */
  iconName: string;
  category: WidgetCategory;
  /** Extra terms the catalogue search box matches on, beyond name and description. */
  keywords?: string[];
  /** Microsoft Graph permission this widget needs, named in the catalogue. */
  requiredPermission?: string;
  /**
   * Puts a refresh button on the tile. Only set it when the widget actually re-reads
   * something in response to `IWidgetContext.refreshToken`.
   */
  isRefreshable?: boolean;
  /** Link out to the full experience, rendered as the tile's footer. */
  footerLink?: WidgetFooterLink;
  /** Widens the settings flyout, for settings that need the room (a query preview). */
  isSettingsWide?: boolean;
  /**
   * Views this widget's content can be drawn in. Declaring more than one puts a view
   * picker in the tile's settings; the widget reads the choice with `viewSetting`
   * and passes it to `WidgetItems`.
   */
  supportedViews?: WidgetItemsView[];
  /** View used until someone picks one. Defaults to the first supported view. */
  defaultView?: WidgetItemsView;
  /** Size used when the widget is first dropped on the grid. */
  defaultSize: IWidgetSize;
  minSize?: IWidgetSize;
  render(context: IWidgetContext): React.ReactElement;
  /**
   * Optional configuration UI, shown in the tile's settings flyout while the
   * dashboard is in edit mode. Use `context.updateSettings` to persist changes.
   */
  renderSettings?(context: IWidgetContext): React.ReactElement;
}
