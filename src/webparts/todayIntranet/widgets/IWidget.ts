import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';

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
  /** Persist new settings for this widget instance. */
  updateSettings(settings: Record<string, unknown>): void;
}

export interface IWidgetSize {
  w: number;
  h: number;
}

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
