/**
 * The vocabulary every widget uses to describe what it wants to show.
 *
 * Widgets map their own data onto these shapes and hand them to the primitives in
 * this folder. Nothing here knows about Microsoft Graph, SharePoint or any other
 * source, so a tile fed by a REST API looks exactly like one fed by Graph.
 */

/** Semantic colour of a piece of content. Mapped to theme colours by the primitives. */
export type WidgetTone = 'neutral' | 'accent' | 'success' | 'warning' | 'danger';

/**
 * How a collection of items is drawn. The items themselves never change — this is
 * purely presentation, chosen per tile by whoever owns the dashboard.
 */
export type WidgetItemsView = 'list' | 'compact' | 'cards' | 'gallery' | 'adaptive';

export interface IWidgetItemsViewInfo {
  id: WidgetItemsView;
  name: string;
  /** Fluent UI icon shown in the view picker. */
  iconName: string;
  /** One line describing the view, used as the picker's tooltip. */
  description: string;
}

/** Order of the view picker. */
export const WIDGET_ITEM_VIEWS: IWidgetItemsViewInfo[] = [
  { id: 'list', name: 'List', iconName: 'BulletedList2', description: 'One row per item.' },
  { id: 'compact', name: 'Compact', iconName: 'AlignLeft', description: 'Tighter rows, so more fits.' },
  { id: 'cards', name: 'Cards', iconName: 'GridViewSmall', description: 'Each item on its own card.' },
  { id: 'gallery', name: 'Gallery', iconName: 'Tiles', description: 'Cards side by side in a grid.' },
  {
    id: 'adaptive',
    name: 'Adaptive card',
    iconName: 'RectangleShape',
    description: 'Rendered as an Adaptive Card, themed from this site.'
  }
];

export function getItemsView(id: string | undefined): IWidgetItemsViewInfo | undefined {
  return WIDGET_ITEM_VIEWS.filter((view) => view.id === id)[0];
}

/** Every view, for the usual case of a widget rendering a plain collection of items. */
export const ALL_ITEM_VIEWS: WidgetItemsView[] = WIDGET_ITEM_VIEWS.map((view) => view.id);

export interface IWidgetBadge {
  text: string;
  tone?: WidgetTone;
}

/** One row rendered by `WidgetList`. */
export interface IWidgetListItem {
  /** Stable identity for React. Falls back to the position in the list. */
  key?: string;
  /** Primary line, truncated to one line. */
  title: string;
  /** Secondary line. Empty and undefined entries are dropped, the rest joined with `·`. */
  meta?: (string | undefined)[];
  /** Optional third line, wrapped over at most two lines. */
  description?: string;
  /** Fluent UI icon shown in front of the text. */
  iconName?: string;
  /** Colours the icon, the accent bar and the badge. Defaults to `accent`. */
  tone?: WidgetTone;
  /** Draws a coloured bar down the leading edge of the row. */
  hasAccentBar?: boolean;
  /** Turns the row into a link. Links open in a new tab. */
  href?: string;
  /** Turns the row into a button. Ignored when `href` is set. */
  onClick?(): void;
  badge?: IWidgetBadge;
  /** Gives the title the tone colour and extra weight, e.g. unread mail. */
  isEmphasized?: boolean;
}

/** Why a widget could not show its data. */
export interface IWidgetError {
  message: string;
  /**
   * Someone has to act before a retry can succeed — missing admin consent, for
   * instance. The message is shown as a warning and no retry is offered.
   */
  isActionRequired?: boolean;
}

/** What a widget shows when it loaded successfully but has nothing to display. */
export interface IWidgetEmptyState {
  iconName: string;
  text: string;
  /** Optional call to action, e.g. "Open Outlook". */
  actionText?: string;
  actionHref?: string;
  onAction?(): void;
}

/**
 * The state of one asynchronous read, in the shape `WidgetView` understands.
 * `useGraphData` returns this, and so should any other data hook added later.
 */
export interface IWidgetDataState<T> {
  status: 'loading' | 'ready' | 'error';
  data?: T;
  error?: IWidgetError;
  /** True while data already on screen is being re-read, e.g. after a refresh. */
  isRefreshing?: boolean;
  /** Re-runs the read. Absent for widgets whose data cannot be re-read. */
  reload?(): void;
}
