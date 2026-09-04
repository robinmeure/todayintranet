/**
 * Placement + configuration of a single widget on the dashboard grid.
 */
export interface IWidgetInstance {
  /** Stable id for this placed widget, unique within a dashboard. */
  id: string;
  /** Key of the widget definition in the widget registry. */
  type: string;
  /** Column (0 based) on a 12 column grid. */
  x: number;
  /** Row (0 based), expressed in grid row units. */
  y: number;
  /** Width in grid columns. */
  w: number;
  /** Height in grid row units. */
  h: number;
  /** Widget specific configuration, owned by the widget itself. */
  settings?: Record<string, unknown>;
}

export interface IDashboardLayout {
  /** Schema version, so stored layouts can be migrated later. */
  version: number;
  widgets: IWidgetInstance[];
  /**
   * Layout preset the widgets were last poured into, or undefined once the user
   * has moved something by hand ("Custom").
   */
  presetId?: string;
  /** Tile height chosen with the preset. See ROW_SIZES in LayoutPresets. */
  rowSizeId?: string;
}

/** v2 added the optional presetId / rowSizeId fields; v1 layouts load unchanged. */
export const CURRENT_LAYOUT_VERSION: number = 2;

export function emptyLayout(): IDashboardLayout {
  return { version: CURRENT_LAYOUT_VERSION, widgets: [] };
}

export function cloneLayout(layout: IDashboardLayout): IDashboardLayout {
  return {
    version: layout.version,
    presetId: layout.presetId,
    rowSizeId: layout.rowSizeId,
    widgets: layout.widgets.map((w) => ({
      ...w,
      settings: w.settings ? { ...w.settings } : undefined
    }))
  };
}
