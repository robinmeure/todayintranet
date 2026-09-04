import { IWidgetInstance } from './IDashboardLayout';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

/** Width of the authoring grid. Narrower breakpoints are derived from it. */
export const GRID_COLUMNS: number = 12;

/**
 * A preset describes one repeating band of slots across the grid. Widgets are
 * poured into the slots in their current order, so applying a preset rearranges
 * the dashboard without changing which widgets are on it.
 */
export interface ILayoutPreset {
  /** Stable id persisted in the layout. Never rename an existing one. */
  id: string;
  name: string;
  description: string;
  /** Column spans of one band, summing to GRID_COLUMNS. */
  columns: number[];
  /** True when the first widget gets a full width row of its own. */
  banner?: boolean;
}

/** Tile height, in grid row units. This is the "rows" half of a layout choice. */
export interface ILayoutRowSize {
  id: string;
  name: string;
  description: string;
  rows: number;
}

export const LAYOUT_PRESETS: ILayoutPreset[] = [
  {
    id: 'single',
    name: 'Single column',
    description: 'One full width widget per row. Easiest to read on any screen.',
    columns: [GRID_COLUMNS]
  },
  {
    id: 'two',
    name: 'Two columns',
    description: 'Two equal widgets per row.',
    columns: [6, 6]
  },
  {
    id: 'three',
    name: 'Three columns',
    description: 'Three equal widgets per row.',
    columns: [4, 4, 4]
  },
  {
    id: 'four',
    name: 'Four columns',
    description: 'Four compact widgets per row. Best with short tiles.',
    columns: [3, 3, 3, 3]
  },
  {
    id: 'main-side',
    name: 'Main and sidebar',
    description: 'A wide widget with a narrow companion on the right.',
    columns: [8, 4]
  },
  {
    id: 'side-main',
    name: 'Sidebar and main',
    description: 'A narrow widget on the left, a wide one on the right.',
    columns: [4, 8]
  },
  {
    id: 'banner-two',
    name: 'Banner and two columns',
    description: 'A full width widget on top, then pairs beneath it.',
    columns: [6, 6],
    banner: true
  },
  {
    id: 'banner-three',
    name: 'Banner and three columns',
    description: 'A full width widget on top, then trios beneath it.',
    columns: [4, 4, 4],
    banner: true
  }
];

export const ROW_SIZES: ILayoutRowSize[] = [
  { id: 'short', name: 'Short', description: 'Fits more on screen', rows: 4 },
  { id: 'medium', name: 'Medium', description: 'The default', rows: 6 },
  { id: 'tall', name: 'Tall', description: 'Shows more per widget', rows: 9 }
];

export const DEFAULT_ROW_SIZE_ID: string = 'medium';

export function getPreset(id: string | undefined): ILayoutPreset | undefined {
  return id ? LAYOUT_PRESETS.filter((preset) => preset.id === id)[0] : undefined;
}

export function getRowSize(id: string | undefined): ILayoutRowSize {
  return ROW_SIZES.filter((size) => size.id === id)[0] ?? ROW_SIZES.filter((size) => size.id === DEFAULT_ROW_SIZE_ID)[0];
}

/**
 * Pours widgets into the preset's slots, left to right and top to bottom, keeping
 * their existing order. A widget whose minimum width is wider than its slot takes
 * the width it needs and the band wraps early rather than overflowing the grid.
 */
export function applyPreset(
  widgets: IWidgetInstance[],
  preset: ILayoutPreset,
  rowSize: ILayoutRowSize
): IWidgetInstance[] {
  const placed: IWidgetInstance[] = [];
  let index = 0;
  let y = 0;

  const place = (widget: IWidgetInstance, x: number, width: number): IWidgetInstance => {
    const definition = WidgetRegistry.get(widget.type);
    const minH = definition?.minSize?.h ?? 2;
    return { ...widget, x, y, w: width, h: Math.max(minH, rowSize.rows) };
  };

  const widthFor = (widget: IWidgetInstance, slot: number): number => {
    const definition = WidgetRegistry.get(widget.type);
    const minW = definition?.minSize?.w ?? 1;
    return Math.min(GRID_COLUMNS, Math.max(minW, slot));
  };

  if (preset.banner && widgets.length > 0) {
    placed.push(place(widgets[0], 0, GRID_COLUMNS));
    index = 1;
    y += placed[0].h;
  }

  while (index < widgets.length) {
    const band: IWidgetInstance[] = [];
    let x = 0;
    let tallest = 0;
    for (let slot = 0; slot < preset.columns.length && index < widgets.length; slot++) {
      const widget = widgets[index];
      const width = widthFor(widget, preset.columns[slot]);
      if (x > 0 && x + width > GRID_COLUMNS) {
        break;
      }
      const next = place(widget, x, width);
      band.push(next);
      tallest = Math.max(tallest, next.h);
      x += width;
      index++;
    }
    // Every tile in a band gets the same height. Ragged bands leave gaps that the
    // grid's vertical compaction would close, moving tiles away from the preset.
    // A band cut short by an oversized widget would leave a gap to its right that
    // the next band floats up into, so the last tile stretches to close it. The
    // final band is left alone: nothing sits below it to float up.
    if (band.length > 0 && x < GRID_COLUMNS && index < widgets.length) {
      band[band.length - 1].w += GRID_COLUMNS - x;
    }
    band.forEach((widget) => placed.push({ ...widget, h: tallest }));
    y += tallest || rowSize.rows;
  }

  return placed;
}
