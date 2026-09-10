import * as React from 'react';
import { Responsive, WidthProvider, Layout, Layouts } from 'react-grid-layout';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ThemeProvider, createTheme, IPartialTheme } from '@fluentui/react/lib/Theme';
import styles from './Dashboard.module.scss';
import { WidgetFrame, INudge } from './WidgetFrame';
import { AddWidgetPanel } from './AddWidgetPanel';
import { LayoutPresetPanel } from './LayoutPresetPanel';
import { IDashboardLayout, IWidgetInstance, CURRENT_LAYOUT_VERSION } from '../model/IDashboardLayout';
import {
  GRID_COLUMNS,
  applyPreset,
  getPreset,
  getRowSize,
  DEFAULT_ROW_SIZE_ID
} from '../model/LayoutPresets';
import { ILayoutStore } from '../services/ILayoutStore';
import { IWidgetContext, IWidgetDefinition } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

const ResponsiveGridLayout = WidthProvider(Responsive);

const BREAKPOINTS: Record<string, number> = { lg: 1200, md: 996, sm: 768, xs: 480, xxs: 0 };
const COLS: Record<string, number> = { lg: GRID_COLUMNS, md: 8, sm: 6, xs: 4, xxs: 2 };
const ROW_HEIGHT: number = 56;
const SAVE_DEBOUNCE_MS: number = 800;

export interface IDashboardProps {
  title: string;
  spContext: WebPartContext;
  store: ILayoutStore;
  /** Layout seeded for users who have never arranged their dashboard. */
  starterLayout: IDashboardLayout;
  /** Current SPFx section / site theme, mapped onto Fluent UI controls. */
  theme?: IReadonlyTheme;
}

/** The canonical layout is authored on 12 columns; narrower breakpoints clamp it. */
function toGridLayouts(widgets: IWidgetInstance[]): Layouts {
  const layouts: Layouts = {};
  Object.keys(COLS).forEach((breakpoint) => {
    const cols = COLS[breakpoint];
    layouts[breakpoint] = widgets.map((widget) => {
      const definition = WidgetRegistry.get(widget.type);
      const w = Math.min(widget.w, cols);
      return {
        i: widget.id,
        x: Math.min(widget.x, Math.max(0, cols - w)),
        y: widget.y,
        w,
        h: widget.h,
        minW: definition?.minSize ? Math.min(definition.minSize.w, cols) : 1,
        minH: definition?.minSize?.h ?? 2
      } as Layout;
    });
  });
  return layouts;
}

function applyGridLayout(widgets: IWidgetInstance[], layout: Layout[]): IWidgetInstance[] {
  const byId = new Map(layout.map((item) => [item.i, item]));
  return widgets.map((widget) => {
    const item = byId.get(widget.id);
    return item ? { ...widget, x: item.x, y: item.y, w: item.w, h: item.h } : widget;
  });
}

function newInstanceId(): string {
  return `w_${Date.now().toString(36)}_${Math.random().toString(36).substring(2, 8)}`;
}

/**
 * The grid compacts vertically, so every tile already sits at its smallest possible
 * `y`. Moving one by a single row is therefore always undone by the next compaction.
 * A keyboard move has to swap the tile with the neighbour it overlaps horizontally.
 * Returns undefined when there is nothing to swap with.
 */
function swapVertically(
  widgets: IWidgetInstance[],
  target: IWidgetInstance,
  direction: number
): IWidgetInstance[] | undefined {
  const overlapping = widgets.filter(
    (w) => w.id !== target.id && w.x < target.x + target.w && target.x < w.x + w.w
  );

  if (direction < 0) {
    const above = overlapping
      .filter((w) => w.y + w.h <= target.y)
      .sort((a, b) => b.y + b.h - (a.y + a.h))[0];
    if (!above) {
      return undefined;
    }
    return widgets.map((w) => {
      if (w.id === target.id) {
        return { ...w, y: above.y };
      }
      return w.id === above.id ? { ...w, y: above.y + target.h } : w;
    });
  }

  const below = overlapping.filter((w) => w.y >= target.y + target.h).sort((a, b) => a.y - b.y)[0];
  if (!below) {
    return undefined;
  }
  return widgets.map((w) => {
    if (w.id === target.id) {
      return { ...w, y: below.y + below.h };
    }
    return w.id === below.id ? { ...w, y: target.y } : w;
  });
}

export const Dashboard: React.FunctionComponent<IDashboardProps> = (props) => {
  const { title, spContext, store, starterLayout, theme } = props;

  const [widgets, setWidgets] = React.useState<IWidgetInstance[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isEditing, setIsEditing] = React.useState<boolean>(false);
  const [isCatalogueOpen, setIsCatalogueOpen] = React.useState<boolean>(false);
  const [isLayoutPanelOpen, setIsLayoutPanelOpen] = React.useState<boolean>(false);
  const [breakpoint, setBreakpoint] = React.useState<string>('lg');
  const [presetId, setPresetId] = React.useState<string | undefined>(undefined);
  const [rowSizeId, setRowSizeId] = React.useState<string>(DEFAULT_ROW_SIZE_ID);
  const [storeMessage, setStoreMessage] = React.useState<string | undefined>(undefined);

  const saveTimer = React.useRef<number | undefined>(undefined);
  /** Layout waiting out the debounce, kept so it can still be written if this goes away. */
  const pendingSave = React.useRef<IDashboardLayout | undefined>(undefined);
  const isMounted = React.useRef<boolean>(true);
  const widgetsRef = React.useRef<IWidgetInstance[]>(widgets);
  const presetRef = React.useRef<string | undefined>(undefined);
  const rowSizeRef = React.useRef<string>(DEFAULT_ROW_SIZE_ID);
  // The grid reports a breakpoint change and a layout change inside the same commit,
  // so this has to be a ref: a state update would still be batched when the layout
  // callback fires and the narrow layout would overwrite the authored one.
  const breakpointRef = React.useRef<string>('lg');

  React.useEffect(() => {
    widgetsRef.current = widgets;
  }, [widgets]);

  React.useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
    };
  }, []);

  React.useEffect(() => {
    // The store is replaced when the author repoints the dashboard id, so show the
    // spinner again rather than leaving the previous scope's widgets on screen.
    setIsLoading(true);
    store
      .load()
      .then((stored) => {
        if (!isMounted.current) {
          return;
        }
        const layout = stored ?? starterLayout;
        widgetsRef.current = layout.widgets;
        presetRef.current = getPreset(layout.presetId)?.id;
        rowSizeRef.current = getRowSize(layout.rowSizeId).id;
        setWidgets(layout.widgets);
        setPresetId(presetRef.current);
        setRowSizeId(rowSizeRef.current);
        setStoreMessage(store.getStatus().message);
        setIsLoading(false);
      })
      .catch(() => {
        if (!isMounted.current) {
          return;
        }
        setWidgets(starterLayout.widgets);
        setIsLoading(false);
      });
    // The store is created once per web part instance.
  }, [store]);

  /** Writes whatever the debounce is holding. Does nothing when nothing is pending. */
  const flushSave = React.useCallback(() => {
    const layout = pendingSave.current;
    if (layout === undefined) {
      return;
    }
    pendingSave.current = undefined;
    store
      .save(layout)
      .then(() => {
        if (isMounted.current) {
          setStoreMessage(store.getStatus().message);
        }
      })
      .catch(() => {
        /* the store already degraded to local storage */
      });
  }, [store]);

  const persist = React.useCallback(
    (next: IWidgetInstance[], preset: string | undefined, rowSize: string) => {
      pendingSave.current = {
        version: CURRENT_LAYOUT_VERSION,
        widgets: next,
        presetId: preset,
        rowSizeId: rowSize
      };
      if (saveTimer.current !== undefined) {
        window.clearTimeout(saveTimer.current);
      }
      saveTimer.current = window.setTimeout(flushSave, SAVE_DEBOUNCE_MS);
    },
    [flushSave]
  );

  // A rearrangement made inside the debounce window would otherwise be dropped when
  // the dashboard unmounts, or when the author repoints it at another dashboard id.
  React.useEffect(
    () => () => {
      if (saveTimer.current !== undefined) {
        window.clearTimeout(saveTimer.current);
        saveTimer.current = undefined;
      }
      flushSave();
    },
    [flushSave]
  );

  /**
   * Single write path. `preset` records which layout preset the widgets currently
   * match; passing undefined marks the arrangement as custom.
   */
  const commit = React.useCallback(
    (next: IWidgetInstance[], preset: string | undefined, rowSize?: string) => {
      widgetsRef.current = next;
      presetRef.current = preset;
      if (rowSize) {
        rowSizeRef.current = rowSize;
        setRowSizeId(rowSize);
      }
      setWidgets(next);
      setPresetId(preset);
      persist(next, preset, rowSizeRef.current);
    },
    [persist]
  );

  /** Geometry changed but the preset it came from still holds. */
  const update = React.useCallback(
    (next: IWidgetInstance[]) => commit(next, presetRef.current),
    [commit]
  );

  /** The user moved something by hand, so the layout no longer matches a preset. */
  const markCustom = React.useCallback(() => {
    if (presetRef.current !== undefined) {
      presetRef.current = undefined;
      setPresetId(undefined);
    }
  }, []);

  const handleLayoutChange = React.useCallback(
    (current: Layout[]) => {
      // Only the 12 column breakpoint is authoritative; narrower ones are derived.
      if (breakpointRef.current !== 'lg' || isLoading) {
        return;
      }
      const previous = widgetsRef.current;
      const next = applyGridLayout(previous, current);
      const changed = next.some((w, i) => {
        const before = previous[i];
        return w.x !== before.x || w.y !== before.y || w.w !== before.w || w.h !== before.h;
      });
      if (changed) {
        update(next);
      }
    },
    [isLoading, update]
  );

  const handleAdd = React.useCallback(
    (definition: IWidgetDefinition) => {
      const added: IWidgetInstance = {
        id: newInstanceId(),
        type: definition.type,
        x: 0,
        y: widgets.reduce((acc, w) => Math.max(acc, w.y + w.h), 0),
        w: definition.defaultSize.w,
        h: definition.defaultSize.h
      };
      const next = [...widgets, added];
      const preset = getPreset(presetRef.current);
      // A preset stays in force, so a new widget drops into the next free slot.
      update(preset ? applyPreset(next, preset, getRowSize(rowSizeRef.current)) : next);
      setIsCatalogueOpen(false);
    },
    [widgets, update]
  );

  const handleRemove = React.useCallback(
    (instanceId: string) => {
      const next = widgets.filter((w) => w.id !== instanceId);
      const preset = getPreset(presetRef.current);
      update(preset ? applyPreset(next, preset, getRowSize(rowSizeRef.current)) : next);
    },
    [widgets, update]
  );

  const handleApplyPreset = React.useCallback(
    (nextPresetId: string, nextRowSizeId: string) => {
      const preset = getPreset(nextPresetId);
      if (!preset) {
        return;
      }
      const rowSize = getRowSize(nextRowSizeId);
      commit(applyPreset(widgetsRef.current, preset, rowSize), preset.id, rowSize.id);
      setIsLayoutPanelOpen(false);
    },
    [commit]
  );

  /**
   * Keyboard equivalent of dragging. Vertical moves swap with the neighbour above or
   * below, because the grid compacts and would otherwise undo a single-row move.
   */
  const handleNudge = React.useCallback(
    (instanceId: string, nudge: INudge) => {
      const cols = COLS.lg;
      const current = widgetsRef.current;
      const target = current.filter((w) => w.id === instanceId)[0];
      if (!target) {
        return;
      }

      const definition = WidgetRegistry.get(target.type);
      const minW = definition?.minSize?.w ?? 1;
      const minH = definition?.minSize?.h ?? 2;

      const w = Math.max(minW, Math.min(cols, target.w + (nudge.dw ?? 0)));
      const h = Math.max(minH, target.h + (nudge.dh ?? 0));
      const x = Math.max(0, Math.min(cols - w, target.x + (nudge.dx ?? 0)));

      const resized: IWidgetInstance = { ...target, x, w, h };
      let next = current.map((item) => (item.id === instanceId ? resized : item));

      const dy = nudge.dy ?? 0;
      if (dy !== 0) {
        next = swapVertically(next, resized, dy) ?? next;
      }

      const changed = next.some((item, i) => {
        const before = current[i];
        return item.x !== before.x || item.y !== before.y || item.w !== before.w || item.h !== before.h;
      });
      if (changed) {
        commit(next, undefined);
      }
    },
    [commit]
  );

  const handleUpdateSettings = React.useCallback(
    (instanceId: string, settings: Record<string, unknown>) =>
      update(widgets.map((w) => (w.id === instanceId ? { ...w, settings } : w))),
    [widgets, update]
  );

  const handleReset = React.useCallback(() => {
    const next = starterLayout.widgets.map((w) => ({ ...w }));
    commit(next, getPreset(starterLayout.presetId)?.id, getRowSize(starterLayout.rowSizeId).id);
  }, [starterLayout, commit]);

  /**
   * A finished drag or resize is the moment the arrangement stops matching a preset.
   * Doing it here rather than in onLayoutChange means the compaction that follows
   * applying a preset does not immediately mark the layout custom.
   */
  const handleGestureStop = React.useCallback(
    (_layout: Layout[], oldItem: Layout, newItem: Layout) => {
      if (
        oldItem.x !== newItem.x ||
        oldItem.y !== newItem.y ||
        oldItem.w !== newItem.w ||
        oldItem.h !== newItem.h
      ) {
        markCustom();
      }
    },
    [markCustom]
  );

  const activePreset = getPreset(presetId);

  const gridLayouts = React.useMemo(() => toGridLayouts(widgets), [widgets]);

  const fluentTheme = React.useMemo(() => {
    if (!theme) {
      return undefined;
    }
    return createTheme({
      palette: theme.palette as IPartialTheme['palette'],
      semanticColors: theme.semanticColors as IPartialTheme['semanticColors'],
      isInverted: !!theme.isInverted
    });
  }, [theme]);

  if (isLoading) {
    return (
      <div className={styles.dashboard}>
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.large} label="Loading your dashboard…" />
        </div>
      </div>
    );
  }

  return (
    <ThemeProvider theme={fluentTheme} applyTo="none" className={styles.dashboard}>
      <div className={styles.toolbar}>
        <h2 className={styles.heading}>{title}</h2>
        {isEditing && (
          <span className={styles.layoutBadge}>
            Layout: {activePreset ? activePreset.name : 'Custom'}
          </span>
        )}
        {isEditing && <DefaultButton iconProps={{ iconName: 'Add' }} text="Add a widget" onClick={() => setIsCatalogueOpen(true)} />}
        {isEditing && (
          <DefaultButton
            iconProps={{ iconName: 'GridViewMedium' }}
            text="Choose a layout"
            onClick={() => setIsLayoutPanelOpen(true)}
          />
        )}
        {isEditing && <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Reset" onClick={handleReset} />}
        {isEditing ? (
          <PrimaryButton iconProps={{ iconName: 'CheckMark' }} text="Done" onClick={() => setIsEditing(false)} />
        ) : (
          <DefaultButton iconProps={{ iconName: 'Edit' }} text="Edit dashboard" onClick={() => setIsEditing(true)} />
        )}
      </div>

      {storeMessage && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
          {storeMessage}
        </MessageBar>
      )}

      {isEditing && breakpoint === 'lg' && (
        <MessageBar messageBarType={MessageBarType.info}>
          Drag a widget by its title bar, or focus a title bar with the Tab key and use the arrow keys to move it —
          hold Shift with the arrow keys to resize.
        </MessageBar>
      )}

      {isEditing && breakpoint !== 'lg' && (
        <MessageBar messageBarType={MessageBarType.info}>
          Widgets stack automatically on narrow screens. Rearrange on a wider window to change the saved layout.
        </MessageBar>
      )}

      {widgets.length === 0 ? (
        <div className={styles.empty}>
          <p>Your dashboard is empty.</p>
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            text="Add a widget"
            onClick={() => {
              setIsEditing(true);
              setIsCatalogueOpen(true);
            }}
          />
        </div>
      ) : (
        <ResponsiveGridLayout
          className="layout"
          layouts={gridLayouts}
          breakpoints={BREAKPOINTS}
          cols={COLS}
          rowHeight={ROW_HEIGHT}
          margin={[16, 16]}
          containerPadding={[0, 0]}
          isDraggable={isEditing}
          isResizable={isEditing}
          draggableHandle=".widget-drag-handle"
          compactType="vertical"
          onBreakpointChange={(next) => {
            breakpointRef.current = next;
            setBreakpoint(next);
          }}
          onLayoutChange={handleLayoutChange}
          onDragStop={handleGestureStop}
          onResizeStop={handleGestureStop}
          measureBeforeMount={false}
          useCSSTransforms={true}
        >
          {widgets.map((widget) => {
            const widgetContext: IWidgetContext = {
              instanceId: widget.id,
              settings: widget.settings ?? {},
              spContext,
              isEditing,
              updateSettings: (settings) => handleUpdateSettings(widget.id, settings)
            };
            return (
              <div key={widget.id}>
                <WidgetFrame
                  instance={widget}
                  widgetContext={widgetContext}
                  isEditing={isEditing}
                  onRemove={handleRemove}
                  onNudge={handleNudge}
                />
              </div>
            );
          })}
        </ResponsiveGridLayout>
      )}

      <AddWidgetPanel
        isOpen={isCatalogueOpen}
        onDismiss={() => setIsCatalogueOpen(false)}
        onAdd={handleAdd}
      />

      <LayoutPresetPanel
        isOpen={isLayoutPanelOpen}
        presetId={presetId}
        rowSizeId={rowSizeId}
        widgetCount={widgets.length}
        onDismiss={() => setIsLayoutPanelOpen(false)}
        onApply={handleApplyPreset}
      />
    </ThemeProvider>
  );
};
