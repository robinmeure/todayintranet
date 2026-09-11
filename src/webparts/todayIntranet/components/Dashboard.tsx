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
import { LayoutStorageStatus } from './LayoutStorageStatus';
import { IDashboardLayout, IWidgetInstance, CURRENT_LAYOUT_VERSION } from '../model/IDashboardLayout';
import {
  GRID_COLUMNS,
  applyPreset,
  getPreset,
  getRowSize,
  DEFAULT_ROW_SIZE_ID
} from '../model/LayoutPresets';
import { ILayoutStore, ILayoutStoreStatus, LayoutStoreAction } from '../services/ILayoutStore';
import { IWidgetHostContext, IWidgetDefinition } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

const ResponsiveGridLayout = WidthProvider(Responsive);

const BREAKPOINTS: Record<string, number> = { lg: 1200, md: 996, sm: 768, xs: 480, xxs: 0 };
const COLS: Record<string, number> = { lg: GRID_COLUMNS, md: 8, sm: 6, xs: 4, xxs: 2 };
const ROW_HEIGHT: number = 56;

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
  const [storeStatus, setStoreStatus] = React.useState<ILayoutStoreStatus>(() => store.getStatus());
  const [operationError, setOperationError] = React.useState<string | undefined>();
  const [isResolving, setIsResolving] = React.useState<boolean>(false);
  const [confirmation, setConfirmation] = React.useState<'reset' | 'import-legacy' | undefined>();
  const session = React.useMemo(() => ({ store, active: true, loaded: false, resolving: false, saveSequence: 0 }), [store]);
  const currentSession = React.useRef(session);
  currentSession.current = session;
  const widgetsRef = React.useRef<IWidgetInstance[]>(widgets);
  const presetRef = React.useRef<string | undefined>(undefined);
  const rowSizeRef = React.useRef<string>(DEFAULT_ROW_SIZE_ID);
  // The grid reports a breakpoint change and a layout change inside the same commit,
  // so this has to be a ref: a state update would still be batched when the layout
  // callback fires and the narrow layout would overwrite the authored one.
  const breakpointRef = React.useRef<string>('lg');

  const applyLayout = React.useCallback((layout: IDashboardLayout) => {
    widgetsRef.current = layout.widgets;
    presetRef.current = getPreset(layout.presetId)?.id;
    rowSizeRef.current = getRowSize(layout.rowSizeId).id;
    setWidgets(layout.widgets);
    setPresetId(presetRef.current);
    setRowSizeId(rowSizeRef.current);
  }, []);

  const isCurrentSession = React.useCallback(
    () => session.active && currentSession.current === session,
    [session]
  );

  React.useEffect(() => {
    let cancelled = false;
    session.active = true;
    const isCurrent = (): boolean => !cancelled && isCurrentSession();
    const refreshStatus = (): void => {
      if (isCurrent()) {
        setStoreStatus(store.getStatus());
      }
    };
    setIsLoading(true);
    setIsEditing(false);
    setIsCatalogueOpen(false);
    setIsLayoutPanelOpen(false);
    setConfirmation(undefined);
    setOperationError(undefined);
    setIsResolving(false);
    const unsubscribe = store.subscribe(refreshStatus);
    refreshStatus();

    const load = async (): Promise<void> => {
      try {
        const stored = await store.load();
        if (isCurrent()) {
          applyLayout(stored ?? starterLayout);
        }
      } catch {
        if (isCurrent()) {
          // A starter is only a visual fallback, never an implicit recovery write.
          applyLayout(starterLayout);
          setOperationError('The saved layout could not be loaded. Use the recovery actions below.');
        }
      } finally {
        if (isCurrent()) {
          session.loaded = true;
          refreshStatus();
          setIsLoading(false);
        }
      }
    };
    load().catch(() => {
      if (isCurrent()) {
        session.loaded = true;
        setIsLoading(false);
        setOperationError('The saved layout could not be loaded. Use the recovery actions below.');
      }
    });

    return () => {
      cancelled = true;
      session.active = false;
      unsubscribe();
      store.dispose();
    };
    // A starter-layout property change must not reload an already open dashboard.
  }, [store, session, applyLayout, isCurrentSession]);

  /**
   * Single write path. `preset` records which layout preset the widgets currently
   * match; passing undefined marks the arrangement as custom.
   */
  const commit = React.useCallback(
    (next: IWidgetInstance[], preset: string | undefined, rowSize?: string) => {
      if (!session.active || !session.loaded || session.resolving || !store.getStatus().canEdit) {
        return;
      }
      widgetsRef.current = next;
      presetRef.current = preset;
      if (rowSize) {
        rowSizeRef.current = rowSize;
      }
      if (isCurrentSession()) {
        setWidgets(next);
        setPresetId(preset);
        setRowSizeId(rowSizeRef.current);
        setOperationError(undefined);
      }
      const sequence = ++session.saveSequence;
      const refresh = (failed: boolean): void => {
        if (isCurrentSession() && sequence === session.saveSequence) {
          setStoreStatus(store.getStatus());
          if (failed) {
            setOperationError('The layout could not be saved. Your changes may not survive closing this page.');
          }
        }
      };
      try {
        // save checkpoints locally before returning; Done publishes the checkpoint.
        store.save({
          version: CURRENT_LAYOUT_VERSION,
          widgets: next,
          presetId: preset,
          rowSizeId: rowSizeRef.current
        }).then(() => refresh(false), () => refresh(true));
      } catch {
        refresh(true);
      }
    },
    [store, session, isCurrentSession]
  );

  /** Geometry changed but the preset it came from still holds. */
  const update = React.useCallback(
    (next: IWidgetInstance[]) => commit(next, presetRef.current),
    [commit]
  );

  /** The user moved something by hand, so the layout no longer matches a preset. */
  const markCustom = React.useCallback(() => {
    if (session.active && session.loaded && !session.resolving && store.getStatus().canEdit && presetRef.current !== undefined) {
      presetRef.current = undefined;
      setPresetId(undefined);
    }
  }, [store, session]);

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
        y: widgetsRef.current.reduce((acc, w) => Math.max(acc, w.y + w.h), 0),
        w: definition.defaultSize.w,
        h: definition.defaultSize.h
      };
      const next = [...widgetsRef.current, added];
      const preset = getPreset(presetRef.current);
      // A preset stays in force, so a new widget drops into the next free slot.
      update(preset ? applyPreset(next, preset, getRowSize(rowSizeRef.current)) : next);
      setIsCatalogueOpen(false);
    },
    [update]
  );

  const handleRemove = React.useCallback(
    (instanceId: string) => {
      const next = widgetsRef.current.filter((w) => w.id !== instanceId);
      const preset = getPreset(presetRef.current);
      update(preset ? applyPreset(next, preset, getRowSize(rowSizeRef.current)) : next);
    },
    [update]
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
      update(widgetsRef.current.map((w) => (w.id === instanceId ? { ...w, settings } : w))),
    [update]
  );

  const handleUpdateTitle = React.useCallback(
    (instanceId: string, widgetTitle: string) =>
      update(widgetsRef.current.map((w) => (w.id === instanceId ? { ...w, title: widgetTitle } : w))),
    [update]
  );

  const handleReset = React.useCallback(() => {
    const next = starterLayout.widgets.map((w) => ({ ...w }));
    commit(next, getPreset(starterLayout.presetId)?.id, getRowSize(starterLayout.rowSizeId).id);
  }, [starterLayout, commit]);

  const executeAction = React.useCallback(async (action: LayoutStoreAction): Promise<void> => {
    if (!isCurrentSession() || session.resolving) {
      return;
    }
    const status = store.getStatus();
    if (action === 'reset' && status.actions.indexOf('reset') < 0 && status.canEdit) {
      handleReset();
      return;
    }
    if (status.actions.indexOf(action) < 0) {
      return;
    }
    setOperationError(undefined);
    if (action === 'export') {
      let url: string | undefined;
      const link = document.createElement('a');
      try {
        url = URL.createObjectURL(new Blob([store.exportRecovery()], { type: 'application/json' }));
        link.href = url;
        link.download = 'dashboard-layout-recovery.json';
        document.body.appendChild(link);
        link.click();
      } catch {
        setOperationError('Recovery data could not be exported. Please try again.');
      } finally {
        link.remove();
        if (url) {
          URL.revokeObjectURL(url);
        }
      }
      return;
    }
    session.resolving = true;
    ++session.saveSequence;
    setIsResolving(true);
    try {
      const layout = await store.resolve(
        action,
        action === 'reset' || action === 'use-remote' ? starterLayout : undefined
      );
      if (isCurrentSession()) {
        applyLayout(layout ?? starterLayout);
      }
    } catch {
      if (isCurrentSession()) {
        setOperationError('The recovery action could not be completed. Review the storage status, then retry or export recovery data.');
      }
    } finally {
      // eslint-disable-next-line require-atomic-updates -- Only one resolution can run per session.
      session.resolving = false;
      if (isCurrentSession()) {
        setStoreStatus(store.getStatus());
        setIsResolving(false);
      }
    }
  }, [store, session, isCurrentSession, handleReset, starterLayout, applyLayout]);

  const runAction = React.useCallback((action: LayoutStoreAction): void => {
    executeAction(action).catch(() => {
      if (isCurrentSession()) {
        setOperationError('The recovery action could not be completed. Please try again.');
      }
    });
  }, [executeAction, isCurrentSession]);

  const handleDone = React.useCallback((): void => {
    setIsEditing(false);
    // Closing edit mode first unmounts any open draft fields. Their layout-effect
    // cleanup checkpoints the final value locally before this task publishes it.
    window.setTimeout(() => {
      if (!isCurrentSession()) {
        return;
      }
      store.publish().catch(() => {
          if (isCurrentSession()) {
            setStoreStatus(store.getStatus());
            setOperationError('The local changes could not be saved to SharePoint. They remain in this browser.');
          }
        });
    }, 0);
  }, [store, isCurrentSession]);

  const requestAction = React.useCallback((action: LayoutStoreAction): void => {
    if (action === 'reset' || action === 'import-legacy') {
      setConfirmation(action);
    } else {
      runAction(action);
    }
  }, [runAction]);

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
  const canEdit = session.loaded && !isLoading && storeStatus.canEdit && !isResolving;
  const editing = isEditing && canEdit;

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

  if (isLoading || !session.loaded) {
    return (
      <div className={styles.dashboard}>
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.large} label="Loading your dashboard…" />
        </div>
        <LayoutStorageStatus
          status={session.loaded ? storeStatus : store.getStatus()}
          busy={true}
          onAction={requestAction}
          onConfirm={() => undefined}
          onCancel={() => undefined}
        />
      </div>
    );
  }

  return (
    <ThemeProvider theme={fluentTheme} applyTo="none" className={styles.dashboard}>
      <div className={styles.toolbar}>
        <h2 className={styles.heading}>{title}</h2>
        {editing && (
          <span className={styles.layoutBadge}>
            Layout: {activePreset ? activePreset.name : 'Custom'}
          </span>
        )}
        {editing && <DefaultButton iconProps={{ iconName: 'Add' }} text="Add a widget" onClick={() => setIsCatalogueOpen(true)} />}
        {editing && (
          <DefaultButton
            iconProps={{ iconName: 'GridViewMedium' }}
            text="Choose a layout"
            onClick={() => setIsLayoutPanelOpen(true)}
          />
        )}
        {editing && <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Reset" onClick={() => requestAction('reset')} />}
        {isEditing ? (
          <PrimaryButton iconProps={{ iconName: 'CheckMark' }} text="Done" onClick={handleDone} />
        ) : (
          <DefaultButton iconProps={{ iconName: 'Edit' }} text="Edit dashboard" disabled={!canEdit} onClick={() => setIsEditing(true)} />
        )}
      </div>

      <LayoutStorageStatus
        status={storeStatus}
        error={operationError}
        busy={isResolving}
        confirmation={confirmation}
        onAction={requestAction}
        onCancel={() => setConfirmation(undefined)}
        onConfirm={() => {
          const action = confirmation;
          setConfirmation(undefined);
          if (action) {
            runAction(action);
          }
        }}
      />

      {editing && breakpoint === 'lg' && (
        <MessageBar messageBarType={MessageBarType.info}>
          Drag a widget by its title bar, or focus a title bar with the Tab key and use the arrow keys to move it —
          hold Shift with the arrow keys to resize.
        </MessageBar>
      )}

      {editing && breakpoint !== 'lg' && (
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
            disabled={!canEdit}
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
          isDraggable={editing}
          isResizable={editing}
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
            const widgetContext: IWidgetHostContext = {
              instanceId: widget.id,
              settings: widget.settings ?? {},
              spContext,
              isEditing: editing,
              updateSettings: (settings) => handleUpdateSettings(widget.id, settings)
            };
            return (
              <div key={widget.id}>
                <WidgetFrame
                  instance={widget}
                  widgetContext={widgetContext}
                  isEditing={editing}
                  onRemove={handleRemove}
                  onNudge={handleNudge}
                  onUpdateTitle={handleUpdateTitle}
                />
              </div>
            );
          })}
        </ResponsiveGridLayout>
      )}

      <AddWidgetPanel
        isOpen={isCatalogueOpen && canEdit}
        existingTypes={widgets.map((widget) => widget.type)}
        onDismiss={() => setIsCatalogueOpen(false)}
        onAdd={handleAdd}
      />

      <LayoutPresetPanel
        isOpen={isLayoutPanelOpen && canEdit}
        presetId={presetId}
        rowSizeId={rowSizeId}
        widgetCount={widgets.length}
        onDismiss={() => setIsLayoutPanelOpen(false)}
        onApply={handleApplyPreset}
      />
    </ThemeProvider>
  );
};
