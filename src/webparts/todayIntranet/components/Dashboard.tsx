import * as React from 'react';
import { Responsive, WidthProvider, Layout, Layouts } from 'react-grid-layout';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './Dashboard.module.scss';
import { WidgetFrame } from './WidgetFrame';
import { AddWidgetPanel } from './AddWidgetPanel';
import { IDashboardLayout, IWidgetInstance, CURRENT_LAYOUT_VERSION } from '../model/IDashboardLayout';
import { ILayoutStore } from '../services/ILayoutStore';
import { IWidgetContext, IWidgetDefinition } from '../widgets/IWidget';
import { WidgetRegistry } from '../widgets/WidgetRegistry';

const ResponsiveGridLayout = WidthProvider(Responsive);

const BREAKPOINTS: Record<string, number> = { lg: 1200, md: 996, sm: 768, xs: 480, xxs: 0 };
const COLS: Record<string, number> = { lg: 12, md: 8, sm: 6, xs: 4, xxs: 2 };
const ROW_HEIGHT: number = 56;
const SAVE_DEBOUNCE_MS: number = 800;

export interface IDashboardProps {
  title: string;
  spContext: WebPartContext;
  store: ILayoutStore;
  /** Layout seeded for users who have never arranged their dashboard. */
  starterLayout: IDashboardLayout;
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

export const Dashboard: React.FunctionComponent<IDashboardProps> = (props) => {
  const { title, spContext, store, starterLayout } = props;

  const [widgets, setWidgets] = React.useState<IWidgetInstance[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [isEditing, setIsEditing] = React.useState<boolean>(false);
  const [isCatalogueOpen, setIsCatalogueOpen] = React.useState<boolean>(false);
  const [breakpoint, setBreakpoint] = React.useState<string>('lg');
  const [storeMessage, setStoreMessage] = React.useState<string | undefined>(undefined);

  const saveTimer = React.useRef<number | undefined>(undefined);
  const isMounted = React.useRef<boolean>(true);
  const widgetsRef = React.useRef<IWidgetInstance[]>(widgets);

  React.useEffect(() => {
    widgetsRef.current = widgets;
  }, [widgets]);

  React.useEffect(() => {
    isMounted.current = true;
    return () => {
      isMounted.current = false;
      if (saveTimer.current !== undefined) {
        window.clearTimeout(saveTimer.current);
      }
    };
  }, []);

  React.useEffect(() => {
    store
      .load()
      .then((stored) => {
        if (!isMounted.current) {
          return;
        }
        setWidgets(stored ? stored.widgets : starterLayout.widgets);
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

  const persist = React.useCallback(
    (next: IWidgetInstance[]) => {
      if (saveTimer.current !== undefined) {
        window.clearTimeout(saveTimer.current);
      }
      saveTimer.current = window.setTimeout(() => {
        const layout: IDashboardLayout = { version: CURRENT_LAYOUT_VERSION, widgets: next };
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
      }, SAVE_DEBOUNCE_MS);
    },
    [store]
  );

  const update = React.useCallback(
    (next: IWidgetInstance[]) => {
      widgetsRef.current = next;
      setWidgets(next);
      persist(next);
    },
    [persist]
  );

  const handleLayoutChange = React.useCallback(
    (current: Layout[]) => {
      // Only the 12 column breakpoint is authoritative; narrower ones are derived.
      if (breakpoint !== 'lg' || isLoading) {
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
    [breakpoint, isLoading, update]
  );

  const handleAdd = React.useCallback(
    (definition: IWidgetDefinition) => {
      const maxY = widgets.reduce((acc, w) => Math.max(acc, w.y + w.h), 0);
      update([
        ...widgets,
        {
          id: newInstanceId(),
          type: definition.type,
          x: 0,
          y: maxY,
          w: definition.defaultSize.w,
          h: definition.defaultSize.h
        }
      ]);
      setIsCatalogueOpen(false);
    },
    [widgets, update]
  );

  const handleRemove = React.useCallback(
    (instanceId: string) => update(widgets.filter((w) => w.id !== instanceId)),
    [widgets, update]
  );

  const handleUpdateSettings = React.useCallback(
    (instanceId: string, settings: Record<string, unknown>) =>
      update(widgets.map((w) => (w.id === instanceId ? { ...w, settings } : w))),
    [widgets, update]
  );

  const handleReset = React.useCallback(
    () => update(starterLayout.widgets.map((w) => ({ ...w }))),
    [starterLayout, update]
  );

  const gridLayouts = React.useMemo(() => toGridLayouts(widgets), [widgets]);

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
    <div className={styles.dashboard}>
      <div className={styles.toolbar}>
        <h2 className={styles.heading}>{title}</h2>
        {isEditing && <DefaultButton iconProps={{ iconName: 'Add' }} text="Add a widget" onClick={() => setIsCatalogueOpen(true)} />}
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
          onBreakpointChange={setBreakpoint}
          onLayoutChange={handleLayoutChange}
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
    </div>
  );
};
