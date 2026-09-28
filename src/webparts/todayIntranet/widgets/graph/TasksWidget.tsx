import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData, IGraphDataResult } from './useGraphData';
import { NumberSetting, SettingsSurface, ToggleSetting, WidgetItemsView, booleanSetting, numberSetting } from '../content';
import { calendarDateKey } from './calendarData';
import {
  advanceTaskHistory, AggregatedTask, emptyTaskHistory, formatTaskDue, ITaskHistory, loadTaskSource, TASK_SOURCE_LABELS,
  TASK_SOURCES, TaskSource, TaskSourceData, taskDueDay, taskSources, TODO_URL
} from './taskData';
import {
  defaultTaskMetadata, emptyTaskDocument, normalizeTaskTags, reorderTask, sortTasks, TASK_PRIORITIES, TaskMetadata
} from './taskOrganization';
import { useTaskOrganization } from './useTaskOrganization';
import { widgetDataCacheKey } from '../data/WidgetDataCache';
import { downloadJson } from '../../common/downloadJson';
import styles from './TasksWidget.module.scss';

const DEFAULT_MAX_ITEMS: number = 6;
const MAX_ITEMS_BOUNDS = { min: 1, max: 20 };
export const TODO_LINK: IWidgetLink = { text: 'Open To Do', href: TODO_URL };
export const TASKS_DEFAULT_VIEW: WidgetItemsView = 'compact';
export const TASKS_VIEWS: WidgetItemsView[] = ['compact', 'list'];

function useTaskSource(context: IWidgetContext, source: TaskSource, enabled: boolean): IGraphDataResult<TaskSourceData> {
  return useGraphData(context, source === 'outlook' ? 'Mail.ReadBasic' : 'Tasks.Read',
    (client) => loadTaskSource(client, source), ['task-source-v1', source], enabled);
}

const InlineTags: React.FunctionComponent<{
  title: string; tags: string[]; disabled: boolean;
  update: (tags: string[]) => boolean; report: (message: string) => void;
}> = ({ title, tags, disabled, update, report }) => {
  const [editing, setEditing] = React.useState<{ original?: string; value: string }>();
  const input = React.useRef<HTMLInputElement>(null);
  const root = React.useRef<HTMLDivElement>(null);
  const cancelled = React.useRef(false);
  const restoreFocus = React.useRef<string>();
  const isEditing = !!editing;
  // Tags arrive as a fresh array whenever the task has no saved metadata, so depend on the values.
  const signature = tags.join('\u0000');
  const missing = editing?.original !== undefined && tags.indexOf(editing.original) < 0;
  React.useEffect(() => {
    if (isEditing) { input.current?.focus(); }
    else if (restoreFocus.current !== undefined) {
      const buttons = Array.from(root.current?.querySelectorAll<HTMLButtonElement>('button[data-tag]') ?? []);
      const match = buttons.filter((button) => button.dataset.tag === restoreFocus.current)[0];
      (match ?? root.current?.querySelector<HTMLButtonElement>('button[data-add]'))?.focus();
      restoreFocus.current = undefined;
    }
  }, [editing?.original, isEditing, signature]);
  React.useEffect(() => { if (disabled) { setEditing(undefined); } }, [disabled]);
  // Another Tasks tile can drop the tag being renamed; close the editor so the cell keeps its add button.
  React.useEffect(() => { if (missing) { setEditing(undefined); } }, [missing]);
  const commit = (keepOpen: boolean): void => {
    if (!editing || cancelled.current) { return; }
    const original = editing.original;
    const value = editing.value.trim().replace(/\s+/g, ' ');
    const close = (): void => { restoreFocus.current = original ?? ''; setEditing(undefined); };
    if (!value) { close(); return; }
    if (tags.some((tag) => tag !== original && tag.toLowerCase() === value.toLowerCase())) {
      // Saving would silently drop one of the two, so report it instead of losing the tag.
      report(`"${value}" is already a tag on ${title}.`);
      if (!keepOpen) { close(); }
      return;
    }
    const next = original === undefined ? [...tags, value] : tags.map((tag) => tag === original ? value : tag);
    if (update(next)) { restoreFocus.current = value; setEditing(undefined); }
  };
  const start = (original?: string): void => {
    cancelled.current = false;
    setEditing({ original, value: original ?? '' });
  };
  const editor = editing && <input ref={input} className={styles.tagInput} value={editing.value} disabled={disabled}
      aria-label={`${editing.original === undefined ? 'Add tag' : `Edit tag ${editing.original}`} for ${title}`}
      placeholder="Tag name" onChange={(event) => setEditing({ ...editing, value: event.target.value })}
      onBlur={() => commit(false)} onKeyDown={(event) => {
        event.stopPropagation();
        if (event.key === 'Enter') { event.preventDefault(); commit(true); }
        if (event.key === 'Escape') {
          event.preventDefault(); cancelled.current = true;
          restoreFocus.current = editing.original ?? ''; setEditing(undefined);
        }
      }} />;
  const renaming = editing?.original !== undefined && !missing;
  return <div ref={root} className={styles.tagList}>
    {tags.map((tag) => editing?.original === tag ? <React.Fragment key={tag.toLowerCase()}>{editor}</React.Fragment>
      : <span className={styles.chip} key={tag.toLowerCase()}>
        <button type="button" disabled={disabled} data-tag={tag} onClick={() => start(tag)}
          aria-label={`Edit tag ${tag} for ${title}`}>{tag}</button>
        <button type="button" disabled={disabled} onClick={() => {
          if (update(tags.filter((value) => value !== tag))) { restoreFocus.current = ''; }
        }} aria-label={`Remove tag ${tag} from ${title}`}><Icon iconName="Cancel" aria-hidden="true" /></button>
      </span>)}
    {editing && editing.original === undefined ? editor : !renaming && <button type="button" data-add="true"
      className={styles.addTag} disabled={disabled || tags.length >= 5}
      aria-label={`Add tag for ${title}`} onClick={() => start()}>
      <Icon iconName="Add" aria-hidden="true" /> Tag</button>}
  </div>;
};

export const TasksWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, MAX_ITEMS_BOUNDS);
  const dueOnly = booleanSetting(context, 'dueOnly', false);
  const configured = JSON.stringify(taskSources(context.settings.sources));
  const [sources, setSources] = React.useState<TaskSource[]>(() => taskSources(context.settings.sources));
  // `configured` is the value key of the setting, so re-derive from it rather than the raw bag.
  React.useEffect(() => { setSources(taskSources(JSON.parse(configured))); }, [configured]);
  const todo = useTaskSource(context, 'todo', sources.indexOf('todo') >= 0);
  const planner = useTaskSource(context, 'planner', sources.indexOf('planner') >= 0);
  const outlook = useTaskSource(context, 'outlook', sources.indexOf('outlook') >= 0);
  const states: Record<TaskSource, IGraphDataResult<TaskSourceData>> = { todo, planner, outlook };
  const organization = useTaskOrganization(context);
  const { state, store } = organization;
  const [error, setError] = React.useState<string>();
  const [announcement, setAnnouncement] = React.useState('');
  const [dragging, setDragging] = React.useState<string>();
  const [drop, setDrop] = React.useState<{ key: string; after: boolean }>();
  const drag = React.useRef<{ key: string; x: number; y: number; active: boolean }>();
  const table = React.useRef<HTMLTableElement>(null);
  const [today, setToday] = React.useState(() => calendarDateKey(new Date()));
  React.useEffect(() => {
    const timer = window.setInterval(() => setToday(calendarDateKey(new Date())), 60000);
    return () => window.clearInterval(timer);
  }, []);
  const scope = widgetDataCacheKey(context, 'task-history', []);
  React.useEffect(() => {
    drag.current = undefined; setDragging(undefined); setDrop(undefined); setError(undefined); setAnnouncement('');
  }, [scope, configured, dueOnly, maxItems, state.canEdit]);
  // Responses are folded into history as state, using React's "adjust state during
  // render" pattern: the transition is pure, and returns the same object once applied.
  const [storedHistory, setHistory] = React.useState<ITaskHistory>(() => emptyTaskHistory(scope));
  const received: Partial<Record<TaskSource, TaskSourceData>> = {};
  sources.forEach((source) => {
    const sourceState = states[source];
    if (sourceState.status === 'ready' && sourceState.data) { received[source] = sourceState.data; }
  });
  const history = advanceTaskHistory(storedHistory, scope, received);
  if (history !== storedHistory) { setHistory(history); }
  const tasks: AggregatedTask[] = [];
  sources.forEach((source) => {
    if (received[source]) { tasks.push(...history.tasks[source]); }
  });
  const ordered = sortTasks(tasks.filter((task) => !dueOnly || task.due), state.document);
  const visible = ordered.slice(0, maxItems);
  const incomplete = sources.some((source) => states[source].status !== 'ready' ||
    !states[source].data?.complete || states[source].refreshError || states[source].isRefreshing);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;
  const toggle = (source: TaskSource): void => {
    setSources((previous) => previous.indexOf(source) >= 0
      ? previous.length === 1 ? previous : previous.filter((item) => item !== source)
      : TASK_SOURCES.filter((item) => item === source || previous.indexOf(item) >= 0));
    drag.current = undefined; setDragging(undefined); setDrop(undefined);
  };
  const update = (task: AggregatedTask, changes: Partial<Pick<TaskMetadata, 'priority' | 'tags'>>): boolean => {
    if (!store || !state.canEdit) { setError('Task organization is not available for editing.'); return false; }
    try {
      const previous = state.document.tasks[task.key] ?? defaultTaskMetadata();
      const priority = changes.priority ?? previous.priority;
      const tags = normalizeTaskTags(changes.tags ?? previous.tags);
      if (priority === previous.priority && JSON.stringify(tags) === JSON.stringify(previous.tags)) {
        setError(undefined); return true;
      }
      store.edit({ version: 1, tasks: { ...state.document.tasks, [task.key]: {
        priority, tags, modified: store.revision(),
        ...(previous.priority === priority && previous.order !== undefined ? { order: previous.order } : {})
      } } });
      setError(undefined); return true;
    } catch (failure) { setError(failure instanceof Error ? failure.message : 'Could not save organization.'); return false; }
  };
  const move = (key: string, target: string, after: boolean): void => {
    if (!store || !state.canEdit) { setError('Task organization is not available for editing.'); return; }
    try {
      store.edit(reorderTask(state.document, ordered, key, target, after, store.revision()));
      setError(undefined);
      setAnnouncement(`Moved ${visible.find((task) => task.key === key)?.title} ${after ? 'after' : 'before'} ${
        visible.find((task) => task.key === target)?.title}; ${state.document.tasks[target]?.priority ?? 'normal'} priority.`);
    } catch (failure) { setError(failure instanceof Error ? failure.message : 'Could not move this task.'); }
  };
  const hitTarget = (x: number, y: number): { key: string; after: boolean } | undefined => {
    const row = document.elementFromPoint(x, y)?.closest<HTMLTableRowElement>('tr[data-task-key]');
    const key = row?.dataset.taskKey;
    if (!row || !key || !table.current?.contains(row) || key === drag.current?.key) { return undefined; }
    const rect = row.getBoundingClientRect();
    return { key, after: y >= rect.top + rect.height / 2 };
  };
  const cancelDrag = (): void => { drag.current = undefined; setDragging(undefined); setDrop(undefined); };

  return (
    <div className={`${styles.tasks} ${context.settings.view === 'list' ? styles.relaxed : ''}`}>
      <div className={styles.filters} role="group" aria-label="Task sources">
        <button type="button" className={sources.length === TASK_SOURCES.length ? styles.activeFilter : ''}
          aria-pressed={sources.length === TASK_SOURCES.length} onClick={() => { setSources([...TASK_SOURCES]); cancelDrag(); }}>All</button>
        {TASK_SOURCES.map((source) => (
          <button type="button" key={source} className={sources.indexOf(source) >= 0 ? styles.activeFilter : ''}
            aria-label={TASK_SOURCE_LABELS[source]} aria-pressed={sources.indexOf(source) >= 0} onClick={() => toggle(source)}>
            {sources.indexOf(source) >= 0 && <Icon iconName="CheckMark" aria-hidden="true" />} {TASK_SOURCE_LABELS[source]}
          </button>
        ))}
      </div>
      {sources.map((source) => {
        const current = states[source];
        const warning = current.error?.message ?? current.refreshError?.message ?? current.data?.warning;
        return warning ? <div className={styles.notice} key={source} role="status">
          <strong>{TASK_SOURCE_LABELS[source]}: </strong>{warning} Results may be incomplete or out of date.{' '}
          <button type="button" onClick={current.reload}>Retry {TASK_SOURCE_LABELS[source]}</button>
        </div> : current.status === 'loading' || current.isRefreshing
          ? <p key={source} className={styles.status} role="status">Loading {TASK_SOURCE_LABELS[source]}...</p> : null;
      })}
      <div className={styles.scroll} tabIndex={0} role="region" aria-label="Open tasks table">
        <table ref={table} className={styles.table}>
          <thead><tr><th scope="col" className={styles.number}><span className={styles.srOnly}>Order</span></th><th scope="col">Task name</th>
            <th scope="col">Quick look</th><th scope="col" className={styles.due}>Due date</th>
            <th scope="col" className={styles.tags}>Tags</th></tr></thead>
          <tbody>
            {visible.map((task, index) => {
              const metadata = state.document.tasks[task.key] ?? defaultTaskMetadata();
              const overdue = (taskDueDay(task) ?? today) < today;
              return <tr key={`${scope}:${task.key}`} data-task-key={task.key}
                className={`${dragging === task.key ? styles.dragging : ''} ${
                  drop?.key === task.key ? drop.after ? styles.dropAfter : styles.dropBefore : ''}`}>
                  <td className={styles.number}>
                    <button type="button" className={styles.dragHandle} disabled={!state.canEdit}
                      aria-label={`Move ${task.title}`} title="Drag to reorder; use Up/Down arrow keys. Adopts the destination priority."
                      onPointerDown={(event) => {
                        if (event.button !== 0 || !event.isPrimary) { return; }
                        event.stopPropagation();
                        event.currentTarget.focus();
                        event.currentTarget.setPointerCapture(event.pointerId);
                        drag.current = { key: task.key, x: event.clientX, y: event.clientY, active: false };
                      }}
                      onPointerMove={(event) => {
                        const current = drag.current;
                        if (!current || current.key !== task.key) { return; }
                        if (!current.active && Math.hypot(event.clientX - current.x, event.clientY - current.y) < 5) { return; }
                        current.active = true; setDragging(task.key);
                        setDrop(hitTarget(event.clientX, event.clientY));
                      }}
                      onPointerUp={(event) => {
                        if (drag.current?.active) {
                          const target = hitTarget(event.clientX, event.clientY);
                          if (target) { move(task.key, target.key, target.after); }
                        }
                        cancelDrag();
                      }}
                      onPointerCancel={cancelDrag} onLostPointerCapture={cancelDrag}
                      onKeyDown={(event) => {
                        if (event.key === 'Escape') { event.stopPropagation(); cancelDrag(); }
                        if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
                          event.preventDefault(); event.stopPropagation();
                          const after = event.key === 'ArrowDown';
                          const target = visible[index + (after ? 1 : -1)];
                          if (target) { move(task.key, target.key, after); }
                        }
                      }}>
                      <Icon iconName="GripperDotsVertical" aria-hidden="true" />
                    </button>
                  </td>
                  <td>
                    <div className={styles.taskName}>
                      <label className={styles.priority} title={`${metadata.priority} personal priority`}>
                        <Icon iconName={metadata.priority === 'high' ? 'FavoriteStarFill' : metadata.priority === 'low' ? 'ChevronDown' : 'FavoriteStar'} aria-hidden="true" />
                        <select aria-label={`Priority for ${task.title}`} value={metadata.priority} disabled={!state.canEdit}
                          onChange={(event) => {
                            const priority = TASK_PRIORITIES.find((value) => value === event.target.value);
                            if (priority) { update(task, { priority }); }
                          }}>
                          {TASK_PRIORITIES.map((value) => <option key={value} value={value}>{value}</option>)}
                        </select>
                      </label>
                      <span className={styles.title} title={task.title}>{task.title}</span>
                    </div>
                    <span className={`${styles.mobileDue} ${overdue ? styles.overdue : ''}`}>{formatTaskDue(task, locale)}{overdue ? ' - Overdue' : ''}</span>
                  </td>
                  <td>{task.sourceUrl ? <a href={task.sourceUrl} target="_blank" rel="noreferrer"
                    aria-label={`${task.linkLabel}: ${task.title}`}>{task.linkLabel} <Icon iconName="OpenInNewWindow" aria-hidden="true" /></a> : 'Link unavailable'}</td>
                  <td className={`${styles.due} ${overdue ? styles.overdue : ''}`}>{formatTaskDue(task, locale)}{overdue && <span className={styles.overdueLabel}>Overdue</span>}</td>
                  <td className={styles.tags}><InlineTags title={task.title} tags={metadata.tags} disabled={!state.canEdit}
                    update={(tags) => update(task, { tags })} report={setError} /></td>
                </tr>;
            })}
          </tbody>
        </table>
      </div>
      {!visible.length && <p role="status" className={styles.status}>{incomplete
        ? 'No tasks to show yet. Some sources are still loading or unavailable.'
        : dueOnly ? 'No open tasks with a due date in these sources.' : 'No open tasks in these sources.'}</p>}
      {visible.length < ordered.length && <p className={styles.status}>Showing {visible.length} of {ordered.length} loaded tasks. Adjust the item limit in tile settings.</p>}
      {error && <p role="alert" className={styles.notice}>{error}</p>}
      <span role="status" aria-live="polite" className={styles.srOnly}>{announcement}</span>
      {(state.status !== 'saved' || context.isEditing) && <details className={styles.organization}>
        <summary>{state.status === 'saved' ? 'Manage personal organization' : <span role="status">{state.message}</span>}</summary>
        <p>Priority, tags and order are personal to this dashboard experience, shared across Tasks tiles on this site web.
          Site owners can read stored organization. Source tasks are never changed.</p>
        <p>Source filters here are temporary. Choose starting sources in tile settings and save with Done.
          Loop task lists synchronized to Planner appear under Planner; separate Loop attribution is not available.</p>
        {store && <div className={styles.editorActions}>
          <button type="button" onClick={() => store.run(() => store.sync())}>Retry organization sync</button>
          <button type="button" onClick={() => {
            try { downloadJson(store.exportRecovery(), 'task-organization-recovery.json'); } catch (failure) {
              setError(`Export failed: ${failure instanceof Error ? failure.message : 'unknown error'}`);
            }
          }}>Export recovery data</button>
          {state.status === 'conflict' && <>
            <button type="button" onClick={() => store.run(() => store.resolveConflict('local'))}>Keep this browser organization</button>
            <button type="button" onClick={() => store.run(() => store.resolveConflict('remote'))}>Use SharePoint organization</button>
          </>}
          {state.status === 'error' && <button type="button" onClick={() => {
            if (window.confirm('Discard browser checkpoints and reload SharePoint organization? Export recovery data first.')) {
              store.run(() => store.clearBrowserRecovery());
            }
          }}>Clear browser checkpoints</button>}
          <button type="button" disabled={!state.canEdit} onClick={() => {
            if (window.confirm('Clear all personal priorities, tags and order, including retained tasks? Source tasks will not change.')) {
              try { store.edit(emptyTaskDocument()); } catch (failure) { setError(failure instanceof Error ? failure.message : 'Could not clear organization.'); }
            }
          }}>Clear organization</button>
        </div>}
      </details>}
    </div>
  );
};

export const TasksWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const selected = taskSources(context.settings.sources);
  return <SettingsSurface description="Only this tile changes. Task organization saves separately, without Done.">
    <NumberSetting context={context} settingKey="maxItems" label="Tasks to show" fallback={DEFAULT_MAX_ITEMS} {...MAX_ITEMS_BOUNDS} />
    <ToggleSetting context={context} settingKey="dueOnly" label="Only tasks with a due date" fallback={false} />
    <fieldset className={styles.settings}><legend>Starting sources</legend>
      {TASK_SOURCES.map((source) => <label key={source}>
        <input type="checkbox" checked={selected.indexOf(source) >= 0}
          disabled={selected.length === 1 && selected[0] === source}
          onChange={() => context.updateSettings({ ...context.settings, sources: selected.indexOf(source) >= 0
            ? selected.filter((value) => value !== source) : [...selected, source] })} />
        {TASK_SOURCE_LABELS[source]}
      </label>)}
    </fieldset>
    <p>To Do includes all lists except Flagged email. Outlook reads flagged messages across your primary mailbox, not shared mailboxes.
      Planner includes synchronized Loop tasks, without a separate Loop filter.</p>
  </SettingsSurface>;
};
