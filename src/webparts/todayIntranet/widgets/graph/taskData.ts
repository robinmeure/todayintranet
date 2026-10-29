import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { DateTimeTimeZone, Message, PlannerTask, TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';
import { calendarDateKey, parseCalendarDate } from './calendarData';
import { isWidgetDataAuthorizationError, widgetDataErrorStatus } from '../data/WidgetDataError';

export type TaskSource = 'todo' | 'planner' | 'outlook';
export const TASK_SOURCES: TaskSource[] = ['todo', 'planner', 'outlook'];
export const TASK_SOURCE_LABELS: Record<TaskSource, string> = { todo: 'To Do', planner: 'Planner', outlook: 'Outlook' };
export const TODO_URL: string = 'https://to-do.office.com/tasks/';
export const PLANNER_URL: string = 'https://planner.cloud.microsoft/';

export interface AggregatedTask {
  key: string;
  source: TaskSource;
  title: string;
  due?: DateTimeTimeZone;
  sourceUrl?: string;
  linkLabel: string;
}

export interface TaskSourceData {
  tasks: AggregatedTask[];
  complete: boolean;
  warning?: string;
  continuation?: TaskContinuation;
}

interface Page<T> { value: T[]; '@odata.nextLink'?: string }
type TaskRecord = Omit<TodoTask, 'dueDateTime'> & Pick<PlannerTask, 'percentComplete'> &
  Pick<Message, 'subject' | 'flag' | 'webLink'> & {
    dueDateTime?: TodoTask['dueDateTime'] | PlannerTask['dueDateTime'];
  };
interface Budget { requests: number; records: number }
interface TaskPageRequest {
  path: string;
  listId?: string;
  initial?: boolean;
  unfiltered?: boolean;
}
export interface TaskContinuation {
  pending: TaskPageRequest[];
  buffered: AggregatedTask[];
  visited: string[];
  discoveryComplete: boolean;
}
export const TASK_BATCH_SIZE: number = 50;
export const TASK_REQUEST_LIMIT: number = 12;
export const TASK_RECORD_LIMIT: number = 1000;

export function taskSources(value: unknown): TaskSource[] {
  if (!Array.isArray(value)) { return ['todo']; }
  const selected = TASK_SOURCES.filter((source) => value.indexOf(source) >= 0);
  return selected.length ? selected : ['todo'];
}

function id(value: string | undefined): string {
  if (!value) { throw new Error('Microsoft 365 returned a task without an identity. Refresh this source.'); }
  return encodeURIComponent(value);
}

function link(value: Message['webLink']): string | undefined {
  if (!value) { return undefined; }
  const url = new URL(value);
  if (url.protocol !== 'https:' || url.username || url.password) {
    throw new Error('Microsoft 365 returned an unsafe task link.');
  }
  return url.href;
}

function due(value: TodoTask['dueDateTime']): DateTimeTimeZone | undefined {
  if (!value?.dateTime) { return undefined; }
  if (!parseCalendarDate(value.dateTime.slice(0, 10))) {
    throw new Error('Microsoft 365 returned an invalid task due date.');
  }
  return { dateTime: value.dateTime, timeZone: value.timeZone ?? undefined };
}

export function taskDueDay(task: AggregatedTask): string | undefined {
  if (!task.due?.dateTime) { return undefined; }
  // To Do and mail flags represent a due calendar day in the supplied zone, not an instant.
  if (task.source !== 'planner') { return task.due.dateTime.slice(0, 10); }
  const date = new Date(task.due.dateTime);
  if (Number.isNaN(date.getTime())) { throw new Error('Planner returned an invalid due timestamp.'); }
  return calendarDateKey(date);
}

export function formatTaskDue(task: AggregatedTask, locale: string | undefined): string {
  const day = taskDueDay(task);
  return day ? parseCalendarDate(day)!.toLocaleDateString(locale, { day: '2-digit', month: 'short', year: 'numeric' }) : 'No due date';
}

async function pages<T>(
  client: MSGraphClientV3, path: string, fields: string | undefined, budget: Budget,
  receive: (items: T[]) => void, outlook: boolean = false, filter?: string
): Promise<boolean> {
  const visited = new Set<string>();
  let next: string | undefined = path;
  while (next) {
    if (budget.requests >= TASK_REQUEST_LIMIT || budget.records >= TASK_RECORD_LIMIT) { return false; }
    if (visited.has(next)) { throw new Error('Microsoft 365 returned a repeated task page.'); }
    visited.add(next);
    let request = client.api(next).version('v1.0');
    if (outlook) { request = request.header('Prefer', 'IdType="ImmutableId"'); }
    // To Do uses bare endpoints because its backend rejects our query options.
    // Never append options to server-provided pagination links.
    if (next === path && fields) {
      request = request.select(fields);
      // Planner does not document $top support for assigned tasks.
      if (path !== '/me/planner/tasks') { request = request.top(100); }
      if (filter) { request = request.filter(filter); }
    }
    budget.requests++;
    const response: Page<T> = await request.get();
    if (!response || !Array.isArray(response.value) ||
        (response['@odata.nextLink'] !== undefined && typeof response['@odata.nextLink'] !== 'string')) {
      throw new Error('Microsoft 365 returned an invalid task page.');
    }
    const available = TASK_RECORD_LIMIT - budget.records;
    receive(response.value.slice(0, available));
    budget.records += response.value.length;
    if (response.value.length > available) { return false; }
    next = response['@odata.nextLink'];
  }
  return true;
}

export async function loadTaskSource(
  client: MSGraphClientV3, source: TaskSource, continuation?: TaskContinuation
): Promise<TaskSourceData> {
  const tasks: AggregatedTask[] = [];
  const budget: Budget = { requests: 0, records: 0 };
  const pending = continuation ? [...continuation.pending] : [];
  const buffered = continuation ? [...continuation.buffered] : [];
  const visited = new Set(continuation?.visited);
  let discoveryComplete = continuation?.discoveryComplete ?? true;
  let complete = discoveryComplete;
  let warning: string | undefined;
  try {
    if (!continuation && source === 'todo') {
      const lists: TodoTaskList[] = [];
      complete = await pages<TodoTaskList>(client, '/me/todo/lists', undefined, budget, (items) => lists.push(...items));
      discoveryComplete = complete;
      for (const list of lists) {
        if (list.wellknownListName === 'flaggedEmails') { continue; }
        const listId = id(list.id);
        pending.push({ path: `/me/todo/lists/${listId}/tasks`, listId, initial: true });
      }
    } else if (!continuation) {
      pending.push({ path: source === 'planner' ? '/me/planner/tasks' : '/me/messages', initial: true });
    }
    tasks.push(...buffered.splice(0, TASK_BATCH_SIZE));
    let taskRecords = tasks.length;
    while (pending.length && taskRecords < TASK_BATCH_SIZE && budget.requests < TASK_REQUEST_LIMIT) {
      const current = pending[0];
      if (visited.has(current.path)) { throw new Error('Microsoft 365 returned a repeated task page.'); }
      let request = client.api(current.path).version('v1.0');
      if (source === 'outlook') {
        request = request.header('Prefer', 'IdType="ImmutableId"');
        if (current.initial) {
          request = request.select('id,subject,flag,webLink').top(TASK_BATCH_SIZE);
          if (!current.unfiltered) { request = request.filter("flag/flagStatus eq 'flagged'"); }
        }
      } else if (source === 'planner' && current.initial) {
        request = request.select('id,title,percentComplete,dueDateTime');
      }
      budget.requests++;
      let response: Page<TaskRecord>;
      try {
        response = await request.get();
      } catch (error) {
        if (source === 'outlook' && current.initial && !current.unfiltered &&
            widgetDataErrorStatus(error) === 400 && !tasks.length) {
          pending[0] = { ...current, unfiltered: true };
          continue;
        }
        throw error;
      }
      if (!response || !Array.isArray(response.value) ||
          (response['@odata.nextLink'] !== undefined && typeof response['@odata.nextLink'] !== 'string')) {
        throw new Error('Microsoft 365 returned an invalid task page.');
      }
      const mapped: AggregatedTask[] = [];
      response.value.forEach((task) => {
        if (source === 'todo') {
          if (task.status !== 'completed') {
            if (typeof task.dueDateTime === 'string') {
              throw new Error('Microsoft 365 returned an invalid task due date.');
            }
            mapped.push({ key: `todo:${current.listId}:${id(task.id)}`, source, title: task.title || '(Untitled task)',
              due: due(task.dueDateTime), sourceUrl: TODO_URL, linkLabel: 'To Do app' });
          }
        } else if (source === 'planner') {
          if (typeof task.percentComplete !== 'number') { throw new Error('Planner returned a task without completion status.'); }
          if (task.percentComplete < 100) {
            if (task.dueDateTime && typeof task.dueDateTime !== 'string') {
              throw new Error('Planner returned an invalid due timestamp.');
            }
            const item: AggregatedTask = { key: `planner:${id(task.id)}`, source, title: task.title || '(Untitled task)',
              due: task.dueDateTime ? due({ dateTime: task.dueDateTime, timeZone: 'UTC' }) : undefined,
              sourceUrl: PLANNER_URL, linkLabel: 'Planner app' };
            taskDueDay(item);
            mapped.push(item);
          }
        } else {
          const message = task;
          if (message.flag?.flagStatus === 'flagged') {
            mapped.push({ key: `outlook:${id(message.id)}`, source, title: message.subject || '(No subject)',
              due: due(message.flag.dueDateTime), sourceUrl: link(message.webLink), linkLabel: 'Open email' });
          }
        }
      });
      visited.add(current.path);
      pending.shift();
      if (response['@odata.nextLink']) {
        pending.unshift({ path: response['@odata.nextLink'], listId: current.listId });
      }
      const available = TASK_BATCH_SIZE - tasks.length;
      tasks.push(...mapped.slice(0, available));
      buffered.push(...mapped.slice(available));
      taskRecords += response.value.length;
    }
    if (pending.length && taskRecords < TASK_BATCH_SIZE && budget.requests >= TASK_REQUEST_LIMIT) {
      warning = `Results are incomplete (limit: ${TASK_REQUEST_LIMIT} requests per batch).`;
    }
  } catch (error) {
    if (isWidgetDataAuthorizationError(error) || (!tasks.length && !continuation)) { throw error; }
    complete = false;
    warning = error instanceof Error ? error.message : 'This source could not finish loading.';
  }
  const unique = new Map<string, AggregatedTask>();
  tasks.forEach((task) => unique.set(task.key, task));
  const next = pending.length || buffered.length
    ? { pending, buffered, visited: Array.from(visited), discoveryComplete } : undefined;
  return { tasks: Array.from(unique.values()), complete: complete && !next,
    continuation: next,
    warning: warning ?? (!complete ? `Results are incomplete (limit: ${TASK_REQUEST_LIMIT} requests / ${TASK_RECORD_LIMIT} records).` : undefined) };
}

export function reconcileTaskSource(previous: AggregatedTask[], next: TaskSourceData): AggregatedTask[] {
  if (next.complete) { return next.tasks; }
  const retained = new Map(previous.map((task) => [task.key, task]));
  next.tasks.forEach((task) => retained.set(task.key, task));
  return Array.from(retained.values());
}

/** Tasks shown per source, and the response each list was last reconciled with. */
export interface ITaskHistory {
  scope: string;
  applied: Partial<Record<TaskSource, TaskSourceData>>;
  tasks: Record<TaskSource, AggregatedTask[]>;
}

export function emptyTaskHistory(scope: string): ITaskHistory {
  return { scope, applied: {}, tasks: { todo: [], planner: [], outlook: [] } };
}

/**
 * Folds newly received source responses into the history. Pure: returns `history`
 * itself when nothing new arrived, so callers can store the result as state.
 * A different scope (another user or site) starts from an empty history.
 */
export function advanceTaskHistory(
  history: ITaskHistory,
  scope: string,
  received: Partial<Record<TaskSource, TaskSourceData>>
): ITaskHistory {
  let next = history.scope === scope ? history : emptyTaskHistory(scope);
  TASK_SOURCES.forEach((source) => {
    const data = received[source];
    if (data && next.applied[source] !== data) {
      next = {
        scope,
        applied: { ...next.applied, [source]: data },
        tasks: { ...next.tasks, [source]: reconcileTaskSource(next.tasks[source], data) }
      };
    }
  });
  return next;
}
