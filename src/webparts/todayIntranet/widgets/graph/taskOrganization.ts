import { AggregatedTask, taskDueDay } from './taskData';

export type TaskPriority = 'high' | 'normal' | 'low';
export const TASK_PRIORITIES: TaskPriority[] = ['high', 'normal', 'low'];
export const TASK_DOCUMENT_LIMIT: number = 60000;
export interface TaskMetadata {
  priority: TaskPriority;
  tags: string[];
  order?: number;
  modified: string;
}
export interface TaskDocument { version: 1; tasks: Record<string, TaskMetadata> }
export const emptyTaskDocument = (): TaskDocument => ({ version: 1, tasks: {} });
export const defaultTaskMetadata = (): TaskMetadata => ({ priority: 'normal', tags: [], modified: '' });

export function normalizeTaskTags(tags: string[]): string[] {
  const result: string[] = [];
  tags.forEach((raw) => {
    const tag = raw.trim().replace(/\s+/g, ' ');
    if (tag.length > 32) { throw new Error('Each tag must be 32 characters or fewer.'); }
    if (tag && !result.some((existing) => existing.toLowerCase() === tag.toLowerCase())) { result.push(tag); }
  });
  if (result.length > 5) { throw new Error('Use at most 5 tags per task.'); }
  return result;
}

export function parseTaskDocument(json: string): TaskDocument {
  if (json.length > TASK_DOCUMENT_LIMIT) { throw new Error('Task organization is full. Export and clear unused entries before saving.'); }
  const value: unknown = JSON.parse(json);
  if (!value || typeof value !== 'object' || Array.isArray(value)) { throw new Error('Invalid task organization document.'); }
  const document = value as Record<string, unknown>;
  if (document.version !== 1 || !document.tasks || typeof document.tasks !== 'object' || Array.isArray(document.tasks)) {
    throw new Error('Task organization uses an invalid or unsupported format.');
  }
  const entries = document.tasks as Record<string, unknown>;
  if (Object.keys(entries).length > 500) { throw new Error('Task organization supports at most 500 entries. Export and clear unused entries.'); }
  const result = emptyTaskDocument();
  Object.keys(entries).sort().forEach((key) => {
    const raw = entries[key];
    if (!/^(todo|planner|outlook):.+$/.test(key) || !raw || typeof raw !== 'object' || Array.isArray(raw)) {
      throw new Error('Invalid task organization entry.');
    }
    const item = raw as Record<string, unknown>;
    if (!TASK_PRIORITIES.some((priority) => priority === item.priority) ||
        !Array.isArray(item.tags) || !item.tags.every((tag): tag is string => typeof tag === 'string') ||
        typeof item.modified !== 'string' || !/^\d{13}:[a-zA-Z0-9-]+$/.test(item.modified) ||
        (item.order !== undefined && (typeof item.order !== 'number' || !Number.isSafeInteger(item.order) || item.order < 0))) {
      throw new Error('Invalid task priority, tags, order or revision.');
    }
    result.tasks[key] = { priority: item.priority as TaskPriority, tags: normalizeTaskTags(item.tags),
      modified: item.modified, ...(item.order === undefined ? {} : { order: item.order as number }) };
  });
  return result;
}

function equal(a: unknown, b: unknown): boolean { return JSON.stringify(a) === JSON.stringify(b); }

function orderGroup(document: TaskDocument, priority: TaskPriority): string {
  return JSON.stringify(Object.keys(document.tasks).sort().filter((key) =>
    document.tasks[key].priority === priority && document.tasks[key].order !== undefined
  ).map((key) => [key, document.tasks[key].order]));
}

export class TaskOrderConflict extends Error {
  public constructor() {
    super('Another device reordered the same priority group. Choose which organization to keep.');
    Object.setPrototypeOf(this, TaskOrderConflict.prototype);
  }
}

export function mergeTaskDocuments(base: TaskDocument, local: TaskDocument, remote: TaskDocument): TaskDocument {
  TASK_PRIORITIES.forEach((priority) => {
    const before = orderGroup(base, priority);
    const ours = orderGroup(local, priority);
    const theirs = orderGroup(remote, priority);
    if (ours !== before && theirs !== before && ours !== theirs) { throw new TaskOrderConflict(); }
  });
  const result = emptyTaskDocument();
  const keys = new Set([...Object.keys(base.tasks), ...Object.keys(local.tasks), ...Object.keys(remote.tasks)]);
  keys.forEach((key) => {
    const before = base.tasks[key];
    const ours = local.tasks[key];
    const theirs = remote.tasks[key];
    const chosen = equal(before, ours) ? theirs : equal(before, theirs) ? ours :
      !ours ? theirs : !theirs ? ours : ours.modified > theirs.modified ? ours : theirs;
    if (chosen) { result.tasks[key] = chosen; }
  });
  return parseTaskDocument(JSON.stringify(result));
}

export function sortTasks(tasks: AggregatedTask[], document: TaskDocument): AggregatedTask[] {
  return [...tasks].sort((a, b) => {
    const first = document.tasks[a.key] ?? defaultTaskMetadata();
    const second = document.tasks[b.key] ?? defaultTaskMetadata();
    const priority = TASK_PRIORITIES.indexOf(first.priority) - TASK_PRIORITIES.indexOf(second.priority);
    if (priority) { return priority; }
    // Explicit positions lead the untouched (due-sorted) tail of each priority group.
    const order = (first.order ?? Infinity) - (second.order ?? Infinity);
    if (order && !Number.isNaN(order)) { return order; }
    const firstDay = taskDueDay(a) ?? '9999-12-31';
    const secondDay = taskDueDay(b) ?? '9999-12-31';
    return firstDay < secondDay ? -1 : firstDay > secondDay ? 1 : a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
  });
}

export function reorderTask(
  document: TaskDocument, tasks: AggregatedTask[], key: string, target: string, after: boolean, modified: string
): TaskDocument {
  if (key === target) { return document; }
  if (!tasks.some((task) => task.key === key) || !tasks.some((task) => task.key === target)) {
    throw new Error('The task is no longer available. Try moving it again.');
  }
  const priority = document.tasks[target]?.priority ?? 'normal';
  const result: TaskDocument = { version: 1, tasks: { ...document.tasks } };
  result.tasks[key] = { ...(document.tasks[key] ?? defaultTaskMetadata()), priority, modified };
  delete result.tasks[key].order;
  // Include filtered and retained tasks so moving a visible row never discards their positions.
  const known = new Set(tasks.map((task) => task.key));
  const hidden: AggregatedTask[] = Object.keys(document.tasks).filter((savedKey) => !known.has(savedKey))
    .map((savedKey) => ({ key: savedKey, source: 'todo', title: '', linkLabel: '' }));
  const sequence = sortTasks([...tasks, ...hidden], document)
    .filter((task) => task.key !== key && (document.tasks[task.key]?.priority ?? 'normal') === priority)
    .map((task) => task.key);
  const index = sequence.indexOf(target) + (after ? 1 : 0);
  sequence.splice(index, 0, key);
  // Renumber only the rows that actually shift. Positions already above the running counter are
  // kept as they are, so a drag never rewrites (or restamps) an untouched task.
  let running = -1;
  for (let position = 0; position < sequence.length; position++) {
    const taskKey = sequence[position];
    const current = taskKey === key ? undefined : document.tasks[taskKey]?.order;
    if (current !== undefined && current > running) {
      running = current;
      if (position > index) { break; }
      continue;
    }
    // Tasks without a position sort after every ordered row, so the tail needs no entries at all.
    if (position > index && current === undefined) { break; }
    running += 1;
    result.tasks[taskKey] = { ...(result.tasks[taskKey] ?? defaultTaskMetadata()), order: running, modified };
  }
  return parseTaskDocument(JSON.stringify(result));
}
