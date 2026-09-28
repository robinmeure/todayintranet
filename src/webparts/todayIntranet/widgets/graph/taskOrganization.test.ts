import {
  defaultTaskMetadata, emptyTaskDocument, mergeTaskDocuments, normalizeTaskTags, parseTaskDocument,
  reorderTask, sortTasks, TASK_DOCUMENT_LIMIT, TaskDocument, TaskMetadata
} from './taskOrganization';
import { AggregatedTask } from './taskData';

const item = (extra?: Partial<TaskMetadata>): TaskMetadata => ({ ...defaultTaskMetadata(), modified: '1790588000000:device', ...extra });
const doc = (tasks: Record<string, TaskMetadata>): TaskDocument => ({ version: 1, tasks });
const task = (key: string, due?: string): AggregatedTask => ({
  key, source: 'todo', title: key, linkLabel: 'To Do app', due: due ? { dateTime: `${due}T00:00:00`, timeZone: 'UTC' } : undefined
});

describe('personal task organization', () => {
  it('normalizes tags without duplicating case variants', () => {
    expect(normalizeTaskTags([' This   week ', 'this week', '', 'Follow up'])).toEqual(['This week', 'Follow up']);
    expect(() => normalizeTaskTags(['x'.repeat(33)])).toThrow('32');
    expect(() => normalizeTaskTags(['a', 'b', 'c', 'd', 'e', 'f'])).toThrow('5');
  });
  it('validates versions, entries, revisions, size and priorities without retaining snapshots', () => {
    expect(() => parseTaskDocument('{"version":2,"tasks":{}}')).toThrow('unsupported');
    expect(() => parseTaskDocument(JSON.stringify(doc({ 'todo:a': item({ order: -1 }) })))).toThrow('Invalid');
    expect(() => parseTaskDocument(' '.repeat(TASK_DOCUMENT_LIMIT + 1))).toThrow('full');
    const parsed = parseTaskDocument(JSON.stringify({ version: 1, tasks: { 'todo:a': { ...item(), title: 'Secret source title' } } }));
    expect(JSON.stringify(parsed)).not.toContain('Secret');
  });
  it('enforces the exact document character and entry limits', () => {
    const empty = JSON.stringify(emptyTaskDocument());
    expect(parseTaskDocument(empty + ' '.repeat(TASK_DOCUMENT_LIMIT - empty.length))).toEqual(emptyTaskDocument());
    expect(() => parseTaskDocument(empty + ' '.repeat(TASK_DOCUMENT_LIMIT - empty.length + 1))).toThrow('full');
    const tasks: Record<string, TaskMetadata> = {};
    for (let index = 0; index < 500; index++) { tasks[`todo:${index}`] = item(); }
    expect(Object.keys(parseTaskDocument(JSON.stringify(doc(tasks))).tasks)).toHaveLength(500);
    tasks['todo:extra'] = item();
    expect(() => parseTaskDocument(JSON.stringify(doc(tasks)))).toThrow('500');
  });
  it('orders by personal priority, explicit position, due day, then identity; undated items follow dated items', () => {
    const metadata = doc({ 'todo:high': item({ priority: 'high' }), 'todo:manual': item({ order: 0 }) });
    expect(sortTasks([task('todo:undated'), task('todo:due', '2026-09-01'), task('todo:high'), task('todo:manual')], metadata)
      .map((value) => value.key)).toEqual(['todo:high', 'todo:manual', 'todo:due', 'todo:undated']);
  });
  it('merges disjoint changes and deterministically picks the newest competing task edit', () => {
    const local = doc({ 'todo:a': item({ tags: ['local'] }) });
    const remote = doc({ 'todo:b': item({ priority: 'high' }) });
    expect(Object.keys(mergeTaskDocuments(emptyTaskDocument(), local, remote).tasks)).toHaveLength(2);
    const newer = doc({ 'todo:a': item({ tags: ['new'], modified: '1790588000001:device' }) });
    expect(mergeTaskDocuments(emptyTaskDocument(), local, newer).tasks['todo:a'].tags).toEqual(['new']);
  });
  it('requires a choice for concurrent ordering changes in the same priority group', () => {
    const base = doc({ 'todo:a': item({ order: 0 }), 'todo:b': item({ order: 1 }) });
    const local = doc({ 'todo:a': item({ order: 1 }), 'todo:b': item({ order: 0 }) });
    const remote = doc({ 'todo:a': item({ order: 2 }), 'todo:b': item({ order: 1 }) });
    expect(() => mergeTaskDocuments(base, local, remote)).toThrow('reordered');
  });
  it('preserves explicit deletion when the remote task was not concurrently edited', () => {
    const base = doc({ 'todo:a': item() });
    expect(mergeTaskDocuments(base, emptyTaskDocument(), base)).toEqual(emptyTaskDocument());
  });
  it('inserts dragged tasks without swapping intervening rows and preserves hidden metadata', () => {
    const base = doc({ 'todo:hidden': item({ order: 1, tags: ['keep'] }), 'todo:a': item({ order: 0 }),
      'todo:b': item({ order: 2 }), 'todo:c': item({ order: 3 }), 'todo:low': item({ priority: 'low' }) });
    const tasks = [task('todo:a'), task('todo:b'), task('todo:c')];
    const moved = reorderTask(base, tasks, 'todo:c', 'todo:a', false, '1790588000001:device');
    expect(sortTasks(tasks, moved).map((value) => value.key)).toEqual(['todo:c', 'todo:a', 'todo:b']);
    expect(moved.tasks['todo:hidden'].tags).toEqual(['keep']);
    expect(moved.tasks['todo:hidden'].order).toBe(2);
    expect(moved.tasks['todo:low']).toEqual(base.tasks['todo:low']);
    const down = reorderTask(moved, tasks, 'todo:c', 'todo:b', true, '1790588000002:device');
    expect(sortTasks(tasks, down).map((value) => value.key)).toEqual(['todo:a', 'todo:b', 'todo:c']);
  });
  it('adopts the drop target priority while preserving tags and source group order', () => {
    const base = doc({ 'todo:a': item({ priority: 'high', order: 0 }), 'todo:b': item({ priority: 'low', tags: ['keep'], order: 0 }),
      'todo:c': item({ priority: 'low', order: 1 }) });
    const tasks = [task('todo:a'), task('todo:b'), task('todo:c')];
    const moved = reorderTask(base, tasks, 'todo:b', 'todo:a', false, '1790588000001:device');
    expect(moved.tasks['todo:b']).toMatchObject({ priority: 'high', tags: ['keep'], order: 0 });
    expect(moved.tasks['todo:c']).toEqual(base.tasks['todo:c']);
    expect(sortTasks(tasks, moved).map((value) => value.key)).toEqual(['todo:b', 'todo:a', 'todo:c']);
    expect(reorderTask(base, tasks, 'todo:a', 'todo:a', false, '1790588000001:device')).toBe(base);
    expect(() => reorderTask(base, tasks, 'todo:missing', 'todo:a', false, '1790588000001:device')).toThrow('no longer available');
  });
  it('only renumbers the rows a drag actually shifts', () => {
    const base = doc({ 'todo:a': item({ order: 0 }), 'todo:b': item({ order: 1 }),
      'todo:c': item({ order: 2 }), 'todo:d': item({ order: 3, tags: ['keep'] }) });
    const tasks = [task('todo:a'), task('todo:b'), task('todo:c'), task('todo:d')];
    const moved = reorderTask(base, tasks, 'todo:c', 'todo:b', false, '1790588000009:device');
    expect(sortTasks(tasks, moved).map((value) => value.key)).toEqual(['todo:a', 'todo:c', 'todo:b', 'todo:d']);
    // Rows outside the moved span keep their revision so a concurrent edit elsewhere still wins the merge.
    expect(moved.tasks['todo:a']).toEqual(base.tasks['todo:a']);
    expect(moved.tasks['todo:d']).toEqual(base.tasks['todo:d']);
    expect(moved.tasks['todo:b'].modified).toBe('1790588000009:device');
  });
  it('keeps a drag inside the document limit for a large priority group', () => {
    const day = (index: number): string => new Date(Date.UTC(2026, 0, 1 + index)).toISOString().slice(0, 10);
    const keys = Array.from({ length: 400 }, (_, index) => `todo:${'k'.repeat(240)}${index}`);
    const tasks = keys.map((key, index) => task(key, day(index)));
    const moved = reorderTask(emptyTaskDocument(), tasks, keys[399], keys[3], false, '1790588000009:device');
    // Only the rows ahead of the drop target need explicit positions; the tail stays due-sorted.
    expect(Object.keys(moved.tasks).length).toBe(4);
    expect(JSON.stringify(moved).length).toBeLessThan(TASK_DOCUMENT_LIMIT);
    expect(sortTasks(tasks, moved).slice(0, 5).map((value) => value.key))
      .toEqual([keys[0], keys[1], keys[2], keys[399], keys[3]]);
  });
});
