import type { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  AggregatedTask, loadTaskSource, reconcileTaskSource, TASK_RECORD_LIMIT, TASK_REQUEST_LIMIT, taskDueDay, taskSources
} from './taskData';

describe('read-only task sources', () => {
  type Request = ReturnType<MSGraphClientV3['api']>;
  const request = new (jest.fn<Request, []>())();
  const get = jest.fn<Promise<unknown>, []>();
  const api = jest.fn<Request, [string]>(() => request);
  const header = jest.fn<Request, [string, string]>(() => request);
  const filter = jest.fn<Request, [string]>(() => request);
  const top = jest.fn<Request, [number]>(() => request);
  const select = jest.fn<Request, [string]>(() => request);
  const client = new (jest.fn<MSGraphClientV3, []>())();
  Object.assign(client, { api });
  // There is deliberately no Graph write surface.
  Object.assign(request, { version: () => request, select, header, filter, top, get });
  beforeEach(() => { jest.clearAllMocks(); get.mockReset(); });

  it('discovers To Do lists without query options and follows the server next link unchanged', async () => {
    const nextLink = 'https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=opaque%2Btoken';
    get.mockImplementationOnce(async () => {
      expect(api).toHaveBeenLastCalledWith('/me/todo/lists');
      expect(select).not.toHaveBeenCalled();
      expect(top).not.toHaveBeenCalled();
      expect(filter).not.toHaveBeenCalled();
      return { value: [], '@odata.nextLink': nextLink };
    }).mockImplementationOnce(async () => {
      expect(api).toHaveBeenLastCalledWith(nextLink);
      expect(select).not.toHaveBeenCalled();
      expect(top).not.toHaveBeenCalled();
      expect(filter).not.toHaveBeenCalled();
      return { value: [{ id: 'list' }] };
    }).mockResolvedValueOnce({ value: [{ id: 'task', title: 'To Do task', status: 'notStarted' }] });
    const result = await loadTaskSource(client, 'todo');
    expect(result).toMatchObject({ complete: true, tasks: [{ key: 'todo:list:task' }] });
    expect(select).not.toHaveBeenCalled();
    expect(top).not.toHaveBeenCalled();
    expect(filter).not.toHaveBeenCalled();
  });

  it('reads bare To Do task endpoints and pages past completed tasks without adding query options', async () => {
    const listId = 'AQMk_example-list';
    const taskPath = `/me/todo/lists/${listId}/tasks`;
    const nextLink = `https://graph.microsoft.com/v1.0${taskPath}?$skiptoken=opaque%2Btoken`;
    get.mockResolvedValueOnce({ value: [{ id: listId }] })
      .mockImplementationOnce(async () => {
        expect(api).toHaveBeenLastCalledWith(taskPath);
        expect(select).not.toHaveBeenCalled();
        expect(top).not.toHaveBeenCalled();
        expect(filter).not.toHaveBeenCalled();
        return { value: [{ id: 'done', status: 'completed' }], '@odata.nextLink': nextLink };
      })
      .mockImplementationOnce(async () => {
        expect(api).toHaveBeenLastCalledWith(nextLink);
        expect(select).not.toHaveBeenCalled();
        expect(top).not.toHaveBeenCalled();
        expect(filter).not.toHaveBeenCalled();
        return { value: [{ id: 'open', title: 'Still open', status: 'notStarted',
          dueDateTime: { dateTime: '2026-09-29T00:00:00', timeZone: 'UTC' } }] };
      });
    const result = await loadTaskSource(client, 'todo');
    expect(result.complete).toBe(true);
    expect(result.tasks).toHaveLength(1);
    expect(result.tasks[0]).toMatchObject({ key: `todo:${listId}:open`, title: 'Still open',
      due: { dateTime: '2026-09-29T00:00:00', timeZone: 'UTC' } });
  });

  it('counts completed To Do tasks against the paging budget and reports incomplete results', async () => {
    let page = 0;
    get.mockResolvedValueOnce({ value: [{ id: 'list' }] })
      .mockImplementation(async () => ({
        value: [{ id: `done-${page}`, status: 'completed' }],
        '@odata.nextLink': `https://graph.microsoft.com/v1.0/me/todo/lists/list/tasks?$skiptoken=${++page}`
      }));
    const result = await loadTaskSource(client, 'todo');
    expect(get).toHaveBeenCalledTimes(TASK_REQUEST_LIMIT);
    expect(result.tasks).toEqual([]);
    expect(result.complete).toBe(false);
    expect(result.warning).toContain('incomplete');
  });

  it('pages all To Do lists and tasks, excludes flagged email, completed tasks and duplicate identities', async () => {
    get.mockResolvedValueOnce({ value: [{ id: 'first' }], '@odata.nextLink': '/lists-next' })
      .mockResolvedValueOnce({ value: [{ id: 'flagged', wellknownListName: 'flaggedEmails' }, { id: 'second' }] })
      .mockResolvedValueOnce({ value: [{ id: 'a', title: 'A' }], '@odata.nextLink': '/tasks-next' })
      .mockResolvedValueOnce({ value: [{ id: 'a', title: 'A updated' }, { id: 'done', status: 'completed' }] })
      .mockResolvedValueOnce({ value: [{ id: 'b', title: 'Same name', dueDateTime: { dateTime: '2026-09-01T00:00:00', timeZone: 'Pacific Standard Time' } }] });
    const result = await loadTaskSource(client, 'todo');
    expect(result.complete).toBe(true);
    expect(result.tasks.map((task) => task.key)).toEqual(['todo:first:a', 'todo:second:b']);
    expect(result.tasks[1].due?.timeZone).toBe('Pacific Standard Time');
    expect(api.mock.calls.map(([path]) => path)).toEqual([
      '/me/todo/lists', '/lists-next', '/me/todo/lists/first/tasks', '/tasks-next', '/me/todo/lists/second/tasks'
    ]);
    expect(top).not.toHaveBeenCalled();
    expect(result.tasks[0].linkLabel).toBe('To Do app');
  });

  it('uses only assigned Planner tasks, pages and filters complete work without guessing Loop attribution', async () => {
    get.mockResolvedValueOnce({ value: [{ id: 'loopish', title: 'Loop', percentComplete: 0 }], '@odata.nextLink': '/planner-next' })
      .mockResolvedValueOnce({ value: [{ id: 'done', percentComplete: 100 }] });
    const result = await loadTaskSource(client, 'planner');
    expect(result.tasks).toHaveLength(1);
    expect(result.tasks[0]).toMatchObject({ key: 'planner:loopish', source: 'planner', linkLabel: 'Planner app' });
    expect(api).toHaveBeenCalledWith('/me/planner/tasks');
    expect(top).not.toHaveBeenCalled();
  });

  it('reads mailbox-wide flags with immutable IDs on every page and does not invent due dates', async () => {
    get.mockResolvedValueOnce({ value: [{ id: 'a', flag: { flagStatus: 'flagged' }, receivedDateTime: '2026-01-01T00:00:00Z' }], '@odata.nextLink': '/mail-next' })
      .mockResolvedValueOnce({ value: [
        { id: 'b', flag: { flagStatus: 'complete' } },
        { id: 'c', flag: { flagStatus: 'flagged', dueDateTime: { dateTime: '2026-09-29T00:00:00', timeZone: 'UTC' } }, webLink: 'https://outlook.office.com/mail/test' }
      ] });
    const result = await loadTaskSource(client, 'outlook');
    expect(result.tasks.map((task) => task.key)).toEqual(['outlook:a', 'outlook:c']);
    expect(result.tasks[0].due).toBeUndefined();
    expect(header).toHaveBeenCalledTimes(2);
    expect(header).toHaveBeenCalledWith('Prefer', 'IdType="ImmutableId"');
    expect(filter).toHaveBeenCalledWith("flag/flagStatus eq 'flagged'");
    expect(select).toHaveBeenCalledWith('id,subject,flag,webLink');
    expect(top).toHaveBeenCalledWith(100);
    expect(api).toHaveBeenCalledWith('/me/messages');
  });

  it('falls back to bounded mailbox paging only when the server rejects filtering', async () => {
    get.mockRejectedValueOnce({ statusCode: 400 }).mockResolvedValueOnce({ value: [
      { id: 'flagged', flag: { flagStatus: 'flagged' } }, { id: 'not-flagged', flag: { flagStatus: 'notFlagged' } }
    ] });
    expect((await loadTaskSource(client, 'outlook')).tasks).toHaveLength(1);
    expect(filter).toHaveBeenCalledTimes(1);
    expect(api).toHaveBeenCalledTimes(2);
  });

  it('does not fall back or return old data on authorization failure', async () => {
    get.mockRejectedValueOnce({ statusCode: 403 });
    await expect(loadTaskSource(client, 'outlook')).rejects.toMatchObject({ statusCode: 403 });
    expect(api).toHaveBeenCalledTimes(1);
  });

  it('bounds requests exactly and marks truncated results incomplete', async () => {
    let page = 0;
    get.mockImplementation(async () => ({ value: [{ id: `${page}`, percentComplete: 0 }], '@odata.nextLink': `/next-${++page}` }));
    const result = await loadTaskSource(client, 'planner');
    expect(get).toHaveBeenCalledTimes(TASK_REQUEST_LIMIT);
    expect(result.tasks).toHaveLength(TASK_REQUEST_LIMIT);
    expect(result.complete).toBe(false);
    expect(result.warning).toContain('incomplete');
  });

  it('bounds records even when a server page exceeds the requested page size', async () => {
    get.mockResolvedValueOnce({ value: Array.from({ length: TASK_RECORD_LIMIT + 1 }, (_, index) => ({ id: String(index), percentComplete: 0 })) });
    const result = await loadTaskSource(client, 'planner');
    expect(result.tasks).toHaveLength(TASK_RECORD_LIMIT);
    expect(result.complete).toBe(false);
  });

  it('reports a mid-pagination failure without losing the successful page', async () => {
    get.mockResolvedValueOnce({ value: [{ id: 'a', percentComplete: 0 }], '@odata.nextLink': '/next' })
      .mockRejectedValueOnce(new Error('Throttled'));
    expect(await loadTaskSource(client, 'planner')).toMatchObject({ complete: false, warning: 'Throttled', tasks: [{ key: 'planner:a' }] });
  });

  it('rejects malformed and repeated pages rather than reporting a successful empty source', async () => {
    get.mockResolvedValueOnce({});
    await expect(loadTaskSource(client, 'todo')).rejects.toThrow('invalid task page');
    get.mockResolvedValue({ value: [], '@odata.nextLink': '/me/planner/tasks' });
    await expect(loadTaskSource(client, 'planner')).rejects.toThrow('repeated');
  });
});

describe('task reconciliation and calendar dates', () => {
  const task: AggregatedTask = { key: 'todo:list:id', source: 'todo', title: 'Keep me', linkLabel: 'To Do app',
    due: { dateTime: '2026-09-28T00:00:00', timeZone: 'Pacific Standard Time' } };
  it('preserves a due calendar day without treating it as UTC midnight', () => {
    expect(taskDueDay(task)).toBe('2026-09-28');
    const local = new Date(2026, 8, 28, 23);
    expect(taskDueDay({ ...task, source: 'planner', due: { dateTime: local.toISOString(), timeZone: 'UTC' } })).toBe('2026-09-28');
  });
  it('removes unseen source tasks only after complete success', () => {
    expect(reconcileTaskSource([task], { tasks: [], complete: false })).toEqual([task]);
    expect(reconcileTaskSource([task], { tasks: [], complete: true })).toEqual([]);
  });
  it('keeps upgraded tiles To Do-only and normalizes invalid filter settings', () => {
    expect(taskSources(undefined)).toEqual(['todo']);
    expect(taskSources(['loop', 'planner', 'planner'])).toEqual(['planner']);
    expect(taskSources([])).toEqual(['todo']);
  });
});
