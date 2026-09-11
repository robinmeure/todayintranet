import * as React from 'react';
import { TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';
import { IWidgetContext, IWidgetLink } from '../IWidget';
import { useGraphData } from './useGraphData';
import {
  IWidgetListItem,
  NumberSetting,
  SettingsSurface,
  ToggleSetting,
  WidgetItems,
  WidgetItemsView,
  WidgetView,
  booleanSetting,
  numberSetting,
  viewSetting
} from '../content';

const SCOPE: string = 'Tasks.Read';
const DEFAULT_MAX_ITEMS: number = 6;
const MAX_ITEMS_BOUNDS = { min: 1, max: 20 };
const TODO_URL: string = 'https://to-do.office.com/tasks/';

export const TODO_LINK: IWidgetLink = { text: 'Open To Do', href: TODO_URL };
export const TASKS_DEFAULT_VIEW: WidgetItemsView = 'compact';

function dueDate(task: TodoTask): Date | undefined {
  return task.dueDateTime?.dateTime ? new Date(task.dueDateTime.dateTime) : undefined;
}

function formatDue(due: Date, today: Date, locale: string | undefined): string {
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  if (due.toDateString() === today.toDateString()) {
    return 'Due today';
  }
  if (due.toDateString() === tomorrow.toDateString()) {
    return 'Due tomorrow';
  }
  return `Due ${due.toLocaleDateString(locale, { day: 'numeric', month: 'short' })}`;
}

/** One task, described in the shared list vocabulary. */
function toItem(task: TodoTask, index: number, today: Date, locale: string | undefined): IWidgetListItem {
  const due = dueDate(task);
  const isOverdue = !!due && due < today;
  const isImportant = task.importance === 'high';

  return {
    key: task.id ?? String(index),
    title: task.title || '(Untitled task)',
    meta: [due ? formatDue(due, today, locale) : undefined],
    iconName: isImportant ? 'Important' : 'CircleRing',
    tone: isOverdue ? 'danger' : isImportant ? 'warning' : 'neutral',
    isEmphasized: isOverdue || isImportant,
    badge: isOverdue ? { text: 'Overdue', tone: 'danger' } : undefined
  };
}

export const TasksWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS, MAX_ITEMS_BOUNDS);
  const dueOnly = booleanSetting(context, 'dueOnly', false);
  const view = viewSetting(context, TASKS_DEFAULT_VIEW);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const state = useGraphData<TodoTask[]>(
    context,
    SCOPE,
    async (client) => {
      const lists = await client
        .api('/me/todo/lists')
        .version('v1.0')
        .select('id,wellknownListName,displayName')
        .get();

      const candidates = (lists.value ?? []) as TodoTaskList[];
      const defaultList =
        candidates.filter((list) => list.wellknownListName === 'defaultList')[0] ?? candidates[0];
      if (!defaultList?.id) {
        return [];
      }

      const response = await client
        .api(`/me/todo/lists/${defaultList.id}/tasks`)
        .version('v1.0')
        .filter("status ne 'completed'")
        .select('id,title,status,importance,dueDateTime')
        .top(maxItems + 10)
        .get();

      let tasks = (response.value ?? []) as TodoTask[];
      if (dueOnly) {
        tasks = tasks.filter((task) => !!task.dueDateTime?.dateTime);
      }

      // Due items first, oldest due date at the top, then everything else.
      tasks.sort((a, b) => {
        const aDue = a.dueDateTime?.dateTime ? Date.parse(a.dueDateTime.dateTime) : Number.MAX_SAFE_INTEGER;
        const bDue = b.dueDateTime?.dateTime ? Date.parse(b.dueDateTime.dateTime) : Number.MAX_SAFE_INTEGER;
        return aDue - bDue;
      });

      return tasks.slice(0, maxItems);
    },
    [maxItems, dueOnly]
  );

  return (
    <WidgetView
      state={state}
      loading={{ label: 'Loading your tasks…', rows: Math.min(maxItems, 4) }}
      empty={{
        iconName: 'CheckboxComposite',
        text: dueOnly ? 'Nothing with a due date. Enjoy it.' : 'Nothing on your task list.'
      }}
    >
      {(tasks) => {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        return (
          <WidgetItems
            view={view}
            ariaLabel="Open tasks"
            items={tasks.map((task, index) => toItem(task, index, today, locale))}
          />
        );
      }}
    </WidgetView>
  );
};

export const TasksWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface description="Only this tile changes. Everyone keeps their own settings.">
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Tasks to show"
      fallback={DEFAULT_MAX_ITEMS}
      {...MAX_ITEMS_BOUNDS}
    />
    <ToggleSetting context={context} settingKey="dueOnly" label="Only tasks with a due date" fallback={false} />
  </SettingsSurface>
);
