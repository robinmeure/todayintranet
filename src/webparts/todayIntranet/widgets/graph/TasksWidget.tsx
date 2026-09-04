import * as React from 'react';
import { TodoTask, TodoTaskList } from '@microsoft/microsoft-graph-types';
import { Icon } from '@fluentui/react/lib/Icon';
import { IWidgetContext } from '../IWidget';
import { useGraphData } from './useGraphData';
import { WidgetLoading, WidgetEmpty, WidgetError } from '../WidgetMessage';
import { NumberSetting, ToggleSetting, SettingsSurface, numberSetting, booleanSetting } from './settings';
import styles from '../WidgetMessage.module.scss';

const SCOPE: string = 'Tasks.Read';
const DEFAULT_MAX_ITEMS: number = 6;

function formatDue(task: TodoTask, locale: string | undefined): string | undefined {
  if (!task.dueDateTime?.dateTime) {
    return undefined;
  }
  const due = new Date(task.dueDateTime.dateTime);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const label = due.toLocaleDateString(locale, { day: 'numeric', month: 'short' });
  return due < today ? `Overdue · ${label}` : `Due ${label}`;
}

export const TasksWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const maxItems = numberSetting(context, 'maxItems', DEFAULT_MAX_ITEMS);
  const dueOnly = booleanSetting(context, 'dueOnly', false);
  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  const result = useGraphData<TodoTask[]>(
    context.spContext,
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

  if (result.status === 'loading') {
    return <WidgetLoading label="Loading your tasks…" />;
  }
  if (result.status === 'error' && result.error) {
    return <WidgetError error={result.error} onRetry={result.reload} />;
  }
  if (!result.data || result.data.length === 0) {
    return <WidgetEmpty iconName="CheckboxComposite" text="Nothing on your task list." />;
  }

  return (
    <ul className={styles.list}>
      {result.data.map((task) => {
        const due = formatDue(task, locale);
        return (
          <li key={task.id} className={styles.item}>
            <Icon
              iconName={task.importance === 'high' ? 'Important' : 'CircleRing'}
              className={styles.itemIcon}
              aria-hidden="true"
            />
            <span className={styles.itemText}>
              <span className={styles.itemTitle}>{task.title || '(Untitled task)'}</span>
              {due && <span className={styles.itemMeta}>{due}</span>}
            </span>
          </li>
        );
      })}
    </ul>
  );
};

export const TasksWidgetSettings: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => (
  <SettingsSurface>
    <NumberSetting
      context={context}
      settingKey="maxItems"
      label="Tasks to show"
      fallback={DEFAULT_MAX_ITEMS}
      min={1}
      max={20}
    />
    <ToggleSetting context={context} settingKey="dueOnly" label="Only tasks with a due date" fallback={false} />
  </SettingsSurface>
);
