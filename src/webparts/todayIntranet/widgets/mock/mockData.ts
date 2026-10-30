import type { Event, Message } from '@microsoft/microsoft-graph-types';
import type { TaskSource, TaskSourceData } from '../graph/taskData';
import { TASK_SOURCE_LABELS } from '../graph/taskData';
import { calendarDateKey } from '../graph/calendarData';
import type { IEngageCommunityCard } from '../graph/vivaEngageData';
import type { ISearchResults } from '../search/SearchService';
import mail from './mail.json';
import calendar from './calendar.json';
import tasks from './tasks.json';
import community from './vivaEngage.json';
import news from './news.json';
import search from './search.json';

export function mockMail(limit: number, unreadOnly: boolean): Message[] {
  return mail.filter((message) => !unreadOnly || !message.isRead).slice(0, limit).map((message, index) => ({
    ...message, receivedDateTime: new Date(Date.now() - index * 3600000).toISOString()
  }));
}

/** Place the fixture inside the selected window, including late-today and historical views. */
export function mockCalendar(start: string, end: string, limit?: number): Event[] {
  const first = new Date(start).getTime();
  const last = new Date(end).getTime();
  const anchor = Math.max(first, Math.min(Date.now(), last - 3600000));
  return calendar.slice(0, limit).map((event) => {
    const begins = Math.min(anchor + event.offsetMinutes * 60000, last - 120000);
    return {
      id: event.id, subject: event.subject, location: event.location,
      start: { dateTime: new Date(begins).toISOString(), timeZone: 'UTC' },
      end: { dateTime: new Date(Math.min(begins + 3600000, last)).toISOString(), timeZone: 'UTC' }
    };
  });
}

export function mockTasks(source: TaskSource): TaskSourceData {
  return {
    complete: true,
    tasks: tasks.map((task) => {
      const due = new Date();
      due.setDate(due.getDate() + task.dueDays);
      return {
        key: `mock-${source}-${task.id}`, source, title: task.title,
        due: { dateTime: `${calendarDateKey(due)}T12:00:00Z`, timeZone: 'UTC' },
        linkLabel: TASK_SOURCE_LABELS[source]
      };
    })
  };
}

export function mockCommunity(limit: number): IEngageCommunityCard {
  return {
    ...community,
    files: community.files.slice(0, limit).map((file, index) => ({
      ...file, lastModifiedDateTime: new Date(Date.now() - index * 86400000).toISOString()
    }))
  };
}

export function mockSearch(kind: 'news' | 'search', limit: number): ISearchResults {
  const rows = kind === 'news' ? news : search;
  return {
    totalRows: rows.length,
    rows: rows.slice(0, limit).map((row, index) => ({
      ...row,
      Created: new Date(Date.now() - index * 86400000).toISOString(),
      LastModifiedTime: new Date(Date.now() - index * 86400000).toISOString()
    }))
  };
}
