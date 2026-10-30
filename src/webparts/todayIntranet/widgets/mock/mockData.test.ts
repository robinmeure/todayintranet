import { mockCalendar, mockCommunity, mockMail, mockSearch, mockTasks } from './mockData';
import { calendarItems, calendarWindow } from '../graph/calendarData';
import { parseLinksSetting } from '../links/linksSettings';
import links from './myLinks.json';

describe('widget mock fixtures', () => {
  afterEach(() => jest.useRealTimers());

  it('keeps inbox samples paged, bounded and consistent with Unread only', () => {
    expect(mockMail(4, false)).toHaveLength(4);
    expect(mockMail(20, true).length).toBeGreaterThan(6);
    expect(mockMail(20, true).every((message) => message.isRead === false)).toBe(true);
    expect(mockMail(20, false).every((message) => !message.webLink)).toBe(true);
  });

  it.each([new Date(2026, 9, 30, 13), new Date(2030, 0, 1, 23, 59, 59)])(
    'keeps meetings visible for today, tomorrow and historical dates at %s', (now) => {
      jest.useFakeTimers();
      jest.setSystemTime(now);
      for (const range of ['today', 'tomorrow', 'date'] as const) {
        const window = calendarWindow(now, range, '2025-01-01', 7);
        const events = mockCalendar(window.start.toISOString(), window.end.toISOString());
        expect(calendarItems(events, window.start, now, 100, 'en-GB')).toHaveLength(8);
        expect(events.every((event) => new Date(event.start!.dateTime!).getTime() >= window.start.getTime()
          && new Date(event.end!.dateTime!).getTime() <= window.end.getTime())).toBe(true);
      }
    }
  );

  it('provides complete, disjoint task sources without live links or continuations', () => {
    const keys = new Set<string>();
    for (const source of ['todo', 'planner', 'outlook'] as const) {
      const data = mockTasks(source);
      expect(data.complete).toBe(true);
      expect(data.continuation).toBeUndefined();
      expect(data.tasks).toHaveLength(8);
      data.tasks.forEach((task) => {
        expect(keys.has(task.key)).toBe(false);
        keys.add(task.key);
        expect(task.source).toBe(source);
        expect(task.sourceUrl).toBeUndefined();
      });
    }
  });

  it('provides bounded community files, distinct search fixtures and valid links', () => {
    expect(mockCommunity(3).files).toHaveLength(3);
    expect(mockCommunity(20).files.length).toBeGreaterThan(5);
    expect(mockSearch('news', 4).rows).toHaveLength(4);
    expect(mockSearch('news', 4).totalRows).toBe(8);
    expect(mockSearch('search', 20).rows).toHaveLength(8);
    expect(mockSearch('search', 20).rows[0].FileType).toBe('docx');
    expect(parseLinksSetting(JSON.stringify(links)).links).toHaveLength(8);
  });
});
