import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import { WidgetStat } from '../content';
import mockClock from '../mock/clock.json';

/**
 * Placeholder widget with live state, so it is obvious that widgets keep
 * running while the dashboard is rearranged.
 */
export const ClockWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const [now, setNow] = React.useState<Date>(new Date());

  React.useEffect(() => {
    if (context.useMockData) { return; }
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, [context.useMockData]);

  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  return (
    <WidgetStat
      value={context.useMockData ? mockClock.value : now.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })}
      caption={context.useMockData ? mockClock.caption
        : now.toLocaleDateString(locale, { weekday: 'long', day: 'numeric', month: 'long' })}
    />
  );
};
