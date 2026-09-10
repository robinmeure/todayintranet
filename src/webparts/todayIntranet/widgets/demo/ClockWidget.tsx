import * as React from 'react';
import { IWidgetContext } from '../IWidget';
import { WidgetStat } from '../content';

/**
 * Placeholder widget with live state, so it is obvious that widgets keep
 * running while the dashboard is rearranged.
 */
export const ClockWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const [now, setNow] = React.useState<Date>(new Date());

  React.useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  const locale = context.spContext.pageContext.cultureInfo.currentUICultureName || undefined;

  return (
    <WidgetStat
      value={now.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })}
      caption={now.toLocaleDateString(locale, { weekday: 'long', day: 'numeric', month: 'long' })}
    />
  );
};
