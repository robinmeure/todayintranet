import * as React from 'react';
import { IWidgetContext } from '../IWidget';

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
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
        gap: '4px'
      }}
    >
      <div style={{ fontSize: '32px', fontWeight: 600, lineHeight: 1 }}>
        {now.toLocaleTimeString(locale, { hour: '2-digit', minute: '2-digit' })}
      </div>
      <div style={{ opacity: 0.75 }}>
        {now.toLocaleDateString(locale, { weekday: 'long', day: 'numeric', month: 'long' })}
      </div>
    </div>
  );
};
