import * as React from 'react';
import { IWidgetContext } from '../IWidget';

const styleSheet: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: '8px',
  height: '100%'
};

/**
 * Placeholder widget that proves settings round-trip through the layout store.
 */
export const WelcomeWidget: React.FunctionComponent<{ context: IWidgetContext }> = ({ context }) => {
  const greeting = (context.settings.greeting as string) ?? 'Good to see you';
  const displayName = context.spContext.pageContext.user.displayName;

  return (
    <div style={styleSheet}>
      <div style={{ fontSize: '20px', fontWeight: 600 }}>
        {greeting}, {displayName.split(' ')[0]}.
      </div>
      <p style={{ margin: 0 }}>
        This tile is a placeholder. Real widgets (calendar, mail, tasks) plug in the same way:
        implement <code>IWidgetDefinition</code> and register it.
      </p>
      {context.isEditing && (
        <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '12px' }}>
          Greeting
          <input
            type="text"
            value={greeting}
            onChange={(e) => context.updateSettings({ ...context.settings, greeting: e.target.value })}
            onMouseDown={(e) => e.stopPropagation()}
          />
        </label>
      )}
    </div>
  );
};
