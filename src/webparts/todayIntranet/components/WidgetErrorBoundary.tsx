import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

export interface IWidgetErrorBoundaryProps {
  widgetName: string;
  children: React.ReactNode;
}

interface IWidgetErrorBoundaryState {
  error?: Error;
}

/**
 * Keeps a crashing widget contained. Without this a single render error unmounts
 * the whole React tree and the dashboard goes blank.
 */
export class WidgetErrorBoundary extends React.Component<IWidgetErrorBoundaryProps, IWidgetErrorBoundaryState> {
  constructor(props: IWidgetErrorBoundaryProps) {
    super(props);
    this.state = {};
  }

  public static getDerivedStateFromError(error: Error): IWidgetErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error(`[TodayIntranet] Widget "${this.props.widgetName}" failed to render.`, error, info);
  }

  public render(): React.ReactNode {
    if (!this.state.error) {
      return this.props.children;
    }

    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
        {this.props.widgetName} could not be displayed. The rest of your dashboard is unaffected.
      </MessageBar>
    );
  }
}
