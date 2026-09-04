import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { initializeIcons } from '@fluentui/react/lib/Icons';

import * as strings from 'TodayIntranetWebPartStrings';
import { Dashboard, IDashboardProps } from './components/Dashboard';
import { IDashboardLayout, CURRENT_LAYOUT_VERSION } from './model/IDashboardLayout';
import { ILayoutStore } from './services/ILayoutStore';
import { SharePointListLayoutStore } from './services/SharePointListLayoutStore';

export interface ITodayIntranetWebPartProps {
  title: string;
}

/** Shown to anyone who has not arranged their dashboard yet. */
const STARTER_LAYOUT: IDashboardLayout = {
  version: CURRENT_LAYOUT_VERSION,
  widgets: [
    { id: 'starter-calendar', type: 'm365.calendar', x: 0, y: 0, w: 4, h: 6 },
    { id: 'starter-mail', type: 'm365.mail', x: 4, y: 0, w: 4, h: 6 },
    { id: 'starter-tasks', type: 'm365.tasks', x: 8, y: 0, w: 4, h: 6 }
  ]
};

export default class TodayIntranetWebPart extends BaseClientSideWebPart<ITodayIntranetWebPartProps> {
  private _store: ILayoutStore | undefined;
  private _theme: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    initializeIcons(undefined, { disableWarnings: true });

    this._store = new SharePointListLayoutStore({
      spHttpClient: this.context.spHttpClient,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      userKey: this.context.pageContext.user.loginName,
      dashboardKey: this.context.instanceId
    });

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDashboardProps> = React.createElement(Dashboard, {
      title: this.properties.title || 'Today',
      spContext: this.context,
      store: this._store!,
      starterLayout: STARTER_LAYOUT,
      theme: this._theme
    });

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._theme = currentTheme;

    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [PropertyPaneTextField('title', { label: strings.TitleFieldLabel })]
            }
          ]
        }
      ]
    };
  }
}
