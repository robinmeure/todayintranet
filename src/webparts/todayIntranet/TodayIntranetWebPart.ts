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
  dashboardId: string;
}

/** Layouts are scoped to this id, so re-adding the web part keeps them. */
const DEFAULT_DASHBOARD_ID: string = 'default';

/**
 * The id ends up in a list item title, an OData filter and a localStorage key, and
 * `|` separates it from the login name, so it is reduced to a safe short slug.
 */
function toDashboardKey(value: string | undefined): string {
  const slug = (value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .substring(0, 50)
    .replace(/^-+|-+$/g, '');
  return slug || DEFAULT_DASHBOARD_ID;
}

/** Shown to anyone who has not arranged their dashboard yet. */
const STARTER_LAYOUT: IDashboardLayout = {
  version: CURRENT_LAYOUT_VERSION,
  presetId: 'three',
  rowSizeId: 'medium',
  widgets: [
    { id: 'starter-calendar', type: 'm365.calendar', x: 0, y: 0, w: 4, h: 6 },
    { id: 'starter-mail', type: 'm365.mail', x: 4, y: 0, w: 4, h: 6 },
    { id: 'starter-tasks', type: 'm365.tasks', x: 8, y: 0, w: 4, h: 6 }
  ]
};

export default class TodayIntranetWebPart extends BaseClientSideWebPart<ITodayIntranetWebPartProps> {
  private _store: ILayoutStore | undefined;
  private _storeKey: string | undefined;
  private _theme: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    initializeIcons(undefined, { disableWarnings: true });
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDashboardProps> = React.createElement(Dashboard, {
      title: this.properties.title || 'Today',
      spContext: this.context,
      store: this._ensureStore(),
      starterLayout: STARTER_LAYOUT,
      theme: this._theme
    });

    ReactDom.render(element, this.domElement);
  }

  /** Rebuilt when the author changes the dashboard id, so the new scope is loaded. */
  private _ensureStore(): ILayoutStore {
    const key = toDashboardKey(this.properties.dashboardId);
    if (!this._store || this._storeKey !== key) {
      this._storeKey = key;
      this._store = new SharePointListLayoutStore({
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        userKey: this.context.pageContext.user.loginName,
        dashboardKey: key
      });
    }
    return this._store;
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

    // SPFx also raises this during initialization, before `properties` is populated,
    // where rendering would throw. Only a later theme switch has to be pushed into
    // the React tree; the first render is driven by SPFx itself.
    if (this.renderedOnce) {
      this.render();
    }
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
            },
            {
              groupName: strings.StorageGroupName,
              groupFields: [
                PropertyPaneTextField('dashboardId', {
                  label: strings.DashboardIdFieldLabel,
                  description: strings.DashboardIdFieldDescription,
                  placeholder: DEFAULT_DASHBOARD_ID,
                  // Repointing the store on every keystroke would read the list
                  // once per intermediate value, so wait for typing to settle.
                  deferredValidationTime: 1500
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
