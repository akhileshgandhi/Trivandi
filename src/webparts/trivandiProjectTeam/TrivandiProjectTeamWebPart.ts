import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TrivandiProjectTeamWebPartStrings';
import TrivandiProjectTeam from './components/TrivandiProjectTeam';
import { getSP } from './loc/pnpjsConfig';
import { spfi, SPFx } from '@pnp/sp/presets/all';
import { initProjectService } from '../../shared/services/projectService';

export interface ITrivandiProjectTeamWebPartProps {
  description: string;
}

export default class TrivandiProjectTeamWebPart
  extends BaseClientSideWebPart<ITrivandiProjectTeamWebPartProps> {

  private _isDarkTheme = false;
  private _environmentMessage = '';

  public render(): void {

    const element = React.createElement(TrivandiProjectTeam, {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      context: this.context,
      siteUrl: this.context.pageContext.web.absoluteUrl
    });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    await super.onInit();
    // ✅ Create SPFI instance using SPFx context
    const sp = spfi().using(SPFx(this.context));

    // ✅ Pass SPFI instance to service
    initProjectService(sp);

    console.log("[WebPart] PnP SPFI initialized");
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(_: IReadonlyTheme | undefined): void { }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
