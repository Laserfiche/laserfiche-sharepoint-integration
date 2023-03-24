import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SendToLaserficheLoginComponentWebPartStrings';
import SendToLaserficheLoginComponent from './components/SendToLaserficheLoginComponent';
import { ISendToLaserficheLoginComponentProps } from './components/ISendToLaserficheLoginComponentProps';

export interface ISendToLaserficheLoginComponentWebPartProps {
  LaserficheRedirectPage: string;
  devMode: boolean;
}

export default class SendToLaserficheLoginComponentWebPart extends BaseClientSideWebPart<ISendToLaserficheLoginComponentWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISendToLaserficheLoginComponentProps> = React.createElement(
      SendToLaserficheLoginComponent,
      {
        laserficheRedirectPage: this.properties.LaserficheRedirectPage,
        context: this.context,
        devMode: this.properties.devMode,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const searchParams = new URLSearchParams(location.search);
    const devMode = searchParams.get('devMode');

    let conditionalGroupFields: IPropertyPaneGroup['groupFields'] = [];

    if (devMode.toLocaleLowerCase() == 'true') {
      conditionalGroupFields = [
        PropertyPaneTextField('LaserficheRedirectPage', {
          label: strings.LaserficheRedirectPage,
        }),
        PropertyPaneToggle('devMode', {
          label: 'Dev Mode',
        }),
      ];
    } else {
      conditionalGroupFields = [
        PropertyPaneTextField('LaserficheRedirectPage', {
          label: strings.LaserficheRedirectPage,
        }),
      ];
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: conditionalGroupFields,
            },
          ],
        },
      ],
    };
  }
}
