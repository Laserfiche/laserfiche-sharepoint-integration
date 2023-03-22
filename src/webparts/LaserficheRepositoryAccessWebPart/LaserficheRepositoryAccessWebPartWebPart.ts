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
import * as strings from 'LaserficheRepositoryAccessWebPartWebPartStrings';
import LaserficheRepositoryAccessWebPart from './components/LaserficheRepositoryAccessWebPart';
import { ILaserficheRepositoryAccessWebPartProps } from './components/ILaserficheRepositoryAccessWebPartProps';

export interface ILaserficheRepositoryAccessWebPartWebPartProps {
  WebPartTitle: string;
  LaserficheRedirectUrl: string;
  devMode: boolean;
}

export default class LaserficheRepositoryAccessWebPartWebPart extends BaseClientSideWebPart<ILaserficheRepositoryAccessWebPartWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ILaserficheRepositoryAccessWebPartProps> = React.createElement(
      LaserficheRepositoryAccessWebPart,
      {
        context: this.context,
        webPartTitle: this.properties.WebPartTitle,
        laserficheRedirectUrl: this.properties.LaserficheRedirectUrl,
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
    const devemode = searchParams.get('devMode');

    let conditionalGroupFields: IPropertyPaneGroup['groupFields'] = [];

    if (devemode === 'true') {
      conditionalGroupFields = [
        PropertyPaneTextField('WebPartTitle', {
          label: strings.WebPartTitle,
        }),
        PropertyPaneTextField('LaserficheRedirectUrl', {
          label: strings.LaserficheRedirectUrl,
        }),
        PropertyPaneToggle('devMode', {
          label: 'Dev Mode',
        }),
      ];
    } else {
      conditionalGroupFields = [
        PropertyPaneTextField('WebPartTitle', {
          label: strings.WebPartTitle,
        }),
        PropertyPaneTextField('LaserficheRedirectUrl', {
          label: strings.LaserficheRedirectUrl,
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
