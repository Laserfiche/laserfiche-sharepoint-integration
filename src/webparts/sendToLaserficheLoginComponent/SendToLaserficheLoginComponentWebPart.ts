// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import SendToLaserficheLoginComponent from './components/SendToLaserficheLoginComponent';


export default class SendToLaserficheLoginComponentWebPart extends BaseClientSideWebPart<{}> {
  public render(): void {
    const element: React.ReactElement = React.createElement(
      SendToLaserficheLoginComponent,
      {
        context: this.context,
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

}
