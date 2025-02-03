// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { BaseDialog } from '@microsoft/sp-dialog';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import { ISPDocumentData } from '../../Utils/Types';
import { Navigation } from 'spfx-navigation';
import SaveToLaserficheCustomDialog from './SaveToLaserficheDialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { GetDocumentDialogData } from './GetDocumentDataDialogUI';

const signInPageRoute = '/SitePages/LaserficheSignIn.aspx';

export class GetDocumentDataCustomDialog extends BaseDialog {
  successful = false;

  constructor(
    private fileInfo: {
      fileName: string;
      spContentType: string;
      spFileUrl: string;
      fileId: string;
    },
    private context: BaseComponentContext
  ) {
    super();
  }

  showNextDialog: (data: ISPDocumentData) => Promise<void> = async (
    data: ISPDocumentData
  ) => {
    const saveToLfDialog = new SaveToLaserficheCustomDialog(data, () =>
      this.close()
    );
    await this.secondaryDialogProvider.show(saveToLfDialog);
    if (!saveToLfDialog.successful) {
      Navigation.navigate(
        this.context.pageContext.web.absoluteUrl + signInPageRoute,
        true
      );
    }
  };

  public render(): void {
    const element: React.ReactElement = (
      <React.StrictMode>
        <GetDocumentDialogData
          spFileInfo={this.fileInfo}
          context={this.context}
          showSaveToDialog={this.showNextDialog}
          handleCancelDialog={this.close}
        />
      </React.StrictMode>
    );
    ReactDOM.render(element, this.domElement);
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
  }
}
