// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { Log } from '@microsoft/sp-core-library';
import { GetDocumentDataCustomDialog } from './GetDocumentDataDialog';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from '@microsoft/sp-listview-extensibility';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { CreateConfigurations } from '../../Utils/CreateConfigurations';
import { getSPListURL } from '../../Utils/Funcs';
import { LASERFICHE_SIGNIN_PAGE_NAME, SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISendToLfCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE = 'SendToLfCommandSet';

export default class SendToLfCommandSet extends BaseListViewCommandSet<ISendToLfCommandSetProperties> {
  hasSignInPage = false;
  hasAdminPage = false;

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SendToLfCommandSet');
    window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
    await CreateConfigurations.ensureAdminConfigListCreatedAsync(this.context);
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand('SAVE_TO_LASERFICHE');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible =
        event.selectedRows.length === 1 &&
        event.selectedRows[0].getValueByName('ContentType') !== 'Folder';
    }
  }

  public async onExecute(
    event: IListViewCommandSetExecuteEventParameters
  ): Promise<void> {
    const spDocumentProperties: RowAccessor = event.selectedRows[0];
    const fileId = spDocumentProperties.getValueByName('ID');
    const fileSize = spDocumentProperties.getValueByName('File_x0020_Size');
    const fileUrl = spDocumentProperties.getValueByName('FileRef');
    const fileName = spDocumentProperties.getValueByName('FileLeafRef');
    const spContentType = spDocumentProperties.getValueByName('ContentType');
    const isCheckedOut =
      spDocumentProperties.getValueByName('CheckoutUser')?.length > 0;

    await this.pageConfigurationCheck();

    const fileExtensionOnly = PathUtils.getFileExtension(fileName);
    const fileNoName = PathUtils.removeFileExtension(fileName);

    if (spContentType === 'Folder') {
      alert('Cannot Send a Folder To Laserfiche');
    } else if (!fileNoName || fileNoName.length === 0) {
      alert(
        'Please add a filename to the selected file before trying to save to Laserfiche.'
      );
    } else if (fileExtensionOnly === 'url') {
      alert('Cannot send the .url file to Laserfiche');
    } else if (isCheckedOut) {
      alert(
        'The selected file is checked out. Please discard the checkout or check the file back in before trying to save to Laserfiche.'
      );
    } else if (fileSize > 100000000) {
      alert('Please select a file below 100MB in size');
    } else if (!this.hasSignInPage) {
      alert(
        'Missing "LaserficheSignIn" SharePoint page. Please refer to the Adding App to SharePoint Site topic in the administration guide for configuration steps.'
      );
    } else {
      await this.trySaveToLaserficheAsync({
        fileName,
        spContentType,
        spFileUrl: fileUrl,
        fileId,
      });
    }
  }

  public async pageConfigurationCheck(): Promise<void> {
    try {
      const res = await fetch(
        `${getSPListURL(this.context, 'Site Pages')}/items`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const sitePages = await res.json();
      for (let o = 0; o < sitePages.value.length; o++) {
        const pageName = sitePages.value[o].Title;
        if (pageName === LASERFICHE_SIGNIN_PAGE_NAME) {
          this.hasSignInPage = true;
        }
      }
    } catch (error) {
      // TODO
    }
  }

  public async trySaveToLaserficheAsync(spFileInfo: {
    fileName: string;
    spContentType: string;
    spFileUrl: string;
    fileId: string;
  }): Promise<void> {
    const saveToDialog = new GetDocumentDataCustomDialog(
      spFileInfo,
      this.context
    );
    await saveToDialog.show();
  }
}
