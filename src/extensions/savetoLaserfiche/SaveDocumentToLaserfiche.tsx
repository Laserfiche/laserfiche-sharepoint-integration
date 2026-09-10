// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import {
  Entry,
  FieldToUpdate,
  FileParameter,
  IImportEntryRequestMetadata,
  ImportEntryRequest,
  ImportEntryRequestMetadata,
} from '@laserfiche/lf-repository-api-client-v2';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { getEntryWebAccessUrl } from '../../Utils/Funcs';
import { ISPDocumentData } from '../../Utils/Types';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';

export interface SavedToLaserficheDocumentData {
  fileLink: string;
  pathBack: string;
  fileName: string;
  action: ActionTypes | undefined;
}

export class SaveDocumentToLaserfiche {
  constructor(
    private spFileMetadata: ISPDocumentData,
    private validRepoClient: IRepositoryApiClientExInternal
  ) {}

  async trySaveDocumentToLaserficheAsync(): Promise<SavedToLaserficheDocumentData> {
    const loginComponent: NgElement & WithProperties<LfLoginComponent> =
      document.querySelector('lf-login');
    const accessToken = loginComponent?.authorization_credentials?.accessToken;
    if (accessToken) {
      const webClientUrl = loginComponent?.account_endpoints.webClientUrl;

      if (this.validRepoClient && this.spFileMetadata) {
        const spFileData = await this.GetFileData();
        const result = await this.saveFileToLaserficheAsync(
          spFileData,
          webClientUrl
        );
        return result;
      } else {
        throw Error(
          'You are not signed in or there was an issue retrieving data from SharePoint. Please try again.'
        );
      }
    } else {
      // user is not logged in
    }
  }

  async GetFileData(): Promise<Blob> {
    const spFileUrl = this.spFileMetadata.fileUrl;
    const fileNameWithExt = this.spFileMetadata.fileName;
    const encodedFileName = encodeURIComponent(fileNameWithExt);
    const encodedSpFileUrl = spFileUrl?.replace(
      fileNameWithExt,
      encodedFileName
    );
    const fullSPDataUrl = window.location.origin + encodedSpFileUrl;
    try {
      const res = await fetch(fullSPDataUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const spFileDataBlob = await res.blob();
      return spFileDataBlob;
    } catch (error) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw error;
    }
  }

  async saveFileToLaserficheAsync(
    spFileData: Blob,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    if (spFileData && this.validRepoClient) {
      const laserficheProfileName = this.spFileMetadata.lfProfile;
      let result: SavedToLaserficheDocumentData | undefined;
      if (laserficheProfileName) {
        result = await this.sendToLaserficheWithMappingAsync(
          spFileData,
          webClientUrl
        );
      } else {
        result = await this.sendToLaserficheNoMappingAsync(
          spFileData,
          webClientUrl
        );
      }
      return result;
    }
    return undefined;
  }

  async sendToLaserficheWithMappingAsync(
    fileData: Blob,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    const metadata: ImportEntryRequestMetadata | undefined = this
      .spFileMetadata.templateName
      ? this.getRequestMetadata()
      : undefined;

    // getCleanedExtension only prepends a '.'; it must be given the extension,
    // not the whole file name.
    const fileExtensionWithPeriod =
      PathUtils.getCleanedExtension(
        PathUtils.getFileExtension(this.spFileMetadata.fileName)
      ) ?? '';
    const filenameWithoutExt = PathUtils.removeFileExtension(
      this.spFileMetadata.fileName
    );
    const docNameIncludesFileName =
      this.spFileMetadata.documentName.includes('FileName');

    const parentEntryId = Number(this.spFileMetadata.entryId);
    const repoId = await this.validRepoClient.getCurrentRepoId();

    let fileName: string;
    if (!this.spFileMetadata.documentName) {
      fileName = filenameWithoutExt;
    } else if (docNameIncludesFileName === false) {
      fileName = this.spFileMetadata.documentName;
    } else {
      fileName = this.spFileMetadata.documentName.replace(
        'FileName',
        filenameWithoutExt
      );
    }
    // The v2 API has no separate `extension` parameter, so the extension has to
    // be part of the electronic document's file name -- exactly once.
    const fileNameInEdoc = fileName + fileExtensionWithPeriod;

    const electronicDocument: FileParameter = {
      fileName: fileNameInEdoc,
      data: fileData,
    };
    const entryRequest = {
      repositoryId: repoId,
      entryId: parentEntryId,
      file: electronicDocument,
      request: new ImportEntryRequest({
        name: fileName,
        autoRename: true,
        // v2 defaults this to false, which would store txt/tif/tiff/bmp/pcx/
        // jpg/jpeg/gif/png files as image pages rather than as the electronic
        // document. v1 always stored the edoc.
        importAsElectronicDocument: true,
        metadata,
      }),
    };

    try {
      const entry: Entry =
        await this.validRepoClient.entriesClient.importEntry(entryRequest);
      const entryId = entry.id ?? 1;
      const fileLink = getEntryWebAccessUrl(
        entryId.toString(),
        webClientUrl,
        false,
        repoId
      );
      const fileUrl = this.spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = window.location.origin + fileUrlWithoutDocName;

      if (this.spFileMetadata.action === ActionTypes.COPY) {
        window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      } else if (this.spFileMetadata.action === ActionTypes.MOVE_AND_DELETE) {
        await this.deleteAndHandleSPFileAsync();
      } else if (this.spFileMetadata.action === ActionTypes.REPLACE) {
        await this.deleteSPFileAndReplaceWithLinkAsync(fileLink);
      } else {
        // TODO what should happen?
      }
      const fileInfo: SavedToLaserficheDocumentData = {
        fileLink,
        pathBack: path,
        fileName,
        action: this.spFileMetadata.action,
      };

      await this.tryUpdateFileNameAsync(repoId, entryId, fileInfo);
      return fileInfo;
    } catch (error) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw error;
    }
  }

  getRequestMetadata(): ImportEntryRequestMetadata {
    const fileMetadata: IImportEntryRequestMetadata =
      this.spFileMetadata.metadata;
    const fields: FieldToUpdate[] = (fileMetadata?.fields ?? []).map(
      (field) => new FieldToUpdate({ name: field.name, values: field.values })
    );
    return new ImportEntryRequestMetadata({
      templateName: fileMetadata?.templateName,
      fields,
    });
  }

  async sendToLaserficheNoMappingAsync(
    fileData: Blob,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    const fileNameWithExt = this.spFileMetadata.fileName;

    const fileNameWithoutExt = PathUtils.removeFileExtension(fileNameWithExt);

    const parentEntryId = 1;

    try {
      const repoId = await this.validRepoClient.getCurrentRepoId();
      const electronicDocument: FileParameter = {
        fileName: fileNameWithExt,
        data: fileData,
      };
      const entryRequest = {
        repositoryId: repoId,
        entryId: parentEntryId,
        file: electronicDocument,
        request: new ImportEntryRequest({
          name: fileNameWithoutExt,
          autoRename: true,
          importAsElectronicDocument: true,
        }),
      };

      const entry: Entry =
        await this.validRepoClient.entriesClient.importEntry(entryRequest);
      const entryId = entry.id;
      const fileLink = getEntryWebAccessUrl(
        entryId.toString(),
        webClientUrl,
        false,
        repoId
      );
      const fileUrl = this.spFileMetadata.fileUrl;
      const fileUrlWithoutDocName = fileUrl.slice(0, fileUrl.lastIndexOf('/'));
      const path = window.location.origin + fileUrlWithoutDocName;

      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      const fileInfo: SavedToLaserficheDocumentData = {
        fileLink,
        pathBack: path,
        fileName: fileNameWithExt,
        action: this.spFileMetadata.action,
      };
      await this.tryUpdateFileNameAsync(repoId, entryId, fileInfo);
      return fileInfo;
    } catch (error) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw error;
    }
  }

  private async tryUpdateFileNameAsync(
    repoId: string,
    entryId: number,
    fileInfo: SavedToLaserficheDocumentData
  ): Promise<void> {
    try {
      const entryInfo: Entry =
        await this.validRepoClient.entriesClient.getEntry({
          repositoryId: repoId,
          entryId,
        });

      fileInfo.fileName = entryInfo.name;
    } catch {
      // do nothing, keep default file name
    }
  }

  async deleteAndHandleSPFileAsync(): Promise<void> {
    const response = await this.deleteFileAsync();
    window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
    if (!response.ok) {
      throw Error(
        `An error occurred while deleting file: ${response.statusText}`
      );
    }
  }

  private async deleteFileAsync(): Promise<Response> {
    const encodedFileName = encodeURIComponent(this.spFileMetadata.fileName);
    const spUrlWithEncodedFileName = this.spFileMetadata.fileUrl.replace(
      this.spFileMetadata.fileName,
      encodedFileName
    );
    const fullSpFileUrl = window.location.origin + spUrlWithEncodedFileName;
    const init: RequestInit = {
      headers: {
        Accept: 'application/json;odata=verbose',
      },
      method: 'DELETE',
    };
    const response = await fetch(fullSpFileUrl, init);
    return response;
  }

  async deleteSPFileAndReplaceWithLinkAsync(
    docFilelink: string
  ): Promise<void> {
    const filenameWithoutExt = PathUtils.removeFileExtension(
      this.spFileMetadata.fileName
    );
    const deleteFile = await this.deleteFileAsync();
    if (deleteFile.ok) {
      await this.replaceFileWithLinkAsync(filenameWithoutExt, docFilelink);
    }
    if (!deleteFile.ok) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw Error(
        `An error occurred while replacing file with link: ${deleteFile.statusText}`
      );
    }
  }

  async replaceFileWithLinkAsync(
    filenameWithoutExt: string,
    docFileLink: string
  ): Promise<void> {
    const resp = await fetch(
      this.spFileMetadata.contextPageAbsoluteUrl + '/_api/contextinfo',
      {
        method: 'POST',
        headers: { accept: 'application/json;odata=verbose' },
      }
    );
    if (resp.ok) {
      const data = await resp.json();
      const FormDigestValue = data.d.GetContextWebInformation.FormDigestValue;
      await this.createLinkAsync(
        filenameWithoutExt,
        docFileLink,
        FormDigestValue
      );
    } else {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw Error(
        `An error occurred while replacing file with link: ${resp.statusText}`
      );
    }
  }

  async createLinkAsync(
    filenameWithoutExt: string,
    docFileLink: string,
    formDigestValue: string
  ): Promise<void> {
    const encodedFileName = encodeURIComponent(filenameWithoutExt);
    const path = this.spFileMetadata.fileUrl.replace(
      this.spFileMetadata.fileName,
      ''
    );
    const AddLinkURL =
      this.spFileMetadata.contextPageAbsoluteUrl +
      `/_api/web/GetFolderByServerRelativeUrl('${path}')/Files/add(url='${encodedFileName}.url',overwrite=true)`;

    const resp = await fetch(AddLinkURL, {
      method: 'POST',
      body: `[InternetShortcut]\nURL=${docFileLink}`,
      headers: {
        'content-type': 'text/plain',
        accept: 'application/json;odata=verbose',
        'X-RequestDigest': formDigestValue,
      },
    });
    window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
    if (!resp.ok) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw Error(
        `An error occurred while replacing file with link: ${resp.statusText}`
      );
    }
  }
}
