import {
  PostEntryWithEdocMetadataRequest,
  FileParameter,
  CreateEntryResult,
  IPostEntryWithEdocMetadataRequest,
  FieldToUpdate,
  ValueToUpdate,
  PutFieldValsRequest,
  Entry,
  SetFields,
  APIServerException,
} from '@laserfiche/lf-repository-api-client';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { getEntryWebAccessUrl } from '../../Utils/Funcs';
import { ISPDocumentData } from '../../Utils/Types';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';
import * as React from 'react';
import styles from './SendToLaserFiche.module.scss';

export interface SavedToLaserficheDocumentData {
  fileLink: string;
  pathBack: string;
  metadataSaved: boolean;
  failedMetadata?: JSX.Element;
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
    let request: PostEntryWithEdocMetadataRequest;
    if (this.spFileMetadata.templateName) {
      request = this.getRequestMetadata(request);
    } else {
      request = new PostEntryWithEdocMetadataRequest({});
    }

    const fileExtensionWithPeriod = PathUtils.getCleanedExtension(
      this.spFileMetadata.fileName
    );
    const filenameWithoutExt = PathUtils.removeFileExtension(
      this.spFileMetadata.fileName
    );
    const docNameIncludesFileName =
      this.spFileMetadata.documentName.includes('FileName');

    const parentEntryId = Number(this.spFileMetadata.entryId);
    const repoId = await this.validRepoClient.getCurrentRepoId();

    let fileName: string | undefined;
    let fileNameInEdoc: string | undefined;
    let extension: string | undefined;
    if (!this.spFileMetadata.documentName) {
      fileName = filenameWithoutExt;
      fileNameInEdoc = this.spFileMetadata.fileName;
      extension = fileExtensionWithPeriod;
    } else if (docNameIncludesFileName === false) {
      fileName = this.spFileMetadata.documentName;
      fileNameInEdoc = this.spFileMetadata.documentName;
      extension = fileExtensionWithPeriod;
    } else {
      const docNameReplacedWithFileName =
        this.spFileMetadata.documentName.replace(
          'FileName',
          filenameWithoutExt
        );
      fileName = docNameReplacedWithFileName;
      fileNameInEdoc =
        docNameReplacedWithFileName + `.${fileExtensionWithPeriod}`;
      extension = fileExtensionWithPeriod;
    }
    const electronicDocument: FileParameter = {
      fileName: fileNameInEdoc,
      data: fileData,
    };
    const entryRequest = {
      repoId,
      parentEntryId,
      fileName,
      autoRename: true,
      electronicDocument,
      request,
      extension,
    };

    try {
      const entryCreateResult: CreateEntryResult =
        await this.validRepoClient.entriesClient.importDocument(entryRequest);
      const entryId = entryCreateResult.operations.entryCreate.entryId ?? 1;
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
        metadataSaved: true,
        fileName,
        action: this.spFileMetadata.action,
      };

      await this.tryUpdateFileNameAsync(repoId, entryCreateResult, fileInfo);
      return fileInfo;
    } catch (error) {
      const conflict409 =
        error.problemDetails.extensions.createEntryResult.operations?.setFields
          ?.exceptions[0].statusCode === 409;
      if (conflict409) {
        const setFields: SetFields =
          error.problemDetails.extensions.createEntryResult.operations
            .setFields;
        const errorMessages = setFields.exceptions.map(
          (value: APIServerException, index: number) => {
            return <li key={index}>{value.message}</li>;
          }
        );
        const entryId =
          error.problemDetails.extensions.createEntryResult.operations
            .entryCreate.entryId;

        const fileLink = getEntryWebAccessUrl(
          entryId.toString(),
          webClientUrl,
          false,
          repoId
        );
        const fileUrl = this.spFileMetadata.fileUrl;
        const fileUrlWithoutDocName = fileUrl.slice(
          0,
          fileUrl.lastIndexOf('/')
        );
        const path = window.location.origin + fileUrlWithoutDocName;
        window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
        const failedMetadata = (
          <ul className={styles.noMargin}>{errorMessages}</ul>
        );
        const fileInfo: SavedToLaserficheDocumentData = {
          fileLink,
          pathBack: path,
          metadataSaved: false,
          failedMetadata,
          fileName,
          action: undefined,
        };

        await this.tryUpdateFileNameAsync(repoId, error, fileInfo);
        return fileInfo;
      } else {
        window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
        throw error;
      }
    }
  }

  getRequestMetadata(
    request: PostEntryWithEdocMetadataRequest
  ): PostEntryWithEdocMetadataRequest {
    const fileMetadata: IPostEntryWithEdocMetadataRequest =
      this.spFileMetadata.metadata;
    const fieldsAlone = fileMetadata.metadata.fields;
    const formattedFieldValues:
      | {
          [key: string]: FieldToUpdate;
        }
      | undefined = {};

    for (const key in fieldsAlone) {
      const value = fieldsAlone[key];
      formattedFieldValues[key] = new FieldToUpdate({
        ...value,
        values: value.values.map((val) => new ValueToUpdate(val)),
      });
    }
    request = new PostEntryWithEdocMetadataRequest({
      template: fileMetadata.template,
      metadata: new PutFieldValsRequest({
        fields: formattedFieldValues,
      }),
    });
    return request;
  }

  async sendToLaserficheNoMappingAsync(
    fileData: Blob,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    const fileNameWithExt = this.spFileMetadata.fileName;

    const fileNameSplitByDot = (fileNameWithExt as string).split('.');
    const fileExtensionWithPeriod = fileNameSplitByDot.pop();
    const fileNameWithoutExt = fileNameSplitByDot.join('.');

    const parentEntryId = 1;

    try {
      const repoId = await this.validRepoClient.getCurrentRepoId();
      const electronicDocument: FileParameter = {
        fileName: fileNameWithExt,
        data: fileData,
      };
      const entryRequest = {
        repoId,
        parentEntryId,
        fileName: fileNameWithoutExt,
        autoRename: true,
        electronicDocument,
        request: new PostEntryWithEdocMetadataRequest({}),
        extension: fileExtensionWithPeriod,
      };

      const entryCreateResult: CreateEntryResult =
        await this.validRepoClient.entriesClient.importDocument(entryRequest);
      const entryId = entryCreateResult.operations.entryCreate.entryId;
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
        metadataSaved: true,
        fileName: fileNameWithExt,
        action: this.spFileMetadata.action,
      };
      await this.tryUpdateFileNameAsync(repoId, entryCreateResult, fileInfo);
      return fileInfo;
    } catch (error) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      throw error;
    }
  }

  private async tryUpdateFileNameAsync(
    repoId: string,
    entryCreateResult: CreateEntryResult,
    fileInfo: SavedToLaserficheDocumentData
  ): Promise<void> {
    try {
      const entryInfo: Entry =
        await this.validRepoClient.entriesClient.getEntry({
          repoId,
          entryId: entryCreateResult.operations.entryCreate.entryId,
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
