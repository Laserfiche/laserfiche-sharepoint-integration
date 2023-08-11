import {
  PostEntryWithEdocMetadataRequest,
  FileParameter,
  CreateEntryResult,
  IPostEntryWithEdocMetadataRequest,
  FieldToUpdate,
  ValueToUpdate,
  PutFieldValsRequest,
  Entry,
} from '@laserfiche/lf-repository-api-client';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
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
  metadataSaved: boolean;
  fileName: string;
}

export class SaveDocumentToLaserfiche {
  constructor(private spFileMetadata: ISPDocumentData) {}

  async trySaveDocumentToLaserficheAsync(): Promise<
    SavedToLaserficheDocumentData | undefined
  > {
    const loginComponent: NgElement & WithProperties<LfLoginComponent> =
      document.querySelector('lf-login');
    const accessToken = loginComponent?.authorization_credentials?.accessToken;
    if (accessToken) {
      const validRepoClient = await this.tryGetValidRepositoryClientAsync();
      const webClientUrl = loginComponent?.account_endpoints.webClientUrl;

      if (validRepoClient && this.spFileMetadata) {
        const spFileData = await this.GetFileData();
        const result = await this.saveFileToLaserficheAsync(
          spFileData,
          validRepoClient,
          webClientUrl
        );
        return result;
      } else {
        return undefined;
      }
    } else {
      // user is not logged in
    }
  }

  async tryGetValidRepositoryClientAsync(): Promise<IRepositoryApiClientExInternal> {
    const repoClientCreator = new RepositoryClientExInternal();
    const newRepoClient = await repoClientCreator.createRepositoryClientAsync();
    try {
      // test accessToken validity
      await newRepoClient.repositoriesClient.getRepositoryList({});
    } catch {
      return undefined;
    }
    return newRepoClient;
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
      console.log('error ocurred' + error);
    }
  }

  async saveFileToLaserficheAsync(
    spFileData: Blob,
    repoClient: IRepositoryApiClientExInternal,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    if (spFileData && repoClient) {
      const laserficheProfileName = this.spFileMetadata.lfProfile;
      let result: SavedToLaserficheDocumentData | undefined;
      if (laserficheProfileName) {
        result = await this.sendToLaserficheWithMappingAsync(
          spFileData,
          repoClient,
          webClientUrl
        );
      } else {
        result = await this.sendToLaserficheNoMappingAsync(
          spFileData,
          repoClient,
          webClientUrl
        );
      }
      return result;
    }
    return undefined;
  }

  async sendToLaserficheWithMappingAsync(
    fileData: Blob,
    repoClient: IRepositoryApiClientExInternal,
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
    const repoId = await repoClient.getCurrentRepoId();

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
        await repoClient.entriesClient.importDocument(entryRequest);
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
        await this.deleteSPFileAsync();
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
      };

      await this.tryUpdateFileNameAsync(
        repoClient,
        repoId,
        entryCreateResult,
        fileInfo
      );
      return fileInfo;
    } catch (error) {
      const conflict409 =
        error.operations.setFields.exceptions[0].statusCode === 409;
      if (conflict409) {
        const entryId = error.operations.entryCreate.entryId;

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
        const fileInfo: SavedToLaserficheDocumentData = {
          fileLink,
          pathBack: path,
          metadataSaved: false,
          fileName,
        };

        await this.tryUpdateFileNameAsync(repoClient, repoId, error, fileInfo);
        return fileInfo;
      } else {
        window.alert(`Error uploading file: ${JSON.stringify(error)}`);
        window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
        return undefined;
      }
    }
  }

  getRequestMetadata(request: PostEntryWithEdocMetadataRequest): PostEntryWithEdocMetadataRequest {
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
    repoClient: IRepositoryApiClientExInternal,
    webClientUrl: string
  ): Promise<SavedToLaserficheDocumentData | undefined> {
    const fileNameWithExt = this.spFileMetadata.fileName;

    const fileNameSplitByDot = (fileNameWithExt as string).split('.');
    const fileExtensionWithPeriod = fileNameSplitByDot.pop();
    const fileNameWithoutExt = fileNameSplitByDot.join('.');

    const parentEntryId = 1;

    try {
      const repoId = await repoClient.getCurrentRepoId();
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
        await repoClient.entriesClient.importDocument(entryRequest);
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
      };
      await this.tryUpdateFileNameAsync(
        repoClient,
        repoId,
        entryCreateResult,
        fileInfo
      );
      return fileInfo;
    } catch (error) {
      window.alert(`Error uploading file: ${JSON.stringify(error)}`);
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      return undefined;
    }
  }

  private async tryUpdateFileNameAsync(
    repoClient: IRepositoryApiClientExInternal,
    repoId: string,
    entryCreateResult: CreateEntryResult,
    fileInfo: SavedToLaserficheDocumentData
  ): Promise<void> {
    try {
      const entryInfo: Entry = await repoClient.entriesClient.getEntry({
        repoId,
        entryId: entryCreateResult.operations.entryCreate.entryId,
      });

      fileInfo.fileName = entryInfo.name;
    } catch {
      // do nothing, keep default file name
    }
  }

  async deleteSPFileAsync(): Promise<void> {
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
    if (response.ok) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      //Perform further activity upon success, like displaying a notification
      alert('File deleted successfully');
    } else {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      console.log('An error occurred. Please try again.');
    }
  }

  async deleteSPFileAndReplaceWithLinkAsync(docFilelink: string): Promise<void> {
    const filenameWithoutExt = PathUtils.removeFileExtension(
      this.spFileMetadata.fileName
    );
    const encodedFileName = encodeURIComponent(this.spFileMetadata.fileName);
    const spFileUrlWithEncodedFileName = this.spFileMetadata.fileUrl.replace(
      this.spFileMetadata.fileName,
      encodedFileName
    );
    const fullSpFileUrl = window.location.origin + spFileUrlWithEncodedFileName;
    const deleteFile = await fetch(fullSpFileUrl, {
      method: 'DELETE',
      headers: {
        Accept: 'application/json;odata=verbose',
      },
    });
    if (deleteFile.ok) {
      alert('File replaced with link successfully');
      await this.replaceFileWithLinkAsync(filenameWithoutExt, docFilelink);
    } else {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      console.log('An error occurred. Please try again.');
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
      await this.createLinkAsync(filenameWithoutExt, docFileLink, FormDigestValue);
    } else {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      console.log('Failed');
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
    if (resp.ok) {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      console.log('Item Inserted..!!');
      console.log(await resp.json());
    } else {
      window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
      console.log('API Error');
      console.log(await resp.json());
    }
  }
}
