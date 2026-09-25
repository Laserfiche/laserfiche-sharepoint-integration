// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// Stand in constructor-shaped fakes for the request models that echo their arguments back, so
// the exact object the source builds for `importEntry` can be inspected in each test.
// @laserfiche/lf-js-utils is NOT mocked: PathUtils' real filename handling is under test too.
vi.mock('@laserfiche/lf-repository-api-client-v2', () => ({
  ImportEntryRequest: vi.fn().mockImplementation(function (args) {
    return { ...args };
  }),
  ImportEntryRequestMetadata: vi.fn().mockImplementation(function (args) {
    return { ...args };
  }),
  FieldToUpdate: vi.fn().mockImplementation(function (args) {
    return { ...args };
  }),
}));

vi.mock('../../Utils/Funcs', () => ({
  getEntryWebAccessUrl: vi.fn().mockReturnValue('https://webclient.example.com/entry/123'),
}));

import type { Mock } from 'vitest';
import { SaveDocumentToLaserfiche } from './SaveDocumentToLaserfiche';
import { ISPDocumentData } from '../../Utils/Types';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';
import { getEntryWebAccessUrl } from '../../Utils/Funcs';

function makeSpFileMetadata(overrides: Partial<ISPDocumentData> = {}): ISPDocumentData {
  return {
    fileName: 'Invoice.pdf',
    documentName: '',
    action: ActionTypes.COPY,
    fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
    entryId: '100',
    contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
    ...overrides,
  };
}

function makeValidRepoClient(
  overrides: {
    importEntryResult?: unknown;
    getEntryResult?: unknown;
  } = {}
): any {
  return {
    getCurrentRepoId: vi.fn().mockResolvedValue('repo-1'),
    entriesClient: {
      importEntry: vi.fn().mockResolvedValue(overrides.importEntryResult ?? { id: 42 }),
      getEntry: vi.fn().mockResolvedValue(overrides.getEntryResult ?? { name: 'final-name.pdf' }),
    },
  };
}

let removeItemMock: Mock;

describe('SaveDocumentToLaserfiche', () => {
  beforeEach(() => {
    removeItemMock = vi.fn();
    Object.defineProperty(window, 'localStorage', {
      value: { removeItem: removeItemMock },
      writable: true,
    });
    document.body.innerHTML = '<lf-login></lf-login>';
    globalThis.fetch = vi.fn() as any;
    (getEntryWebAccessUrl as Mock).mockClear();
  });

  describe('trySaveDocumentToLaserficheAsync', () => {
    test('resolves undefined and never fetches when there is no access token', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const getFileDataSpy = vi.spyOn(instance, 'GetFileData');

      const result = await instance.trySaveDocumentToLaserficheAsync();

      expect(result).toBeUndefined();
      expect(getFileDataSpy).not.toHaveBeenCalled();
      expect(globalThis.fetch).not.toHaveBeenCalled();
    });

    test('throws when the repository client could not be constructed', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const loginEl = document.querySelector('lf-login') as any;
      loginEl.authorization_credentials = { accessToken: 'token-123' };
      loginEl.account_endpoints = {
        webClientUrl: 'https://webclient.example.com',
      };
      loginEl.account_id = 'customer-1';
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, undefined as any);

      await expect(instance.trySaveDocumentToLaserficheAsync()).rejects.toThrow(
        'You are not signed in or there was an issue retrieving data from SharePoint. Please try again.'
      );
    });
  });

  describe('GetFileData', () => {
    test('encodes only the file name segment of the SharePoint file URL', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'my file.pdf',
        fileUrl: '/sites/site1/Shared Documents/my file.pdf',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const blob = { size: 123 };
      (globalThis.fetch as Mock).mockResolvedValue({
        blob: vi.fn().mockResolvedValue(blob),
      });

      const result = await instance.GetFileData();

      const expectedUrl = window.location.origin + '/sites/site1/Shared Documents/my%20file.pdf';
      expect(globalThis.fetch).toHaveBeenCalledWith(expectedUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      expect(result).toBe(blob);
    });

    test('clears SharePoint local storage and rethrows when fetch fails', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const error = new Error('network down');
      (globalThis.fetch as Mock).mockRejectedValue(error);

      await expect(instance.GetFileData()).rejects.toBe(error);
      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    });
  });

  describe('saveFileToLaserficheAsync', () => {
    test('routes to sendToLaserficheWithMappingAsync when lfProfile is set', async () => {
      const spFileMetadata = makeSpFileMetadata({ lfProfile: 'Profile1' });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const withMappingSpy = vi
        .spyOn(instance, 'sendToLaserficheWithMappingAsync')
        .mockResolvedValue({
          fileLink: 'link',
          folderLink: 'folder',
          pathBack: 'back',
          fileName: 'name',
          action: ActionTypes.COPY,
        });
      const noMappingSpy = vi.spyOn(instance, 'sendToLaserficheNoMappingAsync');
      const fileData = {} as any;

      const result = await instance.saveFileToLaserficheAsync(
        fileData,
        'https://webclient.example.com',
        'cust1'
      );

      expect(withMappingSpy).toHaveBeenCalledWith(
        fileData,
        'https://webclient.example.com',
        'cust1'
      );
      expect(noMappingSpy).not.toHaveBeenCalled();
      expect(result?.fileName).toBe('name');
    });

    test('routes to sendToLaserficheNoMappingAsync when lfProfile is unset', async () => {
      const spFileMetadata = makeSpFileMetadata({ lfProfile: undefined });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const withMappingSpy = vi.spyOn(instance, 'sendToLaserficheWithMappingAsync');
      const noMappingSpy = vi.spyOn(instance, 'sendToLaserficheNoMappingAsync').mockResolvedValue({
        fileLink: 'link2',
        folderLink: 'folder2',
        pathBack: 'back2',
        fileName: 'name2',
        action: ActionTypes.COPY,
      });
      const fileData = {} as any;

      const result = await instance.saveFileToLaserficheAsync(
        fileData,
        'https://webclient.example.com',
        'cust1'
      );

      expect(noMappingSpy).toHaveBeenCalledWith(fileData, 'https://webclient.example.com', 'cust1');
      expect(withMappingSpy).not.toHaveBeenCalled();
      expect(result?.fileName).toBe('name2');
    });

    test('resolves undefined without routing when spFileData is falsy', async () => {
      const spFileMetadata = makeSpFileMetadata({ lfProfile: 'Profile1' });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const withMappingSpy = vi.spyOn(instance, 'sendToLaserficheWithMappingAsync');
      const noMappingSpy = vi.spyOn(instance, 'sendToLaserficheNoMappingAsync');

      const result = await instance.saveFileToLaserficheAsync(
        null as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(result).toBeUndefined();
      expect(withMappingSpy).not.toHaveBeenCalled();
      expect(noMappingSpy).not.toHaveBeenCalled();
    });
  });

  describe('sendToLaserficheWithMappingAsync', () => {
    test('links to the destination folder in the web client', async () => {
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(
        makeSpFileMetadata({ entryId: '100' }),
        validRepoClient
      );

      const result = await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
        '100',
        'https://webclient.example.com',
        true,
        'repo-1',
        'cust1'
      );
      expect(result?.folderLink).toBe('https://webclient.example.com/entry/123');
    });

    test('uses the SP file name without extension when documentName is empty', async () => {
      const spFileMetadata = makeSpFileMetadata({
        documentName: '',
        fileName: 'Invoice.pdf',
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.request.name).toBe('Invoice');
      expect(entryRequest.file.fileName).toBe('Invoice.pdf');
      expect(entryRequest.request.importAsElectronicDocument).toBe(true);
      expect(entryRequest.request.autoRename).toBe(true);
    });

    test('uses documentName verbatim when it has no FileName token', async () => {
      const spFileMetadata = makeSpFileMetadata({
        documentName: 'Signed Contract',
        fileName: 'Invoice.pdf',
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.request.name).toBe('Signed Contract');
      expect(entryRequest.file.fileName).toBe('Signed Contract.pdf');
    });

    test('replaces the FileName token with the stripped SP file name', async () => {
      const spFileMetadata = makeSpFileMetadata({
        documentName: 'FileName-Archive',
        fileName: 'Invoice.pdf',
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.request.name).toBe('Invoice-Archive');
      expect(entryRequest.file.fileName).toBe('Invoice-Archive.pdf');
    });

    test('omits metadata when templateName is unset', async () => {
      const spFileMetadata = makeSpFileMetadata({ templateName: undefined });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.request.metadata).toBeUndefined();
    });

    test('maps metadata fields through FieldToUpdate when templateName and fields are present', async () => {
      const spFileMetadata = makeSpFileMetadata({
        templateName: 'Invoice Template',
        metadata: {
          templateName: 'Invoice Template',
          // Real FieldToUpdate is a class with init/toJSON members; the
          // mocked constructor (see vi.mock above) only cares about
          // name/values, so a plain object is enough at runtime.
          fields: [
            { name: 'Field1', values: ['v1'] },
            { name: 'Field2', values: ['v2', 'v3'] },
          ] as never,
        },
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.request.metadata.templateName).toBe('Invoice Template');
      expect(entryRequest.request.metadata.fields).toEqual([
        { name: 'Field1', values: ['v1'] },
        { name: 'Field2', values: ['v2', 'v3'] },
      ]);
    });

    test('clears local storage after a COPY without any extra network calls', async () => {
      const spFileMetadata = makeSpFileMetadata({ action: ActionTypes.COPY });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
      expect(globalThis.fetch).not.toHaveBeenCalled();
    });

    test('deletes the SharePoint file for MOVE_AND_DELETE', async () => {
      const spFileMetadata = makeSpFileMetadata({
        action: ActionTypes.MOVE_AND_DELETE,
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: true,
        statusText: 'OK',
      });

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(globalThis.fetch).toHaveBeenCalledWith(
        window.location.origin + '/sites/site1/Shared Documents/Invoice.pdf',
        expect.objectContaining({ method: 'DELETE' })
      );
    });

    test('deletes and replaces the SharePoint file with a link for REPLACE', async () => {
      const spFileMetadata = makeSpFileMetadata({
        action: ActionTypes.REPLACE,
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
        contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
      });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);
      (globalThis.fetch as Mock).mockImplementation((url: string, init?: RequestInit) => {
        if (init?.method === 'DELETE') {
          return Promise.resolve({ ok: true, statusText: 'OK' });
        }
        if (url.includes('/_api/contextinfo')) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                d: {
                  GetContextWebInformation: { FormDigestValue: 'digest-123' },
                },
              }),
          });
        }
        if (url.includes('/Files/add(')) {
          return Promise.resolve({ ok: true, statusText: 'OK' });
        }
        return Promise.reject(new Error(`Unexpected fetch call: ${url}`));
      });

      await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const addLinkCall = (globalThis.fetch as Mock).mock.calls.find(([url]: string[]) =>
        url.includes('/Files/add(')
      );
      expect(addLinkCall).toBeDefined();
      const [addLinkUrl, addLinkInit] = addLinkCall as [string, RequestInit];
      expect(addLinkUrl).toBe(
        "https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativeUrl('/sites/site1/Shared Documents/')/Files/add(url='Invoice.url',overwrite=true)"
      );
      expect(addLinkInit.body).toBe(
        '[InternetShortcut]\nURL=https://webclient.example.com/entry/123'
      );
      expect(addLinkInit.headers).toEqual(
        expect.objectContaining({ 'X-RequestDigest': 'digest-123' })
      );
    });

    test('keeps the pre-update file name when tryUpdateFileNameAsync fails', async () => {
      const spFileMetadata = makeSpFileMetadata({
        documentName: '',
        fileName: 'Invoice.pdf',
      });
      const validRepoClient = makeValidRepoClient();
      validRepoClient.entriesClient.getEntry.mockRejectedValue(new Error('not found'));
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      const result = await instance.sendToLaserficheWithMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(result?.fileName).toBe('Invoice');
    });

    test('clears local storage and rethrows when importEntry fails', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const validRepoClient = makeValidRepoClient();
      const error = new Error('import failed');
      validRepoClient.entriesClient.importEntry.mockRejectedValue(error);
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await expect(
        instance.sendToLaserficheWithMappingAsync(
          {} as any,
          'https://webclient.example.com',
          'cust1'
        )
      ).rejects.toBe(error);
      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    });
  });

  describe('sendToLaserficheNoMappingAsync', () => {
    test('always sends the hardcoded parent entry id of 1', async () => {
      const spFileMetadata = makeSpFileMetadata({ entryId: '999' });
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheNoMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      const entryRequest = validRepoClient.entriesClient.importEntry.mock.calls[0][0];
      expect(entryRequest.entryId).toBe(1);
    });

    test('links to the root folder in the web client', async () => {
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(
        makeSpFileMetadata({ entryId: '999' }),
        validRepoClient
      );

      const result = await instance.sendToLaserficheNoMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
        '1',
        'https://webclient.example.com',
        true,
        'repo-1',
        'cust1'
      );
      expect(result?.folderLink).toBe('https://webclient.example.com/entry/123');
    });

    test('keeps the full file name including extension when tryUpdateFileNameAsync fails', async () => {
      const spFileMetadata = makeSpFileMetadata({ fileName: 'Invoice.pdf' });
      const validRepoClient = makeValidRepoClient();
      validRepoClient.entriesClient.getEntry.mockRejectedValue(new Error('not found'));
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      const result = await instance.sendToLaserficheNoMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(result?.fileName).toBe('Invoice.pdf');
    });

    test('clears local storage on success', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const validRepoClient = makeValidRepoClient();
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await instance.sendToLaserficheNoMappingAsync(
        {} as any,
        'https://webclient.example.com',
        'cust1'
      );

      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    });

    test('clears local storage and rethrows when importEntry fails', async () => {
      const spFileMetadata = makeSpFileMetadata();
      const validRepoClient = makeValidRepoClient();
      const error = new Error('import failed');
      validRepoClient.entriesClient.importEntry.mockRejectedValue(error);
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, validRepoClient);

      await expect(
        instance.sendToLaserficheNoMappingAsync({} as any, 'https://webclient.example.com', 'cust1')
      ).rejects.toBe(error);
      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    });
  });

  describe('deleteAndHandleSPFileAsync', () => {
    test('throws when the delete request fails', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: false,
        statusText: 'Forbidden',
      });

      await expect(instance.deleteAndHandleSPFileAsync()).rejects.toThrow(
        'An error occurred while deleting file: Forbidden'
      );
    });
  });

  describe('deleteSPFileAndReplaceWithLinkAsync', () => {
    test('replaces the file with a link when delete succeeds', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
        contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const replaceSpy = vi
        .spyOn(instance, 'replaceFileWithLinkAsync')
        .mockResolvedValue(undefined);
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: true,
        statusText: 'OK',
      });

      await instance.deleteSPFileAndReplaceWithLinkAsync('https://webclient.example.com/entry/123');

      expect(replaceSpy).toHaveBeenCalledWith('Invoice', 'https://webclient.example.com/entry/123');
    });

    test('throws and clears local storage without replacing when delete fails', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      const replaceSpy = vi.spyOn(instance, 'replaceFileWithLinkAsync');
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: false,
        statusText: 'Forbidden',
      });

      await expect(
        instance.deleteSPFileAndReplaceWithLinkAsync('https://webclient.example.com/entry/123')
      ).rejects.toThrow('An error occurred while replacing file with link: Forbidden');
      expect(replaceSpy).not.toHaveBeenCalled();
      expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    });
  });

  describe('replaceFileWithLinkAsync', () => {
    test('throws when the contextinfo request fails', async () => {
      const spFileMetadata = makeSpFileMetadata({
        contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: false,
        statusText: 'Unauthorized',
      });

      await expect(
        instance.replaceFileWithLinkAsync('Invoice', 'https://webclient.example.com/entry/123')
      ).rejects.toThrow('An error occurred while replacing file with link: Unauthorized');
    });
  });

  describe('createLinkAsync', () => {
    test('posts the shortcut file to the exact percent-encoded URL', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
        contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: true,
        statusText: 'OK',
      });

      await instance.createLinkAsync(
        'Invoice Report',
        'https://webclient.example.com/entry/123',
        'digest-123'
      );

      expect(globalThis.fetch).toHaveBeenCalledWith(
        "https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativeUrl('/sites/site1/Shared Documents/')/Files/add(url='Invoice%20Report.url',overwrite=true)",
        {
          method: 'POST',
          body: '[InternetShortcut]\nURL=https://webclient.example.com/entry/123',
          headers: {
            'content-type': 'text/plain',
            accept: 'application/json;odata=verbose',
            'X-RequestDigest': 'digest-123',
          },
        }
      );
    });

    test('throws the replace-with-link error when the request fails', async () => {
      const spFileMetadata = makeSpFileMetadata({
        fileName: 'Invoice.pdf',
        fileUrl: '/sites/site1/Shared Documents/Invoice.pdf',
        contextPageAbsoluteUrl: 'https://contoso.sharepoint.com/sites/site1',
      });
      const instance = new SaveDocumentToLaserfiche(spFileMetadata, makeValidRepoClient());
      (globalThis.fetch as Mock).mockResolvedValue({
        ok: false,
        statusText: 'Forbidden',
      });

      await expect(
        instance.createLinkAsync('Invoice', 'https://webclient.example.com/entry/123', 'digest-123')
      ).rejects.toThrow('An error occurred while replacing file with link: Forbidden');
    });
  });
});
