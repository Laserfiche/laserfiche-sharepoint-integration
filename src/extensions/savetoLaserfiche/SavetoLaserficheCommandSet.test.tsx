// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// Each `@microsoft/sp-*` package resolves to its own stub file (see
// vitest.config.mts), so these three factories apply independently.
vi.mock('@microsoft/sp-listview-extensibility', () => ({
  BaseListViewCommandSet: class {
    context: any;
    tryGetCommand(): any {
      return undefined;
    }
  },
}));
vi.mock('@microsoft/sp-core-library', () => ({ Log: { info: vi.fn() } }));
vi.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } },
}));
vi.mock('./GetDocumentDataDialog', () => ({
  GetDocumentDataCustomDialog: vi.fn().mockImplementation(function () {
    return { show: vi.fn().mockResolvedValue(undefined) };
  }),
}));
vi.mock('../../Utils/CreateConfigurations', () => ({
  CreateConfigurations: {
    ensureAdminConfigListCreatedAsync: vi.fn().mockResolvedValue(undefined),
  },
}));

import type { Mock } from 'vitest';
import SendToLfCommandSet from './SavetoLaserficheCommandSet';
import { CreateConfigurations } from '../../Utils/CreateConfigurations';
import { GetDocumentDataCustomDialog } from './GetDocumentDataDialog';
import { LASERFICHE_SIGNIN_PAGE_NAME, SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';

function makeContext(): any {
  return {
    httpClient: { get: vi.fn() },
    pageContext: {
      list: { title: 'Documents' },
      web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/test' },
    },
  };
}

function makeRow(overrides: Record<string, any> = {}): any {
  const values: Record<string, any> = {
    ID: '42',
    File_x0020_Size: 1000,
    FileRef: '/sites/test/Documents/file.pdf',
    FileLeafRef: 'file.pdf',
    CheckoutUser: '',
    ContentType: 'Document',
    ...overrides,
  };
  return { getValueByName: vi.fn((name: string) => values[name]) };
}

afterEach(() => {
  vi.clearAllMocks();
});

describe('onInit', () => {
  test('clears the stashed SharePoint document data and ensures the admin config list exists', async () => {
    const removeItemMock = vi.fn();
    Object.defineProperty(window, 'localStorage', {
      value: { removeItem: removeItemMock },
      writable: true,
    });

    const commandSet = new SendToLfCommandSet();
    (commandSet as any).context = makeContext();

    await commandSet.onInit();

    expect(removeItemMock).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    expect(CreateConfigurations.ensureAdminConfigListCreatedAsync).toHaveBeenCalledWith(
      commandSet.context
    );
  });
});

describe('onListViewUpdated', () => {
  let commandSet: SendToLfCommandSet;

  beforeEach(() => {
    commandSet = new SendToLfCommandSet();
    (commandSet as any).context = makeContext();
  });

  test('shows the command when exactly one non-folder row is selected', () => {
    const command: any = { visible: false };
    commandSet.tryGetCommand = vi.fn().mockReturnValue(command);

    commandSet.onListViewUpdated({
      selectedRows: [makeRow({ ContentType: 'Document' })],
    } as any);

    expect(command.visible).toBe(true);
  });

  test('hides the command when the single selected row is a Folder', () => {
    const command: any = { visible: true };
    commandSet.tryGetCommand = vi.fn().mockReturnValue(command);

    commandSet.onListViewUpdated({
      selectedRows: [makeRow({ ContentType: 'Folder' })],
    } as any);

    expect(command.visible).toBe(false);
  });

  test('hides the command when no rows are selected', () => {
    const command: any = { visible: true };
    commandSet.tryGetCommand = vi.fn().mockReturnValue(command);

    commandSet.onListViewUpdated({ selectedRows: [] } as any);

    expect(command.visible).toBe(false);
  });

  test('hides the command when more than one row is selected', () => {
    const command: any = { visible: true };
    commandSet.tryGetCommand = vi.fn().mockReturnValue(command);

    commandSet.onListViewUpdated({
      selectedRows: [makeRow({ ContentType: 'Document' }), makeRow({ ContentType: 'Document' })],
    } as any);

    expect(command.visible).toBe(false);
  });

  test('does not throw when the command cannot be found', () => {
    commandSet.tryGetCommand = vi.fn().mockReturnValue(undefined);

    expect(() =>
      commandSet.onListViewUpdated({
        selectedRows: [makeRow({ ContentType: 'Document' })],
      } as any)
    ).not.toThrow();
  });
});

describe('onExecute', () => {
  let commandSet: SendToLfCommandSet;

  beforeEach(() => {
    commandSet = new SendToLfCommandSet();
    (commandSet as any).context = makeContext();
    commandSet.hasSignInPage = true;
    commandSet.pageConfigurationCheck = vi.fn().mockResolvedValue(undefined);
    commandSet.trySaveToLaserficheAsync = vi.fn().mockResolvedValue(undefined);
    window.alert = vi.fn();
  });

  test('alerts and does not save when the selected item is a Folder', async () => {
    await commandSet.onExecute({
      selectedRows: [makeRow({ ContentType: 'Folder' })],
    } as any);

    expect(window.alert).toHaveBeenCalledWith(expect.stringContaining('Cannot Send a Folder'));
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('alerts and does not save when the filename is empty', async () => {
    await commandSet.onExecute({
      selectedRows: [makeRow({ FileLeafRef: '' })],
    } as any);

    expect(window.alert).toHaveBeenCalledWith(expect.stringContaining('add a filename'));
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('alerts and does not save for .url files', async () => {
    await commandSet.onExecute({
      selectedRows: [makeRow({ FileLeafRef: 'notes.url' })],
    } as any);

    expect(window.alert).toHaveBeenCalledWith(expect.stringContaining('Cannot send the .url file'));
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('alerts and does not save when the file is checked out', async () => {
    await commandSet.onExecute({
      selectedRows: [makeRow({ CheckoutUser: 'i:0#.f|membership|user@contoso.com' })],
    } as any);

    expect(window.alert).toHaveBeenCalledWith(expect.stringContaining('checked out'));
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('alerts and does not save when the file is over 100MB', async () => {
    await commandSet.onExecute({
      selectedRows: [makeRow({ File_x0020_Size: 200000000 })],
    } as any);

    expect(window.alert).toHaveBeenCalledWith(expect.stringContaining('below 100MB'));
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('alerts and does not save when the LaserficheSignIn page is missing', async () => {
    commandSet.hasSignInPage = false;

    await commandSet.onExecute({ selectedRows: [makeRow()] } as any);

    expect(window.alert).toHaveBeenCalled();
    expect(commandSet.trySaveToLaserficheAsync).not.toHaveBeenCalled();
  });

  test('saves to Laserfiche on the happy path', async () => {
    await commandSet.onExecute({ selectedRows: [makeRow()] } as any);

    expect(commandSet.trySaveToLaserficheAsync).toHaveBeenCalledWith({
      fileName: 'file.pdf',
      spContentType: 'Document',
      spFileUrl: '/sites/test/Documents/file.pdf',
      fileId: '42',
    });
    expect(window.alert).not.toHaveBeenCalled();
  });

  test('looks up the content type when ContentType is empty, then saves', async () => {
    (commandSet.context.httpClient.get as Mock).mockImplementation((url: string) => {
      if (url.includes('/ContentType')) {
        return Promise.resolve({
          json: () => Promise.resolve({ Name: 'Document' }),
        });
      }
      return Promise.reject(new Error('unexpected url'));
    });

    await commandSet.onExecute({
      selectedRows: [makeRow({ ContentType: '' })],
    } as any);

    expect(commandSet.trySaveToLaserficheAsync).toHaveBeenCalledWith({
      fileName: 'file.pdf',
      spContentType: 'Document',
      spFileUrl: '/sites/test/Documents/file.pdf',
      fileId: '42',
    });
    expect(window.alert).not.toHaveBeenCalled();
  });

  test('resolves an undefined content type without throwing when the lookup request fails', async () => {
    (commandSet.context.httpClient.get as Mock).mockRejectedValue(new Error('network error'));

    await expect(
      commandSet.onExecute({
        selectedRows: [makeRow({ ContentType: '' })],
      } as any)
    ).resolves.toBeUndefined();

    expect(commandSet.trySaveToLaserficheAsync).toHaveBeenCalledWith({
      fileName: 'file.pdf',
      spContentType: undefined,
      spFileUrl: '/sites/test/Documents/file.pdf',
      fileId: '42',
    });
  });
});

describe('pageConfigurationCheck', () => {
  let commandSet: SendToLfCommandSet;

  beforeEach(() => {
    commandSet = new SendToLfCommandSet();
    (commandSet as any).context = makeContext();
  });

  test('sets hasSignInPage to true when the Site Pages list contains LaserficheSignIn', async () => {
    window.fetch = vi.fn().mockResolvedValue({
      json: () => Promise.resolve({ value: [{ Title: LASERFICHE_SIGNIN_PAGE_NAME }] }),
    }) as any;

    await commandSet.pageConfigurationCheck();

    expect(commandSet.hasSignInPage).toBe(true);
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining('Site Pages'),
      expect.any(Object)
    );
  });

  test('leaves hasSignInPage false when no page title matches', async () => {
    window.fetch = vi.fn().mockResolvedValue({
      json: () => Promise.resolve({ value: [{ Title: 'SomeOtherPage' }] }),
    }) as any;

    await commandSet.pageConfigurationCheck();

    expect(commandSet.hasSignInPage).toBe(false);
  });

  test('does not throw and leaves hasSignInPage unchanged when fetch rejects', async () => {
    commandSet.hasSignInPage = true;
    window.fetch = vi.fn().mockRejectedValue(new Error('network error')) as any;

    await expect(commandSet.pageConfigurationCheck()).resolves.toBeUndefined();

    expect(commandSet.hasSignInPage).toBe(true);
  });
});

describe('trySaveToLaserficheAsync', () => {
  test('opens the GetDocumentDataCustomDialog and awaits show()', async () => {
    const commandSet = new SendToLfCommandSet();
    (commandSet as any).context = makeContext();
    const spFileInfo = {
      fileName: 'file.pdf',
      spContentType: 'Document',
      spFileUrl: '/sites/test/Documents/file.pdf',
      fileId: '42',
    };

    await commandSet.trySaveToLaserficheAsync(spFileInfo);

    expect(GetDocumentDataCustomDialog).toHaveBeenCalledWith(spFileInfo, commandSet.context);
    const instance = (GetDocumentDataCustomDialog as unknown as Mock).mock.results[0].value;
    expect(instance.show).toHaveBeenCalled();
  });
});
