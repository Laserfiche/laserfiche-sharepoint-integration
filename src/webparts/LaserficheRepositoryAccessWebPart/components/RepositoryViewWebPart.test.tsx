// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('@laserfiche/lf-repository-api-client-v2', () => {
  return {
    EntryType: {
      Folder: 'Folder',
      Shortcut: 'Shortcut',
      RecordSeries: 'RecordSeries',
      Document: 'Document',
    },
  };
});

vi.mock('../../../Utils/Funcs', () => ({
  getEntryWebAccessUrl: vi.fn(),
}));

import type { Mock } from 'vitest';
import { LfRepoTreeNode, LfRepoTreeNodeService } from '@laserfiche/lf-ui-components-services';
import * as React from 'react';
import { fireEvent, render, waitFor, screen, within } from '@testing-library/react';
import RepositoryViewWebPart, {
  isNodeSelectable,
  openNode,
  RepositoryBrowserToolbar,
} from './RepositoryViewWebPart';
import { EntryType } from '@laserfiche/lf-repository-api-client-v2';
import { getEntryWebAccessUrl } from '../../../Utils/Funcs';
import {
  CANNOT_IMPORT_INTO_RECORD_SERIES,
  UPLOAD_FILE_TO_LASERFICHE,
  UPLOAD_FILE_TO_LASERFICHE_TITLE,
} from '../../strings';
import { LASERFICHE_ICON_URL } from '../../constants';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';

describe('RepositoryViewWebPart', () => {
  let repoClient: IRepositoryApiClientExInternal;

  beforeEach(async () => {
    repoClient = {
      // Mock the necessary methods and properties of repoClient
    } as IRepositoryApiClientExInternal;
  });

  test('should assign correct entryTypes to viewableEntryTypes', async () => {
    // Arrange
    let setViewableEntryTypes: EntryType[] = [];
    const LfRepoTreeNodeServiceMock = LfRepoTreeNodeService as Mock;
    LfRepoTreeNodeServiceMock.mockImplementation(function () {
      return {
        get viewableEntryTypes() {
          return setViewableEntryTypes;
        },
        set viewableEntryTypes(value) {
          setViewableEntryTypes = value;
        },
      };
    });

    // Act
    render(
      <RepositoryViewWebPart
        repoClient={repoClient}
        webClientUrl={''}
        customerId={''}
        loggedIn={false}
      />
    );

    // Assert
    await waitFor(() => {
      expect(setViewableEntryTypes).toEqual([
        EntryType.Folder,
        EntryType.Shortcut,
        EntryType.Document,
        EntryType.RecordSeries,
      ]);
    });
  });

  test('isNodeSelectable should return true for Folder node', async () => {
    // Arrange
    const folderTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.Folder,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(folderTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return true for Document node', async () => {
    // Arrange
    const documentTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.Document,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(documentTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return true for Shortcut to Folder node', async () => {
    // Arrange
    const shortcutFolderTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.Shortcut,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      targetType: EntryType.Folder,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(shortcutFolderTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return true for Shortcut to Document node', async () => {
    // Arrange
    const documentShortcutTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.Shortcut,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      targetType: EntryType.Document,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(documentShortcutTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return true for Shortcut to RecordSeries node', async () => {
    // Arrange
    const recordSeriesShortcutTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.Shortcut,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      targetType: EntryType.RecordSeries,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(recordSeriesShortcutTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return true for RecordSeries node', async () => {
    // Arrange
    const recordSeriesTreeNode: LfRepoTreeNode = {
      id: '1',
      entryType: EntryType.RecordSeries,
      icon: 'folder',
      isContainer: true,
      isLeaf: false,
      name: 'Test Folder',
      path: '/Test Folder',
    };

    // Act
    const isNodeSelectableResult = isNodeSelectable(recordSeriesTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('button is disabled when parentItem.entryType is RecordSeries', () => {
    // Arrange
    const props = {
      parentItem: {
        entryType: EntryType.RecordSeries,
      } as LfRepoTreeNode,
      repoClient,
      webClientUrl: '',
      customerId: '',
      selectedItem: {} as LfRepoTreeNode,
      loggedIn: false,
      refreshFolderBrowserAsync: async () => {},
    };

    // Act
    render(<RepositoryBrowserToolbar {...props} />);

    // Assert
    const button = screen.getByTitle(CANNOT_IMPORT_INTO_RECORD_SERIES);

    expect(button).toBeDisabled();
  });

  test('button is enabled when parentItem.entryType is not RecordSeries', () => {
    // Arrange
    const props = {
      parentItem: {
        entryType: EntryType.Folder,
      } as LfRepoTreeNode,
      repoClient,
      webClientUrl: '',
      customerId: '',
      selectedItem: {} as LfRepoTreeNode,
      loggedIn: false,
      refreshFolderBrowserAsync: async () => {},
    };

    // Act
    render(<RepositoryBrowserToolbar {...props} />);

    // Assert
    const button = screen.getByTitle(UPLOAD_FILE_TO_LASERFICHE);

    expect(button).toBeEnabled();
  });

  // No repoClient, so the dialog skips loading templates and tags.
  function openUploadDialog(): void {
    render(
      <RepositoryBrowserToolbar
        parentItem={{ entryType: EntryType.Folder } as LfRepoTreeNode}
        repoClient={undefined}
        webClientUrl=''
        customerId=''
        selectedItem={{} as LfRepoTreeNode}
        loggedIn={false}
        refreshFolderBrowserAsync={async () => {}}
      />
    );
    fireEvent.click(screen.getByTitle(UPLOAD_FILE_TO_LASERFICHE));
  }

  test('upload dialog title shows the Laserfiche icon', () => {
    // Act
    openUploadDialog();

    // Assert
    const header = screen
      .getByText(UPLOAD_FILE_TO_LASERFICHE_TITLE)
      .closest('.modal-header') as HTMLElement;
    expect(within(header).getByRole('img')).toHaveAttribute('src', LASERFICHE_ICON_URL);
  });

  // The real <lf-field-container> isn't loaded in tests, so the attributes it
  // is given are as close to its collapsed panels as this suite can get.
  test('upload dialog starts with the Template and Fields sections collapsed', () => {
    // Act
    openUploadDialog();

    // Assert
    const fieldContainer = document.querySelector('lf-field-container');
    expect(fieldContainer).toHaveAttribute('collapsible', 'true');
    expect(fieldContainer).toHaveAttribute('start_collapsed', 'true');
  });

  test('upload dialog has no import options', () => {
    // Act
    openUploadDialog();

    // Assert
    expect(screen.getByText(UPLOAD_FILE_TO_LASERFICHE_TITLE)).toBeVisible();
    expect(screen.queryByText('Import options')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Generate text')).not.toBeInTheDocument();
  });
});

describe('openNode', () => {
  let setParentItem: Mock;
  let props: {
    repoClient: IRepositoryApiClientExInternal;
    webClientUrl: string;
    customerId: string;
    loggedIn: boolean;
  };

  beforeEach(() => {
    setParentItem = vi.fn();
    props = {
      repoClient: {
        getCurrentRepoId: vi.fn().mockResolvedValue('repoId'),
      } as unknown as IRepositoryApiClientExInternal,
      webClientUrl: 'http://webclient.url',
      customerId: 'customerId',
      loggedIn: true,
    };
  });

  test('should set parent item if entryType is Folder', async () => {
    // Arrante
    const openedNode = { entryType: EntryType.Folder } as LfRepoTreeNode;
    window.open = vi.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).toHaveBeenCalledWith(openedNode);
    expect(window.open).not.toHaveBeenCalled();
  });

  test('should set parent item if entryType is RecordSeries', async () => {
    // Arrange
    const openedNode = { entryType: EntryType.RecordSeries } as LfRepoTreeNode;
    window.open = vi.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).toHaveBeenCalledWith(openedNode);
    expect(window.open).not.toHaveBeenCalled();
  });

  test('should open web client URL if entryType is Document', async () => {
    // Arrange
    const openedNode = {
      entryType: EntryType.Document,
      id: 'nodeId',
      isContainer: false,
    } as LfRepoTreeNode;

    window.open = vi.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).not.toHaveBeenCalled();
    expect(props.repoClient.getCurrentRepoId).toHaveBeenCalled();
    expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
      'nodeId',
      'http://webclient.url',
      false,
      'repoId',
      'customerId'
    );
    expect(window.open).toHaveBeenCalled();
  });

  test('should open web client URL if entryType is Shortcut to Document', async () => {
    // Arrange
    const openedNode = {
      entryType: EntryType.Shortcut,
      targetType: EntryType.Document,
      id: 'nodeId',
      isContainer: false,
    } as LfRepoTreeNode;

    window.open = vi.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).not.toHaveBeenCalled();
    expect(props.repoClient.getCurrentRepoId).toHaveBeenCalled();
    expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
      'nodeId',
      'http://webclient.url',
      false,
      'repoId',
      'customerId'
    );
    expect(window.open).toHaveBeenCalled();
  });
});
