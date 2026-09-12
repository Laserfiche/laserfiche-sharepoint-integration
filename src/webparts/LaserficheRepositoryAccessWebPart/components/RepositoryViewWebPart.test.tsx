// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

jest.mock('@laserfiche/lf-repository-api-client-v2', () => {
  return {
    EntryType: {
      Folder: 'Folder',
      Shortcut: 'Shortcut',
      RecordSeries: 'RecordSeries',
      Document: 'Document',
    },
  };
});

jest.mock('../../../Utils/Funcs', () => ({
  getEntryWebAccessUrl: jest.fn(),
}));

import {
  LfRepoTreeNode,
  LfRepoTreeNodeService,
} from '@laserfiche/lf-ui-components-services';
import * as React from 'react';
import { render, waitFor, screen } from '@testing-library/react';
import '@testing-library/jest-dom';
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
} from '../../strings';
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
    const LfRepoTreeNodeServiceMock = LfRepoTreeNodeService as jest.Mock;
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
    const isNodeSelectableResult = isNodeSelectable(
      recordSeriesShortcutTreeNode
    );

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
});

describe('openNode', () => {
  let setParentItem: jest.Mock;
  let props: {
    repoClient: IRepositoryApiClientExInternal;
    webClientUrl: string;
    loggedIn: boolean;
  };

  beforeEach(() => {
    setParentItem = jest.fn();
    props = {
      repoClient: {
        getCurrentRepoId: jest.fn().mockResolvedValue('repoId'),
      } as unknown as IRepositoryApiClientExInternal,
      webClientUrl: 'http://webclient.url',
      loggedIn: true,
    };
  });

  test('should set parent item if entryType is Folder', async () => {
    // Arrante
    const openedNode = { entryType: EntryType.Folder } as LfRepoTreeNode;
    window.open = jest.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).toHaveBeenCalledWith(openedNode);
    expect(window.open).not.toHaveBeenCalled();
  });

  test('should set parent item if entryType is RecordSeries', async () => {
    // Arrange
    const openedNode = { entryType: EntryType.RecordSeries } as LfRepoTreeNode;
    window.open = jest.fn();

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

    window.open = jest.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).not.toHaveBeenCalled();
    expect(props.repoClient.getCurrentRepoId).toHaveBeenCalled();
    expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
      'nodeId',
      'http://webclient.url',
      false,
      'repoId'
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

    window.open = jest.fn();

    // Act
    await openNode(openedNode, setParentItem, props);

    // Assert
    expect(setParentItem).not.toHaveBeenCalled();
    expect(props.repoClient.getCurrentRepoId).toHaveBeenCalled();
    expect(getEntryWebAccessUrl).toHaveBeenCalledWith(
      'nodeId',
      'http://webclient.url',
      false,
      'repoId'
    );
    expect(window.open).toHaveBeenCalled();
  });
});
