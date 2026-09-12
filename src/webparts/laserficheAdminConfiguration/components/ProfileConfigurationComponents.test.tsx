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

import {
  LfRepoTreeNode,
  LfRepoTreeNodeService,
} from '@laserfiche/lf-ui-components-services';
import * as React from 'react';
import { render, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import {
  isNodeSelectable,
  RepositoryBrowserModal,
} from './ProfileConfigurationComponents';
import { EntryType } from '@laserfiche/lf-repository-api-client-v2';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';

describe('ProfileConfigurationComponents', () => {
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
      <RepositoryBrowserModal
        CloseFolderBrowserUp={jest.fn()}
        SelectFolder={jest.fn()}
        selectedEntryNodePath=''
        repoClient={repoClient}
      />
    );

    // Assert
    await waitFor(() => {
      expect(setViewableEntryTypes).toEqual([
        EntryType.Folder,
        EntryType.Shortcut,
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

  test('isNodeSelectable should return false for Document node', async () => {
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
    expect(isNodeSelectableResult).toBe(false);
  });

  test('isNodeSelectable should return true for Shortcut to Folder node', async () => {
    // Arrange
    const folderShortcutTreeNode: LfRepoTreeNode = {
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
    const isNodeSelectableResult = isNodeSelectable(folderShortcutTreeNode);

    // Assert
    expect(isNodeSelectableResult).toBe(true);
  });

  test('isNodeSelectable should return false for Shortcut to Document node', async () => {
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
    expect(isNodeSelectableResult).toBe(false);
  });

  test('isNodeSelectable should return false for Shortcut to RecordSeries node', async () => {
    // Arrange
    const shortcutRecordSeriesTreeNode: LfRepoTreeNode = {
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
      shortcutRecordSeriesTreeNode
    );

    // Assert
    expect(isNodeSelectableResult).toBe(false);
  });

  test('isNodeSelectable should return false for RecordSeries node', async () => {
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
    expect(isNodeSelectableResult).toBe(false);
  });
});
