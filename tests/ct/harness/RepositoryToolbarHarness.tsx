import * as React from 'react';
import type { LfRepoTreeNode } from '@laserfiche/lf-ui-components-services';
import { RepositoryBrowserToolbar } from '../../../src/webparts/LaserficheRepositoryAccessWebPart/components/RepositoryViewWebPart';
import { makeRepoClient, RepoClientScenario } from './repoClient';

export interface RepositoryToolbarHarnessProps {
  scenario?: RepoClientScenario;
  parentEntryType?: 'Folder' | 'RecordSeries';
  parentId?: string;
  selectedId?: string;
  hasParent?: boolean;
  importDelayMs?: number;
  fieldContainerShouldFailInit?: boolean;
  templatesLoadDelayMs?: number;
}

export default function RepositoryToolbarHarness(
  props: RepositoryToolbarHarnessProps
): JSX.Element {
  const repoClient = makeRepoClient(props.scenario ?? 'default', props.importDelayMs);
  // Set before the modal mounts (and therefore before its effect calls the
  // fake <lf-field-container>'s initAsync), so there's no race between this
  // flag and that call -- see tests/ct/playwright/index.ts.
  window.__fieldContainerShouldFailInit = props.fieldContainerShouldFailInit ?? false;
  window.__templatesLoadDelayMs = props.templatesLoadDelayMs ?? 0;

  const parentItem: LfRepoTreeNode | undefined =
    props.hasParent === false
      ? undefined
      : ({
          id: props.parentId ?? '1',
          entryType: props.parentEntryType ?? 'Folder',
          isContainer: true,
          name: 'Parent Folder',
          path: '/Parent Folder',
        } as LfRepoTreeNode);

  const selectedItem: LfRepoTreeNode | undefined = props.selectedId
    ? ({
        id: props.selectedId,
        entryType: 'Document',
        isContainer: false,
        name: 'Selected Item',
        path: '/Parent Folder/Selected Item',
      } as LfRepoTreeNode)
    : undefined;

  return (
    <RepositoryBrowserToolbar
      repoClient={repoClient}
      webClientUrl='https://webclient.example.com'
      customerId='customer-1'
      selectedItem={selectedItem}
      parentItem={parentItem}
      loggedIn={true}
      refreshFolderBrowserAsync={async () => {
        window.__repoClientCalls = window.__repoClientCalls ?? [];
        window.__repoClientCalls.push({ method: 'refreshFolderBrowserAsync', args: [] });
      }}
    />
  );
}
