// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { ISPDocumentData } from '../../Utils/Types';
import { LF_UI_COMPONENTS_URL, SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';
import LoadingDialog, {
  LaserficheDialogTitle,
  SavedToLaserficheSuccessDialogButtons,
  SavedToLaserficheSuccessDialogText,
  useConfirm,
} from './CommonDialogs';
import {
  SaveDocumentToLaserfiche,
  SavedToLaserficheDocumentData,
} from './SaveDocumentToLaserfiche';
import styles from './SendToLaserFiche.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as ReactDOM from 'react-dom';
import { BaseDialog } from '@microsoft/sp-dialog';
import { LaserficheLogin } from '../../Utils/LaserficheLogin';
import { Entry } from '@laserfiche/lf-repository-api-client-v2';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { CANCEL, DOCUMENT_ALREADY_EXISTS, LASERFICHE } from '../../webparts/strings';

export default class SaveToLaserficheCustomDialog extends BaseDialog {
  successful = false;

  handleSuccessSave: (successful: boolean) => void = (successful: boolean) => {
    this.successful = successful;
  };

  closeClick: (success?: SavedToLaserficheDocumentData) => Promise<void> = async (
    success?: SavedToLaserficheDocumentData
  ) => {
    await this.close();
    if (this.closeParent) {
      await this.closeParent(success);
    }
  };

  constructor(
    private spFileData: ISPDocumentData,
    private closeParent?: (success?: SavedToLaserficheDocumentData) => Promise<void>
  ) {
    super();
  }

  public render(): void {
    const element: React.ReactElement = (
      <React.StrictMode>
        <SaveToLaserficheDialog
          spFileMetadata={this.spFileData}
          isSuccessfulLoggedIn={this.handleSuccessSave}
          closeClick={this.closeClick}
        />
      </React.StrictMode>
    );
    ReactDOM.render(element, this.domElement);
  }

  protected override async onAfterClose(): Promise<void> {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
    if (this.closeParent) {
      await this.closeParent();
    }
  }
}

const ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME =
  'An entry with the same name already exists in the specified folder. If you continue, Laserfiche will automatically rename the new document.';
function SaveToLaserficheDialog(props: {
  isSuccessfulLoggedIn: (success: boolean) => void;
  closeClick: (success?: SavedToLaserficheDocumentData) => Promise<void>;
  spFileMetadata: ISPDocumentData;
}): JSX.Element {
  // useRef, not createRef: createRef hands back a new ref on every render and
  // React nulls the old one, which the mount effect below would then read.
  const loginComponent = React.useRef<NgElement & WithProperties<LfLoginComponent>>();

  const [success, setSuccess] = React.useState<SavedToLaserficheDocumentData | undefined>();
  const [error, setError] = React.useState<JSX.Element | undefined>();
  const [showSaveTo, setShowSaveTo] = React.useState<boolean>(true);
  const [getConfirmation, Confirmation] = useConfirm();

  const saveToDialogCloseClick: () => Promise<void> = async () => {
    await props.closeClick(success);
  };

  async function tryGetValidRepositoryClientAsync(): Promise<
    IRepositoryApiClientExInternal | undefined
  > {
    const repoClientCreator = new RepositoryClientExInternal();
    const newRepoClient = await repoClientCreator.createRepositoryClientAsync();
    try {
      // test accessToken validity
      await newRepoClient.repositoriesClient.listRepositories({});
    } catch {
      return undefined;
    }
    return newRepoClient;
  }

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
      try {
        // lf-login restores authorization_credentials from local storage, so
        // they can be present but no longer valid (expired or revoked, and the
        // refresh fails). tryGetValidRepositoryClientAsync then yields
        // undefined: treat that as signed out too, so the sign-in flow starts
        // instead of the save failing on an undefined client.
        const validRepoClient = loginComponent.current?.authorization_credentials
          ? await tryGetValidRepositoryClientAsync()
          : undefined;
        if (validRepoClient) {
          const saveToLF = new SaveDocumentToLaserfiche(props.spFileMetadata, validRepoClient);
          try {
            try {
              const repoId = await validRepoClient.getCurrentRepoId();
              const entryInfo: Entry = await validRepoClient.entriesClient.getEntry({
                repositoryId: repoId,
                entryId: Number.parseInt(props.spFileMetadata.entryId, 10),
              });
              const entryWithPath = await validRepoClient.entriesClient.getEntryByPath({
                repositoryId: repoId,
                fullPath: PathUtils.combinePaths(
                  entryInfo.fullPath,
                  PathUtils.removeFileExtension(props.spFileMetadata.fileName)
                ),
              });
              // v2 returns a GetEntryByPathResponse for a successful lookup, so
              // the response object itself is always truthy: test the entry.
              if (entryWithPath?.entry) {
                setShowSaveTo(false);
                const confirmSave = await getConfirmation(
                  ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME
                );
                if (confirmSave) {
                  setShowSaveTo(true);
                  await continueSavingDocumentAsync(saveToLF);
                } else {
                  window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
                  props.isSuccessfulLoggedIn(true);
                  await props.closeClick();
                }
              } else {
                // No entry at that path: same outcome as the 404 below.
                await continueSavingDocumentAsync(saveToLF);
              }
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
            } catch (err: any) {
              const docDoesNotAlreadyExists = err.status === 404;
              if (docDoesNotAlreadyExists) {
                await continueSavingDocumentAsync(saveToLF);
              } else {
                throw err;
              }
            }
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
          } catch (err: any) {
            if (err.status === 401 || err.status === 403) {
              props.isSuccessfulLoggedIn(false);
              await props.closeClick();
            } else if (err.status === 404) {
              props.isSuccessfulLoggedIn(true);
              setError(
                <>
                  <span>{err.message}.</span>
                  <div>{`Verify that an entry with ID "${props.spFileMetadata.entryId}" exists and that you have access to it.`}</div>
                </>
              );
              console.error(err);
            } else {
              props.isSuccessfulLoggedIn(true);
              setError(<span>{err.message}</span>);
              console.error(err);
            }
          }
        } else {
          props.isSuccessfulLoggedIn(false);
          await props.closeClick();
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        console.error(`Error initializing dialog: ${err}`);
      }
    };

    void initializeComponentAsync();
  }, []);

  return (
    <div className={styles.wrapper}>
      <LaserficheLogin ref={loginComponent} redirectUri='' />
      <div className={`${styles.header}${showSaveTo ? '' : ` ${styles.hideImport}`}`}>
        <LaserficheDialogTitle title={LASERFICHE} />

        <button className={styles.lfCloseButton} title='close' onClick={saveToDialogCloseClick}>
          <span className='material-icons-outlined'> close </span>
        </button>
      </div>

      <div className={`${styles.contentBox}${showSaveTo ? '' : ` ${styles.hideImport}`}`}>
        {!success && !error && <LoadingDialog />}
        {success && (
          <SavedToLaserficheSuccessDialogText successfulSave={success} action={success.action} />
        )}
        {error && (
          <span>
            {`Error saving:`} {error}
          </span>
        )}
      </div>

      <div className={`${styles.footer}${showSaveTo ? '' : ` ${styles.hideImport}`}`}>
        <SavedToLaserficheSuccessDialogButtons closeClick={saveToDialogCloseClick} />
      </div>
      <Confirmation cancelButtonText={CANCEL} headerText={DOCUMENT_ALREADY_EXISTS} />
    </div>
  );

  async function continueSavingDocumentAsync(saveToLF: SaveDocumentToLaserfiche): Promise<void> {
    const successSaveToLF = await saveToLF.trySaveDocumentToLaserficheAsync();
    if (!successSaveToLF) {
      // The save resolves undefined only when lf-login has no access token by
      // the time it runs: everything else on this path either returns a
      // document or throws. Treat it the way the 401/403 branch above does,
      // rather than reporting a save that never happened and leaving the
      // dialog on its loading spinner forever.
      props.isSuccessfulLoggedIn(false);
      await props.closeClick();
      return;
    }
    props.isSuccessfulLoggedIn(true);
    setSuccess(successSaveToLF);
  }
}
