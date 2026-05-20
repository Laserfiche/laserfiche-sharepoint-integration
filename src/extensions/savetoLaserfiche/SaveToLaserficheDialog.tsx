// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { ISPDocumentData } from '../../Utils/Types';
import {
  clientId,
  LF_UI_COMPONENTS_URL,
  SP_LOCAL_STORAGE_KEY,
  ZONE_JS_URL,
} from '../../webparts/constants';
import LoadingDialog, {
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
import { getRegion } from '../../Utils/Funcs';
import { Entry } from '@laserfiche/lf-repository-api-client';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { PathUtils } from '@laserfiche/lf-js-utils';
import { CANCEL, DOCUMENT_ALREADY_EXISTS } from '../../webparts/strings';

export default class SaveToLaserficheCustomDialog extends BaseDialog {
  successful = false;

  handleSuccessSave: (successful: boolean) => void = (successful: boolean) => {
    this.successful = successful;
  };

  closeClick: (success?: SavedToLaserficheDocumentData) => Promise<void> =
    async (success?: SavedToLaserficheDocumentData) => {
      await this.close();
      if (this.closeParent) {
        await this.closeParent(success);
      }
    };

  constructor(
    private spFileData: ISPDocumentData,
    private closeParent?: (
      success?: SavedToLaserficheDocumentData
    ) => Promise<void>
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

  protected async onAfterClose(): Promise<void> {
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
  const loginComponent = React.createRef<
    NgElement & WithProperties<LfLoginComponent>
  >();

  const region = getRegion();
  const [success, setSuccess] = React.useState<
    SavedToLaserficheDocumentData | undefined
  >();
  const [error, setError] = React.useState<JSX.Element | undefined>();
  const [showSaveTo, setShowSaveTo] = React.useState<boolean>(true);
  const [getConfirmation, Confirmation] = useConfirm();

  const saveToDialogCloseClick: () => Promise<void> = async () => {
    await props.closeClick(success);
  };

  async function tryGetValidRepositoryClientAsync(): Promise<IRepositoryApiClientExInternal> {
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

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      await SPComponentLoader.loadScript(ZONE_JS_URL);
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
      try {
        if (loginComponent.current?.authorization_credentials) {
          const validRepoClient = await tryGetValidRepositoryClientAsync();
          const saveToLF = new SaveDocumentToLaserfiche(
            props.spFileMetadata,
            validRepoClient
          );
          try {
            try {
              const repoId = await validRepoClient.getCurrentRepoId();
              const entryInfo: Entry =
                await validRepoClient.entriesClient.getEntry({
                  repoId,
                  entryId: Number.parseInt(props.spFileMetadata.entryId, 10),
                });
              const entryWithPathExists =
                await validRepoClient.entriesClient.getEntryByPath({
                  repoId,
                  fullPath: PathUtils.combinePaths(
                    entryInfo.fullPath,
                    PathUtils.removeFileExtension(props.spFileMetadata.fileName)
                  ),
                });
              if (entryWithPathExists) {
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
      <div
        className={`${styles.header}${
          showSaveTo ? '' : ` ${styles.hideImport}`
        }`}
      >
        <div className={styles.logoHeader}>
          <lf-login
            hidden
            redirect_uri=''
            authorize_url_host_name={region}
            redirect_behavior='Replace'
            client_id={clientId}
            ref={loginComponent}
          />
          <img
            src='data:image/png;base64,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'
            width='30'
            height='30'
          />
          <span className={styles.paddingLeft}>Laserfiche</span>
        </div>

        <button
          className={styles.lfCloseButton}
          title='close'
          onClick={saveToDialogCloseClick}
        >
          <span className='material-icons-outlined'> close </span>
        </button>
      </div>

      <div
        className={`${styles.contentBox}${
          showSaveTo ? '' : ` ${styles.hideImport}`
        }`}
      >
        {!success && !error && <LoadingDialog />}
        {success && (
          <SavedToLaserficheSuccessDialogText successfulSave={success} />
        )}
        {error && (
          <span>
            {`Error saving:`} {error}
          </span>
        )}
      </div>

      <div
        className={`${styles.footer}${
          showSaveTo ? '' : ` ${styles.hideImport}`
        }`}
      >
        <SavedToLaserficheSuccessDialogButtons
          successfulSave={success}
          closeClick={saveToDialogCloseClick}
        />
      </div>
      <Confirmation
        cancelButtonText={CANCEL}
        headerText={DOCUMENT_ALREADY_EXISTS}
      />
    </div>
  );

  async function continueSavingDocumentAsync(
    saveToLF: SaveDocumentToLaserfiche
  ): Promise<void> {
    const successSaveToLF = await saveToLF.trySaveDocumentToLaserficheAsync();
    props.isSuccessfulLoggedIn(true);
    setSuccess(successSaveToLF);
  }
}
