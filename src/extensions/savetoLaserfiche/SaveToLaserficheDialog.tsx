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
import {
  APIServerException,
  Entry,
} from '@laserfiche/lf-repository-api-client';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../repository-client/repository-client-types';
import { PathUtils } from '@laserfiche/lf-js-utils';

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

const ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME = 'An entry with the same name already exists in the specified folder. If you continue, Laserfiche will automatically rename the new document.';
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
      try {
        await SPComponentLoader.loadScript(ZONE_JS_URL);
        await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
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
              const entryWithPathExists = validRepoClient.entriesClient.getEntryByPath({
                repoId,
                fullPath: PathUtils.combinePaths(
                  entryInfo.fullPath,
                  props.spFileMetadata.fileName
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
            } catch (err) {
              const docDoesNotAlreadyExists = (err as APIServerException).statusCode === 404;
              if (docDoesNotAlreadyExists) {
                await continueSavingDocumentAsync(saveToLF);
              }
              else {
                throw err;
              }
            }
          } catch (err) {
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
      } catch (err) {
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
            src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
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
      <Confirmation cancelButtonText='Cancel' />
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
