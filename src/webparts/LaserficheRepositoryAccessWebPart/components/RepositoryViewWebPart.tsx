// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import {
  CreateEntryRequest,
  CreateEntryRequestEntryType,
  Entry,
  EntryType,
  FieldToUpdate,
  FileParameter,
  ImportEntryRequest,
  ImportEntryRequestMetadata,
} from '@laserfiche/lf-repository-api-client-v2';
import {
  LfRepoTreeNodeService,
  LfFieldsService,
  LfRepoTreeNode,
} from '@laserfiche/lf-ui-components-services';
import {
  ColumnDef,
  LfFieldContainerComponent,
  LfRepositoryBrowserComponent,
} from '@laserfiche/types-lf-ui-components';
import { PathUtils } from '@laserfiche/lf-js-utils';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { ChangeEvent } from 'react';
import { getEntryWebAccessUrl, getErrorDetails } from '../../../Utils/Funcs';
import { SavedLaserficheDocument } from '../../../Utils/Types';
import styles from './LaserficheRepositoryAccess.module.scss';
import {
  DelayedSpinner,
  LaserficheDialogTitle,
  SavedToLaserficheSuccessDialog,
  useConfirm,
} from './../../../extensions/savetoLaserfiche/CommonDialogs';
import { LfTagsPicker } from './LfTagsPicker';
import {
  CANCEL,
  CANNOT_IMPORT_INTO_RECORD_SERIES,
  CLOSE,
  CREATE_FOLDER,
  DOCUMENT_ALREADY_EXISTS,
  ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME,
  FOLDER_NAME,
  GO_BACK,
  LASERFICHE_REPOSITORY_EXPLORER,
  LOADING,
  NAME,
  OK,
  PLEASE_SELECT_FILE_FOLDER_TO_OPEN,
  SUBMIT,
  UNABLE_TO_LOAD_TEMPLATES_AND_FIELDS,
  UNKNOWN_ERROR,
  UPLOAD_FILE_TO_LASERFICHE,
  UPLOAD_FILE_TO_LASERFICHE_TITLE,
  UPLOADING,
} from '../../strings';
import laserficheLogoUrl from './../../../Assets/Images/laserfiche-logo.png';
import waIconsUrl from './../../../Assets/Images/waicons.svg';
import './../../../Assets/CSS/commonStyles.css';

const cols: ColumnDef[] = [
  {
    id: 'creationTime',
    displayName: 'Creation Date',
    defaultWidth: '100px',
    resizable: true,
    sortable: true,
  },
  {
    id: 'lastModifiedTime',
    displayName: 'Last Modified Date',
    defaultWidth: '100px',
    resizable: true,
    sortable: true,
  },
  {
    id: 'pageCount',
    displayName: 'Pages',
    defaultWidth: '100px',
    resizable: true,
    sortable: true,
  },
  {
    id: 'templateName',
    displayName: 'Template Name',
    defaultWidth: '100px',
    resizable: true,
    sortable: true,
  },
];

const fileValidation = 'Please select the file to upload';
const fileSizeValidation = 'Please select a file below 100MB in size';
const fileNameValidation = 'Please provide a valid filename';
const fileNameWithBacklash = 'Please provide a valid filename without backslash';
const folderValidation = 'Please provide a folder name';
const folderBackslashNameValidation = 'Entry names cannot contain backslash';
const folderExists = 'Object already exists';
const requiredFieldsValidation = 'Please provide values for all required fields';

export const isNodeSelectable: (node: LfRepoTreeNode) => boolean = (node: LfRepoTreeNode) => {
  if (
    node?.entryType === EntryType.Folder ||
    node?.entryType === EntryType.Document ||
    node?.entryType === EntryType.RecordSeries
  ) {
    return true;
  } else if (
    (node?.entryType === EntryType.Shortcut && node?.targetType === EntryType.Folder) ||
    (node?.entryType === EntryType.Shortcut && node?.targetType === EntryType.Document) ||
    (node?.entryType === EntryType.Shortcut && node?.targetType === EntryType.RecordSeries)
  ) {
    return true;
  } else {
    return false;
  }
};

export default function RepositoryViewComponent(props: {
  repoClient: IRepositoryApiClientExInternal;
  webClientUrl: string;
  customerId: string;
  loggedIn: boolean;
}): JSX.Element {
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef<NgElement & WithProperties<LfRepositoryBrowserComponent>>();
  let lfRepoTreeService: LfRepoTreeNodeService;

  const [parentItem, setParentItem] = React.useState<LfRepoTreeNode | undefined>(undefined);
  const [selectedItem, setSelectedItem] = React.useState<LfRepoTreeNode | undefined>(undefined);

  React.useEffect(() => {
    const onEntrySelected: EventListener = (event: Event) => {
      const customEvent = event as CustomEvent<LfRepoTreeNode[] | undefined>;
      const selectedNode = customEvent.detail ? customEvent.detail[0] : undefined;
      setSelectedItem(selectedNode);
    };

    const onEntryOpened: EventListener = async (event: Event) => {
      const customEvent = event as CustomEvent<LfRepoTreeNode[] | undefined>;
      const openedNode = customEvent.detail ? customEvent.detail[0] : undefined;
      await openNode(openedNode, setParentItem, props);
    };

    const initializeTreeAsync: () => Promise<void> = async () => {
      const repoBrowser = repositoryBrowser.current;
      lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
      lfRepoTreeService.viewableEntryTypes = [
        EntryType.Folder,
        EntryType.Shortcut,
        EntryType.Document,
        EntryType.RecordSeries,
      ];
      repoBrowser?.addEventListener('entrySelected', onEntrySelected);
      repoBrowser?.addEventListener('entryDblClicked', onEntryOpened);
      if (lfRepoTreeService) {
        lfRepoTreeService.columnIds = [
          'creationTime',
          'lastModifiedTime',
          'pageCount',
          'templateName',
        ];
        try {
          await repoBrowser?.initAsync(lfRepoTreeService);
          setParentItem(repoBrowser?.currentFolder as LfRepoTreeNode);
          repoBrowser?.setColumnsToDisplay(cols);
          await repoBrowser?.refreshAsync();
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
        } catch (err: any) {
          console.error(err);
        }
      } else {
        console.debug('Unable to initialize tree, lfRepoTreeService is undefined');
      }
    };
    if (props.repoClient) {
      void initializeTreeAsync();
    }
  }, [props.repoClient, props.loggedIn]);

  const refreshFolderBrowserAsync: () => Promise<void> = async () => {
    await repositoryBrowser.current.refreshAsync(false);
  };

  return (
    <>
      <div>
        <main className='bg-white'>
          <div style={{ margin: '10px 0px' }}>
            <img style={{ width: '30px' }} src={laserficheLogoUrl} alt='' />
            <span className={styles.browserTitle}>{LASERFICHE_REPOSITORY_EXPLORER}</span>
          </div>
          {props.loggedIn && (
            <>
              <RepositoryBrowserToolbar
                repoClient={props.repoClient}
                selectedItem={selectedItem}
                parentItem={parentItem}
                loggedIn={props.loggedIn}
                webClientUrl={props.webClientUrl}
                customerId={props.customerId}
                refreshFolderBrowserAsync={refreshFolderBrowserAsync}
              />
              <div className={styles.repositoryBrowserContainer}>
                <lf-repository-browser
                  ref={repositoryBrowser}
                  ok_button_text='Okay'
                  cancel_button_text='Cancel'
                  multiple='false'
                  isSelectable={isNodeSelectable}
                />
              </div>
            </>
          )}
        </main>
      </div>
    </>
  );
}

export async function openNode(
  openedNode: LfRepoTreeNode,
  setParentItem: React.Dispatch<React.SetStateAction<LfRepoTreeNode>>,
  props: {
    repoClient: IRepositoryApiClientExInternal;
    webClientUrl: string;
    customerId: string;
    loggedIn: boolean;
  }
): Promise<void> {
  const entryType =
    openedNode.entryType === EntryType.Shortcut ? openedNode.targetType : openedNode.entryType;
  if (entryType === EntryType.Folder || entryType === EntryType.RecordSeries) {
    setParentItem(openedNode);
  } else {
    const repoId = await props.repoClient.getCurrentRepoId();

    if (openedNode?.id) {
      const webClientNodeUrl = getEntryWebAccessUrl(
        openedNode.id,
        props.webClientUrl,
        openedNode.isContainer,
        repoId,
        props.customerId
      );
      window.open(webClientNodeUrl);
    }
  }
}

export function RepositoryBrowserToolbar(props: {
  repoClient: IRepositoryApiClientExInternal;
  webClientUrl: string;
  customerId: string;
  selectedItem: LfRepoTreeNode;
  parentItem: LfRepoTreeNode;
  loggedIn: boolean;
  refreshFolderBrowserAsync: () => Promise<void>;
}): JSX.Element {
  const [showUploadModal, setShowUploadModal] = React.useState(false);
  const [uploadedDocument, setUploadedDocument] = React.useState<
    SavedLaserficheDocument | undefined
  >(undefined);
  const [showCreateModal, setShowCreateModal] = React.useState(false);
  const [showAlertModal, setShowAlertModal] = React.useState(false);

  const openNewFolderModal: () => void = () => {
    setShowCreateModal(true);
  };

  const openImportFileModal: () => void = () => {
    setShowUploadModal(true);
  };

  const openFileOrFolder: () => void = async () => {
    const repoId = await props.repoClient.getCurrentRepoId();

    if (props.selectedItem?.id) {
      const webClientNodeUrl = getEntryWebAccessUrl(
        props.selectedItem.id,
        props.webClientUrl,
        props.selectedItem.isContainer,
        repoId,
        props.customerId
      );
      window.open(webClientNodeUrl);
    } else if (props.parentItem?.id) {
      const webClientNodeUrl = getEntryWebAccessUrl(
        props.parentItem.id,
        props.webClientUrl,
        props.parentItem.isContainer,
        repoId,
        props.customerId
      );
      window.open(webClientNodeUrl);
    } else {
      setShowAlertModal(true);
    }
  };

  const confirmAlertButton: () => void = () => {
    setShowAlertModal(false);
  };

  const showUploadedDocument: (savedDocument: SavedLaserficheDocument) => void = (
    savedDocument: SavedLaserficheDocument
  ) => {
    setShowUploadModal(false);
    setUploadedDocument(savedDocument);
  };

  const closeUploadedDocument: () => Promise<void> = async () => {
    setUploadedDocument(undefined);
  };

  return (
    <>
      <div id='mainWebpartContent'>
        <div className={styles.buttonContainer}>
          <button
            className={styles.lfMaterialIconButton}
            title='Open entry in Laserfiche'
            onClick={openFileOrFolder}
          >
            <img className={styles.waIcon} src={`${waIconsUrl}#open`} alt='' />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title={
              props?.parentItem?.entryType === EntryType.RecordSeries
                ? CANNOT_IMPORT_INTO_RECORD_SERIES
                : UPLOAD_FILE_TO_LASERFICHE
            }
            disabled={props?.parentItem?.entryType === EntryType.RecordSeries}
            onClick={openImportFileModal}
          >
            <img className={styles.waIcon} src={`${waIconsUrl}#upload`} alt='' />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title='Create folder in Laserfiche'
            onClick={openNewFolderModal}
          >
            <img className={styles.waIcon} src={`${waIconsUrl}#add-folder`} alt='' />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title='Refresh Laserfiche folder'
            onClick={props.refreshFolderBrowserAsync}
          >
            <img className={styles.waIcon} src={`${waIconsUrl}#refresh`} alt='' />
          </button>
        </div>
      </div>
      {showUploadModal && (
        <div className={styles.modal} id='uploadModal' data-backdrop='static' data-keyboard='false'>
          {showUploadModal && (
            <ImportFileModal
              repoClient={props.repoClient}
              loggedIn={props.loggedIn}
              parentItem={props.parentItem}
              webClientUrl={props.webClientUrl}
              customerId={props.customerId}
              closeImportModal={() => setShowUploadModal(false)}
              onImported={showUploadedDocument}
              refreshFolderBrowserAsync={props.refreshFolderBrowserAsync}
            />
          )}
        </div>
      )}
      {uploadedDocument && (
        <div
          className={styles.modal}
          id='uploadSuccessModal'
          data-backdrop='static'
          data-keyboard='false'
        >
          <SavedToLaserficheSuccessDialog
            successfulSave={uploadedDocument}
            closeClick={closeUploadedDocument}
          />
        </div>
      )}
      {showCreateModal && (
        <div className={styles.modal} id='createModal' data-backdrop='static' data-keyboard='false'>
          <CreateFolderModal
            repoClient={props.repoClient}
            closeCreateFolderModal={() => setShowCreateModal(false)}
            parentItem={props.parentItem}
            refreshFolderBrowserAsync={props.refreshFolderBrowserAsync}
          />
        </div>
      )}
      {showAlertModal && (
        <div className={styles.modal} id='AlertModal' data-backdrop='static' data-keyboard='false'>
          <div className='modal-dialog'>
            <div className={`modal-content ${styles.modalContent} ${styles.wrapper}`}>
              <div className='modal-body'>{PLEASE_SELECT_FILE_FOLDER_TO_OPEN}</div>
              <div className='modal-footer'>
                <button
                  type='button'
                  className='lf-button primary-button'
                  onClick={confirmAlertButton}
                >
                  {OK}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// Called once an entry has been created, so a failed refresh must not be
// reported as a failed create: retrying would then create a duplicate.
async function tryRefreshFolderBrowserAsync(
  refreshFolderBrowserAsync: () => Promise<void>
): Promise<void> {
  try {
    await refreshFolderBrowserAsync();
  } catch (err) {
    console.error('Unable to refresh the folder list:', err);
  }
}

// <lf-field-container> fetches the template list itself, the first time its
// Template dropdown opens, and shows nothing while it waits. Wrapping the
// service it fetches through is the only way to tell that load is running.
function trackTemplatesLoading(
  service: LfFieldsService,
  onLoadingChange: (loading: boolean) => void
): LfFieldsService {
  const getAvailableTemplatesAsync = service.getAvailableTemplatesAsync.bind(service);
  service.getAvailableTemplatesAsync = async () => {
    onLoadingChange(true);
    try {
      return await getAvailableTemplatesAsync();
    } finally {
      onLoadingChange(false);
    }
  };
  return service;
}

// <lf-field-container> renders its "Template" section label inside its first
// mat-panel-title, so the templates-loading spinner is portaled in there to
// sit beside that label.
function findTemplateHeader(fieldContainer: HTMLElement | undefined): HTMLElement | undefined {
  return fieldContainer?.querySelector<HTMLElement>('mat-panel-title') ?? undefined;
}

function ImportFileModal(props: {
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
  parentItem?: LfRepoTreeNode;
  webClientUrl: string;
  customerId: string;
  closeImportModal: () => void;
  onImported: (savedDocument: SavedLaserficheDocument) => void;
  refreshFolderBrowserAsync: () => Promise<void>;
}): JSX.Element {
  const fieldContainer: React.RefObject<NgElement & WithProperties<LfFieldContainerComponent>> =
    React.useRef();

  const [importFileValidationMessage, setImportFileValidationMessage] = React.useState<
    string | undefined
  >(undefined);
  const [fileUploadPercentage, setFileUploadPercentage] = React.useState(0);
  const [file, setFile] = React.useState<File | undefined>(undefined);
  const [fileName, setFileName] = React.useState<string | undefined>(undefined);
  const [adhocDialogOpened, setAdhocDialogOpened] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  // No template is pre-selected on open (initAsync is called without a
  // templateIdentifier below), so there's nothing to be invalid about until
  // the user picks a template/field -- true is a safe default.
  const [fieldsAreValid, setFieldsAreValid] = React.useState<boolean>(true);
  const [selectedTagNames, setSelectedTagNames] = React.useState<string[]>([]);
  const [templatesLoading, setTemplatesLoading] = React.useState<boolean>(false);
  const [templateHeader, setTemplateHeader] = React.useState<HTMLElement | undefined>(undefined);

  const [showImport, setShowImport] = React.useState<boolean>(true);
  const [getConfirmation, Confirmation] = useConfirm();

  const onDialogOpened: () => void = () => {
    setAdhocDialogOpened(true);
  };

  // Adding/removing an ad hoc field doesn't itself emit fieldValuesChanged,
  // so force a validity check when the ad hoc dialog closes -- otherwise a
  // newly-added required field stays untracked until the user touches it.
  const onDialogClosed: () => void = () => {
    setAdhocDialogOpened(false);
    setFieldsAreValid(fieldContainer.current?.forceValidation() ?? true);
  };

  const onFieldValuesChanged: EventListener = (event: Event) => {
    setFieldsAreValid((event as CustomEvent<boolean>).detail);
  };

  // templateSelectedChanged fires on selection, but (unlike editing a field)
  // doesn't carry a validity payload and isn't followed by fieldValuesChanged
  // until the user touches a field -- force a validity check so a template
  // with an empty required field disables OK immediately on selection.
  const onTemplateSelectedChanged: EventListener = () => {
    setFieldsAreValid(fieldContainer.current?.forceValidation() ?? true);
  };

  React.useEffect(() => {
    // A load still running from before this effect re-ran (or after the
    // modal closed) must not touch the spinner.
    let cancelled = false;
    setTemplatesLoading(false);
    const setTemplatesLoadingIfCurrent: (loading: boolean) => void = (loading: boolean) => {
      if (!cancelled) {
        setTemplatesLoading(loading);
      }
    };

    const initializeFieldContainerAsync: () => Promise<void> = async () => {
      try {
        fieldContainer.current.addEventListener('dialogOpened', onDialogOpened);
        fieldContainer.current.addEventListener('dialogClosed', onDialogClosed);
        fieldContainer.current.addEventListener('fieldValuesChanged', onFieldValuesChanged);
        fieldContainer.current.addEventListener(
          'templateSelectedChanged',
          onTemplateSelectedChanged
        );

        await fieldContainer.current.initAsync(
          trackTemplatesLoading(new LfFieldsService(props.repoClient), setTemplatesLoadingIfCurrent)
        );
        if (!cancelled) {
          setTemplateHeader(findTemplateHeader(fieldContainer.current));
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        console.error(err);
        setImportFileValidationMessage(
          `${UNABLE_TO_LOAD_TEMPLATES_AND_FIELDS} ${getErrorDetails(err) ?? UNKNOWN_ERROR}`
        );
      }
    };
    if (props.repoClient) {
      void initializeFieldContainerAsync();
    }
    return () => {
      cancelled = true;
      fieldContainer.current?.removeEventListener('dialogOpened', onDialogOpened);
      fieldContainer.current?.removeEventListener('dialogClosed', onDialogClosed);
      fieldContainer.current?.removeEventListener('fieldValuesChanged', onFieldValuesChanged);
      fieldContainer.current?.removeEventListener(
        'templateSelectedChanged',
        onTemplateSelectedChanged
      );
    };
  }, [props.repoClient, props.loggedIn]);

  const closeImportFileModal: () => void = () => {
    props.closeImportModal();
  };

  const importFileToRepositoryAsync: () => Promise<void> = async () => {
    try {
      const fileData = file;
      const repoId = await props.repoClient.getCurrentRepoId();
      setFileUploadPercentage(5);
      setImportFileValidationMessage(undefined);
      if (!fileData) {
        setFileUploadPercentage(0);
        setImportFileValidationMessage(fileValidation);
        return;
      }
      const fileDataSize = fileData.size;
      if (fileDataSize > 100000000) {
        setFileUploadPercentage(0);
        setImportFileValidationMessage(fileSizeValidation);
        return;
      }
      if (!fileName) {
        setFileUploadPercentage(0);
        setImportFileValidationMessage(fileNameValidation);
        return;
      }
      const extension =
        PathUtils.getCleanedExtension(PathUtils.getFileExtension(fileData.name)) ?? '';
      const renamedFile = new File([fileData], fileName + extension);
      const fileContainsBackslash = fileName.includes('\\');
      try {
        const entryWithPath = await props.repoClient.entriesClient.getEntryByPath({
          repositoryId: repoId,
          fullPath: PathUtils.combinePaths(props.parentItem.path, fileName),
        });
        // v2 returns a GetEntryByPathResponse on success, so the response is
        // always truthy: test the entry itself.
        if (entryWithPath?.entry) {
          setShowImport(false);
          const confirmUpload = await getConfirmation(
            ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME
          );
          setShowImport(true);
          if (confirmUpload) {
            // continue
          } else {
            setFileUploadPercentage(0);
            return;
          }
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        const docDoesNotAlreadyExists = err.status === 404;
        if (docDoesNotAlreadyExists) {
          // doesn't exist, good to go
        } else {
          throw err;
        }
      }
      if (fileContainsBackslash) {
        setFileUploadPercentage(0);
        setImportFileValidationMessage(fileNameWithBacklash);
        return;
      }
      await continueImportAsync(extension, renamedFile, repoId);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (err: any) {
      setFileUploadPercentage(0);
      setError(getErrorDetails(err) ?? UNKNOWN_ERROR);
      console.error(error);
    }
  };

  async function continueImportAsync(
    extension: string,
    renamedFile: File,
    repoId: string
  ): Promise<void> {
    const fieldValidation = fieldContainer.current?.forceValidation();
    if (fieldValidation) {
      const fieldValues = fieldContainer.current.getFieldValues();
      const formattedFieldValues: FieldToUpdate[] = [];

      for (const key in fieldValues) {
        const value = fieldValues[key];
        formattedFieldValues.push(
          new FieldToUpdate({
            name: key,
            values: value?.values?.map((val) => val.value),
          })
        );
      }

      const templateValue = getTemplateName();
      let templateName;
      if (templateValue) {
        templateName = templateValue;
      }

      setFileUploadPercentage(80);
      const fieldsmetadata: ImportEntryRequestMetadata = new ImportEntryRequestMetadata({
        templateName,
        fields: formattedFieldValues,
        // Part of the import rather than a separate call afterwards, so a tag
        // the repository rejects fails the upload instead of being dropped.
        tags: selectedTagNames.length > 0 ? selectedTagNames : undefined,
      });
      // v2 has no separate `extension` parameter: the extension has to be part
      // of the electronic document's file name.
      const fileNameWithExt = fileName + extension;
      const parentEntryId = props.parentItem.id;

      const file: FileParameter = {
        data: renamedFile,
        fileName: fileNameWithExt,
      };
      const requestParameters = {
        repositoryId: repoId,
        entryId: Number.parseInt(parentEntryId, 10),
        file,
        request: new ImportEntryRequest({
          name: fileName,
          autoRename: true,
          // v2 defaults this to false, which would store txt/tif/tiff/bmp/pcx/
          // jpg/jpeg/gif/png files as image pages rather than as the
          // electronic document. v1 always stored the edoc.
          importAsElectronicDocument: true,
          metadata: fieldsmetadata,
        }),
      };

      const importedEntry = await props.repoClient.entriesClient.importEntry(requestParameters);
      setFileUploadPercentage(100);
      await tryRefreshFolderBrowserAsync(props.refreshFolderBrowserAsync);
      props.onImported({
        fileName: importedEntry.name ?? fileName,
        fileLink: getEntryWebAccessUrl(
          importedEntry.id?.toString(),
          props.webClientUrl,
          false,
          repoId,
          props.customerId
        ),
        folderLink: getEntryWebAccessUrl(
          parentEntryId,
          props.webClientUrl,
          true,
          repoId,
          props.customerId
        ),
      });
    } else {
      setFileUploadPercentage(0);
      setImportFileValidationMessage(requiredFieldsValidation);
    }
  }

  function getTemplateName(): string {
    const templateValue = fieldContainer.current.getTemplateValue();
    if (templateValue) {
      return templateValue.name;
    }
    return undefined;
  }

  function setFileToImport(e: ChangeEvent<HTMLInputElement>): void {
    const inputFile = e.target.files[0];
    const filePath = e.target.value;
    const fileSize = inputFile.size;
    const newFileName = PathUtils.getLastPathSegment(filePath);
    const withoutExtension = PathUtils.removeFileExtension(newFileName);
    if (fileSize < 100000000) {
      setImportFileValidationMessage(undefined);
      setFile(inputFile);
      setFileName(withoutExtension);
    } else {
      setFileUploadPercentage(0);
      setImportFileValidationMessage(fileSizeValidation);
    }
  }

  const setNewFileName: (e: ChangeEvent<HTMLInputElement>) => void = (
    e: ChangeEvent<HTMLInputElement>
  ) => {
    const newFileName = e.target.value;

    setFileName(newFileName);
  };

  const validationError = importFileValidationMessage ? (
    <div style={{ color: 'red' }}>
      <span>{importFileValidationMessage}</span>
    </div>
  ) : undefined;

  const templatesSpinner = (
    <DelayedSpinner loading={templatesLoading} label={LOADING} className='ms-2 flex-shrink-0' />
  );

  return (
    <div className='modal-dialog modal-dialog-scrollable modal-lg'>
      <div className={`modal-content ${styles.modalContent} ${styles.wrapper}`}>
        <div hidden={!showImport} className={`modal-header ${styles.header}`}>
          <div className='modal-title' id='ModalLabel'>
            <LaserficheDialogTitle title={UPLOAD_FILE_TO_LASERFICHE_TITLE} />
          </div>
          <button
            type='button'
            className='btn-close'
            aria-label='Close'
            onClick={closeImportFileModal}
          />
          <div
            className='progress'
            style={{
              display: fileUploadPercentage > 0 ? 'block' : 'none',
              width: '100%',
            }}
          >
            <div
              className='progress-bar progress-bar-striped active'
              style={{
                width: fileUploadPercentage + '%',
                backgroundColor: 'orange',
                height: 'inherit',
              }}
            >
              {UPLOADING}
            </div>
          </div>
        </div>
        <div hidden={!showImport} className={`modal-body ${styles.contentBox}`}>
          {!error && (
            <>
              <div className='input-group mb-3'>
                <input
                  type='file'
                  className='form-control'
                  id='importFile'
                  onChange={setFileToImport}
                  aria-label='Choose a file to import'
                />
              </div>
              {validationError}
              <div className='row mb-3'>
                <label className='col-sm-3 col-form-label lf-text-label'>{NAME}</label>
                <div className='col-sm-9'>
                  <input
                    type='text'
                    className='form-control'
                    id='uploadFileID'
                    onChange={setNewFileName}
                    value={fileName}
                  />
                </div>
              </div>
              <div
                className={`${styles.lfComponentContainer}${
                  adhocDialogOpened ? ` ${styles.lfAdhocMinHeight}` : ''
                }`}
              >
                <lf-field-container
                  collapsible='true'
                  start_collapsed='true'
                  ref={fieldContainer}
                />
                {templateHeader
                  ? ReactDOM.createPortal(templatesSpinner, templateHeader)
                  : templatesSpinner}
              </div>
              {props.repoClient && (
                <LfTagsPicker
                  repoClient={props.repoClient}
                  onSelectedTagNamesChange={setSelectedTagNames}
                />
              )}
            </>
          )}
          {error && <span style={{ justifyContent: 'center' }}>{`Error uploading: ${error}`}</span>}
        </div>
        <div hidden={!showImport} className={`modal-footer ${styles.footer}`}>
          <button
            type='button'
            className='lf-button primary-button'
            disabled={fileUploadPercentage > 0 || !fieldsAreValid}
            onClick={error ? closeImportFileModal : importFileToRepositoryAsync}
          >
            {OK}
          </button>
          <button type='button' className='lf-button sec-button' onClick={closeImportFileModal}>
            {CANCEL}
          </button>
        </div>
        <Confirmation cancelButtonText={GO_BACK} headerText={DOCUMENT_ALREADY_EXISTS} />
      </div>
    </div>
  );
}

function CreateFolderModal(props: {
  repoClient: IRepositoryApiClientExInternal;
  closeCreateFolderModal: () => void;
  parentItem: LfRepoTreeNode;
  refreshFolderBrowserAsync: () => Promise<void>;
}): JSX.Element {
  const [folderName, setFolderName] = React.useState('');
  const [createFolderNameValidationMessage, setCreateFolderNameValidationMessage] = React.useState<
    string | undefined
  >(undefined);

  const closeNewFolderModal: () => void = () => {
    setCreateFolderNameValidationMessage(undefined);
    setFolderName('');
    props.closeCreateFolderModal();
  };

  const createNewFolderAsync: () => Promise<void> = async () => {
    if (folderName) {
      const fileContainsBackslash = folderName.includes('\\');
      if (fileContainsBackslash) {
        setCreateFolderNameValidationMessage(folderBackslashNameValidation);
      } else {
        setCreateFolderNameValidationMessage(undefined);

        const repoId = await props.repoClient.getCurrentRepoId();
        const createEntryRequest: CreateEntryRequest = new CreateEntryRequest({
          entryType: CreateEntryRequestEntryType.Folder,
          name: folderName,
        });
        const requestParameters = {
          repositoryId: repoId,
          entryId: Number.parseInt(props.parentItem.id, 10),
          request: createEntryRequest,
        };
        try {
          const array = [];
          const newFolderEntry: Entry =
            await props.repoClient.entriesClient.createEntry(requestParameters);

          array.push(newFolderEntry);
          await tryRefreshFolderBrowserAsync(props.refreshFolderBrowserAsync);
          props.closeCreateFolderModal();
          setFolderName('');
        } catch {
          setCreateFolderNameValidationMessage(folderExists);
        }
      }
    } else {
      setCreateFolderNameValidationMessage(folderValidation);
    }
  };

  function handleFolderNameChange(e: ChangeEvent<HTMLInputElement>): void {
    setFolderName(e.target.value);
  }

  return (
    <div className='modal-dialog'>
      <div className={`modal-content ${styles.modalContent} ${styles.wrapper}`}>
        <div className='modal-header'>
          <h5 className='modal-title' id='ModalLabel'>
            {CREATE_FOLDER}
          </h5>
          <button
            type='button'
            className='btn-close'
            aria-label='Close'
            onClick={props.closeCreateFolderModal}
          />
        </div>
        <div className='modal-body'>
          <div className='mb-3'>
            <label>{FOLDER_NAME}</label>
            <input
              type='text'
              className='form-control'
              id='folderName'
              placeholder='Name'
              onChange={handleFolderNameChange}
            />
          </div>
          <div style={{ color: 'red' }}>
            <span>{createFolderNameValidationMessage}</span>
          </div>
        </div>
        <div className='modal-footer'>
          <button type='button' className='lf-button primary-button' onClick={createNewFolderAsync}>
            {SUBMIT}
          </button>
          <button type='button' className='lf-button sec-button' onClick={closeNewFolderModal}>
            {CLOSE}
          </button>
        </div>
      </div>
    </div>
  );
}
