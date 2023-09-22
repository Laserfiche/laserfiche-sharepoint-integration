import { NgElement, WithProperties } from '@angular/elements';
import {
  EntryType,
  PostEntryChildrenRequest,
  PostEntryChildrenEntryType,
  Entry,
  FieldToUpdate,
  ValueToUpdate,
  PostEntryWithEdocMetadataRequest,
  PutFieldValsRequest,
  FileParameter,
} from '@laserfiche/lf-repository-api-client';
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
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { ChangeEvent } from 'react';
import { getEntryWebAccessUrl } from '../../../Utils/Funcs';
import styles from './LaserficheRepositoryAccess.module.scss';
require('./../../../Assets/CSS/commonStyles.css');

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
    displayName: 'Page',
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
const fileSizeValidation = 'Please select a file below 100MB size';
const fileNameValidation = 'Please provide proper name of the file';
const fileNameWithBacklash =
  'Please provide proper name of the file without backslash';
const folderValidation = 'Please provide folder name';
const folderNameValidation = 'Invalid Name, only alphanumeric are allowed.';
const folderExists = 'Object already exists';

export default function RepositoryViewComponent(props: {
  repoClient: IRepositoryApiClientExInternal;
  webClientUrl: string;
  loggedIn: boolean;
}): JSX.Element {
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef<NgElement & WithProperties<LfRepositoryBrowserComponent>>();
  let lfRepoTreeService: LfRepoTreeNodeService;

  const [parentItem, setParentItem] = React.useState<
    LfRepoTreeNode | undefined
  >(undefined);
  const [selectedItem, setSelectedItem] = React.useState<
    LfRepoTreeNode | undefined
  >(undefined);

  React.useEffect(() => {
    const onEntrySelected: (
      event: CustomEvent<LfRepoTreeNode[] | undefined>
    ) => void = (event: CustomEvent<LfRepoTreeNode[] | undefined>) => {
      const selectedNode = event.detail ? event.detail[0] : undefined;
      setSelectedItem(selectedNode);
    };

    const onEntryOpened: (
      event: CustomEvent<LfRepoTreeNode[] | undefined>
    ) => Promise<void> = async (
      event: CustomEvent<LfRepoTreeNode[] | undefined>
    ) => {
      const openedNode = event.detail ? event.detail[0] : undefined;
      const entryType =
        openedNode.entryType === EntryType.Shortcut
          ? openedNode.targetType
          : openedNode.entryType;
      if (
        entryType === EntryType.Folder ||
        entryType === EntryType.RecordSeries
      ) {
        setParentItem(openedNode);
      } else {
        const repoId = await props.repoClient.getCurrentRepoId();

        if (openedNode?.id) {
          const webClientNodeUrl = getEntryWebAccessUrl(
            openedNode.id,
            props.webClientUrl,
            openedNode.isContainer,
            repoId
          );
          window.open(webClientNodeUrl);
        }
      }
    };

    const initializeTreeAsync: () => Promise<void> = async () => {
      const repoBrowser = repositoryBrowser.current;
      lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
      lfRepoTreeService.viewableEntryTypes = [
        EntryType.Folder,
        EntryType.Shortcut,
        EntryType.Document,
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
        } catch (err) {
          console.error(err);
        }
      } else {
        console.debug(
          'Unable to initialize tree, lfRepoTreeService is undefined'
        );
      }
    };
    if (props.repoClient) {
      void initializeTreeAsync();
    }
  }, [props.repoClient, props.loggedIn]);

  const isNodeSelectable: (node: LfRepoTreeNode) => boolean = (
    node: LfRepoTreeNode
  ) => {
    if (
      node?.entryType === EntryType.Folder ||
      node?.entryType === EntryType.Document
    ) {
      return true;
    } else if (
      (node?.entryType === EntryType.Shortcut &&
        node?.targetType === EntryType.Folder) ||
      (node?.entryType === EntryType.Shortcut &&
        node?.targetType === EntryType.Document)
    ) {
      return true;
    } else {
      return false;
    }
  };

  const refreshFolderBrowserAsync: () => Promise<void> = async () => {
    await repositoryBrowser.current.refreshAsync(false);
  };

  return (
    <>
      <div>
        <main className='bg-white'>
          <div style={{ margin: '10px 0px' }}>
            <img
              style={{ width: '30px' }}
              src={require('./../../../Assets/Images/laserfiche-logo.png')}
            />
            <span className={styles.browserTitle}>
              Laserfiche Repository Explorer
            </span>
          </div>
          {props.loggedIn && (
            <>
              <RepositoryBrowserToolbar
                repoClient={props.repoClient}
                selectedItem={selectedItem}
                parentItem={parentItem}
                loggedIn={props.loggedIn}
                webClientUrl={props.webClientUrl}
                refreshFolderBrowserAsync={refreshFolderBrowserAsync}
              />
              <div
                className='lf-folder-browser-sample-container'
                style={{ height: '400px' }}
              >
                <div className='repository-browser'>
                  <lf-repository-browser
                    ref={repositoryBrowser}
                    ok_button_text='Okay'
                    cancel_button_text='Cancel'
                    multiple='false'
                    style={{ height: '420px' }}
                    isSelectable={isNodeSelectable}
                  />
                </div>
              </div>
            </>
          )}
        </main>
      </div>
    </>
  );
}

function RepositoryBrowserToolbar(props: {
  repoClient: IRepositoryApiClientExInternal;
  webClientUrl: string;
  selectedItem: LfRepoTreeNode;
  parentItem: LfRepoTreeNode;
  loggedIn: boolean;
  refreshFolderBrowserAsync: () => Promise<void>;
}): JSX.Element {
  const [showUploadModal, setShowUploadModal] = React.useState(false);
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
        repoId
      );
      window.open(webClientNodeUrl);
    } else {
      setShowAlertModal(true);
    }
  };

  const confirmAlertButton: () => void = () => {
    setShowAlertModal(false);
  };

  return (
    <>
      <div id='mainWebpartContent'>
        <div className={styles.buttonContainer}>
          <button
            className={styles.lfMaterialIconButton}
            title='Open File'
            onClick={openFileOrFolder}
          >
            <img
              className={styles.waIcon}
              src={`${require('./../../../Assets/Images/waicons.svg')}#open`}
            />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title='Upload File'
            onClick={openImportFileModal}
          >
            <img
              className={styles.waIcon}
              src={`${require('./../../../Assets/Images/waicons.svg')}#upload`}
            />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title='Create Folder'
            onClick={openNewFolderModal}
          >
            <img
              className={styles.waIcon}
              src={`${require('./../../../Assets/Images/waicons.svg')}#add-folder`}
            />
          </button>
          <button
            className={styles.lfMaterialIconButton}
            title='Refresh'
            onClick={props.refreshFolderBrowserAsync}
          >
            <img
              className={styles.waIcon}
              src={`${require('./../../../Assets/Images/waicons.svg')}#refresh`}
            />
          </button>
        </div>
      </div>
      <div
        className={styles.modal}
        id='uploadModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showUploadModal}
      >
        {showUploadModal && (
          <ImportFileModal
            repoClient={props.repoClient}
            loggedIn={props.loggedIn}
            parentItem={props.parentItem}
            closeImportModal={() => setShowUploadModal(false)}
          />
        )}
      </div>
      <div
        className={styles.modal}
        id='createModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showCreateModal}
      >
        <CreateFolderModal
          repoClient={props.repoClient}
          closeCreateFolderModal={() => setShowCreateModal(false)}
          parentItem={props.parentItem}
        />
      </div>
      <div
        className={styles.modal}
        id='AlertModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showAlertModal}
      >
        <div className='modal-dialog'>
          <div className={`${styles.modalContent} ${styles.wrapper}`}>
            <div className='modal-body'>Please select file/folder to open</div>
            <div className='modal-footer'>
              <button
                type='button'
                className='lf-button primary-button'
                data-dismiss='modal'
                onClick={confirmAlertButton}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

function ImportFileModal(props: {
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
  parentItem?: LfRepoTreeNode;
  closeImportModal: () => void;
}): JSX.Element {
  const fieldContainer: React.RefObject<
    NgElement & WithProperties<LfFieldContainerComponent>
  > = React.useRef();
  let lfFieldsService: LfFieldsService;

  const [importFileValidationMessage, setImportFileValidationMessage] =
    React.useState<string | undefined>(undefined);
  const [fileUploadPercentage, setFileUploadPercentage] = React.useState(0);
  const [file, setFile] = React.useState<File | undefined>(undefined);
  const [fileName, setFileName] = React.useState<string | undefined>(undefined);
  const [adhocDialogOpened, setAdhocDialogOpened] =
    React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const onDialogOpened: () => void = () => {
    setAdhocDialogOpened(true);
  };

  const onDialogClosed: () => void = () => {
    setAdhocDialogOpened(false);
  };

  React.useEffect(() => {
    const initializeFieldContainerAsync: () => Promise<void> = async () => {
      try {
        fieldContainer.current.addEventListener('dialogOpened', onDialogOpened);
        fieldContainer.current.addEventListener('dialogClosed', onDialogClosed);

        lfFieldsService = new LfFieldsService(props.repoClient);
        await fieldContainer.current.initAsync(lfFieldsService);
      } catch (err) {
        setFileUploadPercentage(0);
        setError(err.message);
        console.error(error);
      }
    };
    if (props.repoClient) {
      void initializeFieldContainerAsync();
    }
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
      const extension = PathUtils.getCleanedExtension(fileData.name);
      const renamedFile = new File([fileData], fileName + extension);
      const fileContainsBacklash = fileName.includes('\\');
      if (fileContainsBacklash) {
        setFileUploadPercentage(0);
        setImportFileValidationMessage(fileNameWithBacklash);
        return;
      }
      const fieldValidation = fieldContainer.current.forceValidation();
      if (fieldValidation) {
        const fieldValues = fieldContainer.current.getFieldValues();
        const formattedFieldValues:
          | {
              [key: string]: FieldToUpdate;
            }
          | undefined = {};

        for (const key in fieldValues) {
          const value = fieldValues[key];
          formattedFieldValues[key] = new FieldToUpdate({
            ...value,
            values: value.values.map((val) => new ValueToUpdate(val)),
          });
        }

        const templateValue = getTemplateName();
        let templateName;
        if (templateValue) {
          templateName = templateValue;
        }

        setFileUploadPercentage(100);
        const fieldsmetadata: PostEntryWithEdocMetadataRequest =
          new PostEntryWithEdocMetadataRequest({
            template: templateName,
            metadata: new PutFieldValsRequest({
              fields: formattedFieldValues,
            }),
          });
        const fileNameWithExt = fileName + extension;
        const fileextensionperiod = extension;
        const fileNameNoPeriod = fileName;
        const parentEntryId = props.parentItem.id;

        const file: FileParameter = {
          data: renamedFile,
          fileName: fileNameWithExt,
        };
        const requestParameters = {
          repoId,
          parentEntryId: Number.parseInt(parentEntryId, 10),
          electronicDocument: file,
          autoRename: true,
          fileName: fileNameNoPeriod,
          request: fieldsmetadata,
          extension: fileextensionperiod,
        };

        await props.repoClient.entriesClient.importDocument(requestParameters);
        props.closeImportModal();
      } else {
        fieldContainer.current.forceValidation();
      }
    } catch (err) {
      setFileUploadPercentage(0);
      setError(err.message);
      console.error(error);
    }
  };

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

  return (
    <div className='modal-dialog modal-dialog-scrollable modal-lg'>
      <div className={`modal-content ${styles.modalContent} ${styles.wrapper}`}>
        <div className={`modal-header ${styles.header}`}>
          <div className='modal-title' id='ModalLabel'>
            Upload File
          </div>
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
              Uploading
            </div>
          </div>
        </div>
        <div className={`modal-body ${styles.contentBox}`}>
          {!error && (
            <>
              <div className='input-group mb-3'>
                <div className='custom-file'>
                  <input
                    type='file'
                    className='custom-file-input'
                    id='importFile'
                    onChange={setFileToImport}
                    aria-describedby='inputGroupFileAddon04'
                    placeholder='Choose file'
                  />
                  <label className='custom-file-label' id='importFileName'>
                    {file?.name ? file.name : 'Choose a file'}
                  </label>
                </div>
              </div>
              {validationError}
              <div className='form-group row mb-3'>
                <label className='col-sm-3 col-form-label'>Name</label>
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
                className={`lf-component-container${
                  adhocDialogOpened ? ' lfAdhocMinHeight' : ''
                }`}
              >
                <lf-field-container
                  collapsible='true'
                  startCollapsed='true'
                  ref={fieldContainer}
                />
              </div>
            </>
          )}
          {error && (
            <span
              style={{ justifyContent: 'center' }}
            >{`Error uploading: ${error}`}</span>
          )}
        </div>
        <div className={`modal-footer ${styles.footer}`}>
          <button
            type='button'
            className='lf-button primary-button'
            disabled={fileUploadPercentage > 0}
            onClick={importFileToRepositoryAsync}
          >
            OK
          </button>
          <button
            type='button'
            className='lf-button sec-button'
            onClick={closeImportFileModal}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

function CreateFolderModal(props: {
  repoClient: IRepositoryApiClientExInternal;
  closeCreateFolderModal: () => void;
  parentItem: LfRepoTreeNode;
}): JSX.Element {
  const [folderName, setFolderName] = React.useState('');
  const [
    createFolderNameValidationMessage,
    setCreateFolderNameValidationMessage,
  ] = React.useState<string | undefined>(undefined);

  const closeNewFolderModal: () => void = () => {
    setCreateFolderNameValidationMessage(undefined);
    setFolderName('');
    props.closeCreateFolderModal();
  };

  const createNewFolderAsync: () => Promise<void> = async () => {
    if (folderName) {
      if (/[^ A-Za-z0-9]/.test(folderName)) {
        setCreateFolderNameValidationMessage(folderNameValidation);
      } else {
        setCreateFolderNameValidationMessage(undefined);

        const repoId = await props.repoClient.getCurrentRepoId();
        const postEntryChildrenRequest: PostEntryChildrenRequest =
          new PostEntryChildrenRequest({
            entryType: PostEntryChildrenEntryType.Folder,
            name: folderName,
          });
        const requestParameters = {
          repoId,
          entryId: Number.parseInt(props.parentItem.id, 10),
          request: postEntryChildrenRequest,
        };
        try {
          const array = [];
          const newFolderEntry: Entry =
            await props.repoClient.entriesClient.createOrCopyEntry(
              requestParameters
            );

          array.push(newFolderEntry);
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
            Create Folder
          </h5>
          <button
            type='button'
            className='close'
            data-dismiss='modal'
            aria-label='Close'
            onClick={props.closeCreateFolderModal}
          >
            <span aria-hidden='true'>&times;</span>
          </button>
        </div>
        <div className='modal-body'>
          <div className='form-group'>
            <label>Folder Name</label>
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
          <button
            type='button'
            className='lf-button primary-button'
            data-dismiss='modal'
            onClick={createNewFolderAsync}
          >
            Submit
          </button>
          <button
            type='button'
            className='lf-button sec-button'
            data-dismiss='modal'
            onClick={closeNewFolderModal}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}
