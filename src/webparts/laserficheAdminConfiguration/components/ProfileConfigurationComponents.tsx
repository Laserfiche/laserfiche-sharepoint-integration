import { NgElement, WithProperties } from '@angular/elements';
import {
  EntryType,
  TemplateFieldInfo,
  WTemplateInfo,
} from '@laserfiche/lf-repository-api-client';
import {
  LfRepoTreeNode,
  LfRepoTreeNodeService,
} from '@laserfiche/lf-ui-components-services';
import { LfRepositoryBrowserComponent } from '@laserfiche/types-lf-ui-components';
import * as React from 'react';
import { ChangeEvent, useState } from 'react';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';

export interface ProfileConfiguration {
  ConfigurationName: string;
  DocumentName: string;
  selectedTemplateName?: string;
  selectedFolder?: LfFolder;
  Action: string;
  mappedFields: MappedFields[];
}

export interface LfFolder {
  path: string;
  id: string;
}

export interface SPProfileConfigurationData {
  Title: string;
  TypeAsString: string;
  InternalName: string;
}

export interface MappedFields {
  id: string;
  lfField: TemplateFieldInfo | undefined;
  spField: SPProfileConfigurationData | undefined;
}

export enum ActionTypes {
  'COPY'= 'COPY',
  'MOVE_AND_DELETE' = 'MOVE_AND_DELETE',
  'REPLACE' = 'REPLACE'
}

export function ProfileHeader(props: { configurationName: string }) {
  return (
    <h6 className='mb-0'>
      Profile :{' '}
      <span className='h5' id='configurationTitle'>
        {props.configurationName}
      </span>
    </h6>
  );
}

export function ConfigurationBody(props: {
  availableLfTemplates: WTemplateInfo[];
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
  profileConfig: ProfileConfiguration;
  handleProfileConfigUpdate: (config: ProfileConfiguration) => void;
  handleTemplateChange: (templateName: string) => void;
}) {
  const [showFolderModal, setShowFolderModal] = useState(false);

  const selectedEntryNodePath = props.profileConfig.selectedFolder?.path;

  const onSelectFolder = async (selectedNode: LfRepoTreeNode | undefined) => {
    if (!props.repoClient) {
      throw new Error('Repo Client is undefined.');
    }
    const config = { ...props.profileConfig };
    const lfFolder: LfFolder = {
      path: selectedNode.path,
      id: selectedNode.id,
    };
    config.selectedFolder = lfFolder;
    props.handleProfileConfigUpdate(config);
    setShowFolderModal(false);
  };

  const handleTemplateChange = (event) => {
    const value = (event.target as HTMLSelectElement).value;
    const templateName = value;
    const profileConfig = { ...props.profileConfig };
    profileConfig.selectedTemplateName = templateName;
    props.handleProfileConfigUpdate(profileConfig);
    props.handleTemplateChange(templateName);
  };

  const handleActionTypeChange = (event) => {
    const value = (event.target as HTMLSelectElement).value;
    const actionName = value;
    const profileConfig = { ...props.profileConfig };
    profileConfig.Action = actionName;
    props.handleProfileConfigUpdate(profileConfig);
  }

  function CloseFolderModalUp() {
    setShowFolderModal(false);
  }

  async function OpenFoldersModal() {
    setShowFolderModal(true);
  }
  return (
    <>
      <div className='form-group row'>
        <DocumentName
          documentName={props.profileConfig?.DocumentName}
        />
      </div>
      <div className='form-group row'>
        <TemplateSelector
          availableLfTemplates={props.availableLfTemplates}
          selectedTemplateName={props.profileConfig?.selectedTemplateName}
          repoClient={props.repoClient}
          onChangeTemplate={handleTemplateChange}
        />
      </div>
      <div className='form-group row'>
        <label htmlFor='txt3' className='col-sm-2 col-form-label'>
          Laserfiche Destination
        </label>
        <div className='col-sm-6'>
          <input
            type='text'
            className='form-control'
            id='destinationPath'
            placeholder='(Path in Laserfiche) Example: \folder\subfolder'
            disabled
            value={props.profileConfig.selectedFolder?.path}
          />
          <div>
            <span>Use the Browse button to select a path</span>
          </div>
        </div>
        <div className='col-sm-2' id='folderModal' style={{ marginTop: '2px' }}>
          <a
            href='javascript:;'
            className='btn btn-primary btn-sm'
            data-toggle='modal'
            data-target='#folderModal'
            onClick={OpenFoldersModal}
          >
            Browse
          </a>
        </div>
      </div>
      <div className='form-group row'>
        <label htmlFor='dwl4' className='col-sm-2 col-form-label'>
          After import
        </label>
        <div className='col-sm-6'>
          <select onChange={handleActionTypeChange} defaultValue={props.profileConfig.Action} className='custom-select' id='action'>
            <option value={ActionTypes.COPY}>
              Leave a copy of the file in SharePoint
            </option>
            <option value={ActionTypes.REPLACE}>
              Replace SharePoint file with a link to the document in Laserfiche
            </option>
            <option value={ActionTypes.MOVE_AND_DELETE}>Delete SharePoint file</option>
          </select>
        </div>
        <div className='col-sm-2'>
          {/* <div className="custom-control custom-checkbox mt-2" style={{ paddingLeft: "3px !important", "marginLeft": "-23px" }}>
                          <a data-toggle="tooltip" style={{ "color": "#0062cc" }}><span className="fa fa-question-circle fa-2"></span></a>
                        </div> */}
        </div>
      </div>
      <div
        className='modal'
        id='folderModal'
        data-backdrop='static'
        data-keyboard='false'
        hidden={!showFolderModal}
      >
        {showFolderModal && (
          <RepositoryBrowserModal
            repoClient={props.repoClient}
            CloseFolderBrowserUp={CloseFolderModalUp}
            selectedEntryNodePath={selectedEntryNodePath}
            SelectFolder={onSelectFolder}
          />
        )}
      </div>
    </>
  );
}

export function RepositoryBrowserModal(props: {
  CloseFolderBrowserUp: () => void;
  SelectFolder: (node: LfRepoTreeNode | undefined) => void;
  selectedEntryNodePath: string;
  repoClient: IRepositoryApiClientExInternal;
}) {
  const [shouldShowOpen, setShouldShowOpen] = useState(false);
  const [shouldShowSelect, setShouldShowSelect] = useState(false);
  const [shouldDisableSelect, setShouldDisableSelect] = useState(false);

  const [entrySelected, setEntrySelected] = useState<
    LfRepoTreeNode | undefined
  >(undefined);
  const onEntrySelected = (event: CustomEvent<LfRepoTreeNode[]>) => {
    const treeNodesSelected: LfRepoTreeNode[] = event.detail;
    const selectedNode =
      treeNodesSelected?.length > 0 ? treeNodesSelected[0] : undefined;
    setEntrySelected(selectedNode);
    setShouldShowOpen(selectedNode && selectedNode.isContainer);
    setShouldShowSelect(
      !selectedNode && !!repositoryBrowser?.current?.currentFolder
    );
    setShouldDisableSelect(getShouldDisableSelect());
  };
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef();
  let lfRepoTreeService;
  React.useEffect(() => {
    if (props.repoClient) {
      lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
      lfRepoTreeService.viewableEntryTypes = [
        EntryType.Folder,
        EntryType.Shortcut,
      ];
      initializeTreeAsync();
    }
  }, [props.repoClient]);
  const isNodeSelectable = (node: LfRepoTreeNode) => {
    if (node?.entryType == EntryType.Folder) {
      return true;
    } else if (
      node?.entryType == EntryType.Shortcut &&
      node?.targetType == EntryType.Folder
    ) {
      return true;
    } else {
      return false;
    }
  };

  async function initializeTreeAsync() {
    if (!props.repoClient) {
      throw new Error('RepoId is undefined');
    }
    repositoryBrowser.current?.addEventListener(
      'entrySelected',
      onEntrySelected
    );
    let focusedNode: LfRepoTreeNode | undefined;
    if (props.selectedEntryNodePath) {
      const repoId = await props.repoClient.getCurrentRepoId();
      const focusedNodeByPath =
        await props.repoClient.entriesClient.getEntryByPath({
          repoId: repoId,
          fullPath: props.selectedEntryNodePath,
        });
      const repoName = await props.repoClient.getCurrentRepoName();
      const focusedNodeEntry = focusedNodeByPath?.entry;
      if (focusedNodeEntry) {
        focusedNode = lfRepoTreeService?.createLfRepoTreeNode(
          focusedNodeEntry,
          repoName
        );
      }
    }
    if (lfRepoTreeService) {
      await repositoryBrowser?.current?.initAsync(
        lfRepoTreeService,
        focusedNode
      );
    } else {
      console.debug(
        'Unable to initialize tree, lfRepoTreeService is undefined'
      );
    }
  }
  function getShouldShowSelect(): boolean {
    const showSelect =
      !entrySelected && !!repositoryBrowser?.current?.currentFolder;
    return showSelect;
  }

  function getShouldShowOpen(): boolean {
    return !!entrySelected;
  }

  function getShouldDisableSelect(): boolean {
    return !isNodeSelectable(
      repositoryBrowser?.current?.currentFolder as LfRepoTreeNode
    );
  }

  const onSelectFolder = () => {
    props.SelectFolder(
      repositoryBrowser?.current?.currentFolder as LfRepoTreeNode
    );
  };

  const onOpenNode = async () => {
    await repositoryBrowser?.current?.openSelectedNodesAsync();
    setShouldShowOpen(getShouldShowOpen());
    setShouldShowSelect(getShouldShowSelect());
  };

  return (
    <>
      <div className='modal-dialog modal-dialog-centered'>
        <div className='modal-content'>
          <div className='modal-header'>
            <h5 className='modal-title' id='ModalLabel'>
              Select folder for saving to Laserfiche
            </h5>
            <button
              type='button'
              className='close'
              data-dismiss='modal'
              aria-label='Close'
              onClick={props.CloseFolderBrowserUp}
            >
              <span aria-hidden='true'>&times;</span>
            </button>
          </div>
          <div className='modal-body'>
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
                <div className='repository-browser-button-containers'>
                  <span>
                    <button
                      className='lf-button primary-button'
                      onClick={onOpenNode}
                      hidden={!shouldShowOpen}
                    >
                      OPEN
                    </button>
                    <button
                      className='lf-button primary-button'
                      onClick={onSelectFolder}
                      hidden={!shouldShowSelect}
                      disabled={shouldDisableSelect}
                    >
                      Select
                    </button>
                    <button
                      className='sec-button lf-button margin-left-button'
                      onClick={props.CloseFolderBrowserUp}
                    >
                      CANCEL
                    </button>
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}

export function DocumentName(props: { documentName: string }) {
  return (
    <>
      <label htmlFor='txt1' className='col-sm-2 col-form-label'>
        Document Name
      </label>
      <div className='col-sm-6'>
        <input
          type='text'
          className='form-control'
          id='documentName'
          placeholder='Document Name'
          disabled
          value={props.documentName}
        />
      </div>
    </>
  );
}

export function TemplateSelector(props: {
  availableLfTemplates: WTemplateInfo[];
  selectedTemplateName: string;
  repoClient: IRepositoryApiClientExInternal;
  onChangeTemplate: (event: ChangeEvent<HTMLSelectElement>) => void;
}) {
  const laserficheTemplateOptions = props.availableLfTemplates?.map((item) => (
    <option key={item.id} value={item.displayName}>{item.displayName}</option>
  ));
  return (
    <>
      <label htmlFor='dwl2' className='col-sm-2 col-form-label'>
        Laserfiche Template
      </label>
      <div className='col-sm-6'>
        <select
          className='custom-select'
          id='documentTemplate'
          onChange={(e) => props.onChangeTemplate(e)}
          value={props.selectedTemplateName}
        >
          <option>None</option>
          {laserficheTemplateOptions}
        </select>
      </div>
    </>
  );
}
export function SharePointLaserficheColumnMatching(props: {
  profileConfig: ProfileConfiguration;
  availableSPFields: SPProfileConfigurationData[];
  lfFieldsForSelectedTemplate: TemplateFieldInfo[];
  validate: boolean;
  handleProfileConfigUpdate: (profileConfig: ProfileConfiguration) => void;
}) {
  const [deleteModal, setDeleteModal] = useState<JSX.Element | undefined>(
    undefined
  );
  const handleSpFieldChange = (
    e: ChangeEvent<HTMLSelectElement>,
    mapping: MappedFields
  ) => {
    const targetElement = e.target as HTMLSelectElement;
    const newConfig = { ...props.profileConfig };
    const rowsArray = [...newConfig.mappedFields];
    const currentRow = rowsArray.findIndex((row) => {
      return mapping === row;
    });
    const spField = props.availableSPFields.find(
      (data) => data.InternalName === targetElement.value
    );
    if (spField) {
      rowsArray[currentRow].spField = spField;
    }
    newConfig.mappedFields = rowsArray;
    props.handleProfileConfigUpdate(newConfig);
  };
  const handleLfFieldChange = (
    e: ChangeEvent<HTMLSelectElement>,
    mapping: MappedFields
  ) => {
    const targetElement = e.target as HTMLSelectElement;
    const newConfig = { ...props.profileConfig };
    const rowsArray = [...newConfig.mappedFields];
    const currentRow = rowsArray.findIndex((row) => {
      return mapping === row;
    });
    const lfField = props.lfFieldsForSelectedTemplate?.find(
      (data) => data.id.toString() === targetElement.value
    );
    if (lfField) {
      rowsArray[currentRow].lfField = lfField;
    }
    newConfig.mappedFields = rowsArray;
    props.handleProfileConfigUpdate(newConfig);
  };

  function CloseModalUp() {
    setDeleteModal(undefined);
  }
  const RemoveSpecificMapping = (idx) => {
    const del = (
      <DeleteModal
        configurationName='the field mapping'
        onCancel={CloseModalUp}
        onConfirmDelete={() => DeleteMapping(idx)}
      />
    );
    setDeleteModal(del);
  };
  function DeleteMapping(id: number) {
    const newConfig = { ...props.profileConfig };
    const rows = [...props.profileConfig.mappedFields];
    rows.splice(id, 1);
    newConfig.mappedFields = rows;
    props.handleProfileConfigUpdate(newConfig);
    setDeleteModal(undefined);
  }

  const AddNewMappingFields = () => {
    if (props.profileConfig.selectedTemplateName) {
      const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(
        36
      );
      const item = {
        id: id,
        spField: undefined,
        lfField: undefined,
      };
      const profileConfig = { ...props.profileConfig };
      profileConfig.mappedFields = [...profileConfig.mappedFields, item];
      props.handleProfileConfigUpdate(profileConfig);
    }
  };
  const spFields = props.availableSPFields?.slice()?.map((field) => (
    <option key={field.InternalName} value={field.InternalName}>
      {field.Title} ({field.TypeAsString})
    </option>
  ));
  const laserficheFields = props.lfFieldsForSelectedTemplate?.map((items) => {
    return (
      <option key={items.id} value={items.id}>
        {items.name} ({items.fieldType})
      </option>
    );
  });

  let fullValidationError = undefined;
  if (props.validate) {
    if (
      (props.profileConfig.mappedFields.some((item) => !item.spField) ||
        props.profileConfig.mappedFields.some((items) => !items.lfField)) &&
      props.profileConfig.selectedTemplateName
    ) {
      fullValidationError = (
        <span>Please ensure all fields are correctly mapped</span>
      );
    }
  }
  // TODO also filter for ones that are already mapped
  const optionalFields = props.lfFieldsForSelectedTemplate
    ?.filter((field) => !field.isRequired)
    ?.map((items) => {
      return (
        <option key={items.id} value={items.id}>
          {items.name} ({items.fieldType})
        </option>
      );
    });
  const mappedList = props.profileConfig.mappedFields?.map(
    (fieldMapping, index) => {
      const errorMessageMapping: JSX.Element | undefined =
        getMappingErrorMessage(fieldMapping);
      return (
        <tr id={index.toString()} key={index}>
          <td>
            <select
              name='SharePointField'
              className='custom-select'
              value={fieldMapping.spField?.InternalName ?? 'Select'}
              id={fieldMapping.id}
              onChange={(e) => handleSpFieldChange(e, fieldMapping)}
            >
              <option>Select</option>
              {spFields}
            </select>
          </td>
          <td>
            <select
              name='LaserficheField'
              className='custom-select'
              value={fieldMapping.lfField?.id ?? 'Select'}
              id={fieldMapping.id}
              disabled={fieldMapping.lfField?.isRequired}
              onChange={(e) => handleLfFieldChange(e, fieldMapping)}
            >
              <option>Select</option>
              {fieldMapping.lfField?.isRequired
                ? laserficheFields
                : optionalFields}
            </select>
          </td>
          <td>
            {fieldMapping.lfField?.isRequired ? (
              <span style={{ fontSize: '13px', color: 'red' }}>
                *Required field in Laserfiche
              </span>
            ) : (
              <a
                href='javascript:;'
                className='ml-3'
                onClick={() => RemoveSpecificMapping(index)}
              >
                <span className='material-icons'>delete</span>
              </a>
            )}
            {errorMessageMapping}
          </td>
        </tr>
      );
    }
  );

  return (
    <>
      {props.profileConfig.selectedTemplateName ? (
        <>
          <table className='table table-sm'>
            <thead>
              <tr>
                <th className='text-center' style={{ width: '39%' }}>
                  SharePoint Column
                </th>
                <th className='text-center' style={{ width: '38%' }}>
                  Laserfiche Field
                </th>
              </tr>
            </thead>
            <tbody id='tableEditBodyId'>{mappedList}</tbody>
          </table>
          {fullValidationError}
          <a
            onClick={AddNewMappingFields}
            className='btn btn-primary pl-5 pr-5 float-right ml-2'
          >
            Add Field
          </a>
        </>
      ) : (
        <span>Please select a template above to map fields</span>
      )}
      <div
        className='modal'
        id='deleteModal'
        hidden={!deleteModal}
        data-backdrop='static'
        data-keyboard='false'
      >
        {deleteModal}
      </div>
    </>
  );
}

export function DeleteModal(props: {
  configurationName: string;
  onConfirmDelete: () => void;
  onCancel: () => void;
}) {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className='modal-content'>
        <div className='modal-header'>
          <h5 className='modal-title' id='ModalLabel'>
            Delete Confirmation
          </h5>
          <button
            type='button'
            className='close'
            data-dismiss='modal'
            aria-label='Close'
            onClick={props.onCancel}
          >
            <span aria-hidden='true'>&times;</span>
          </button>
        </div>
        <div className='modal-body'>
          Do you want to permanently delete &quot;
          {props.configurationName}&quot;?
        </div>
        <div className='modal-footer'>
          <button
            type='button'
            className='btn btn-primary btn-sm'
            data-dismiss='modal'
            onClick={props.onConfirmDelete}
          >
            OK
          </button>
          <button
            type='button'
            className='btn btn-secondary btn-sm'
            data-dismiss='modal'
            onClick={props.onCancel}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
function getMappingErrorMessage(
  mappedField: MappedFields
): JSX.Element | undefined {
  if (mappedField.lfField && mappedField.spField) {
    const spFieldtype = mappedField.spField.TypeAsString;
    const lfFieldtype = mappedField.lfField.fieldType;
    const hasMismatch = hasFieldTypeMismatch(mappedField);

    if (hasMismatch) {
      return (
        <div
          style={{
            color: 'red',
            fontSize: '13px',
            marginLeft: '10px',
          }}
          title={`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${lfFieldtype}`}
        >
          SharePoint field type of {spFieldtype} cannot be mapped with
          Laserfiche field type of {lfFieldtype}
          <span className='material-icons'>warning</span>Data types mismatch
        </div>
      );
    } else {
      return undefined;
    }
  }
  return undefined;
}

export function hasFieldTypeMismatch(mapped: MappedFields) {
  const lfFieldType = mapped.lfField.fieldType;
  const spFieldType = mapped.spField.TypeAsString;
  if (
    lfFieldType == 'DateTime' ||
    lfFieldType == 'Date' ||
    lfFieldType == 'Time'
  ) {
    if (spFieldType != 'DateTime') {
      return true;
    }
  } else if (lfFieldType == 'LongInteger' || lfFieldType == 'ShortInteger') {
    if (spFieldType != 'Number') {
      return true;
    }
  } else if (lfFieldType == 'Number') {
    if (spFieldType != 'Number' && spFieldType != 'Currency') {
      return true;
    }
  } else if (lfFieldType == 'List') {
    if (spFieldType != 'Choice') {
      return true;
    }
  }
  return false;
}

export function validateNewConfiguration(profileConfig: ProfileConfiguration) {
  if(!profileConfig.ConfigurationName || profileConfig.ConfigurationName.length === 0) {
    return false;
  }
  for (const mapped of profileConfig.mappedFields) {
    if (!mapped.spField || !mapped.lfField) {
      return false;
    }
    if(hasFieldTypeMismatch(mapped)) {
      return false;
    }
  }
  return true;
}