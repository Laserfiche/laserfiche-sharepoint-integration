import { NgElement, WithProperties } from '@angular/elements';
import {
  EntryType,
  TemplateFieldInfo,
  WFieldType,
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
import { getCorrespondingTypeFieldName } from '../../../Utils/Funcs';
import styles from './LaserficheAdminConfiguration.module.scss';

export interface ProfileConfiguration {
  ConfigurationName: string;
  DocumentName: string;
  selectedTemplateName?: string;
  selectedFolder?: LfFolder;
  Action: ActionTypes;
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
  'COPY' = 'COPY',
  'MOVE_AND_DELETE' = 'MOVE_AND_DELETE',
  'REPLACE' = 'REPLACE',
}

export function ProfileHeader(props: {
  configurationName: string;
}): JSX.Element {
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
}): JSX.Element {
  const [showFolderModal, setShowFolderModal] = useState(false);

  const selectedEntryNodePath = props.profileConfig.selectedFolder?.path;

  const onSelectFolderAsync: (
    selectedNode: LfRepoTreeNode | undefined
  ) => Promise<void> = async (selectedNode: LfRepoTreeNode | undefined) => {
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

  const handleTemplateChange: (
    event: React.ChangeEvent<HTMLSelectElement>
  ) => void = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = (event.target as HTMLSelectElement).value;
    const templateName = value;
    const profileConfig = { ...props.profileConfig };
    profileConfig.selectedTemplateName = templateName;
    props.handleProfileConfigUpdate(profileConfig);
    props.handleTemplateChange(templateName);
  };

  const handleActionTypeChange: (
    event: React.ChangeEvent<HTMLSelectElement>
  ) => void = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = (event.target as HTMLSelectElement).value;
    const actionName = value;
    const profileConfig = { ...props.profileConfig };
    profileConfig.Action = actionName as ActionTypes;
    props.handleProfileConfigUpdate(profileConfig);
  };

  function closeFolderModalUp(): void {
    setShowFolderModal(false);
  }

  function openFolderModal(): void {
    setShowFolderModal(true);
  }
  return (
    <>
      {/* Do not need document name for now as the document will always be saved with the SharePoint document name until tokens are supported */}
      {/* <div className={`${styles.formGroupRow} form-group row`}>
        <DocumentName documentName={props.profileConfig?.DocumentName} />
      </div> */}
      <div className={`${styles.formGroupRow} form-group row`}>
        <TemplateSelector
          availableLfTemplates={props.availableLfTemplates}
          selectedTemplateName={props.profileConfig?.selectedTemplateName}
          repoClient={props.repoClient}
          onChangeTemplate={handleTemplateChange}
        />
      </div>
      <div className={`${styles.formGroupRow} form-group row`}>
        <label htmlFor='txt3' className='col-sm-3 col-form-label'>
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
        </div>
        <div className='col-sm-2' id='folderModal' style={{ marginTop: '2px' }}>
          <button className='lf-button sec-button' onClick={openFolderModal}>
            Browse
          </button>
        </div>
      </div>
      <div className={`${styles.formGroupRow} form-group row`}>
        <label htmlFor='dwl4' className='col-sm-3 col-form-label'>
          After import
        </label>
        <div className='col-sm-6'>
          <select
            onChange={handleActionTypeChange}
            defaultValue={props.profileConfig.Action}
            className='custom-select'
            id='action'
          >
            <option value={ActionTypes.COPY}>
              Leave a copy of the file in SharePoint
            </option>
            <option value={ActionTypes.REPLACE}>
              Replace SharePoint file with a link to the document in Laserfiche
            </option>
            <option value={ActionTypes.MOVE_AND_DELETE}>
              Delete SharePoint file
            </option>
          </select>
        </div>
      </div>
      {showFolderModal && (
        <div
          className={styles.modal}
          id='folderModal'
          data-backdrop='static'
          data-keyboard='false'
        >
          <RepositoryBrowserModal
            repoClient={props.repoClient}
            CloseFolderBrowserUp={closeFolderModalUp}
            selectedEntryNodePath={selectedEntryNodePath}
            SelectFolder={onSelectFolderAsync}
          />
        </div>
      )}
    </>
  );
}

export function RepositoryBrowserModal(props: {
  CloseFolderBrowserUp: () => void;
  SelectFolder: (node: LfRepoTreeNode | undefined) => void;
  selectedEntryNodePath: string;
  repoClient: IRepositoryApiClientExInternal;
}): JSX.Element {
  const [shouldShowOpen, setShouldShowOpen] = useState(false);
  const [shouldShowSelect, setShouldShowSelect] = useState(false);
  const [shouldDisableSelect, setShouldDisableSelect] = useState(false);

  const [entrySelected, setEntrySelected] = useState<
    LfRepoTreeNode | undefined
  >(undefined);
  const repositoryBrowser: React.RefObject<
    NgElement & WithProperties<LfRepositoryBrowserComponent>
  > = React.useRef();
  const onEntrySelected: (event: CustomEvent<LfRepoTreeNode[]>) => void = (
    event: CustomEvent<LfRepoTreeNode[]>
  ) => {
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
  let lfRepoTreeService: LfRepoTreeNodeService;

  React.useEffect(() => {
    if (props.repoClient) {
      lfRepoTreeService = new LfRepoTreeNodeService(props.repoClient);
      lfRepoTreeService.viewableEntryTypes = [
        EntryType.Folder,
        EntryType.Shortcut,
      ];
      void initializeTreeAsync();
    }
  }, [props.repoClient]);

  const isNodeSelectable: (node: LfRepoTreeNode) => boolean = (
    node: LfRepoTreeNode
  ) => {
    if (node?.entryType === EntryType.Folder) {
      return true;
    } else if (
      node?.entryType === EntryType.Shortcut &&
      node?.targetType === EntryType.Folder
    ) {
      return true;
    } else {
      return false;
    }
  };

  async function initializeTreeAsync(): Promise<void> {
    try {
      if (!props.repoClient) {
        throw new Error('RepoId is undefined');
      }
      repositoryBrowser.current?.addEventListener(
        'entrySelected',
        onEntrySelected
      );

      if (lfRepoTreeService) {
        await repositoryBrowser?.current?.initAsync(
          lfRepoTreeService,
          props.selectedEntryNodePath
        );
      } else {
        console.debug(
          'Unable to initialize tree, lfRepoTreeService is undefined'
        );
      }
    } catch (err) {
      console.error(`Unable to initialize repository browser: ${err}`);
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

  const onSelectFolder: () => void = () => {
    props.SelectFolder(
      repositoryBrowser?.current?.currentFolder as LfRepoTreeNode
    );
  };

  const onOpenNode: () => Promise<void> = async () => {
    await repositoryBrowser?.current?.openSelectedNodesAsync();
    setShouldShowOpen(getShouldShowOpen());
    setShouldShowSelect(getShouldShowSelect());
  };

  return (
    <>
      <div className={`${styles.wrapper}`}>
        <div className={styles.header}>
          <div className={styles.logoHeader}>
            <div>Select folder</div>
          </div>

          <button
            className={styles.lfCloseButton}
            title='close'
            onClick={props.CloseFolderBrowserUp}
          >
            <span className='material-icons-outlined'> close </span>
          </button>
        </div>

        <div className={styles.contentBox}>
          <lf-repository-browser
            ref={repositoryBrowser}
            multiple='false'
            isSelectable={isNodeSelectable}
          />
        </div>

        <div className={styles.footer}>
          {shouldShowOpen && (
            <button className={`lf-button primary-button`} onClick={onOpenNode}>
              Open
            </button>
          )}
          {shouldShowSelect && (
            <button
              className='lf-button primary-button'
              onClick={onSelectFolder}
              disabled={shouldDisableSelect}
            >
              Select
            </button>
          )}
          <button
            className={`sec-button lf-button ${styles.marginLeftButton}`}
            onClick={props.CloseFolderBrowserUp}
          >
            Cancel
          </button>
        </div>
      </div>
    </>
  );
}

export function DocumentName(props: { documentName: string }): JSX.Element {
  return (
    <>
      <label htmlFor='txt1' className='col-sm-3 col-form-label'>
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
}): JSX.Element {
  const laserficheTemplateOptions = props.availableLfTemplates?.map((item) => (
    <option key={item.id} value={item.displayName}>
      {item.displayName}
    </option>
  ));
  return (
    <>
      <label htmlFor='dwl2' className='col-sm-3 col-form-label'>
        Laserfiche Template
      </label>
      <div className='col-sm-6'>
        <select
          className='custom-select'
          id='documentTemplate'
          onChange={(e) => props.onChangeTemplate(e)}
          value={props.selectedTemplateName}
        >
          <option value=''>None</option>
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
  hasError: (hasError: boolean) => void;
  handleProfileConfigUpdate: (profileConfig: ProfileConfiguration) => void;
}): JSX.Element {
  const [deleteModal, setDeleteModal] = useState<JSX.Element | undefined>(
    undefined
  );
  const handleSpFieldChange: (
    e: ChangeEvent<HTMLSelectElement>,
    mapping: MappedFields
  ) => void = (e: ChangeEvent<HTMLSelectElement>, mapping: MappedFields) => {
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
  const handleLfFieldChange: (
    e: ChangeEvent<HTMLSelectElement>,
    mapping: MappedFields
  ) => void = (e: ChangeEvent<HTMLSelectElement>, mapping: MappedFields) => {
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

  function closeModalUp(): void {
    setDeleteModal(undefined);
  }
  const removeSpecificMapping: (idx: number) => void = (idx: number) => {
    const del = (
      <DeleteModal
        configurationName='the field mapping'
        onCancel={closeModalUp}
        onConfirmDelete={() => deleteMapping(idx)}
      />
    );
    setDeleteModal(del);
  };
  function deleteMapping(id: number): void {
    const newConfig = { ...props.profileConfig };
    const rows = [...props.profileConfig.mappedFields];
    rows.splice(id, 1);
    newConfig.mappedFields = rows;
    props.handleProfileConfigUpdate(newConfig);
    setDeleteModal(undefined);
  }

  const addNewMappingFields: () => void = () => {
    if (props.profileConfig.selectedTemplateName) {
      const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(
        36
      );
      const item: MappedFields = {
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
        {items.name} ({getCorrespondingTypeFieldName(items.fieldType)})
      </option>
    );
  });

  let fullValidationError = undefined;
  if (props.validate) {
    if (
      (props.profileConfig.mappedFields?.some((item) => !item.spField) ||
        props.profileConfig.mappedFields?.some((items) => !items.lfField)) &&
      props.profileConfig.selectedTemplateName
    ) {
      fullValidationError = (
        <span>Please ensure all fields are correctly mapped</span>
      );
    }
  }

  function getAvailableOptionalFields(
    fieldMapping: MappedFields
  ): React.ReactNode {
    return props.lfFieldsForSelectedTemplate
      ?.filter(
        (field) =>
          !field.isRequired &&
          (field.id === fieldMapping.lfField?.id ||
            !props.profileConfig.mappedFields.find(
              (item) => item?.lfField?.id === field?.id
            ))
      )
      ?.map((items) => {
        return (
          <option key={items.id} value={items.id}>
            {items.name} ({getCorrespondingTypeFieldName(items.fieldType)})
          </option>
        );
      });
  }

  const mappedList = props.profileConfig.mappedFields?.map(
    (fieldMapping, index) => {
      const errorMessageMapping: JSX.Element | undefined =
        getMappingErrorMessage(fieldMapping);
      if (errorMessageMapping) {
        props.hasError(true);
      } else {
        props.hasError(false);
      }
      return (
        <>
          <div className={styles.rowDiv} id={index.toString()} key={index}>
            <span className={styles.dataCellWidth}>
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
            </span>
            <span className={styles.dataCellWidth}>
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
                  : getAvailableOptionalFields(fieldMapping)}
              </select>
            </span>
            <span>
              {!fieldMapping.lfField?.isRequired && (
                <button
                  className={styles.lfMaterialIconButton}
                  onClick={() => removeSpecificMapping(index)}
                >
                  <span className='material-icons-outlined'> close </span>
                </button>
              )}
            </span>
          </div>
          {fieldMapping.lfField?.isRequired && (
            <div style={{ display: 'flex' }}>
              <span className={styles.dataCellWidth} />
              <span
                className={styles.dataCellWidth}
                style={{ fontSize: '13px', color: 'red' }}
              >
                *Required field in Laserfiche
              </span>
            </div>
          )}

          {errorMessageMapping}
          <hr />
        </>
      );
    }
  );

  return (
    <>
      {props.profileConfig.selectedTemplateName ? (
        <>
          <div>
            <div className={styles.rowDiv}>
              <span className={styles.dataCellWidth}>SharePoint Column</span>
              <span className={styles.dataCellWidth}>Laserfiche Field</span>
            </div>
            <div id='tableEditBodyId'>{mappedList}</div>
          </div>
          {fullValidationError}
          <div className={styles.footerIcons}>
            <button
              onClick={addNewMappingFields}
              className='lf-button primary-button'
            >
              Add Field
            </button>
          </div>
        </>
      ) : (
        <span>Please select a template above to map fields</span>
      )}
      {deleteModal !== undefined && (
        <div
          className={styles.modal}
          id='deleteModal'
          data-backdrop='static'
          data-keyboard='false'
        >
          {deleteModal}
        </div>
      )}
    </>
  );
}

export function DeleteModal(props: {
  configurationName: string;
  onConfirmDelete: () => void;
  onCancel: () => void;
}): JSX.Element {
  return (
    <div className='modal-dialog modal-dialog-centered'>
      <div className={`modal-content ${styles.wrapper}`}>
        <div className={styles.header}>
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
        <div className={styles.contentBox}>
          Do you want to permanently delete &quot;
          {props.configurationName}&quot;?
        </div>
        <div className={styles.footer}>
          <button
            type='button'
            className='lf-button primary-button'
            data-dismiss='modal'
            onClick={props.onConfirmDelete}
          >
            OK
          </button>
          <button
            type='button'
            className={`lf-button sec-button ${styles.marginLeftButton}`}
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
    const lfFieldTypeDisplayName = getCorrespondingTypeFieldName(mappedField.lfField.fieldType);
    const hasMismatch = hasFieldTypeMismatch(mappedField);

    if (hasMismatch) {
      return (
        <div
          style={{
            fontSize: '13px',
            marginLeft: '10px',
            display: 'flex',
            alignItems: 'center',
          }}
          title={`SharePoint field type of ${spFieldtype} cannot be mapped with Laserfiche field type of ${getCorrespondingTypeFieldName}}`}
        >
          <span
            className='material-icons-outlined'
            style={{
              color: 'red',
            }}
          >
            warning
          </span>
          Data types mismatch. SharePoint field type of {spFieldtype} cannot be
          mapped with Laserfiche field type of {lfFieldTypeDisplayName}
        </div>
      );
    } else {
      return undefined;
    }
  }
  return undefined;
}

export function hasFieldTypeMismatch(mapped: MappedFields): boolean {
  const lfFieldType = mapped.lfField.fieldType;
  const spFieldType = mapped.spField.TypeAsString;
  if (
    lfFieldType === WFieldType.DateTime ||
    lfFieldType === WFieldType.Date ||
    lfFieldType === WFieldType.Time
  ) {
    if (spFieldType !== 'DateTime') {
      return true;
    }
  } else if (
    lfFieldType === WFieldType.LongInteger ||
    lfFieldType === WFieldType.ShortInteger
  ) {
    if (spFieldType !== 'Number') {
      return true;
    }
  } else if (lfFieldType === WFieldType.Number) {
    if (spFieldType !== 'Number' && spFieldType !== 'Currency') {
      return true;
    }
  } else if (lfFieldType === WFieldType.List) {
    if (spFieldType !== 'Choice') {
      return true;
    }
  }
  return false;
}

export function validateNewConfiguration(
  profileConfig: ProfileConfiguration
): boolean {
  const profileNameContainsSpecialCharacters = /[^ A-Za-z0-9]/.test(
    profileConfig.ConfigurationName
  );
  if (
    !profileConfig.ConfigurationName ||
    profileConfig.ConfigurationName.length === 0 ||
    profileNameContainsSpecialCharacters
  ) {
    return false;
  }
  if (profileConfig.mappedFields) {
    for (const mapped of profileConfig.mappedFields) {
      if (!mapped.spField || !mapped.lfField) {
        return false;
      }
      if (hasFieldTypeMismatch(mapped)) {
        return false;
      }
    }
  }
  return true;
}
