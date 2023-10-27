// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { BaseDialog } from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import {
  LASERFICHE_ADMIN_CONFIGURATION_NAME,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  MANAGE_CONFIGURATIONS,
  MANAGE_MAPPING,
  SP_LOCAL_STORAGE_KEY,
} from '../../webparts/constants';
import {
  ISPDocumentData,
  ProfileMappingConfiguration,
} from '../../Utils/Types';
import { Navigation } from 'spfx-navigation';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  ProfileConfiguration,
  SPProfileConfigurationData,
} from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import {
  IPostEntryWithEdocMetadataRequest,
  FieldToUpdate,
  IPutFieldValsRequest,
  PutFieldValsRequest,
  IValueToUpdate,
  TemplateFieldInfo,
  ValueToUpdate,
  WFieldType,
} from '@laserfiche/lf-repository-api-client';
import { IListItem } from '../../webparts/laserficheAdminConfiguration/components/IListItem';
import { getSPListURL } from '../../Utils/Funcs';
import SaveToLaserficheCustomDialog from './SaveToLaserficheDialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import LoadingDialog from './CommonDialogs';
import { SPComponentLoader } from '@microsoft/sp-loader';

const signInPageRoute = '/SitePages/LaserficheSignIn.aspx';

export class GetDocumentDataCustomDialog extends BaseDialog {
  successful = false;

  constructor(
    private fileInfo: {
      fileName: string;
      spContentType: string;
      spFileUrl: string;
      fileId: string;
    },
    private context: BaseComponentContext
  ) {
    super();
  }

  showNextDialog: (data: ISPDocumentData) => Promise<void> = async (
    data: ISPDocumentData
  ) => {
    const saveToLfDialog = new SaveToLaserficheCustomDialog(data, () =>
      this.close()
    );
    await this.secondaryDialogProvider.show(saveToLfDialog);
    if (!saveToLfDialog.successful) {
      Navigation.navigate(
        this.context.pageContext.web.absoluteUrl + signInPageRoute,
        true
      );
    }
  };

  public render(): void {
    const element: React.ReactElement = (
      <React.StrictMode>
        <GetDocumentDialogData
          spFileInfo={this.fileInfo}
          context={this.context}
          showSaveToDialog={this.showNextDialog}
          handleCancelDialog={this.close}
        />
      </React.StrictMode>
    );
    ReactDOM.render(element, this.domElement);
  }

  protected onAfterClose(): void {
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onAfterClose();
  }
}

const FOLLOWING_SP_FIELDS_NO_VALUE_FOR_DOC_BUT_REQUIRED_IN_LASERFICHE_BASED_ON_MAPPINGS =
  'The following SharePoint fields are mapped to required fields in Laserfiche and must have valid values:';
const PLEASE_ENSURE_FIELDS_EXIST_FOR_DOCUMENT_AND_TRY_AGAIN =
  'Please fill out the required fields and try again.';
const CANCEL = 'Cancel';
const NO_SP_CONTENT_TYPE_EXISTS_AND_NO_DEFAULT_MAPPING =
  'No SharePoint Content Type exists for this document and no default mapping exists.';
const PLEASE_UPDATE_CONTENT_TYPE_OR_CONTACT_ADMIN_FOR_DEFAULT_MAPPING =
  'Please update the Content Type or contact your administrator to set up a default mapping.';

function GetDocumentDialogData(props: {
  showSaveToDialog: (fileData: ISPDocumentData) => void;
  handleCancelDialog: () => Promise<void>;
  spFileInfo: {
    fileName: string;
    spContentType: string;
    spFileUrl: string;
    fileId: string;
  };
  context: BaseComponentContext;
}): JSX.Element {
  const [missingFields, setMissingFields] = React.useState<
    undefined | SPProfileConfigurationData[]
  >(undefined);

  const [error, setError] = React.useState<JSX.Element | undefined>(undefined);

  const listFields = (
    <ul>
      {missingFields?.map((field) => (
        <li key={field.Title}>{field.Title}</li>
      ))}
    </ul>
  );

  React.useEffect(() => {
    SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
    SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);

    void saveDocumentToLaserficheAsync();
  }, []);

  async function saveDocumentToLaserficheAsync(): Promise<void> {
    try {
      const libraryUrl = props.context.pageContext.list.title;
      const allSPFieldValues: { [key: string]: string } =
        await getAllFieldsValuesAsync(libraryUrl, props.spFileInfo.fileId);
      const allSPFieldProperties: SPProfileConfigurationData[] =
        await getAllFieldsPropertiesAsync(libraryUrl);
      const docData = await getDocumentDataAsync(
        allSPFieldValues,
        allSPFieldProperties
      );

      if (docData) {
        window.localStorage.setItem(
          SP_LOCAL_STORAGE_KEY,
          JSON.stringify(docData)
        );

        props.showSaveToDialog(docData);
      }
    } catch (err) {
      setError(<div>{`Error saving: ${err.message}`}</div>);
      console.error(err);
    }
  }

  async function getAllFieldsPropertiesAsync(
    libraryUrl: string
  ): Promise<SPProfileConfigurationData[]> {
    const res = await fetch(
      `${getSPListURL(
        props.context,
        libraryUrl
      )}/Fields?$filter=Group ne '_Hidden'`,
      {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      }
    );
    const results = await res.json();
    const spFieldNameDefs: SPProfileConfigurationData[] = results.value;
    return spFieldNameDefs;
  }

  async function getDocumentDataAsync(
    allSpFieldValues: { [key: string]: string },
    allSPFieldProperties: SPProfileConfigurationData[]
  ): Promise<ISPDocumentData | undefined> {
    const response: SPHttpClientResponse = await props.context.spHttpClient.get(
      `${getSPListURL(
        props.context,
        LASERFICHE_ADMIN_CONFIGURATION_NAME
      )}/items?$filter=Title eq '${MANAGE_MAPPING}'&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
        },
      }
    );

    const itemsWithTitleManageMapping = await response.json();
    let matchingMapping = undefined;
    if (itemsWithTitleManageMapping.value?.length > 0) {
      const manageMappingListItem: IListItem =
        itemsWithTitleManageMapping.value[0];
      const manageMappingDetails: ProfileMappingConfiguration[] = JSON.parse(
        manageMappingListItem.JsonValue
      );
      matchingMapping = manageMappingDetails.find(
        (el) => el.SharePointContentType === props.spFileInfo.spContentType
      );
      if (!matchingMapping) {
        matchingMapping = manageMappingDetails.find(
          (el) => el.SharePointContentType === 'DEFAULT'
        );
      }
    }

    let docData: ISPDocumentData;
    if (!matchingMapping) {
      if (!props.spFileInfo.spContentType) {
        setError(
          <>
            <div>{`${NO_SP_CONTENT_TYPE_EXISTS_AND_NO_DEFAULT_MAPPING}`}</div>
            <div>{`${PLEASE_UPDATE_CONTENT_TYPE_OR_CONTACT_ADMIN_FOR_DEFAULT_MAPPING}`}</div>
          </>
        );
      } else {
        const NO_MAPPING_EXISTS = `No mapping exists for SharePoint Content Type "${props.spFileInfo.spContentType}" and no default mapping exists.`;
        setError(
          <>
            <div>{`${NO_MAPPING_EXISTS}`}</div>
            <div>{`${PLEASE_UPDATE_CONTENT_TYPE_OR_CONTACT_ADMIN_FOR_DEFAULT_MAPPING}`}</div>
          </>
        );
      }
    } else {
      docData = await getDocumentDataWithMapping(
        matchingMapping,
        allSpFieldValues,
        allSPFieldProperties
      );
    }
    return docData;
  }

  async function getDocumentDataWithMapping(
    matchingMapping: ProfileMappingConfiguration,
    allSpFieldValues: { [key: string]: string },
    allSPFieldProperties: SPProfileConfigurationData[]
  ): Promise<ISPDocumentData> {
    const laserficheProfile = matchingMapping.LaserficheContentType;

    const adminConfigList = await props.context.spHttpClient.get(
      `${getSPListURL(
        props.context,
        LASERFICHE_ADMIN_CONFIGURATION_NAME
      )}/items?$filter=Title eq '${MANAGE_CONFIGURATIONS}'&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
        },
      }
    );
    const adminConfigListJson = await adminConfigList.json();

    const allConfigs: ProfileConfiguration[] = JSON.parse(
      adminConfigListJson.value[0].JsonValue
    );
    const matchingLFConfig = allConfigs.find(
      (lfConfig) => lfConfig.ConfigurationName === laserficheProfile
    );
    if (matchingLFConfig.selectedTemplateName?.length > 0) {
      const metadata: IPostEntryWithEdocMetadataRequest = {
        template: matchingLFConfig.selectedTemplateName,
      };
      const missingRequiredFields: SPProfileConfigurationData[] = [];
      const fields: { [key: string]: FieldToUpdate } = {};
      formatMetadata(
        matchingLFConfig,
        missingRequiredFields,
        allSpFieldValues,
        allSPFieldProperties,
        fields
      );

      if (missingRequiredFields.length === 0) {
        const fileData = getDocumentDataWithMetadata(
          fields,
          metadata,
          matchingLFConfig,
          laserficheProfile
        );
        return fileData;
      } else {
        setMissingFields(missingRequiredFields);
        return undefined;
      }
    } else {
      const fileData: ISPDocumentData = {
        action: matchingLFConfig.Action,
        fileName: props.spFileInfo.fileName,
        fileUrl: props.spFileInfo.spFileUrl,
        documentName: matchingLFConfig.DocumentName,
        entryId: matchingLFConfig.selectedFolder.id,
        contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
        lfProfile: laserficheProfile,
      };
      return fileData;
    }
  }

  async function getAllFieldsValuesAsync(
    libraryUrl: string,
    fileId: string
  ): Promise<{ [key: string]: string }> {
    const res = await props.context.spHttpClient.get(
      `${getSPListURL(
        props.context,
        libraryUrl
      )}/items(${fileId})/FieldValuesForEdit`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      }
    );

    const allSpFieldValues = await res.json();
    return allSpFieldValues;
  }

  function getDocumentDataWithMetadata(
    fields: { [key: string]: FieldToUpdate },
    metadata: IPostEntryWithEdocMetadataRequest,
    matchingLFConfig: ProfileConfiguration,
    laserficheProfileName: string
  ): ISPDocumentData {
    const metadataFields: IPutFieldValsRequest = {
      fields,
    };
    metadata.metadata = new PutFieldValsRequest(metadataFields);

    const fileData: ISPDocumentData = {
      action: matchingLFConfig.Action,
      contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
      documentName: matchingLFConfig.DocumentName,
      templateName: matchingLFConfig.selectedTemplateName,
      entryId: matchingLFConfig.selectedFolder.id,
      fileUrl: props.spFileInfo.spFileUrl,
      fileName: props.spFileInfo.fileName,
      metadata,
      lfProfile: laserficheProfileName,
    };

    return fileData;
  }

  function formatMetadata(
    matchingLFConfig: ProfileConfiguration,
    missingRequiredFields: SPProfileConfigurationData[],
    allSpFieldValues: { [key: string]: string },
    allSPFieldProperties: SPProfileConfigurationData[],
    fields: { [key: string]: FieldToUpdate }
  ): void {
    for (const mapping of matchingLFConfig.mappedFields) {
      const spFieldName = mapping.spField.InternalName;
      // TODO which one to use?
      let spDocFieldValue: string =
        allSpFieldValues[spFieldName] ??
        allSpFieldValues[mapping.spField.Title];

      if (spDocFieldValue?.length > 0) {
        const lfField = mapping.lfField;

        spDocFieldValue = forceTruncateToFieldTypeLength(
          lfField,
          spDocFieldValue
        );
        spDocFieldValue = spDocFieldValue.replace(/[\\]/g, `\\\\`);
        spDocFieldValue = spDocFieldValue.replace(/["]/g, `\\"`);

        if (
          lfField.isRequired &&
          (!spDocFieldValue || spDocFieldValue.length === 0)
        ) {
          const currentField: SPProfileConfigurationData | undefined =
            allSPFieldProperties.find(
              (prop) => prop.InternalName === mapping.spField.InternalName
            );
          missingRequiredFields.push(currentField);
        }

        const valueToUpdate: IValueToUpdate = {
          value: spDocFieldValue,
          position: 1,
        };
        const newValueToUpdate = new ValueToUpdate(valueToUpdate);
        fields[lfField.name] = new FieldToUpdate({
          values: [newValueToUpdate],
        });
      } else {
        if (mapping.lfField.isRequired) {
          missingRequiredFields.push(mapping.spField);
        }
      }
    }
  }

  function forceTruncateToFieldTypeLength(
    lfField: TemplateFieldInfo,
    spDocFieldValue: string
  ): string {
    if (lfField.length !== 0) {
      if (spDocFieldValue.length > lfField.length) {
        // automatically trims length to match constraint
        spDocFieldValue = spDocFieldValue.slice(0, lfField.length);
      }
    } else if (
      lfField.fieldType === WFieldType.ShortInteger ||
      lfField.fieldType === WFieldType.LongInteger ||
      lfField.fieldType === WFieldType.Number
    ) {
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9.]/g, '');
      spDocFieldValue = extractOnlyNumbers;
    }
    return spDocFieldValue;
  }

  return (
    <div className={styles.wrapper}>
      <div className={styles.header}>
        <div className={styles.logoHeader}>
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
          onClick={props.handleCancelDialog}
        >
          <span className='material-icons-outlined'> close </span>
        </button>
      </div>

      <div className={styles.contentBox}>
        {!(missingFields?.length > 0) && !error && <LoadingDialog />}
        {missingFields?.length > 0 && (
          <MissingFieldsDialog missingFields={listFields} />
        )}
        {error}
      </div>

      <div className={styles.footer}>
        <button
          onClick={props.handleCancelDialog}
          className='lf-button sec-button'
        >
          {CANCEL}
        </button>
      </div>
    </div>
  );
}

function MissingFieldsDialog(props: {
  missingFields: JSX.Element;
}): JSX.Element {
  const textInside = (
    <span>
      {
        FOLLOWING_SP_FIELDS_NO_VALUE_FOR_DOC_BUT_REQUIRED_IN_LASERFICHE_BASED_ON_MAPPINGS
      }
      {props.missingFields}
      {PLEASE_ENSURE_FIELDS_EXIST_FOR_DOCUMENT_AND_TRY_AGAIN}
    </span>
  );

  return (
    <div>
      <p>{textInside}</p>
    </div>
  );
}
