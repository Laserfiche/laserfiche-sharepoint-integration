// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import styles from './SendToLaserFiche.module.scss';
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
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  ProfileConfiguration,
  SPProfileConfigurationData,
} from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import {
  FieldType,
  FieldToUpdate,
  IImportEntryRequestMetadata,
  TemplateFieldDefinition,
} from '@laserfiche/lf-repository-api-client-v2';
import { IListItem } from '../../webparts/laserficheAdminConfiguration/components/IListItem';
import { getSPListURL } from '../../Utils/Funcs';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import LoadingDialog, { useConfirm } from './CommonDialogs';
import { COULD_NOT_DETERMINE_CONTENT_TYPE, THERE_WAS_AN_ISSUE_DETERMINING_CONTENT_TYPE_OF_ITEM_DEFAULT_MAPPING_WILL_BE_USED } from '../../webparts/strings';

const CANCEL = 'Cancel';
const NO_SP_CONTENT_TYPE_EXISTS_AND_NO_DEFAULT_MAPPING =
  'No SharePoint Content Type exists for this document and no default mapping exists.';
const PLEASE_UPDATE_CONTENT_TYPE_OR_CONTACT_ADMIN_FOR_DEFAULT_MAPPING =
  'Please update the Content Type or contact your administrator to set up a default mapping.';

const FOLLOWING_SP_FIELDS_NO_VALUE_FOR_DOC_BUT_REQUIRED_IN_LASERFICHE_BASED_ON_MAPPINGS =
  'The following SharePoint fields are mapped to required fields in Laserfiche and must have valid values:';
const PLEASE_ENSURE_FIELDS_EXIST_FOR_DOCUMENT_AND_TRY_AGAIN =
  'Please fill out the required fields and try again.';

export function GetDocumentDialogData(props: {
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
  const [getConfirmation, Confirmation] = useConfirm();

  const [showLoading, setShowLoading] = React.useState<boolean>(false);

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

      if (!props.spFileInfo.spContentType) {
        const warn = await getConfirmation(THERE_WAS_AN_ISSUE_DETERMINING_CONTENT_TYPE_OF_ITEM_DEFAULT_MAPPING_WILL_BE_USED);
        if (!warn) {
          console.warn('Content type could not be determined. User chose to cancel operation.');
          await props.handleCancelDialog();
          return;
        } else {
          setShowLoading(true);
        }
      }
      else {
        setShowLoading(true);
      }

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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } catch (err: any) {
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
      const metadata: IImportEntryRequestMetadata = {
        templateName: matchingLFConfig.selectedTemplateName,
      };
      const missingRequiredFields: SPProfileConfigurationData[] = [];
      const fields: FieldToUpdate[] = [];
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
    fields: FieldToUpdate[],
    metadata: IImportEntryRequestMetadata,
    matchingLFConfig: ProfileConfiguration,
    laserficheProfileName: string
  ): ISPDocumentData {
    metadata.fields = fields;

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
    fields: FieldToUpdate[]
  ): void {
    for (const mapping of matchingLFConfig.mappedFields) {
      const spFieldName = mapping.spField.EntityPropertyName;

      let spDocFieldValue = allSpFieldValues[spFieldName];
      if (!spDocFieldValue) {
        const valueForTrimmedKey = Object.entries(allSpFieldValues).find(
          ([key]) => key.replace(/x005f_/g, '') === spFieldName
        )?.[1];

        spDocFieldValue = valueForTrimmedKey;
      }

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

        fields.push(
          new FieldToUpdate({
            name: lfField.name,
            values: [spDocFieldValue],
          })
        );
      } else {
        if (mapping.lfField.isRequired) {
          missingRequiredFields.push(mapping.spField);
        }
      }
    }
  }

  function forceTruncateToFieldTypeLength(
    lfField: TemplateFieldDefinition,
    spDocFieldValue: string
  ): string {
    if (lfField.length !== 0) {
      if (spDocFieldValue.length > lfField.length) {
        // automatically trims length to match constraint
        spDocFieldValue = spDocFieldValue.slice(0, lfField.length);
      }
    } else if (
      lfField.fieldType === FieldType.ShortInteger ||
      lfField.fieldType === FieldType.LongInteger ||
      lfField.fieldType === FieldType.Number
    ) {
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9.]/g, '');
      spDocFieldValue = extractOnlyNumbers;
    }
    return spDocFieldValue;
  }

  return (
    <>
      {showLoading && (
        <div className={styles.wrapper}>
          <div className={styles.header}>
            <div className={styles.logoHeader}>
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
      )}
      <Confirmation cancelButtonText={CANCEL} headerText={COULD_NOT_DETERMINE_CONTENT_TYPE}/>
    </>
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
