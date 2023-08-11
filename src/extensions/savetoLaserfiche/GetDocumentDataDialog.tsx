import { BaseDialog } from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import {
  ADMIN_CONFIGURATION_LIST,
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
  ActionTypes,
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
  ProblemDetails,
} from '@laserfiche/lf-repository-api-client';
import { IListItem } from '../../webparts/laserficheAdminConfiguration/components/IListItem';
import { getSPListURL } from '../../Utils/Funcs';
import SaveToLaserficheCustomDialog from './SaveToLaserficheDialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import LoadingDialog from './CommonDialogs';
import { SPComponentLoader } from '@microsoft/sp-loader';

const signInPageRoute = '/SitePages/LaserficheSpSignIn.aspx';

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

  showNextDialog: (data: ISPDocumentData) => Promise<void> = async (data: ISPDocumentData) => {
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

const FOLLOWING_SP_FIELDS_BLANK_MAPPED_TO_REQUIRED_LF_FIELDS =
  'The following SharePoint field values are blank and are mapped to required Laserfiche fields:';
const PLEASE_FILL_OUT_REQUIRED_FIELDS_TRY_AGAIN =
  'Please fill out these required fields and try again.';

const CANCEL = 'Cancel';

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

  if (missingFields) {
    <span>
      {FOLLOWING_SP_FIELDS_BLANK_MAPPED_TO_REQUIRED_LF_FIELDS}
      {missingFields}
      {PLEASE_FILL_OUT_REQUIRED_FIELDS_TRY_AGAIN}
    </span>;
  }

  const listFields = missingFields?.map((field) => (
    <div key={field.Title}>- {field.Title}</div>
  ));

  React.useEffect(() => {
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
    );
    SPComponentLoader.loadCss(
      'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
    );

    saveDocumentToLaserficheAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, []);

  async function saveDocumentToLaserficheAsync(): Promise<void> {
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
  }

  async function getAllFieldsPropertiesAsync(
    libraryUrl: string
  ): Promise<SPProfileConfigurationData[]> {
    try {
      const res = await fetch(
        `${getSPListURL(
          this.context,
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
      const spFieldNameDefs: SPProfileConfigurationData[] = JSON.parse(
        results.value
      );
      return spFieldNameDefs;
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  async function getDocumentDataAsync(
    allSpFieldValues: { [key: string]: string },
    allSPFieldProperties: SPProfileConfigurationData[]
  ): Promise<ISPDocumentData | undefined> {
    const response: SPHttpClientResponse = await props.context.spHttpClient.get(
      `${getSPListURL(
        props.context,
        ADMIN_CONFIGURATION_LIST
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
    }

    let docData: ISPDocumentData;
    if (!matchingMapping) {
      docData = getDocumentDataNoMetadata();
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
        ADMIN_CONFIGURATION_LIST
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
    try {
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
    } catch (error) {
      console.log('error ocurred' + error);
    }
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

  function getDocumentDataNoMetadata(): ISPDocumentData {
    const fileData: ISPDocumentData = {
      fileName: props.spFileInfo.fileName,
      documentName: props.spFileInfo.fileName,
      fileUrl: props.spFileInfo.spFileUrl,
      entryId: '1',
      contextPageAbsoluteUrl: props.context.pageContext.web.absoluteUrl,
      action: ActionTypes.COPY,
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
      const spFieldName = mapping.spField.Title;
      let spDocFieldValue: string = allSpFieldValues[spFieldName];

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
    } else if (lfField.fieldType === WFieldType.ShortInteger) {
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9]/g, '');
      const valueAsNumber = Number.parseInt(extractOnlyNumbers, 10);
      if (valueAsNumber > 64999 || valueAsNumber < 0) {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      } else {
        spDocFieldValue = extractOnlyNumbers;
      }
    } else if (lfField.fieldType === WFieldType.LongInteger) {
      const extractOnlyNumbers = spDocFieldValue.replace(/[^0-9]/g, '');
      const valueAsNumber = Number.parseInt(extractOnlyNumbers, 10);
      if (valueAsNumber > 3999999999 || valueAsNumber < 0) {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      } else {
        spDocFieldValue = extractOnlyNumbers;
      }
    } else if (lfField.fieldType === WFieldType.Number) {
      const valueRemoveNonNumbers = spDocFieldValue.replace(/[^0-9.]/g, '');
      if (!isNaN(Number.parseFloat(valueRemoveNonNumbers))) {
        spDocFieldValue = valueRemoveNonNumbers;
      } else {
        // TODO invalid field -- should it truncate???
        spDocFieldValue = '';
      }
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
          <p className={styles.dialogTitle}>Laserfiche</p>
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
        {!(missingFields?.length > 0) && <LoadingDialog />}
        {missingFields?.length > 0 && (
          <MissingFieldsDialog missingFields={listFields} />
        )}
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

function MissingFieldsDialog(props: { missingFields: JSX.Element[] }): JSX.Element {
  const textInside = (
    <span>
      The following SharePoint field values are blank and are mapped to required
      Laserfiche fields:
      {props.missingFields}Please fill out these required fields and try again.
    </span>
  );

  return (
    <div>
      <p>{textInside}</p>
    </div>
  );
}
