// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import {
  ODataValueContextOfIListOfWTemplateInfo,
  ODataValueOfIListOfTemplateFieldInfo,
  TemplateFieldInfo,
  WTemplateInfo,
} from '@laserfiche/lf-repository-api-client';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import * as React from 'react';
import { useState } from 'react';
import { NavLink } from 'react-router-dom';
import { IManageConfigurationProps } from './ManageConfigurationProps';
import {
  ConfigurationBody,
  SharePointLaserficheColumnMatching,
  SPProfileConfigurationData,
} from './ProfileConfigurationComponents';
import styles from './LaserficheAdminConfiguration.module.scss';

export default function ManageConfiguration(
  props: IManageConfigurationProps
): JSX.Element {
  const [availableLfTemplates, setAvailableLfTemplates] = useState<
    WTemplateInfo[] | undefined
  >([]);
  const [lfFieldsForSelectedTemplate, setLfFieldsForSelectedTemplate] =
    useState<TemplateFieldInfo[] | undefined>(undefined);
  const [availableSPFields, setAvailableSPFields] = useState<
    SPProfileConfigurationData[] | undefined
  >(undefined);
  const [showConfirmModal, setShowConfirmModal] = useState<boolean>(false);
  const [showErrorModal, setShowErrorModal] = useState<string | undefined>();
  const [saveDisabled, setSaveDisabled] = useState<boolean>(false);

  async function getAllAvailableTemplates(): Promise<WTemplateInfo[]> {
    const repoId = await props.repoClient.getCurrentRepoId();
    const templateInfo: WTemplateInfo[] = [];
    await props.repoClient.templateDefinitionsClient.getTemplateDefinitionsForEach(
      {
        callback: async (response: ODataValueContextOfIListOfWTemplateInfo) => {
          if (response.value) {
            templateInfo.push(...response.value);
          }
          return true;
        },
        repoId,
      }
    );
    return templateInfo;
  }

  const getLaserficheFieldsAsync: (
    templateName: string
  ) => Promise<TemplateFieldInfo[]> = async (templateName: string) => {
    if (templateName?.length > 0) {
      const repoId = await props.repoClient.getCurrentRepoId();
      const apiTemplateResponse: ODataValueOfIListOfTemplateFieldInfo =
        await props.repoClient.templateDefinitionsClient.getTemplateFieldDefinitionsByTemplateName(
          { repoId, templateName: templateName }
        );
      const fieldsValues: TemplateFieldInfo[] = apiTemplateResponse.value;
      return fieldsValues;
    } else {
      return null;
    }
  };

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      try {
        const templates: WTemplateInfo[] = await getAllAvailableTemplates();
        templates.sort();
        setAvailableLfTemplates(templates);
        if (props.profileConfig.selectedTemplateName) {
          const templateFields: TemplateFieldInfo[] =
            await getLaserficheFieldsAsync(
              props.profileConfig.selectedTemplateName
            );
          setLfFieldsForSelectedTemplate(templateFields);
        }
        const spColumns: SPProfileConfigurationData[] =
          await getAllSharePointSiteColumnsAsync();
        spColumns.sort((a, b) => (a.Title > b.Title ? 1 : -1));
        setAvailableSPFields(spColumns);
      } catch (err) {
        console.error(
          `Error initializing configuration component: ${err}`
        );
      }
    };
    if (props.repoClient) {
      void initializeComponentAsync();
    }
  }, [props.repoClient]);

  async function getAllSharePointSiteColumnsAsync(): Promise<
    SPProfileConfigurationData[]
  > {
    const restApiUrl: string =
      props.context.pageContext.web.absoluteUrl +
      "/_api/web/fields?$filter=(Hidden ne true and Group ne '_Hidden')";
    const res = await fetch(restApiUrl, {
      method: 'GET',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
    });
    const results = (await res.json()).value as SPProfileConfigurationData[];

    return results;
  }

  const onChangeTemplateAsync: (templateName: string) => Promise<void> = async (
    templateName: string
  ) => {
    const templateFields = await getLaserficheFieldsAsync(templateName);
    if (templateFields) {
      const array = [];
      for (let index = 0; index < templateFields.length; index++) {
        const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(
          36
        );
        const laserficheField = templateFields[index];
        if (laserficheField.isRequired) {
          array.push({
            id: id,
            spField: undefined,
            lfField: templateFields[index],
          });
        }
      }
      const profileConfig = { ...props.profileConfig };
      profileConfig.mappedFields = array;
      profileConfig.selectedTemplateName = templateName;
      setLfFieldsForSelectedTemplate(templateFields);
      props.handleProfileConfigUpdate(profileConfig);
    } else {
      const profileConfig = { ...props.profileConfig };
      profileConfig.selectedTemplateName = templateName;
      profileConfig.mappedFields = undefined;
      setLfFieldsForSelectedTemplate([]);
      props.handleProfileConfigUpdate(profileConfig);
    }
  };

  function onClickConfirmButton(): void {
    history.back();
    setShowConfirmModal(false);
  }

  function onClickErrorButton(): void {
    setShowErrorModal(undefined);
  }

  async function saveConfigurationAsync(): Promise<void> {
    try {
      await props.saveConfiguration();
      setShowConfirmModal(true);
    } catch (err) {
      setShowErrorModal(err.message);
    }
  }

  function hasError(hasError: boolean): void {
    if (hasError) {
      setSaveDisabled(true);
    } else {
      setSaveDisabled(false);
    }
  }

  return (
    <div>
      <div className='p-3'>
        <main className='bg-white shadow-sm'>
          <div className='addPageSpinloader' hidden={props.loadingContent}>
            {!props.loadingContent && (
              <Spinner size={SpinnerSize.large} label='loading' />
            )}
            ,
          </div>
          <div hidden={!props.loadingContent}>
            <div className='card rounded-0'>
              <div className='card-header d-flex justify-content-between'>
                {props.header}
              </div>
              <div className='card-body'>
                {props.extraConfiguration}
                <ConfigurationBody
                  availableLfTemplates={availableLfTemplates}
                  repoClient={props.repoClient}
                  loggedIn={props.loggedIn}
                  handleTemplateChange={onChangeTemplateAsync}
                  profileConfig={props.profileConfig}
                  handleProfileConfigUpdate={props.handleProfileConfigUpdate}
                />
              </div>
              <h6 className='card-header border-top'>
                Mappings from SharePoint Column to Laserfiche Field Values
              </h6>
              <div className='card-body'>
                <SharePointLaserficheColumnMatching
                  profileConfig={props.profileConfig}
                  availableSPFields={availableSPFields}
                  lfFieldsForSelectedTemplate={lfFieldsForSelectedTemplate}
                  handleProfileConfigUpdate={props.handleProfileConfigUpdate}
                  validate={props.validate}
                  hasError={hasError}
                />
              </div>
              <div
                className={`${styles.footerIcons} card-footer bg-transparent`}
              >
                {props.loggedIn && (
                  <NavLink
                    id='navid'
                    to='/ManageConfigurationsPage'
                    className={styles.navLink}
                  >
                    <button className='lf-button sec-button'>Back</button>
                  </NavLink>
                )}
                <button
                  className={`${styles.marginLeftButton} lf-button primary-button`}
                  onClick={saveConfigurationAsync}
                  disabled={saveDisabled}
                >
                  Save
                </button>
              </div>
            </div>
          </div>
        </main>
      </div>
      {showConfirmModal && <div
        className={styles.modal}
        data-backdrop='static'
        data-keyboard='false'
        id='ConfirmModal'
      >
        <div className='modal-dialog modal-dialog-centered'>
          <div
            className={`modal-content ${styles.modalContent} ${styles.wrapper}`}
          >
            <div className='modal-body'>
              {props.createNew ? 'Profile Added' : 'Profile Updated'}
            </div>
            <div className='modal-footer'>
              <button
                type='button'
                className='lf-button primary-button'
                data-dismiss='modal'
                onClick={onClickConfirmButton}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      </div>}
      {showErrorModal && <div
        className={styles.modal}
        data-backdrop='static'
        data-keyboard='false'
        id='ErrorModal'
      >
        <div className='modal-dialog modal-dialog-centered'>
          <div
            className={`modal-content ${styles.modalContent} ${styles.wrapper}`}
          >
            <div className={`modal-header ${styles.header}`}>
              Error {props.createNew ? 'Saving' : 'Updating'} Profile
            </div>
            <div className={`modal-body ${styles.contentBox}`}>
              {showErrorModal}
            </div>
            <div className={`modal-footer ${styles.footer}`}>
              <button
                type='button'
                className='lf-button primary-button'
                data-dismiss='modal'
                onClick={onClickErrorButton}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      </div>}
    </div>
  );
}
