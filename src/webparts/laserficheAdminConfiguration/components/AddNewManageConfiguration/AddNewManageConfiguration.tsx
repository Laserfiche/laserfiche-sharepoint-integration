// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { IAddNewManageConfigurationProps } from './IAddNewManageConfigurationProps';
import ManageConfiguration from '../ManageConfigurationComponent';
import { useState } from 'react';
import {
  ActionTypes,
  LfFolder,
  ProfileConfiguration,
  validateNewConfiguration,
} from '../ProfileConfigurationComponents';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListItem } from '../IListItem';
import {
  LASERFICHE_ADMIN_CONFIGURATION_NAME,
  MANAGE_CONFIGURATIONS,
} from '../../../constants';
import { getSPListURL } from '../../../../Utils/Funcs';
import styles from './../LaserficheAdminConfiguration.module.scss';
require('../../../../Assets/CSS/bootstrap.min.css');
require('./../../../../Assets/CSS/commonStyles.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-repository-browser']: any;
    }
  }
}

const rootFolder: LfFolder = {
  id: '1',
  path: '\\',
};

const initialConfig: ProfileConfiguration = {
  selectedFolder: rootFolder,
  DocumentName: 'FileName',
  ConfigurationName: '',
  mappedFields: [],
  Action: ActionTypes.COPY,
};

export default function AddNewManageConfiguration(
  props: IAddNewManageConfigurationProps
): JSX.Element {
  const [profileConfig, setProfileConfig] = useState(initialConfig);
  const [validate, setValidate] = useState(false);
  const [configNameError, setConfigNameError] = useState(undefined);
  const handleProfileConfigUpdate: (
    profileConfig: ProfileConfiguration
  ) => void = (profileConfig: ProfileConfiguration) => {
    setValidate(false);
    setProfileConfig(profileConfig);
  };
  function handleProfileConfigNameChange(e: React.ChangeEvent): void {
    const newName = (e.target as HTMLInputElement).value;
    const profileConfiguration = { ...profileConfig };
    profileConfiguration.ConfigurationName = newName;
    setValidate(false);
    setProfileConfig(profileConfiguration);
  }

  async function saveSPConfigurationsAsync(
    Id: string,
    configsToSave: ProfileConfiguration[]
  ): Promise<void> {
    const restApiUrl = `${getSPListURL(
      props.context,
      LASERFICHE_ADMIN_CONFIGURATION_NAME
    )}/items(${Id})`;
    const body: string = JSON.stringify({
      Title: MANAGE_CONFIGURATIONS,
      JsonValue: JSON.stringify(configsToSave),
    });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: body,
    };
    const response = await props.context.spHttpClient.post(
      restApiUrl,
      SPHttpClient.configurations.v1,
      options
    );
    if (!response.ok) {
      throw Error(response.statusText);
    }
  }

  async function saveNewManageConfigurationAsync(): Promise<void> {
    setValidate(true);
    setConfigNameError(undefined);
    const validate = validateNewConfiguration(profileConfig);
    if (validate) {
      const manageConfigurationConfig: IListItem[] = await GetItemIdForManageConfigurations();
      if (manageConfigurationConfig?.length > 0) {
        const configWithCurrentName = manageConfigurationConfig[0];
        const savedProfileConfigurations: ProfileConfiguration[] =
          JSON.parse(configWithCurrentName.JsonValue) ?? [];
        const profileExists = savedProfileConfigurations.find(
          (config) =>
            config.ConfigurationName === profileConfig.ConfigurationName
        );
        if (!profileExists) {
          const allConfigurations =
            savedProfileConfigurations.concat(profileConfig);
          await saveSPConfigurationsAsync(
            configWithCurrentName.Id,
            allConfigurations
          );
        } else {
          setConfigNameError(
            <span>
              Profile with this name already exists, please provide different
              name
            </span>
          );
        }
      } else {
        await saveNewPageConfigurationAsync();
      }
    } else {
      throw Error('Invalid configuration. Please review any errors.');
    }
  }

  async function GetItemIdForManageConfigurations(): Promise<IListItem[]> {
    const restApiUrl = `${getSPListURL(
      props.context,
      LASERFICHE_ADMIN_CONFIGURATION_NAME
    )}/Items?$select=Id,Title,JsonValue&$filter=Title eq '${MANAGE_CONFIGURATIONS}'`;
    const res = await fetch(restApiUrl, {
      method: 'GET',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    });
    const results = await res.json();
    if (results.value.length > 0) {
      return results.value as IListItem[];
    } else {
      return null;
    }
  }

  async function saveNewPageConfigurationAsync(): Promise<void> {
    const profileConfigAsString = JSON.stringify([profileConfig]);
    const restApiUrl = `${getSPListURL(
      props.context,
      LASERFICHE_ADMIN_CONFIGURATION_NAME
    )}/items`;
    const body: string = JSON.stringify({
      Title: MANAGE_CONFIGURATIONS,
      JsonValue: profileConfigAsString,
    });
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'content-type': 'application/json;odata=nometadata',
        'odata-version': '',
      },
      body: body,
    };
    const response = await props.context.spHttpClient.post(
      restApiUrl,
      SPHttpClient.configurations.v1,
      options
    );
    if (!response.ok) {
      throw Error(response.statusText);
    }
  }

  let configNameValidation: JSX.Element | undefined;
  if (validate) {
    if (configNameError) {
      configNameValidation = configNameError;
    } else if (
      !profileConfig.ConfigurationName ||
      profileConfig.ConfigurationName.length === 0
    ) {
      configNameValidation = (
        <span>Please specify a name for this configuration</span>
      );
    } else if (/[^ A-Za-z0-9]/.test(profileConfig.ConfigurationName)) {
      // TODO can we allow special characters
      configNameValidation = (
        <span>Invalid Name, only alphanumeric or space are allowed.</span>
      );
    }
  }

  const header = (
    <div>
      <h6 className='mb-0'>Add New Profile</h6>
    </div>
  );
  const extraConfiguration = (
    <>
      <div className={`${styles.formGroupRow} form-group row`}>
        <label htmlFor='txt0' className='col-sm-3 col-form-label'>
          Profile Name <span style={{ color: 'red' }}>*</span>
        </label>
        <div className='col-sm-6'>
          <input
            type='text'
            className='form-control'
            id='configurationName'
            onChange={handleProfileConfigNameChange}
            placeholder='Profile Name'
          />
          <div
            id='configurationExists'
            hidden={!configNameValidation}
            style={{ color: 'red' }}
          >
            {configNameValidation}
          </div>
        </div>
      </div>
    </>
  );
  return (
    <ManageConfiguration
      header={header}
      extraConfiguration={extraConfiguration}
      repoClient={props.repoClient}
      loggedIn={props.loggedIn}
      profileConfig={profileConfig}
      loadingContent={true}
      createNew={true}
      context={props.context}
      handleProfileConfigUpdate={handleProfileConfigUpdate}
      saveConfiguration={saveNewManageConfigurationAsync}
      validate={validate}
    />
  );
}
