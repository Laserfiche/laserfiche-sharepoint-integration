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
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
} from '../../../constants';
import { getSPListURL } from '../../../../Utils/Funcs';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');
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
  const handleProfileConfigUpdate: (profileConfig: ProfileConfiguration) => void = (profileConfig: ProfileConfiguration) => {
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
  ): Promise<boolean> {
    const restApiUrl = `${getSPListURL(
      props.context,
      ADMIN_CONFIGURATION_LIST
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
    if (response.ok) {
      return true;
    } else {
      return false;
    }
    // TODO should this really throw?
  }

  async function saveNewManageConfigurationAsync(): Promise<boolean> {
    setValidate(true);
    setConfigNameError(undefined);
    const validate = validateNewConfiguration(profileConfig);
    if (validate) {
      const manageConfigurationConfig: IListItem[] = await GetItemIdByTitle();
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
          const succeeeded = await saveSPConfigurationsAsync(
            configWithCurrentName.Id,
            allConfigurations
          );
          return succeeeded;
        } else {
          setConfigNameError(
            <span>
              Profile with this name already exists, please provide different
              name
            </span>
          );
        }
      } else {
        const suceeded = await saveNewPageConfigurationAsync();
        return suceeded;
      }
    }
    return false;
  }

  async function GetItemIdByTitle(): Promise<IListItem[]> {
    const restApiUrl = `${getSPListURL(
      props.context,
      ADMIN_CONFIGURATION_LIST
    )}/Items?$select=Id,Title,JsonValue&$filter=Title eq '${MANAGE_CONFIGURATIONS}'`;
    try {
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
    } catch (error) {
      console.log('error occured' + error);
    }
  }

  async function saveNewPageConfigurationAsync(): Promise<boolean> {
    const profileConfigAsString = JSON.stringify([profileConfig]);
    const restApiUrl = `${getSPListURL(
      props.context,
      ADMIN_CONFIGURATION_LIST
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
    if (response.ok) {
      return true;
    } else {
      return false;
    }
  }

  let configNameValidation: JSX.Element | undefined;
  if (validate) {
    if (configNameError) {
      configNameValidation = configNameError;
    } else if (!profileConfig.ConfigurationName || profileConfig.ConfigurationName.length === 0) {
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
      <div className='form-group row'>
        <label
          htmlFor='txt0'
          className='col-sm-2 col-form-label'
          style={{ width: '165px' }}
        >
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
