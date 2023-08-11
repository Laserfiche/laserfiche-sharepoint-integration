import * as React from 'react';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IEditManageConfigurationProps } from './IEditManageConfigurationProps';
import { IListItem } from '../IListItem';

import { useEffect, useState } from 'react';
import {
  ProfileHeader,
  validateNewConfiguration,
} from '../ProfileConfigurationComponents';
import ManageConfiguration from '../ManageConfigurationComponent';
import { ProfileConfiguration } from '../ProfileConfigurationComponents';
import {
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
} from '../../../constants';
import { getSPListURL } from '../../../../Utils/Funcs';
import { ProblemDetails } from '@laserfiche/lf-repository-api-client';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
      // eslint-disable-next-line
      ['lf-repository-browser']: any;
    }
  }
}

export default function EditManageConfiguration(
  props: IEditManageConfigurationProps
): JSX.Element {
  const [profileConfig, setProfileConfig] = useState<
    ProfileConfiguration | undefined
  >(undefined);

  const [validate, setValidate] = useState(false);
  const handleProfileConfigUpdate: (
    profileConfig: ProfileConfiguration
  ) => void = (profileConfig: ProfileConfiguration) => {
    setValidate(false);
    setProfileConfig(profileConfig);
  };

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

  useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      const results = await GetItemIdByTitle();
      const configurationName = props.match.params.name;
      if (results?.length > 0) {
        const profileConfigs = JSON.parse(results[0].JsonValue);
        if (profileConfigs.length > 0) {
          for (let i = 0; i < profileConfigs.length; i++) {
            if (profileConfigs[i].ConfigurationName === configurationName) {
              const selectedConfig: ProfileConfiguration = profileConfigs[i];
              setProfileConfig(selectedConfig);
            }
          }
        }
      }
    };
    initializeComponentAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, []);

  async function saveEditExistingConfigurationAsync(): Promise<boolean> {
    setValidate(true);
    const validate = validateNewConfiguration(profileConfig);
    if (validate) {
      const manageConfigurationConfig: IListItem[] = await GetItemIdByTitle();
      if (manageConfigurationConfig?.length > 0) {
        const configWithCurrentName = manageConfigurationConfig[0];
        const savedProfileConfigurations: ProfileConfiguration[] = JSON.parse(
          configWithCurrentName.JsonValue
        );
        const profileIndex = savedProfileConfigurations.findIndex(
          (config) =>
            config.ConfigurationName === profileConfig.ConfigurationName
        );
        if (profileIndex !== -1) {
          savedProfileConfigurations[profileIndex] = profileConfig;
          const configsToSave = savedProfileConfigurations;
          const succeeded = await saveSPConfigurationsAsync(
            configWithCurrentName.Id,
            configsToSave
          );
          return succeeded;
        } else {
          // error this config should exist
        }
      }
    }
    return false;
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

  const header = (
    <div>
      <ProfileHeader configurationName={profileConfig?.ConfigurationName} />
    </div>
  );
  return profileConfig ? (
    <ManageConfiguration
      header={header}
      repoClient={props.repoClient}
      loggedIn={props.loggedIn}
      profileConfig={profileConfig}
      loadingContent={true}
      createNew={false}
      context={props.context}
      handleProfileConfigUpdate={handleProfileConfigUpdate}
      saveConfiguration={saveEditExistingConfigurationAsync}
      validate={validate}
    />
  ) : (
    <span>Nothing to see</span>
  );
}
