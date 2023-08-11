import * as React from 'react';
import { useEffect, useState } from 'react';
import { NavLink } from 'react-router-dom';
import { IManageConfigurationPageProps } from './IManageConfigurationPageProps';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListItem } from '../IListItem';
import {
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
} from '../../../constants';
import { getSPListURL } from '../../../../Utils/Funcs';
import {
  DeleteModal,
  ProfileConfiguration,
} from '../ProfileConfigurationComponents';
import { ProblemDetails } from '@laserfiche/lf-repository-api-client';
import styles from './../LaserficheAdminConfiguration.module.scss';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

const ADD_PROFILE = 'Add Profile';
const PROFILE_NAME = 'Profile Name';
const ACTION = 'Action';

export default function ManageConfigurationsPage(
  props: IManageConfigurationPageProps
): JSX.Element {
  const [configRows, setConfigRows] = useState<ProfileConfiguration[]>([]);
  const [deleteModal, setDeleteModal] = useState<JSX.Element | undefined>(
    undefined
  );

  useEffect(() => {
    const updateConfigurationsAsync: () => Promise<void> = async () => {
      const configurations: { id: string; configs: ProfileConfiguration[] } =
        await getManageConfigurationsAsync();
      if (configurations?.configs.length > 0) {
        setConfigRows(configRows.concat(...configurations.configs));
      }
    };
    updateConfigurationsAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, []);

  async function getManageConfigurationsAsync(): Promise<{
    id: string;
    configs: ProfileConfiguration[];
  }> {
    const array: IListItem[] = [];
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
        for (let i = 0; i < results.value.length; i++) {
          array.push(results.value[i]);
        }
        return { id: array[0].Id, configs: JSON.parse(array[0].JsonValue) };
      } else {
        return null;
      }
    } catch (error) {
      console.log('error occurred' + error);
    }
  }

  function removeSpecificConfiguration(idx: number): void {
    const rows = [...configRows];
    const configName = rows[idx].ConfigurationName;
    const deleteModal = (
      <DeleteModal
        configurationName={configName}
        onConfirmDelete={() => removeRowAsync(idx)}
        onCancel={closeModal}
      />
    );
    setDeleteModal(deleteModal);
  }

  async function removeRowAsync(id: number): Promise<void> {
    const rows = [...configRows];
    const deleteRows = [...configRows];
    rows.splice(id, 1);
    await deleteMappingAsync(deleteRows, id);
    setDeleteModal(undefined);
  }

  function closeModal(): void {
    setDeleteModal(undefined);
  }

  async function deleteMappingAsync(
    rows: ProfileConfiguration[],
    idx: number
  ): Promise<void> {
    const manageConfigs: { id: string; configs: ProfileConfiguration[] } =
      await getManageConfigurationsAsync();
    if (manageConfigs.configs?.length > 0) {
      const indexOfProfileToRemove = manageConfigs.configs.findIndex(
        (config) => config.ConfigurationName === rows[idx].ConfigurationName
      );
      if (indexOfProfileToRemove !== -1) {
        manageConfigs.configs.splice(indexOfProfileToRemove, 1);
        setConfigRows(manageConfigs.configs);
        const restApiUrl = `${getSPListURL(
          props.context,
          ADMIN_CONFIGURATION_LIST
        )}/items(${manageConfigs.id})`;

        const updatedConfigurations = JSON.stringify(manageConfigs.configs);
        const body: string = JSON.stringify({
          Title: MANAGE_CONFIGURATIONS,
          JsonValue: updatedConfigurations,
        });
        const options: ISPHttpClientOptions = {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
          },
          body,
        };
        await props.context.spHttpClient.post(
          restApiUrl,
          SPHttpClient.configurations.v1,
          options
        );
      }
    }
  }

  const tableData = configRows.map((item, index) => {
    return (
      <tr id='addr0' key={index}>
        <td>{item.ConfigurationName}</td>
        <td className='text-center'>
          <div className={styles.iconsContainer}>
            <NavLink to={'/EditManageConfiguration/' + item.ConfigurationName} 
                    className={styles.navLinkNoUnderline}>
              <button className={styles.lfMaterialIconButton}>
                <span className='material-icons-outlined'>edit</span>
              </button>
            </NavLink>
            <button
              className={styles.lfMaterialIconButton}
              onClick={() => removeSpecificConfiguration(index)}
            >
              <span
                className={`${styles.marginLeftButton} material-icons-outlined`}
              >
                delete
              </span>
            </button>
          </div>
        </td>
      </tr>
    );
  });

  return (
    <div>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '85%', marginLeft: '-26px' }}
      >
        <main className='bg-white shadow-sm'>
          <div className='p-3'>
            <div className='card rounded-0'>
              <div className='card-header d-flex justify-content-between pt-1 pb-1'>
                <NavLink
                  to='/AddNewManageConfiguration'
                  style={{
                    marginRight: '18px',
                    fontWeight: '500',
                    fontSize: '15px',
                  }}
                >
                  <a className='btn btn-primary pl-5 pr-5'>{ADD_PROFILE}</a>
                </NavLink>
              </div>
              <div className='card-body'>
                <table className='table table-bordered table-striped table-hover'>
                  <thead>
                    <tr>
                      <th className='text-center'>{PROFILE_NAME}</th>
                      <th className='text-center' style={{ width: '30%' }}>
                        {ACTION}
                      </th>
                    </tr>
                  </thead>
                  <tbody>{tableData}</tbody>
                </table>
              </div>
            </div>
          </div>
        </main>
      </div>
      <div>
        <div
          className='modal'
          id='deleteModal'
          hidden={!deleteModal}
          data-backdrop='static'
          data-keyboard='false'
        >
          {deleteModal}
        </div>
      </div>
    </div>
  );
}
