import * as React from 'react';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import {
  DeleteModal,
  ProfileConfiguration,
} from '../ProfileConfigurationComponents';
import { ChangeEvent, useState } from 'react';
import { IManageMappingsPageProps } from './IManageMappingsPageProps';
import { IListItem } from '../IListItem';
import {
  ADMIN_CONFIGURATION_LIST,
  MANAGE_CONFIGURATIONS,
  MANAGE_MAPPING,
} from '../../../constants';
import { getSPListURL } from '../../../../Utils/Funcs';
import { ProfileMappingConfiguration } from '../../../../Utils/Types';
import { ProblemDetails } from '@laserfiche/lf-repository-api-client';
import styles from './../LaserficheAdminConfiguration.module.scss';
require('../../../../Assets/CSS/bootstrap.min.css');
require('../../adminConfig.css');
require('../../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');

interface SPContentType {
  ID: string;
  Name: string;
  Description: string;
}

const sharepointValidationMapping =
  'Please select a content type from the SharePoint Content Type drop down';
const laserficheValidationMapping =
  'Please select a content type from the Laserfiche Profile dropdown';
const validationOf = 'Already Mapping exists for this SharePoint content type';

export default function ManageMappingsPage(
  props: IManageMappingsPageProps
): JSX.Element {
  const [mappingRows, setMappingRows] = useState([]);
  const [sharePointContentTypes, setSharePointContentTypes] = useState<
    string[]
  >([]);
  const [laserficheContentTypes, setLaserficheContentTypes] = useState<
    string[]
  >([]);
  const [deleteModal, setDeleteModal] = useState(undefined);
  const [validationMessage, setValidationMessage] = useState(undefined);

  React.useEffect(() => {
    getAllMappingsAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, [props.repoClient]);

  async function getAllMappingsAsync(): Promise<void> {
    await getAllSharePointContentTypesAsync();
    await getAllLaserficheContentTypesAsync();
    const results: { id: string; mappings: ProfileMappingConfiguration[] } =
      await getManageMappingsAsync();
    if (results?.mappings.length > 0) {
      setMappingRows(mappingRows.concat(results.mappings));
    }
  }

  async function getAllLaserficheContentTypesAsync(): Promise<void> {
    const array: string[] = [];
    const results: { id: string; configs: ProfileConfiguration[] } =
      await getManageConfigurationsAsync();
    if (results?.configs.length > 0) {
      const configs = results.configs;
      for (let i = 0; i < configs.length; i++) {
        array.push(configs[i].ConfigurationName);
      }
      setLaserficheContentTypes(array);
    }
  }

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

  async function getAllSharePointContentTypesAsync(): Promise<void> {
    const restApiUrl =
      props.context.pageContext.web.absoluteUrl + '/_api/web/contenttypes';
    try {
      const res = await fetch(restApiUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
      const results = await res.json();
      const array: string[] = results.value.map(
        (contentType: SPContentType) => contentType.Name
      );
      array.sort((a, b) => (a > b ? 1 : -1));
      setSharePointContentTypes(array);
    } catch (error) {
      console.log('error occurred' + error);
    }
  }

  async function createNewMappingAsync(
    idx: number,
    rows: ProfileMappingConfiguration[]
  ): Promise<void> {
    setValidationMessage(undefined);
    if (rows[idx].SharePointContentType === 'Select') {
      setValidationMessage(sharepointValidationMapping);
    } else if (rows[idx].LaserficheContentType === 'Select') {
      setValidationMessage(laserficheValidationMapping);
    } else {
      const existingMappings: {
        id: string;
        mappings: ProfileMappingConfiguration[];
      } = await getManageMappingsAsync();
      if (existingMappings) {
        if (existingMappings?.mappings.length > 0) {
          const mappingExists = existingMappings.mappings.find(
            (mapping) => mapping.id === rows[idx].id
          );
          if (mappingExists !== undefined) {
            await updateExistingMappingAsync(
              existingMappings.mappings,
              rows,
              idx,
              existingMappings.id
            );
          } else {
            await addNewInExistingMappingAsync(
              existingMappings.mappings,
              rows,
              idx,
              existingMappings.id
            );
          }
        } else {
          const restApiUrl = `${getSPListURL(
            props.context,
            ADMIN_CONFIGURATION_LIST
          )}/items(${existingMappings.id})`;
          const row = [...mappingRows];
          const newjsonValue = [
            {
              id: row[idx].id,
              SharePointContentType: row[idx].SharePointContentType,
              LaserficheContentType: row[idx].LaserficheContentType,
              toggle: true,
            },
          ];
          const jsonObject = JSON.stringify(newjsonValue);
          const body: string = JSON.stringify({
            Title: MANAGE_MAPPING,
            JsonValue: jsonObject,
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
          await props.context.spHttpClient.post(
            restApiUrl,
            SPHttpClient.configurations.v1,
            options
          );
          rows[idx].toggle = !rows[idx].toggle;
          setMappingRows(rows);
        }
      } else {
        const restApiUrl = `${getSPListURL(
          props.context,
          ADMIN_CONFIGURATION_LIST
        )}/items`;
        const newRow = [...mappingRows];
        const jsonValues = [
          {
            id: newRow[idx].id,
            SharePointContentType: newRow[idx].SharePointContentType,
            LaserficheContentType: newRow[idx].LaserficheContentType,
            toggle: true,
          },
        ];
        const jsonObject = JSON.stringify(jsonValues);
        const body: string = JSON.stringify({
          Title: MANAGE_MAPPING,
          JsonValue: jsonObject,
        });
        const options: ISPHttpClientOptions = {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata',
            'odata-version': '',
          },
          body: body,
        };
        await props.context.spHttpClient.post(
          restApiUrl,
          SPHttpClient.configurations.v1,
          options
        );
        rows[idx].toggle = !rows[idx].toggle;
        setMappingRows(rows);
      }
    }
  }

  async function addNewInExistingMappingAsync(
    jsonValue: ProfileMappingConfiguration[],
    rows: ProfileMappingConfiguration[],
    idx: number,
    itemId: string
  ): Promise<void> {
    let exitEntry = false;
    for (let i = 0; i < jsonValue.length; i++) {
      if (
        jsonValue[i].SharePointContentType === rows[idx].SharePointContentType
      ) {
        exitEntry = true;
        break;
      }
    }
    if (!exitEntry) {
      const restApiUrl = `${getSPListURL(
        props.context,
        ADMIN_CONFIGURATION_LIST
      )}/items(${itemId})`;
      const newJsonValue = [
        ...jsonValue,
        {
          id: rows[idx].id,
          SharePointContentType: rows[idx].SharePointContentType,
          LaserficheContentType: rows[idx].LaserficheContentType,
          toggle: true,
        },
      ];
      const jsonObject = JSON.stringify(newJsonValue);
      const body: string = JSON.stringify({
        Title: MANAGE_MAPPING,
        JsonValue: jsonObject,
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
      await props.context.spHttpClient.post(
        restApiUrl,
        SPHttpClient.configurations.v1,
        options
      );
      rows[idx].toggle = !rows[idx].toggle;
      setMappingRows(rows);
      if (jsonValue.length + 1 === rows.length) {
        setValidationMessage(undefined);
      }
    } else {
      setValidationMessage(validationOf);
    }
  }

  async function updateExistingMappingAsync(
    jsonValue: ProfileMappingConfiguration[],
    rows: ProfileMappingConfiguration[],
    idx: number,
    itemId: string
  ): Promise<void> {
    const spContentTypeMatch = jsonValue.find(
      (mapping) =>
        mapping.SharePointContentType === rows[idx].SharePointContentType
    );
    if (!spContentTypeMatch || spContentTypeMatch.id === rows[idx].id) {
      const matchingId = jsonValue.findIndex(
        (mapping) => mapping.id === rows[idx].id
      );
      jsonValue[matchingId] = { ...rows[idx] };
      rows[idx].toggle = !rows[idx].toggle;
      jsonValue[matchingId].toggle = rows[idx].toggle;
      const restApiUrl = `${getSPListURL(
        props.context,
        ADMIN_CONFIGURATION_LIST
      )}/items(${itemId})`;
      const newJsonValue = [...jsonValue];
      const jsonObject = JSON.stringify(newJsonValue);
      const body: string = JSON.stringify({
        Title: MANAGE_MAPPING,
        JsonValue: jsonObject,
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
      await props.context.spHttpClient.post(
        restApiUrl,
        SPHttpClient.configurations.v1,
        options
      );
      setMappingRows(rows);
      if (
        rows.some(
          (item: ProfileMappingConfiguration) =>
            item.SharePointContentType === 'Select'
        )
      ) {
        setValidationMessage(sharepointValidationMapping);
      } else if (
        rows.some(
          (item: ProfileMappingConfiguration) =>
            item.LaserficheContentType === 'Select'
        )
      ) {
        setValidationMessage(laserficheValidationMapping);
      } else {
        setValidationMessage(undefined);
      }
    } else {
      setValidationMessage(validationOf);
    }
  }

  async function deleteMappingAsync(
    rows: ProfileMappingConfiguration[],
    idx: number
  ): Promise<void> {
    const results: { id: string; mappings: ProfileMappingConfiguration[] } =
      await getManageMappingsAsync();
    if (results) {
      const itemId = results.id;
      const mappings = results.mappings;
      const matchingMappingIndex = mappings.findIndex(
        (mapping) => mapping.id === rows[idx].id
      );
      if (matchingMappingIndex > -1) {
        mappings.splice(matchingMappingIndex, 1);
        const restApiUrl = `${getSPListURL(
          props.context,
          ADMIN_CONFIGURATION_LIST
        )}/items(${itemId})`;
        const newMappings = [...mappings];
        const jsonObject = JSON.stringify(newMappings);
        const body: string = JSON.stringify({
          Title: MANAGE_MAPPING,
          JsonValue: jsonObject,
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
        await props.context.spHttpClient.post(
          restApiUrl,
          SPHttpClient.configurations.v1,
          options
        );
        const existingSPContentType = rows.find(
          (mapping) =>
            mapping.SharePointContentType === rows[idx].SharePointContentType
        );
        if (!existingSPContentType) {
          setValidationMessage(undefined);
        } else {
          setValidationMessage(validationOf);
        }
      } else {
        if (mappings.length + 1 === rows.length) {
          setValidationMessage(undefined);
        } else {
          const selectSPContentType = mappings.find(
            (mapping) => mapping.SharePointContentType === 'Select'
          );
          if (selectSPContentType) {
            setValidationMessage(sharepointValidationMapping);
          } else {
            const selectLfContentType = mappings.find(
              (mapping) => mapping.LaserficheContentType === 'Select'
            );
            if (selectLfContentType) {
              setValidationMessage(laserficheValidationMapping);
            }
          }
        }
      }
    }
  }

  async function getManageMappingsAsync(): Promise<{
    id: string;
    mappings: ProfileMappingConfiguration[];
  }> {
    const array: IListItem[] = [];
    const restApiUrl = `${getSPListURL(
      props.context,
      ADMIN_CONFIGURATION_LIST
    )}/Items?$select=Id,Title,JsonValue&$filter=Title eq '${MANAGE_MAPPING}'`;
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
        return { id: array[0].Id, mappings: JSON.parse(array[0].JsonValue) };
      } else {
        return undefined;
      }
    } catch (error) {
      console.log('error occurred' + error);
    }
  }

  const addNewMapping: () => void = () => {
    const id = (+new Date() + Math.floor(Math.random() * 999999)).toString(36);
    const item = {
      id,
      SharePointContentType: 'Select',
      LaserficheContentType: 'Select',
      toggle: false,
    };
    setMappingRows([...mappingRows, item]);
  };

  const removeSpecificMapping: (idx: number) => void = (idx: number) => {
    const rows = [...mappingRows];
    const delModal = (
      <DeleteModal
        onCancel={closeModalUp}
        onConfirmDelete={() => removeRowAsync(idx)}
        configurationName={rows[idx].SharePointContentType}
      />
    );
    setDeleteModal(delModal);
  };

  async function removeRowAsync(id: number): Promise<void> {
    const rows = [...mappingRows];
    const deleteRows = [...mappingRows];
    rows.splice(id, 1);
    setMappingRows(rows);
    await deleteMappingAsync(deleteRows, id);
    setDeleteModal(undefined);
  }

  const editSpecificMapping: (idx: number) => void = (idx: number) => {
    const rows = [...mappingRows];
    rows[idx].toggle = !rows[idx].toggle;
    setMappingRows(rows);
  };

  const saveSpecificMappingAsync: (idx: number) => Promise<void> = async (
    idx: number
  ) => {
    const rows = [...mappingRows];
    await createNewMappingAsync(idx, rows);
  };

  const handleChange: (
    event: ChangeEvent<HTMLSelectElement>,
    idx: number
  ) => void = (event: ChangeEvent<HTMLSelectElement>, idx: number) => {
    const item = {
      id: event.target.id,
      name: event.target.name,
      value: event.target.value,
    };
    const newRows = [...mappingRows];
    if (item.name === 'SharePointContentType') {
      newRows[idx].SharePointContentType = item.value;
    } else if (item.name === 'LaserficheContentType') {
      newRows[idx].LaserficheContentType = item.value;
    }
    setMappingRows(newRows);
  };

  function closeModalUp(): void {
    setDeleteModal(undefined);
  }

  const resetAsync: () => Promise<void> = async () => {
    setDeleteModal(undefined);
    await getAllSharePointContentTypesAsync();
    await getAllLaserficheContentTypesAsync();
    const results: { id: string; mappings: ProfileMappingConfiguration[] } =
      await getManageMappingsAsync();
    if (results?.mappings.length > 0) {
      setMappingRows(results.mappings);
    }
    setValidationMessage(undefined);
  };

  const sharePointContentTypesDisplay = sharePointContentTypes.map(
    (contentType) => (
      <option key={contentType} value={contentType}>
        {contentType}
      </option>
    )
  );
  const lfContentTypesDisplay = laserficheContentTypes.map((contentType) => (
    <option key={contentType} value={contentType}>
      {contentType}
    </option>
  ));
  const renderTableData = mappingRows.map((item, index) => {
    if (item.toggle) {
      return (
        <tr id='addr0' key={index}>
          <td>
            <select
              name='SharePointContentType'
              disabled
              className='custom-select'
              value={mappingRows[index].SharePointContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {sharePointContentTypesDisplay}
            </select>
          </td>
          <td>
            <select
              name='LaserficheContentType'
              disabled
              className='custom-select'
              value={mappingRows[index].LaserficheContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {lfContentTypesDisplay}
            </select>
          </td>
          <td className={styles.iconTableCell}>
            <div className={styles.iconsContainer}>
              <button
                className={styles.lfMaterialIconButton}
                onClick={() => editSpecificMapping(index)}
              >
                <span className='material-icons-outlined'>edit</span>
              </button>
              <button
                className={`${styles.lfMaterialIconButton} ${styles.marginLeftButton}`}
                onClick={() => removeSpecificMapping(index)}
              >
                <span className='material-icons-outlined'>delete</span>
              </button>
            </div>
          </td>
        </tr>
      );
    } else {
      return (
        <tr id='addr0' key={index}>
          <td>
            <select
              name='SharePointContentType'
              className='custom-select'
              value={mappingRows[index].SharePointContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {sharePointContentTypesDisplay}
            </select>
          </td>
          <td>
            <select
              name='LaserficheContentType'
              className='custom-select'
              value={mappingRows[index].LaserficheContentType}
              id={mappingRows[index].id}
              onChange={(e) => handleChange(e, index)}
            >
              <option>Select</option>
              {lfContentTypesDisplay}
            </select>
          </td>
          <td className={styles.iconTableCell}>
            <div className={styles.iconsContainer}>
              <button
                className={styles.lfMaterialIconButton}
                onClick={() => saveSpecificMappingAsync(index)}
              >
                <span className='material-icons-outlined'>save</span>
              </button>
              <button
                className={`${styles.lfMaterialIconButton} ${styles.marginLeftButton}`}
                onClick={() => removeSpecificMapping(index)}
              >
                <span className='material-icons-outlined'>delete</span>
              </button>
            </div>
          </td>
        </tr>
      );
    }
  });
  const viewSharePointContentTypes =
    props.context.pageContext.web.absoluteUrl + '/_layouts/15/mngctype.aspx';

  return (
    <>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '85%', marginLeft: '-26px' }}
      >
        <div className='p-3'>
          <div className='card rounded-0'>
            <div className='card-header d-flex justify-content-between'>
              <div>
                <h6 className='mb-0'>Content Type Mappings Laserfiche</h6>
              </div>
              <div>
                <a
                  href=''
                  onClick={() => window.open(viewSharePointContentTypes)}
                  target='_blank'
                >
                  View SharePoint Content Types
                </a>
              </div>
            </div>
            <div className='card-body'>
              <table className='table table-sm'>
                <thead>
                  <tr>
                    <th className='text-center' style={{ width: '45%' }}>
                      SharePoint Content Type
                    </th>
                    <th className='text-center' style={{ width: '45%' }}>
                      Laserfiche Profile
                    </th>
                    <th className='text-center'>Action</th>
                  </tr>
                </thead>
                <tbody>{renderTableData}</tbody>
              </table>
            </div>

            {validationMessage && (
              <div id='sharePointValidationMapping' style={{ color: 'red' }}>
                <span>{validationMessage}</span>
              </div>
            )}
            <div className={`${styles.footerIcons} card-footer bg-transparent`}>
              <button className='lf-button sec-button' onClick={resetAsync}>
                Reset
              </button>
              <button
                className={`${styles.marginLeftButton} lf-button primary-button`}
                onClick={addNewMapping}
              >
                Add
              </button>
            </div>
          </div>
        </div>
      </div>
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
