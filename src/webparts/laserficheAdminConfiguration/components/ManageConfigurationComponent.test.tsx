// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

jest.mock('@laserfiche/lf-repository-api-client', () => ({
  WTemplateInfo: jest.fn().mockImplementation(({ name, displayName }) => ({
    name,
    displayName,
  })),
  WFieldType: {
    Date: 'Date',
    List: 'List',
    Time: 'Time',
    Number: 'Number',
  },
}));

import { render, waitFor, screen } from '@testing-library/react';
import ManageConfigurationComponent from './ManageConfigurationComponent';
import { ProfileConfiguration } from './ProfileConfigurationComponents';
import * as React from 'react';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import {
  TemplateDefinitionsClient,
} from '@laserfiche/lf-repository-api-client';
import mockWebPartContext from '../../../__mocks__/@microsoft/sp-webpart-base';
import { ProfileConfigContext } from './LaserficheAdminConfiguration';
import { HashRouter } from 'react-router-dom';
import { TEMPLATE_NO_LONGER_VALID_METADATA_WILL_NOT_BE_SAVED } from '../../strings';

describe('ManageConfigurationComponent', () => {
  test('template warning appears if selected template name no longer exists in list of templates and save button is still enabled', async () => {
    // Arrange
    const mockProfileConfig: ProfileConfiguration = {
      selectedTemplateName: 'NonExistentTemplate',
      // other properties
    } as ProfileConfiguration;

    const repoClient = {
      getCurrentRepoId: jest.fn().mockResolvedValue('1234'),
      templateDefinitionsClient: {
        getTemplateDefinitionsForEach: jest
          .fn()
          .mockImplementation(({ callback }) =>
            callback({
              value: [
                { name: 'Template1', id: '1' },
                { name: 'Template2', id: '2' },
              ],
            })
          ),

        getTemplateFieldDefinitionsByTemplateName: jest.fn().mockResolvedValue({
          value: [
            { name: 'Field1', id: '1', required: true },
            { name: 'Field2', id: '2' },
          ],
        }),
      } as unknown as TemplateDefinitionsClient,
    } as unknown as IRepositoryApiClientExInternal;
    window.fetch = jest.fn().mockResolvedValue({
      json: jest.fn().mockResolvedValue({ value: [] }),
    });

    // Act
    render(
      <HashRouter>
        <ProfileConfigContext.Provider
          value={{ setSaveDisabled: jest.fn(), saveDisabled: false }}
        >
          <ManageConfigurationComponent
            repoClient={repoClient}
            profileConfig={mockProfileConfig}
            handleProfileConfigUpdate={jest.fn()}
            validate={true}
            loggedIn={true}
            loadingContent={false}
            createNew={false}
            context={mockWebPartContext}
            saveConfiguration={jest.fn()}
          />
        </ProfileConfigContext.Provider>
      </HashRouter>
    );

    // Assert
    await waitFor(() => {
      expect(
        screen.getByText(TEMPLATE_NO_LONGER_VALID_METADATA_WILL_NOT_BE_SAVED)
      ).toBeInTheDocument();
      const saveButtonElement = screen.queryByTestId('saveButton');
      expect(saveButtonElement).toBeEnabled();
    });
  });

  test('save button is enabled if selected template exists in template list', async () => {
    // Arrange
    const mockProfileConfig: ProfileConfiguration = {
      selectedTemplateName: 'Template1',
      // other properties
    } as ProfileConfiguration;

    const repoClient = {
      getCurrentRepoId: jest.fn().mockResolvedValue('1234'),
      templateDefinitionsClient: {
        getTemplateDefinitionsForEach: jest
          .fn()
          .mockImplementation(({ callback }) =>
            callback({
              value: [
                { name: 'Template1', id: '1' },
                { name: 'Template2', id: '2' },
              ],
            })
          ),

        getTemplateFieldDefinitionsByTemplateName: jest.fn().mockResolvedValue({
          value: [
            { name: 'Field1', id: '1' },
            { name: 'Field2', id: '2' },
          ],
        }),
      } as unknown as TemplateDefinitionsClient,
    } as unknown as IRepositoryApiClientExInternal;
    window.fetch = jest.fn().mockResolvedValue({
      json: jest.fn().mockResolvedValue({ value: [] }),
    });

    // Act
    render(
      <HashRouter>
        <ProfileConfigContext.Provider
          value={{ setSaveDisabled: jest.fn(), saveDisabled: false }}
        >
          <ManageConfigurationComponent
            repoClient={repoClient}
            profileConfig={mockProfileConfig}
            handleProfileConfigUpdate={jest.fn()}
            validate={true}
            loggedIn={true}
            loadingContent={false}
            createNew={false}
            context={mockWebPartContext}
            saveConfiguration={jest.fn()}
          />
        </ProfileConfigContext.Provider>
      </HashRouter>
    );

    // Assert
    await waitFor(() => {
      const templateNotValidElement = screen.queryByText(
        TEMPLATE_NO_LONGER_VALID_METADATA_WILL_NOT_BE_SAVED
      );
      expect(templateNotValidElement).not.toBeInTheDocument();
      const saveButtonElement = screen.queryByTestId('saveButton');
      expect(saveButtonElement).toBeEnabled();
    });
  });
});
