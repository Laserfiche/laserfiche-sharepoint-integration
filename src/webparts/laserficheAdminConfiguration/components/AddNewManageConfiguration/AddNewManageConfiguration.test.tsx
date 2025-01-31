// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import '@testing-library/jest-dom';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import AddNewManageConfiguration from './AddNewManageConfiguration';
import { RepositoryClientExInternal } from '../../../../repository-client/repository-client';
import { BrowserRouter } from 'react-router-dom';
import { MANAGE_CONFIGURATIONS_PAGE_TITLE, PROFILE_WITH_NAME_ALREADY_EXISTS_PROVIDE_DIFFERENT_NAME } from '../../../strings';
import {
  ProfileConfigContext,
  ProfileConfigContextProps,
} from '../LaserficheAdminConfiguration';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http-base';
import { SPWeb, PageContext, SPPermission } from '@microsoft/sp-page-context';

class SaveDisabledStateMock implements ProfileConfigContextProps {
  private _saveDisabled: boolean = false;

  get saveDisabled(): boolean {
    return this._saveDisabled;
  }
  setSaveDisabled: React.Dispatch<React.SetStateAction<boolean>> = (val) => {
    if (typeof val === 'function') {
      this._saveDisabled = (val as (prevState: boolean) => boolean)(
        this._saveDisabled
      );
    } else {
      this._saveDisabled = val;
    }
  };
}

describe('AddNewManageConfiguration', () => {
  let mockWebPartContext: WebPartContext;

  beforeEach(() => {
    const spHttpClient = {
      post: jest.fn(),
    };
    mockWebPartContext = {
      spHttpClient: spHttpClient as unknown as SPHttpClient,
      pageContext: {
        web: {
          permissions: new SPPermission({ High: 0, Low: 0 }),
        } as SPWeb,
      } as PageContext,
    } as WebPartContext;
  });

  test('renders title of page', () => {
    // Arrange/Act
    render(
      <BrowserRouter>
        <ProfileConfigContext.Provider
          value={{ setSaveDisabled: jest.fn(), saveDisabled: false }}
        >
          <AddNewManageConfiguration
            context={mockWebPartContext}
            repoClient={new RepositoryClientExInternal().repoClient}
            loggedIn={true}
          />
        </ProfileConfigContext.Provider>
      </BrowserRouter>
    );

    // Assert
    const linkElement = screen.getByText(MANAGE_CONFIGURATIONS_PAGE_TITLE);
    expect(linkElement).toBeInTheDocument();
  });

  test('savebutton is disabled if profile already exists in savedProfileConfigurations', async () => {
    // Arrange
    const mockProfileConfig = {
      ConfigurationName: 'ExistingProfile',
      // other properties
    };

    const mockManageConfigurationConfig = [
      {
        Id: '1',
        JsonValue: JSON.stringify([mockProfileConfig]),
      },
    ];
    window.fetch = jest.fn().mockResolvedValue({
      json: jest
        .fn()
        .mockResolvedValue({ value: mockManageConfigurationConfig }),
    });

    // Act
    render(
      <BrowserRouter>
        <ProfileConfigContext.Provider value={new SaveDisabledStateMock()}>
          <AddNewManageConfiguration
            context={mockWebPartContext}
            repoClient={new RepositoryClientExInternal().repoClient}
            loggedIn={true}
          />
        </ProfileConfigContext.Provider>
      </BrowserRouter>
    );

    fireEvent.change(screen.getByLabelText(/Profile Name/i), {
      target: { value: 'ExistingProfile' },
    });

    fireEvent.click(screen.getByText(/Save/i));

    // Assert
    await waitFor(() => {
      expect(
        screen.getByText(
          PROFILE_WITH_NAME_ALREADY_EXISTS_PROVIDE_DIFFERENT_NAME
        )
      ).toBeInTheDocument();
      expect(screen.getByTestId('saveButton')).toBeDisabled();
    });
  });

  test('savebutton is enabled on render', async () => {
    // Arrange
    const mockProfileConfig = {
      ConfigurationName: 'ExistingProfile',
      // other properties
    };

    const mockManageConfigurationConfig = [
      {
        Id: '1',
        JsonValue: JSON.stringify([mockProfileConfig]),
      },
    ];
    window.fetch = jest.fn().mockResolvedValue({
      json: jest
        .fn()
        .mockResolvedValue({ value: mockManageConfigurationConfig }),
    });

    // Act
    render(
      <BrowserRouter>
        <ProfileConfigContext.Provider value={new SaveDisabledStateMock()}>
          <AddNewManageConfiguration
            context={mockWebPartContext}
            repoClient={new RepositoryClientExInternal().repoClient}
            loggedIn={true}
          />
        </ProfileConfigContext.Provider>
      </BrowserRouter>
    );

    // Assert
    expect(screen.getByTestId('saveButton')).toBeEnabled();
  });
});
