// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { useState } from 'react';
import { ILaserficheAdminConfigurationProps } from './ILaserficheAdminConfigurationProps';
import { HashRouter, Route, Switch } from 'react-router-dom';
import { Stack, StackItem } from 'office-ui-fabric-react';
import AdminMainPage from '../components/AdminMainPage/AdminMainPage';
import HomePage from './HomePage/HomePage';
import ManageConfigurationsPage from './ManageConfigurationsPage/ManageConfigurationsPage';
import ManageMappingsPage from './ManageMappingsPage/ManageMappingsPage';
import EditManageConfiguration from './EditManageConfiguration/EditManageConfiguration';
import AddNewManageConfiguration from './AddNewManageConfiguration/AddNewManageConfiguration';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import styles from './LaserficheAdminConfiguration.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  LoggedOutMessageWrapper,
  LoginComponent,
  NoAdminRightsMessage,
} from './AdminConfigurationUtilComponents';

export interface ProfileConfigContextProps {
  saveDisabled: boolean;
  setSaveDisabled: React.Dispatch<React.SetStateAction<boolean>>;
}
export const ProfileConfigContext = React.createContext<
  ProfileConfigContextProps | undefined
>(undefined);

const ProfileConfigStateProvider = (
  props: React.PropsWithChildren<{}>
): JSX.Element => {
  const [saveDisabled, setSaveDisabled] = useState<boolean>(false);

  const contextValue = {
    saveDisabled,
    setSaveDisabled,
  };

  return (
    <ProfileConfigContext.Provider value={contextValue}>
      {props.children}
    </ProfileConfigContext.Provider>
  );
};

export default function LaserficheAdminConfiguration(
  props: ILaserficheAdminConfigurationProps
): JSX.Element {
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);
  const [messageErrorModal, setMessageErrorModal] = useState<
    JSX.Element | undefined
  >(undefined);

  function isAdmin(): boolean {
    const permission = new SPPermission(
      props.context.pageContext.web.permissions.value
    );
    const isFullControl = permission.hasPermission(SPPermission.manageWeb);
    return isFullControl;
  }

  async function ensureRepoClientInitializedAsync(): Promise<void> {
    if (!repoClient) {
      const repoClientCreator = new RepositoryClientExInternal();
      const newRepoClient =
        await repoClientCreator.createRepositoryClientAsync();
      setRepoClient(newRepoClient);
    }
  }

  return (
    <React.StrictMode>
      <HashRouter>
        <Stack>
          {isAdmin() && (
            <>
              <LoginComponent
                loggedIn={loggedIn}
                setLoggedIn={setLoggedIn}
                setMessageErrorModal={setMessageErrorModal}
                ensureRepoClientInitializedAsync={
                  ensureRepoClientInitializedAsync
                }
                context={props.context}
              />
              <AdminMainPage
                context={props.context}
                loggedIn={loggedIn}
                repoClient={repoClient}
              />
            </>
          )}
          <StackItem>
            <Switch>
              <Route
                exact={true}
                component={() => <HomePage />}
                path='/HomePage'
              />
              <Route exact={true} component={() => <HomePage />} path='/' />

              <Route
                component={() => (
                  <>
                    <LoggedOutMessageWrapper loggedIn={loggedIn}>
                      <ManageConfigurationsPage context={props.context} />
                    </LoggedOutMessageWrapper>
                  </>
                )}
                path='/ManageConfigurationsPage'
              />
              <Route
                exact={true}
                component={() => (
                  <LoggedOutMessageWrapper loggedIn={loggedIn}>
                    <ManageMappingsPage
                      context={props.context}
                      isLoggedIn={loggedIn}
                      repoClient={repoClient}
                    />
                  </LoggedOutMessageWrapper>
                )}
                path='/ManageMappingsPage'
              />
              <Route
                exact={true}
                component={() => (
                  <ProfileConfigStateProvider>
                    <LoggedOutMessageWrapper loggedIn={loggedIn}>
                      <AddNewManageConfiguration
                        context={props.context}
                        loggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    </LoggedOutMessageWrapper>
                  </ProfileConfigStateProvider>
                )}
                path='/AddNewManageConfiguration'
              />
              <Route
                exact={true}
                render={(properties) => (
                  <ProfileConfigStateProvider>
                    <LoggedOutMessageWrapper loggedIn={loggedIn}>
                      <EditManageConfiguration
                        {...properties}
                        context={props.context}
                        loggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    </LoggedOutMessageWrapper>
                  </ProfileConfigStateProvider>
                )}
                path='/EditManageConfiguration/:name'
              />
            </Switch>
          </StackItem>
          {!isAdmin() && <NoAdminRightsMessage />}
          {messageErrorModal !== undefined && (
            <div
              className={styles.modal}
              id='messageErrorModal'
              data-backdrop='static'
              data-keyboard='false'
            >
              {messageErrorModal}
            </div>
          )}
        </Stack>
      </HashRouter>
    </React.StrictMode>
  );
}
