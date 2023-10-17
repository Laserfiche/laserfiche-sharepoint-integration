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
import {
  clientId,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  LOGIN_WINDOW_SUCCESS,
  ZONE_JS_URL,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import {
  AbortedLoginError,
  LfLoginComponent,
} from '@laserfiche/types-lf-ui-components';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { getRegion, getSPListURL } from '../../../Utils/Funcs';
import styles from './LaserficheAdminConfiguration.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';
import { MessageDialog } from '../../../extensions/savetoLaserfiche/CommonDialogs';

const YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN =
  'You do not have the necessary rights to view or edit the Laserfiche SharePoint Integration configuration. Please contact your administrator for help.';

const needLaserficheSignInPage = `Missing "LaserficheSignIn" SharePoint page. Please refer to the Adding App to SharePoint Site topic in the administration guide for configuration steps.`;
export default function LaserficheAdminConfiguration(
  props: ILaserficheAdminConfigurationProps
): JSX.Element {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.createRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);
  const [messageErrorModal, setMessageErrorModal] = useState<
    JSX.Element | undefined
  >(undefined);

  const region = getRegion();

  const redirectPage = window.location.origin + window.location.pathname;

  function isAdmin(): boolean {
    const permission = new SPPermission(
      props.context.pageContext.web.permissions.value
    );
    const isFullControl = permission.hasPermission(SPPermission.manageWeb);
    return isFullControl;
  }

  async function getAndInitializeRepositoryClientAndServicesAsync(): Promise<void> {
    const accessToken =
      loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await ensureRepoClientInitializedAsync();
    } else {
      // user is not logged in
    }
  }

  async function ensureRepoClientInitializedAsync(): Promise<void> {
    if (!repoClient) {
      const repoClientCreator = new RepositoryClientExInternal();
      const newRepoClient =
        await repoClientCreator.createRepositoryClientAsync();
      setRepoClient(newRepoClient);
    }
  }

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
      SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
      await SPComponentLoader.loadScript(ZONE_JS_URL);
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
      try {
        const loginCompleted: () => Promise<void> = async () => {
          await getAndInitializeRepositoryClientAndServicesAsync();
          setLoggedIn(true);
        };
        const logoutCompleted: () => Promise<void> = async () => {
          setLoggedIn(false);
        };

        loginComponent.current.addEventListener(
          'loginCompleted',
          loginCompleted
        );
        loginComponent.current.addEventListener(
          'logoutCompleted',
          logoutCompleted
        );
        if (loginComponent.current.authorization_credentials) {
          await getAndInitializeRepositoryClientAndServicesAsync();
          setLoggedIn(true);
        }
      } catch (err) {
        console.error(`Error initializing configuration page: ${err}`);
      }
    };

    void initializeComponentAsync();
  }, []);

  async function pageConfigurationCheck(): Promise<boolean> {
    try {
      const res = await fetch(
        `${getSPListURL(props.context, 'Site Pages')}/items`,
        {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json',
          },
        }
      );
      const sitePages = await res.json();
      for (let o = 0; o < sitePages.value.length; o++) {
        const pageName = sitePages.value[o].Title;
        if (pageName === 'LaserficheSignIn') {
          return true;
        }
      }
    } catch (error) {
      // TODO
    }
    return false;
  }

  async function clickLogin(): Promise<void> {
    const url =
      props.context.pageContext.web.absoluteUrl +
      '/SitePages/LaserficheSignIn.aspx?autologin';
    const hasSignIn = await pageConfigurationCheck();
    if (!hasSignIn) {
      const mes = (
        <MessageDialog
          title='Sign In Failed'
          message={needLaserficheSignInPage}
          clickOkay={() => {
            setMessageErrorModal(undefined);
          }}
        />
      );
      setMessageErrorModal(mes);
      return;
    }
    const loginWindow = window.open(url, 'loginWindow', 'popup');
    loginWindow.resizeTo(800, 600);
    window.addEventListener('message', (event) => {
      if (event.origin === window.origin) {
        if (event.data === LOGIN_WINDOW_SUCCESS) {
          loginWindow.close();
        } else if (event.data) {
          const parsedError: AbortedLoginError = event.data;
          if (parsedError.ErrorMessage && parsedError.ErrorType) {
            loginWindow.close();
            const mes = (
              <MessageDialog
                title='Sign In Failed'
                message={`Sign in failed, please try again. Details: ${parsedError.ErrorMessage}`}
                clickOkay={() => {
                  setMessageErrorModal(undefined);
                }}
              />
            );
            setMessageErrorModal(mes);
          }
        }
      }
    });
  }

  return (
    <React.StrictMode>
      <HashRouter>
        <Stack>
          {isAdmin() && (
            <>
              <div className={styles.loginButton}>
                <lf-login
                  redirect_uri={redirectPage}
                  authorize_url_host_name={region}
                  redirect_behavior='Replace'
                  client_id={clientId}
                  ref={loginComponent}
                  hidden
                />
                <button
                  onClick={clickLogin}
                  className={`lf-button login-button ${
                    loggedIn ? 'sec-button' : 'primary-button'
                  }`}
                >
                  {loggedIn ? 'Sign out' : 'Sign in'}
                </button>
              </div>
              <AdminMainPage
                context={props.context}
                loggedIn={loggedIn}
                repoClient={repoClient}
              />
              <StackItem>
                <Switch>
                  <Route
                    exact={true}
                    component={() => <HomePage />}
                    path='/HomePage'
                  />
                  <Route exact={true} component={() => <HomePage />} path='/' />
                  <Route
                    exact={true}
                    component={() => (
                      <ManageConfigurationsPage context={props.context} />
                    )}
                    path='/ManageConfigurationsPage'
                  />
                  <Route
                    exact={true}
                    component={() => (
                      <ManageMappingsPage
                        context={props.context}
                        isLoggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    )}
                    path='/ManageMappingsPage'
                  />
                  <Route
                    exact={true}
                    component={() => (
                      <AddNewManageConfiguration
                        context={props.context}
                        loggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    )}
                    path='/AddNewManageConfiguration'
                  />
                  <Route
                    exact={true}
                    render={(properties) => (
                      <EditManageConfiguration
                        {...properties}
                        context={props.context}
                        loggedIn={loggedIn}
                        repoClient={repoClient}
                      />
                    )}
                    path='/EditManageConfiguration/:name'
                  />
                </Switch>
              </StackItem>
            </>
          )}
          {!isAdmin() && (
            <span>
              <b>
                {YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN}
              </b>
            </span>
          )}
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
