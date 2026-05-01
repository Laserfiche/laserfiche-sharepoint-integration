// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import {
  LfLoginComponent,
  AbortedLoginError,
} from '@laserfiche/types-lf-ui-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import React from 'react';
import { MessageDialog } from '../../../extensions/savetoLaserfiche/CommonDialogs';
import { getRegion, getSPListURL } from '../../../Utils/Funcs';
import {
  LASERFICHE_SIGNIN_PAGE_NAME,
  LOGIN_WINDOW_SUCCESS,
  clientId,
} from '../../constants';
import '../../../Utils/loadLfUiComponents';
import {
  YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN,
  PLEASE_LOGIN_TO_LASERFICHE,
  YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART,
  FOR_MORE_INFO_VISIT,
  SIGN_IN_FAILED,
  needLaserficheSignInPage,
  SIGN_OUT,
  SIGN_IN,
} from '../../strings';
import styles from './LaserficheAdminConfiguration.module.scss';

const LoggedOutMessage: React.FC = () => {
  return (
    <div className='pt-2'>
      {`${PLEASE_LOGIN_TO_LASERFICHE}`}
      <div className='pt-2'>
        {` ${YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART} ${FOR_MORE_INFO_VISIT} `}
        <a href='https://www.laserfiche.com/products/pricing'>laserfiche.com</a>
        .
      </div>
    </div>
  );
};

export const LoggedOutMessageWrapper: React.FC<{ loggedIn: boolean }> = (
  props
) => {
  return <>{!props.loggedIn ? <LoggedOutMessage /> : props.children}</>;
};

export const NoAdminRightsMessage: React.FC = () => {
  return (
    <span>
      <b>{YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN}</b>
    </span>
  );
};

export const LoginComponent: React.FC<{
  setLoggedIn: (isLoggedIn: boolean) => void;
  loggedIn: boolean;
  context: WebPartContext;
  setMessageErrorModal: (val: JSX.Element | undefined) => void;
  ensureRepoClientInitializedAsync: () => Promise<void>;
}> = (props) => {
  const region = getRegion();

  const redirectPage = window.location.origin + window.location.pathname;

  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.createRef();

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      try {
        const loginCompleted: () => Promise<void> = async () => {
          await getAndInitializeRepositoryClientAndServicesAsync();
          props.setLoggedIn(true);
        };
        const logoutCompleted: () => Promise<void> = async () => {
          props.setLoggedIn(false);
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
          props.setLoggedIn(true);
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        console.error(`Error initializing configuration page: ${err}`);
      }
    };

    void initializeComponentAsync();
  }, []);

  async function getAndInitializeRepositoryClientAndServicesAsync(): Promise<void> {
    const accessToken =
      loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await props.ensureRepoClientInitializedAsync();
    } else {
      // user is not logged in
    }
  }

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
        if (pageName === LASERFICHE_SIGNIN_PAGE_NAME) {
          return true;
        }
      }
    } catch (error) {
      console.warn(
        `Unable to determine if a SharePoint Page with name ${LASERFICHE_SIGNIN_PAGE_NAME} exists.`,
        error
      );
      return false;
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
          title={SIGN_IN_FAILED}
          message={needLaserficheSignInPage}
          clickOkay={() => {
            props.setMessageErrorModal(undefined);
          }}
        />
      );
      props.setMessageErrorModal(mes);
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
                title={SIGN_IN_FAILED}
                message={`Sign in failed, please try again. Details: ${parsedError.ErrorMessage}`}
                clickOkay={() => {
                  props.setMessageErrorModal(undefined);
                }}
              />
            );
            props.setMessageErrorModal(mes);
          }
        }
      }
    });
  }

  return (
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
          props.loggedIn ? 'sec-button' : 'primary-button'
        }`}
      >
        {props.loggedIn ? SIGN_OUT : SIGN_IN}
      </button>
    </div>
  );
};
