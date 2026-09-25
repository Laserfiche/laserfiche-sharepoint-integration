// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent, LoginType } from '@laserfiche/types-lf-ui-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import React from 'react';
import { getRegion } from '../../../Utils/Funcs';
import { useSignInPopup } from '../../../Utils/useSignInPopup';
import {
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  clientId,
  repositoryScopes,
} from '../../constants';
import {
  YOU_DO_NOT_HAVE_RIGHTS_FOR_ADMIN_CONFIG_PLEASE_CONTACT_ADMIN,
  PLEASE_LOGIN_TO_LASERFICHE,
  YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART,
  FOR_MORE_INFO_VISIT,
  SIGN_OUT,
  SIGN_IN,
} from '../../strings';
import styles from './LaserficheAdminConfiguration.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';

const LoggedOutMessage: React.FC = () => {
  return (
    <div className='pt-2'>
      {`${PLEASE_LOGIN_TO_LASERFICHE}`}
      <div className='pt-2'>
        {` ${YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART} ${FOR_MORE_INFO_VISIT} `}
        <a href='https://www.laserfiche.com/products/pricing'>laserfiche.com</a>.
      </div>
    </div>
  );
};

export const LoggedOutMessageWrapper: React.FC<{ loggedIn: boolean }> = (props) => {
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

  // useRef, not createRef: createRef hands back a new ref on every render and
  // React nulls the old one, so the listeners the mount effect registers would
  // read a dead ref after any re-render and skip initializing the repo client.
  const loginComponent: React.RefObject<NgElement & WithProperties<LfLoginComponent>> =
    React.useRef();

  const signInOrOutAsync = useSignInPopup({
    context: props.context,
    loggedIn: props.loggedIn,
    setMessageModal: props.setMessageErrorModal,
    onSignedOut: () => props.setLoggedIn(false),
  });

  React.useEffect(() => {
    const initializeComponentAsync: () => Promise<void> = async () => {
      SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
      SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
      // Unmounted while the script was loading: nothing left to set up.
      if (!loginComponent.current) {
        return;
      }
      try {
        const loginCompleted: () => Promise<void> = async () => {
          await getAndInitializeRepositoryClientAndServicesAsync();
          props.setLoggedIn(true);
        };
        const logoutCompleted: () => Promise<void> = async () => {
          props.setLoggedIn(false);
        };

        loginComponent.current.addEventListener('loginCompleted', loginCompleted);
        loginComponent.current.addEventListener('logoutCompleted', logoutCompleted);
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
    const accessToken = loginComponent?.current?.authorization_credentials?.accessToken;
    if (accessToken) {
      await props.ensureRepoClientInitializedAsync();
    } else {
      // user is not logged in
    }
  }

  return (
    <div className={styles.loginButton}>
      <lf-login
        redirect_uri={redirectPage}
        authorize_url_host_name={region}
        redirect_behavior='Replace'
        client_id={clientId}
        scope={repositoryScopes}
        login_type={LoginType.Cloud}
        ref={loginComponent}
        hidden
      />
      <button
        onClick={signInOrOutAsync}
        className={`lf-button login-button ${props.loggedIn ? 'sec-button' : 'primary-button'}`}
      >
        {props.loggedIn ? SIGN_OUT : SIGN_IN}
      </button>
    </div>
  );
};
