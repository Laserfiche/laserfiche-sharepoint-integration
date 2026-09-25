// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import SvgHtmlIcons from '../components/SVGHtmlIcons';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { LfLoginComponent, LoginState, LoginType } from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import {
  clientId,
  repositoryScopes,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { useEffect, useState } from 'react';
import RepositoryViewComponent from './RepositoryViewWebPart';
import './LaserficheRepositoryAccess.module.scss';
import { ILaserficheRepositoryAccessWebPartProps } from './ILaserficheRepositoryAccessWebPartProps';
import { getRegion } from '../../../Utils/Funcs';
import { useSignInPopup } from '../../../Utils/useSignInPopup';
import styles from './LaserficheRepositoryAccess.module.scss';
import { SIGN_IN, SIGN_OUT } from '../../strings';
import '../../../Assets/CSS/bootstrap.min.css';

const YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART =
  'You must be a currently licensed Laserfiche Cloud user to use this web part.';
const FOR_MORE_INFO_VISIT = 'For more information visit';
const ONCE_SIGNED_IN_YOULL_SEE_REPOSITORY =
  "Once signed in you'll be able to view your Laserfiche repository.";

export default function LaserficheRepositoryAccessWebPart(
  props: ILaserficheRepositoryAccessWebPartProps
): JSX.Element {
  const [webClientUrl, setWebClientUrl] = React.useState('');
  const [customerId, setCustomerId] = React.useState('');
  const loginComponent: React.RefObject<NgElement & WithProperties<LfLoginComponent>> =
    React.useRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<IRepositoryApiClientExInternal | undefined>(
    undefined
  );
  const [messageErrorModal, setMessageErrorModal] = useState<JSX.Element | undefined>(undefined);
  // Assigned inside the effect so the popup handler can reconcile the signed-in
  // state without duplicating the repository client setup.
  const syncSignedInState = React.useRef<(() => Promise<void>) | undefined>(undefined);

  const signInOrOutAsync = useSignInPopup({
    context: props.context,
    loggedIn,
    setMessageModal: setMessageErrorModal,
    // Do not ask this element: the popup signed out on its own, and this one
    // keeps reporting LoggedIn (and keeps its cached authorization_credentials)
    // until the popup's storage writes reach it. Syncing here raced
    // logoutCompleted and re-asserted "signed in", so whichever landed last
    // decided the button.
    onSignedOut: () => setLoggedIn(false),
    // The popup may have signed in on its own element without touching this
    // one, so ask directly rather than waiting for an event that may not come.
    onSignedIn: async () => {
      await syncSignedInState.current?.();
    },
  });

  const region = getRegion();

  const redirectPage = window.location.origin + window.location.pathname;

  useEffect(() => {
    const ensureRepoClientInitializedAsync: () => Promise<void> = async () => {
      if (!repoClient) {
        const repoClientCreator = new RepositoryClientExInternal();
        const repoClient = await repoClientCreator.createRepositoryClientAsync();
        setRepoClient(repoClient);
      }
    };

    const getAndInitializeRepositoryClientAndServicesAsync: () => Promise<void> = async () => {
      const accessToken = loginComponent?.current?.authorization_credentials?.accessToken;
      setWebClientUrl(loginComponent?.current?.account_endpoints.webClientUrl);
      setCustomerId(loginComponent?.current?.account_id);
      if (accessToken) {
        await ensureRepoClientInitializedAsync();
      } else {
        // user is not logged in
      }
    };

    // A sign-in can finish without this element ever raising loginCompleted: it
    // may have restored the session before we subscribed, or the popup may have
    // found an existing session and changed no storage at all. So reconcile
    // explicitly at the points we know something happened.
    const syncSignedInStateAsync: () => Promise<void> = async () => {
      const signedIn =
        loginComponent.current?.state === LoginState.LoggedIn ||
        !!loginComponent.current?.authorization_credentials;
      if (signedIn) {
        await getAndInitializeRepositoryClientAndServicesAsync();
      }
      setLoggedIn(signedIn);
    };
    syncSignedInState.current = syncSignedInStateAsync;

    const initializeComponentAsync: () => Promise<void> = async () => {
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
      SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
      SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
      try {
        const loginCompleted: () => Promise<void> = async () => {
          await getAndInitializeRepositoryClientAndServicesAsync();
          setLoggedIn(true);
        };
        const logoutCompleted: () => Promise<void> = async () => {
          setLoggedIn(false);
        };

        loginComponent.current.addEventListener('loginCompleted', loginCompleted);
        loginComponent.current.addEventListener('logoutCompleted', logoutCompleted);
        await syncSignedInStateAsync();
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        console.error(`Unable to initialize repository explorer: ${err}`);
      }
    };

    void initializeComponentAsync();
  }, []);

  return (
    <React.StrictMode>
      <div style={{ display: 'none' }}>
        <SvgHtmlIcons />
      </div>
      <div className='p-3'>
        <div className={styles.loginButton}>
          <lf-login
            redirect_uri={redirectPage}
            redirect_behavior='Replace'
            client_id={clientId}
            scope={repositoryScopes}
            authorize_url_host_name={region}
            login_type={LoginType.Cloud}
            ref={loginComponent}
            hidden
          />
          <button
            onClick={signInOrOutAsync}
            className={`lf-button login-button ${loggedIn ? 'sec-button' : 'primary-button'}`}
          >
            {loggedIn ? SIGN_OUT : SIGN_IN}
          </button>
        </div>
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
        <RepositoryViewComponent
          webClientUrl={webClientUrl}
          customerId={customerId}
          repoClient={repoClient}
          loggedIn={loggedIn}
        />
        {!loggedIn && (
          <span>
            {`${YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART} ${FOR_MORE_INFO_VISIT} `}
            <a href='https://www.laserfiche.com/products/pricing'>laserfiche.com</a>
            {`. ${ONCE_SIGNED_IN_YOULL_SEE_REPOSITORY}`}
          </span>
        )}
      </div>
    </React.StrictMode>
  );
}
