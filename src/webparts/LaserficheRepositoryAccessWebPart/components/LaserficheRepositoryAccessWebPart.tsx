// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import SvgHtmlIcons from '../components/SVGHtmlIcons';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  AbortedLoginError,
  LfLoginComponent,
  LoginState,
  LoginType,
} from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import {
  clientId,
  repositoryScopes,
  LASERFICHE_SIGNIN_PAGE_NAME,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  LOGIN_WINDOW_SUCCESS,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { useEffect, useState } from 'react';
import RepositoryViewComponent from './RepositoryViewWebPart';
import './LaserficheRepositoryAccess.module.scss';
import { ILaserficheRepositoryAccessWebPartProps } from './ILaserficheRepositoryAccessWebPartProps';
import { getRegion, getSPListURL, openLoginWindow } from '../../../Utils/Funcs';
import styles from './LaserficheRepositoryAccess.module.scss';
import { MessageDialog } from '../../../extensions/savetoLaserfiche/CommonDialogs';
import { POPUP_BLOCKED } from '../../strings';
import '../../../Assets/CSS/bootstrap.min.css';

const YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART =
  'You must be a currently licensed Laserfiche Cloud user to use this web part.';
const FOR_MORE_INFO_VISIT = 'For more information visit';
const ONCE_SIGNED_IN_YOULL_SEE_REPOSITORY =
  "Once signed in you'll be able to view your Laserfiche repository.";

const needLaserficheSignInPage = `Missing ${LASERFICHE_SIGNIN_PAGE_NAME} SharePoint page. Please refer to the Adding App to SharePoint Site topic in the administration guide for configuration steps.`;
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
  const loginWindowRef = React.useRef<Window | undefined>(undefined);
  // What the last click asked the popup to do. The popup reports the same
  // success message either way, so this is the only thing that says which.
  const requestedAction = React.useRef<'login' | 'logout'>('login');
  // Assigned inside the effect so the popup handler can reconcile the signed-in
  // state without duplicating the repository client setup.
  const syncSignedInState = React.useRef<(() => Promise<void>) | undefined>(undefined);

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

  // Tied to the component's lifetime: registering on click leaked a listener
  // per mount, because nothing removed it when the web part went away.
  useEffect(() => {
    const handleMessage: (event: MessageEvent) => void = (event) => {
      void handlePopupMessageAsync(event);
    };
    window.addEventListener('message', handleMessage);
    return () => {
      window.removeEventListener('message', handleMessage);
    };
  }, []);

  async function pageConfigurationCheck(): Promise<boolean> {
    try {
      const res = await fetch(`${getSPListURL(props.context, 'Site Pages')}/items`, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json',
        },
      });
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
    // The popup cannot tell a sign-in apart from a sign-out by looking at its
    // own element state, so say which one this click means.
    const action = loggedIn ? 'logout' : 'login';
    requestedAction.current = action;
    const url =
      props.context.pageContext.web.absoluteUrl +
      `/SitePages/LaserficheSignIn.aspx?autologin&action=${action}`;
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
    const loginWindow = openLoginWindow(url);
    if (!loginWindow) {
      // A blocked pop-up returns null, so bail out instead of throwing.
      const mes = (
        <MessageDialog
          title='Sign In Failed'
          message={POPUP_BLOCKED}
          clickOkay={() => {
            setMessageErrorModal(undefined);
          }}
        />
      );
      setMessageErrorModal(mes);
      return;
    }
    loginWindowRef.current = loginWindow;
  }

  async function handlePopupMessageAsync(event: MessageEvent): Promise<void> {
    if (event.origin !== window.origin) {
      return;
    }
    if (event.data === LOGIN_WINDOW_SUCCESS) {
      loginWindowRef.current?.close();
      loginWindowRef.current = undefined;
      if (requestedAction.current === 'logout') {
        // Do not ask this element: the popup signed out on its own, and this
        // one keeps reporting LoggedIn (and keeps its cached
        // authorization_credentials) until the popup's storage writes reach it.
        // Syncing here raced logoutCompleted and re-asserted "signed in", so
        // whichever landed last decided the button.
        setLoggedIn(false);
        return;
      }
      // The popup may have signed in on its own element without touching this
      // one, so ask directly rather than waiting for an event that may not come.
      await syncSignedInState.current?.();
      return;
    }
    if (event.data) {
      const parsedError: AbortedLoginError = event.data;
      if (parsedError.ErrorMessage && parsedError.ErrorType) {
        loginWindowRef.current?.close();
        loginWindowRef.current = undefined;
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
            onClick={clickLogin}
            className={`lf-button login-button ${loggedIn ? 'sec-button' : 'primary-button'}`}
          >
            {loggedIn ? 'Sign out' : 'Sign in'}
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
