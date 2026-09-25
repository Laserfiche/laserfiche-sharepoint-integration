// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Navigation } from 'spfx-navigation';
import {
  AbortedLoginError,
  LfLoginComponent,
  LoginState,
  LoginType,
} from '@laserfiche/types-lf-ui-components';
import {
  clientId,
  repositoryScopes,
  LASERFICHE_ICON_URL,
  LASERFICHE_SIGNIN_PAGE_NAME,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  LOGIN_WINDOW_SUCCESS,
  SP_LOCAL_STORAGE_KEY,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';
import {
  getEntryWebAccessUrl,
  getRegion,
  getSPDocumentDataFromLocalStorage,
  getSPListURL,
  openLoginWindow,
} from '../../../Utils/Funcs';
import styles from './SendToLaserficheLoginComponent.module.scss';
import { MessageDialog } from '../../../extensions/savetoLaserfiche/CommonDialogs';
import { LASERFICHE, POPUP_BLOCKED } from '../../strings';

const CANCEL = 'Cancel';
// How long the popup waits for the component to finish exchanging the
// authorization code before it reports failure to the opener.
const POPUP_LOGIN_TIMEOUT_MS = 30000;
const SIGN_IN_DID_NOT_COMPLETE = 'Timed out waiting for sign in to complete.';
const NOTE_THIS_WEB_PART_IS_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE =
  '*Note: This web part is only needed if you are attempting to save a document to Laserfiche.';
const YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART =
  'You must be a currently licensed Laserfiche Cloud user to use this web part.';
const FOR_MORE_INFO_VISIT = 'For more information visit';

// SPFx records the gulp serve manifest URL here once a page is loaded with
// debugManifestsFile, and keeps using it for the rest of the session.
const SPFX_DEBUG_SESSION_STORAGE_KEY = 'spfx-debug';

// The sign-in popup gets a URL of our own making, so nothing carries the local
// debug build into it. A popup inherits the opener's session storage, but that
// on its own is not enough: it asks for the localhost manifests and fails with
// "Error loading debug manifests". Repeating the parameters in the popup URL is
// what actually loads them, so forward them whenever the opener is itself
// running against a debug manifest. Returns '' in a normal deployment, where
// the key is absent, leaving the popup URL untouched.
function debugManifestsQueryString(): string {
  try {
    const debugSettings = sessionStorage.getItem(SPFX_DEBUG_SESSION_STORAGE_KEY);
    if (!debugSettings) {
      return '';
    }
    const manifestsFileUrl: string | undefined = JSON.parse(debugSettings).manifestsFileUrl;
    if (!manifestsFileUrl) {
      return '';
    }
    return `&debugManifestsFile=${encodeURIComponent(
      manifestsFileUrl
    )}&loadSPFX=true&debug=true&noredir=true`;
  } catch {
    // Session storage can throw, and the entry is not ours to assume is valid
    // JSON. Debug parameters are a developer convenience, so fall back to the
    // plain URL rather than break sign-in over them.
    return '';
  }
}

const needLaserficheSignInPage = `Missing ${LASERFICHE_SIGNIN_PAGE_NAME} SharePoint page. Please refer to the Adding App to SharePoint Site topic in the administration guide for configuration steps.`;
export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
): JSX.Element {
  const loginComponent: React.RefObject<NgElement & WithProperties<LfLoginComponent>> =
    React.useRef();

  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);
  const [messageErrorModal, setMessageErrorModal] = React.useState<JSX.Element | undefined>(
    undefined
  );

  // Held in a ref: a plain local resets on every render, so the "post once"
  // guard could let the popup message the opener more than once.
  const sentPostMessage = React.useRef(false);
  const popupTimeout = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  const logoutRequested = React.useRef(false);

  // Which action this page's own button asked the popup for, and the popup it
  // opened. Both are refs so the message handler, attached once, reads the
  // current click rather than the one that registered it.
  const requestedAction = React.useRef<'login' | 'logout'>('login');
  const loginWindowRef = React.useRef<Window | undefined>(undefined);

  const postToOpenerOnce: (message: unknown) => void = (message) => {
    if (sentPostMessage.current) {
      return;
    }
    sentPostMessage.current = true;
    window.opener?.postMessage(message, window.origin);
  };

  const clearPopupTimeout: () => void = () => {
    if (popupTimeout.current !== undefined) {
      clearTimeout(popupTimeout.current);
      popupTimeout.current = undefined;
    }
  };

  const startPopupTimeout: () => void = () => {
    clearPopupTimeout();
    popupTimeout.current = setTimeout(() => {
      postToOpenerOnce({
        ErrorType: 'TokenExchangeTimeout',
        ErrorMessage: SIGN_IN_DID_NOT_COMPLETE,
      } as AbortedLoginError);
    }, POPUP_LOGIN_TIMEOUT_MS);
  };

  const region = getRegion();

  const spFileMetadata = getSPDocumentDataFromLocalStorage();

  let webClientUrl: string | undefined;
  if (loggedIn) {
    webClientUrl = getEntryWebAccessUrl(
      '1',
      loginComponent.current?.account_endpoints.webClientUrl,
      true,
      undefined,
      loginComponent.current?.account_id
    );
  }
  const loginText: JSX.Element | undefined = getLoginText();

  const loginCompletedInPopup: () => Promise<void> = async () => {
    clearPopupTimeout();
    postToOpenerOnce(LOGIN_WINDOW_SUCCESS);
  };

  const loginCompletedInMainWindow: () => Promise<void> = async () => {
    setLoggedIn(true);
    if (spFileMetadata) {
      const dialog = new SaveToLaserficheCustomDialog(spFileMetadata, async (success) => {
        if (success) {
          Navigation.navigate(success.pathBack, true);
        }
      });
      await dialog.show();
      if (!dialog.successful) {
        console.warn('Could not sign in successfully');
      }
    }
  };

  const logoutCompletedInMainWindow: () => void = () => {
    setLoggedIn(false);
  };

  const logoutCompletedInPopup: (ev: Event) => void = (ev: Event) => {
    const errorOccurred = (ev as CustomEvent).detail as AbortedLoginError | undefined;

    // lf-login also raises this with no detail when it decides nobody is signed
    // in yet, which is the normal opening move of a sign-in. Releasing the popup
    // on that would close it before it ever reaches the sign-in page, so only an
    // aborted login or a logout we asked for ends the popup here.
    if (!errorOccurred && !logoutRequested.current) {
      return;
    }

    clearPopupTimeout();
    postToOpenerOnce(errorOccurred ?? LOGIN_WINDOW_SUCCESS);
  };

  React.useEffect(() => {
    // Held here rather than read from the ref at cleanup time: React 17 clears
    // refs before it runs an unmounting component's effect cleanups, so
    // loginComponent.current is already null by then.
    const loginElement = loginComponent.current;
    const cleanUpFunction: () => void = () => {
      clearPopupTimeout();
      loginElement.removeEventListener('loginCompleted', loginCompletedInMainWindow);
      loginElement.removeEventListener('loginCompleted', loginCompletedInPopup);
      loginElement.removeEventListener('logoutCompleted', logoutCompletedInPopup);
      loginElement.removeEventListener('logoutCompleted', logoutCompletedInMainWindow);
    };

    const setUpLoginComponentAsync: () => Promise<void> = async () => {
      SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
      SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
      loginComponent.current.addEventListener('logoutCompleted', logoutCompletedInPopup);
      await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);

      try {
        if (window.location.href.includes('autologin')) {
          document.body.style.display = 'none';
          await handleLoginOrLogoutInPopupAsync();
        } else {
          await handleLoginOrLogoutInMainWindowAsync();
        }
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
      } catch (err: any) {
        console.error(`Unable to initialize sign-in page: ${err}`);
      }
    };

    cleanUpFunction();
    void setUpLoginComponentAsync();

    return cleanUpFunction;
  }, []);

  // Kept apart from the effect above, whose cleanup is about the lf-login
  // element. Tied to the component's lifetime: registering on click leaked a
  // listener per mount, because nothing removed it when this page went away.
  React.useEffect(() => {
    const handleMessage: (event: MessageEvent) => void = (event) => {
      handlePopupMessage(event);
    };
    window.addEventListener('message', handleMessage);
    return () => {
      window.removeEventListener('message', handleMessage);
    };
  }, []);

  async function handleLoginOrLogoutInMainWindowAsync(): Promise<void> {
    loginComponent.current.addEventListener('loginCompleted', loginCompletedInMainWindow);
    loginComponent.current.addEventListener('logoutCompleted', logoutCompletedInMainWindow);
    const isLoggedIn: boolean = loginComponent.current.state === LoginState.LoggedIn;

    setLoggedIn(isLoggedIn);
    if (isLoggedIn && spFileMetadata) {
      await trySaveToLaserficheAsync();
    }
  }

  async function trySaveToLaserficheAsync(): Promise<void> {
    const dialog = new SaveToLaserficheCustomDialog(spFileMetadata, async (success) => {
      if (success) {
        Navigation.navigate(success.pathBack, true);
      }
    });
    await dialog.show();
    if (!dialog.successful) {
      console.warn('Could not sign in successfully');
    }
  }

  async function handleLoginOrLogoutInPopupAsync(): Promise<void> {
    if (loginComponent.current.state === LoginState.LoggedIn) {
      const wantsLogout = new URLSearchParams(window.location.search).get('action') === 'logout';
      if (!wantsLogout) {
        // The opener asked for a sign-in and this element already holds a
        // session. Signing out here is what made clicking "Sign in" sign the
        // user out; report success instead and let the opener adopt it.
        postToOpenerOnce(LOGIN_WINDOW_SUCCESS);
        return;
      }
      logoutRequested.current = true;
      const logoutButton = loginComponent.current.querySelector(
        '.login-button'
      ) as HTMLButtonElement;
      logoutButton.click();
      return;
    }

    loginComponent.current.addEventListener('loginCompleted', loginCompletedInPopup);

    const redirectedFromACS =
      document.referrer.includes('accounts.') || document.referrer.includes('signin.');
    if (!redirectedFromACS) {
      await loginComponent.current.initLoginFlowAsync();
      return;
    }

    // Back from the sign-in page the component is still exchanging the
    // authorization code for a token, so its state reads LoggedOut for a moment.
    // Reporting success here closed the popup mid-exchange, the token never
    // reached localStorage, and the opener never saw the storage event that
    // makes lf-login emit loginCompleted. Wait for the real event instead, with
    // a timeout so a silent failure still releases the popup.
    startPopupTimeout();
  }

  function getLoginText(): JSX.Element {
    let loginText: JSX.Element | undefined;
    if (!spFileMetadata) {
      loginText = (
        <>
          <p>{NOTE_THIS_WEB_PART_IS_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE}</p>
          {loggedIn ? (
            <p>
              {'Welcome to Laserfiche.'}
              {webClientUrl && (
                <>
                  {' Go to '}
                  <a
                    href={webClientUrl}
                    target='_blank'
                    rel='noreferrer'
                    style={{ color: '#0079d6' }}
                  >
                    your Laserfiche repository
                  </a>
                </>
              )}
            </p>
          ) : (
            <div>
              <p>
                {`${YOU_MUST_BE_CLOUD_USER_TO_USE_WEB_PART} ${FOR_MORE_INFO_VISIT} `}
                <a href='https://www.laserfiche.com/products/pricing'>laserfiche.com</a>
                {`.`}
              </p>
              <p>You are not signed in. You can sign in using the following button.</p>
            </div>
          )}
        </>
      );
    } else if (spFileMetadata?.fileUrl && !loggedIn) {
      loginText = (
        <>
          <div>
            {`You are not signed in. Please sign in to continue saving ${spFileMetadata?.fileName}.`}
          </div>
          <br />
        </>
      );
    } else if (spFileMetadata?.fileUrl && loggedIn) {
      loginText = (
        <>
          <div>{`You are now signed in. Attempting to save ${spFileMetadata?.fileName}.`}</div>
          <br />
        </>
      );
    } else {
      <p>{NOTE_THIS_WEB_PART_IS_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE}</p>;
    }
    return loginText;
  }

  function redirect(): void {
    const spFileUrl = spFileMetadata.fileUrl;
    const fileNameWithExtension = spFileMetadata.fileName;
    const spFileUrlWithoutFileName = spFileUrl.replace(fileNameWithExtension, '');
    const path = window.location.origin + spFileUrlWithoutFileName;
    window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
    Navigation.navigate(path, true);
  }

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
    // own element state, so say which one this click means. Without it the
    // popup takes the click for a sign-in, reports success and signs nobody
    // out, which is what made "Sign out" here do nothing.
    const action = loggedIn ? 'logout' : 'login';
    requestedAction.current = action;
    const url =
      props.context.pageContext.web.absoluteUrl +
      `/SitePages/LaserficheSignIn.aspx?autologin&action=${action}` +
      debugManifestsQueryString();
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

  function handlePopupMessage(event: MessageEvent): void {
    if (event.origin !== window.origin) {
      return;
    }
    if (event.data === LOGIN_WINDOW_SUCCESS) {
      loginWindowRef.current?.close();
      loginWindowRef.current = undefined;
      if (requestedAction.current === 'logout') {
        // The popup signs out on its own lf-login element, so
        // logoutCompleted does not necessarily reach the one on this page.
        setLoggedIn(false);
      }
    } else if (event.data) {
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

  const redirectURL = window.location.origin + window.location.pathname + '?autologin';

  return (
    <React.StrictMode>
      <div className={styles.signInHeader}>
        <img src={LASERFICHE_ICON_URL} className={styles.laserficheLogo} />
        <span className={styles.signInHeaderText}>{LASERFICHE}</span>
      </div>

      <div className={styles.signInLabel}>{loginText}</div>
      <div className={styles.loginButton}>
        <lf-login
          redirect_uri={redirectURL}
          authorize_url_host_name={region}
          redirect_behavior='Replace'
          client_id={clientId}
          scope={repositoryScopes}
          login_type={LoginType.Cloud}
          ref={loginComponent}
          hidden
        />
        <div className={styles.buttonRow}>
          <button
            onClick={clickLogin}
            className={`lf-button login-button ${loggedIn ? 'sec-button' : 'primary-button'}`}
          >
            {loggedIn ? 'Sign out' : 'Sign in'}
          </button>
          {spFileMetadata?.fileUrl && (
            <button className='lf-button sec-button' onClick={redirect}>
              {CANCEL}
            </button>
          )}
        </div>
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
    </React.StrictMode>
  );
}
