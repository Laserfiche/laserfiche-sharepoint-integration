// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Navigation } from 'spfx-navigation';
import {
  AbortedLoginError,
  LfLoginComponent,
  LoginState,
} from '@laserfiche/types-lf-ui-components';
import {
  clientId,
  LASERFICHE_SIGNIN_PAGE_NAME,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  LOGIN_WINDOW_SUCCESS,
  SP_LOCAL_STORAGE_KEY,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISPDocumentData } from '../../../Utils/Types';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';
import {
  getEntryWebAccessUrl,
  getRegion,
  getSPListURL,
} from '../../../Utils/Funcs';
import styles from './SendToLaserficheLoginComponent.module.scss';
import { MessageDialog } from '../../../extensions/savetoLaserfiche/CommonDialogs';
declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

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

const needLaserficheSignInPage = `Missing ${LASERFICHE_SIGNIN_PAGE_NAME} SharePoint page. Please refer to the Adding App to SharePoint Site topic in the administration guide for configuration steps.`;
export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
): JSX.Element {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();

  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);
  const [messageErrorModal, setMessageErrorModal] = React.useState<
    JSX.Element | undefined
  >(undefined);

  // Held in a ref: a plain local resets on every render, so the "post once"
  // guard could let the popup message the opener more than once.
  const sentPostMessage = React.useRef(false);
  const popupTimeout = React.useRef<ReturnType<typeof setTimeout> | undefined>(
    undefined
  );

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

  const spFileMetadata = JSON.parse(
    window.localStorage.getItem(SP_LOCAL_STORAGE_KEY)
  ) as ISPDocumentData;

  let webClientUrl: string | undefined;
  if (loggedIn) {
    webClientUrl = getEntryWebAccessUrl(
      '1',
      loginComponent.current?.account_endpoints.webClientUrl,
      true
    );
  }
  const loginText: JSX.Element | undefined = getLoginText();

  const loginCompletedInPopup: () => Promise<void> = async () => {
    // eslint-disable-next-line no-debugger -- temporary debugging aid, remove with the statement below
    debugger; //TODO: Remove this debugger statement after testing

    clearPopupTimeout();
    postToOpenerOnce(LOGIN_WINDOW_SUCCESS);
  };

  const loginCompletedInMainWindow: () => Promise<void> = async () => {
    setLoggedIn(true);
    if (spFileMetadata) {
      const dialog = new SaveToLaserficheCustomDialog(
        spFileMetadata,
        async (success) => {
          if (success) {
            Navigation.navigate(success.pathBack, true);
          }
        }
      );
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
    const errorOccurred = (ev as CustomEvent).detail as
      | AbortedLoginError
      | undefined;
    clearPopupTimeout();
    // A clean logout releases the popup; an aborted login reports why, so the
    // opener can show the failure instead of just closing the window.
    postToOpenerOnce(errorOccurred ?? LOGIN_WINDOW_SUCCESS);
  };

  React.useEffect(() => {
    const cleanUpFunction: () => void = () => {
      clearPopupTimeout();
      loginComponent.current.removeEventListener(
        'loginCompleted',
        loginCompletedInMainWindow
      );
      loginComponent.current.removeEventListener(
        'loginCompleted',
        loginCompletedInPopup
      );
      loginComponent.current.removeEventListener(
        'logoutCompleted',
        logoutCompletedInPopup
      );
      loginComponent.current.removeEventListener(
        'logoutCompleted',
        logoutCompletedInMainWindow
      );
    };

    const setUpLoginComponentAsync: () => Promise<void> = async () => {
      SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
      SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
      loginComponent.current.addEventListener(
        'logoutCompleted',
        logoutCompletedInPopup
      );
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

  async function handleLoginOrLogoutInMainWindowAsync(): Promise<void> {
    loginComponent.current.addEventListener(
      'loginCompleted',
      loginCompletedInMainWindow
    );
    loginComponent.current.addEventListener(
      'logoutCompleted',
      logoutCompletedInMainWindow
    );
    const isLoggedIn: boolean =
      loginComponent.current.state === LoginState.LoggedIn;

    setLoggedIn(isLoggedIn);
    if (isLoggedIn && spFileMetadata) {
      await trySaveToLaserficheAsync();
    }
  }

  async function trySaveToLaserficheAsync(): Promise<void> {
    const dialog = new SaveToLaserficheCustomDialog(
      spFileMetadata,
      async (success) => {
        if (success) {
          Navigation.navigate(success.pathBack, true);
        }
      }
    );
    await dialog.show();
    if (!dialog.successful) {
      console.warn('Could not sign in successfully');
    }
  }

  async function handleLoginOrLogoutInPopupAsync(): Promise<void> {
    // eslint-disable-next-line no-debugger -- temporary debugging aid, remove with the statement below
    debugger; //TODO: Remove this debugger statement after testing

    if (loginComponent.current.state === LoginState.LoggedIn) {
      const logoutButton = loginComponent.current.querySelector(
        '.login-button'
      ) as HTMLButtonElement;
      logoutButton.click();
      return;
    }

    loginComponent.current.addEventListener(
      'loginCompleted',
      loginCompletedInPopup
    );

    const redirectedFromACS =
      document.referrer.includes('accounts.') ||
      document.referrer.includes('signin.');
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
                <a href='https://www.laserfiche.com/products/pricing'>
                  laserfiche.com
                </a>
                {`.`}
              </p>
              <p>
                You are not signed in. You can sign in using the following
                button.
              </p>
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
          <div>
            {`You are now signed in. Attempting to save ${spFileMetadata?.fileName}.`}
          </div>
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
    const spFileUrlWithoutFileName = spFileUrl.replace(
      fileNameWithExtension,
      ''
    );
    const path = window.location.origin + spFileUrlWithoutFileName;
    window.localStorage.removeItem(SP_LOCAL_STORAGE_KEY);
    Navigation.navigate(path, true);
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

  const redirectURL =
    window.location.origin + window.location.pathname + '?autologin';

  return (
    <React.StrictMode>
      <div className={styles.signInHeader}>
        <img
          src='data:image/png;base64,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'
          className={styles.laserficheLogo}
        />
        <span className={styles.signInHeaderText}>Laserfiche</span>
      </div>

      <div className={styles.signInLabel}>{loginText}</div>
      <div className={styles.loginButton}>
        <lf-login
          redirect_uri={redirectURL}
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
        <br />
        {spFileMetadata?.fileUrl && (
          <button className='lf-button sec-button' onClick={redirect}>
            {CANCEL}
          </button>
        )}
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
