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
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  SP_LOCAL_STORAGE_KEY,
  ZONE_JS_URL,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISPDocumentData } from '../../../Utils/Types';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';
import { getEntryWebAccessUrl, getRegion } from '../../../Utils/Funcs';
import styles from './SendToLaserficheLoginComponent.module.scss';

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
const NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE =
  '*Note: This page should only be needed if you are attempting to save a document to Laserfiche.';

export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
): JSX.Element {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();

  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);

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

  const autoLoginCompleted: () => Promise<void> = async () => {
    window.opener.postMessage('loginWindowSuccess', window.origin);
  };

  const loginCompleted: () => Promise<void> = async () => {
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

  const logoutCompleted: (
    ev: CustomEvent<void | AbortedLoginError>,
    autoLogout?: boolean
  ) => void = (
    ev: CustomEvent<void | AbortedLoginError>,
    autoLogout: boolean = false
  ) => {
    const logOutError = ev.detail;
    if (autoLogout && !logOutError) {
      window.opener.postMessage('loginWindowSuccess', window.origin);
    } else if (logOutError) {
      window.opener.postMessage(logOutError, window.origin);
    } else {
      setLoggedIn(false);
    }
  };

  React.useEffect(() => {
    const logoutCompleteCallBackTrue: (ev: Event) => void = (ev: Event) => {
      const errorOccurred = (ev as CustomEvent).detail;
      if (errorOccurred) {
        logoutCompleted(ev as CustomEvent, true);
      }
    };

    const logoutCompleteCallBackFalse: (ev: Event) => void = (ev: Event) => {
      logoutCompleted(ev as CustomEvent);
    };

    const cleanUpFunction: () => void = () => {
      loginComponent.current.removeEventListener(
        'loginCompleted',
        loginCompleted
      );
      loginComponent.current.removeEventListener(
        'loginCompleted',
        autoLoginCompleted
      );
      loginComponent.current.removeEventListener(
        'logoutCompleted',
        logoutCompleteCallBackTrue
      );
      loginComponent.current.removeEventListener(
        'logoutCompleted',
        logoutCompleteCallBackFalse
      );
    };

    const setUpLoginComponentAsync: () => Promise<void> = async () => {
      try {
        SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
        SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
        loginComponent.current.addEventListener(
          'logoutCompleted',
          logoutCompleteCallBackTrue
        );
        await SPComponentLoader.loadScript(ZONE_JS_URL);
        await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);

        if (window.location.href.includes('autologin')) {
          document.body.style.display = 'none';
          if (loginComponent.current.state !== LoginState.LoggedIn) {
            if (
              !document.referrer.includes('accounts.') &&
              !document.referrer.includes('signin.')
            ) {
              await loginComponent.current.initLoginFlowAsync();
              loginComponent.current.addEventListener(
                'loginCompleted',
                autoLoginCompleted
              );
            } else if (loginComponent.current.state === LoginState.LoggedOut) {
              window.opener.postMessage('loginWindowSuccess', window.origin);
            } else {
              loginComponent.current.addEventListener(
                'loginCompleted',
                autoLoginCompleted
              );
            }
          } else {
            const loginbutton = loginComponent.current.querySelector(
              '.login-button'
            ) as HTMLButtonElement;
            loginbutton.click();
          }
        } else {
          loginComponent.current.addEventListener(
            'loginCompleted',
            loginCompleted
          );
          loginComponent.current.addEventListener('logoutCompleted', (ev) =>
            logoutCompleted(ev as CustomEvent)
          );
          const isLoggedIn: boolean =
            loginComponent.current.state === LoginState.LoggedIn;

          setLoggedIn(isLoggedIn);
          if (isLoggedIn && spFileMetadata) {
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
        }
      } catch (err) {
        console.error(`Unable to initialize sign-in page: ${err}`);
      }
    };

    void setUpLoginComponentAsync();

    return cleanUpFunction;
  }, []);

  function getLoginText(): JSX.Element {
    let loginText: JSX.Element | undefined;
    if (!spFileMetadata) {
      loginText = (
        <>
          <p>{NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE}</p>
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
            <p>
              You are not signed in. You can sign in using the following button.
            </p>
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
      <p>{NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_SAVING_TO_LASERFICHE}</p>;
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

  function clickLogin(): void {
    const url =
      props.context.pageContext.web.absoluteUrl +
      '/SitePages/LaserficheSignIn.aspx?autologin';

    const loginWindow = window.open(url, '_blank', 'popup');
    window.addEventListener('message', (event) => {
      if (event.origin === window.origin) {
        if (event.data === 'loginWindowSuccess') {
          loginWindow.close();
        } else if (event.data) {
          const parsedError: AbortedLoginError = event.data;
          loginWindow.close();
          window.alert(
            `Error retrieving login credentials: ${parsedError.ErrorMessage}. Please try again.`
          );
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
          src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII='
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
    </React.StrictMode>
  );
}
