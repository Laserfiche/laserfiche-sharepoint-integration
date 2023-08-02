import * as React from 'react';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Navigation } from 'spfx-navigation';
import {
  LfLoginComponent,
  LoginState,
} from '@laserfiche/types-lf-ui-components';
import { clientId, SP_LOCAL_STORAGE_KEY } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';
import { ISPDocumentData } from '../../../Utils/Types';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';
import { getRegion } from '../../../Utils/Funcs';
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

const SIGN_IN = 'Sign In';
const SIGN_OUT = 'Sign Out';
const CANCEL = 'Cancel';
const NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_ATTEMPTING_TO_SAVE_AND_NOT_LOGGED_IN =
  '*Note: This page should only be needed if you are attempting to save a document to Laserfiche and not signed in.';

const DOCUMENT_SAVED_TO_LASERFICHE_SUCCESSFULLY =
  'The document was saved to Laserfiche successfully.';
export default function SendToLaserficheLoginComponent(
  props: ISendToLaserficheLoginComponentProps
) {
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();

  const [loggedIn, setLoggedIn] = React.useState<boolean>(false);

  const region = getRegion();

  const spFileMetadata = JSON.parse(
    window.localStorage.getItem(SP_LOCAL_STORAGE_KEY)
  ) as ISPDocumentData;

  const [loginText, setLoginText] = React.useState<JSX.Element | undefined>(
    <p>
      {NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_ATTEMPTING_TO_SAVE_AND_NOT_LOGGED_IN}
    </p>
  );

  React.useEffect(() => {
    const setUpLoginComponent = async () => {
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
      );
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
      );
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
      );
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      );
      loginComponent.current.addEventListener('loginCompleted', loginCompleted);
      loginComponent.current.addEventListener(
        'logoutCompleted',
        logoutCompleted
      );

      const isLoggedIn: boolean =
        loginComponent.current.state === LoginState.LoggedIn;

      setLoggedIn(isLoggedIn);
      if (!spFileMetadata) {
        setLoginText(
          <>
            <p>
              {
                NOTE_THIS_PAGE_ONLY_NEEDED_WHEN_ATTEMPTING_TO_SAVE_AND_NOT_LOGGED_IN
              }
            </p>
            {isLoggedIn ? (
              <p>
                {'Welcome to Laserfiche. Go to '}
                <a href={loginComponent.current.account_endpoints.webClientUrl}>
                  your Laserfiche repository
                </a>
              </p>
            ) : (
              <p>
                You are not signed in. You can sign in using the button below.
              </p>
            )}
          </>
        );
      } else if (spFileMetadata?.fileUrl && !loggedIn) {
        setLoginText(
          <>
            <div>
              You are not signed in. Please sign in to continue saving{' '}
              {spFileMetadata?.fileName}.
            </div>
            <br />
          </>
        );
      } else if (spFileMetadata?.fileUrl && loggedIn) {
        setLoginText(
          <>
            <div>
              You are now signed in. Attempting to save{' '}
              {spFileMetadata?.fileName}.
            </div>
            <br />
          </>
        );
      }
      if (isLoggedIn && spFileMetadata) {
        const dialog = new SaveToLaserficheCustomDialog(spFileMetadata, true);

        await dialog.show();
        if (!dialog.successful) {
          console.warn('Could not login successfully');
        } else {
          // TODO should be able to view in Laserfiche or SharePoint?
          setLoginText(<p>{DOCUMENT_SAVED_TO_LASERFICHE_SUCCESSFULLY}</p>);
        }
      }
    };

    setUpLoginComponent();
  }, []);

  const loginCompleted = async () => {
    setLoggedIn(true);
    if (spFileMetadata) {
      const dialog = new SaveToLaserficheCustomDialog(spFileMetadata, true);
      await dialog.show();
      if (!dialog.successful) {
        console.warn('Could not login successfully');
      } else {
        // TODO should be able to view in Laserfiche or SharePoint?
        setLoginText(<p>{DOCUMENT_SAVED_TO_LASERFICHE_SUCCESSFULLY}</p>);
      }
    }
  };

  const logoutCompleted = () => {
    setLoggedIn(false);
    window.location.href =
      props.context.pageContext.web.absoluteUrl + props.laserficheRedirectUrl;
  };

  function redirect() {
    const spFileUrl = spFileMetadata.fileUrl;
    const fileNameWithExtension = spFileMetadata.fileName;
    const spFileUrlWithoutFileName = spFileUrl.replace(
      fileNameWithExtension,
      ''
    );
    const path = window.location.origin + spFileUrlWithoutFileName;
    Navigation.navigate(path, true);
  }

  return (
    <div>
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
          redirect_uri={
            props.context.pageContext.web.absoluteUrl +
            props.laserficheRedirectUrl
          }
          authorize_url_host_name={region}
          redirect_behavior='Replace'
          client_id={clientId}
          sign_in_text={SIGN_IN}
          sign_out_text={SIGN_OUT}
          ref={loginComponent}
        />
        <br />
        {spFileMetadata?.fileUrl && (
          <button className='lf-button sec-button' onClick={redirect}>
            {CANCEL}
          </button>
        )}
      </div>
    </div>
  );
}
