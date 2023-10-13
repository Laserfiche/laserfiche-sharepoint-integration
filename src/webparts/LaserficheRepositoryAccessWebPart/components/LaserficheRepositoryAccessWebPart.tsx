import * as React from 'react';
import SvgHtmlIcons from '../components/SVGHtmlIcons';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  AbortedLoginError,
  LfLoginComponent,
} from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import {
  clientId,
  LF_INDIGO_PINK_CSS_URL,
  LF_MS_OFFICE_LITE_CSS_URL,
  LF_UI_COMPONENTS_URL,
  ZONE_JS_URL,
} from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { useEffect, useState } from 'react';
import RepositoryViewComponent from './RepositoryViewWebPart';
require('../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');
require('../../../Assets/CSS/bootstrap.min.css');
import './LaserficheRepositoryAccess.module.scss';
import { ILaserficheRepositoryAccessWebPartProps } from './ILaserficheRepositoryAccessWebPartProps';
import { getRegion } from '../../../Utils/Funcs';
import styles from './LaserficheRepositoryAccess.module.scss';

declare global {
  // eslint-disable-next-line
  namespace JSX {
    interface IntrinsicElements {
      // eslint-disable-next-line
      ['lf-field-container']: any;
      // eslint-disable-next-line
      ['lf-login']: any;
    }
  }
}

export default function LaserficheRepositoryAccessWebPart(
  props: ILaserficheRepositoryAccessWebPartProps
): JSX.Element {
  const [webClientUrl, setWebClientUrl] = React.useState('');
  const loginComponent: React.RefObject<
    NgElement & WithProperties<LfLoginComponent>
  > = React.useRef();
  const [loggedIn, setLoggedIn] = useState<boolean>(false);
  const [repoClient, setRepoClient] = useState<
    IRepositoryApiClientExInternal | undefined
  >(undefined);

  const region = getRegion();

  const redirectPage = window.location.origin + window.location.pathname;

  useEffect(() => {
    const ensureRepoClientInitializedAsync: () => Promise<void> = async () => {
      if (!repoClient) {
        const repoClientCreator = new RepositoryClientExInternal();
        const repoClient =
          await repoClientCreator.createRepositoryClientAsync();
        setRepoClient(repoClient);
      }
    };

    const getAndInitializeRepositoryClientAndServicesAsync: () => Promise<void> =
      async () => {
        const accessToken =
          loginComponent?.current?.authorization_credentials?.accessToken;
        setWebClientUrl(
          loginComponent?.current?.account_endpoints.webClientUrl
        );
        if (accessToken) {
          await ensureRepoClientInitializedAsync();
        } else {
          // user is not logged in
        }
      };

    const initializeComponentAsync: () => Promise<void> = async () => {
      try {
        await SPComponentLoader.loadScript(ZONE_JS_URL);
        await SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL);
        SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL);
        SPComponentLoader.loadCss(LF_MS_OFFICE_LITE_CSS_URL);
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
        console.error(`Unable to initialize repository explorer: ${err}`);
      }
    };

    void initializeComponentAsync();
  }, []);

  function clickLogin(): void {
    const url =
      props.context.pageContext.web.absoluteUrl +
      '/SitePages/LaserficheSignIn.aspx?autologin';
    const loginWindow = window.open(url, 'loginWindow', 'popup');
    loginWindow.resizeTo(800, 600);
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
            authorize_url_host_name={region}
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
        <RepositoryViewComponent
          webClientUrl={webClientUrl}
          repoClient={repoClient}
          loggedIn={loggedIn}
        />
      </div>
    </React.StrictMode>
  );
}
