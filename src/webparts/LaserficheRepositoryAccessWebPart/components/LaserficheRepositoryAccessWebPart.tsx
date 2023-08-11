import * as React from 'react';
import SvgHtmlIcons from '../components/SVGHtmlIcons';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { IRepositoryApiClientExInternal } from '../../../repository-client/repository-client-types';
import { RepositoryClientExInternal } from '../../../repository-client/repository-client';
import { clientId } from '../../constants';
import { NgElement, WithProperties } from '@angular/elements';
import { useEffect, useState } from 'react';
import RepositoryViewComponent from './RepositoryViewWebPart';
require('../../../../node_modules/bootstrap/dist/js/bootstrap.min.js');
require('../../../Assets/CSS/bootstrap.min.css');
require('../../../Assets/CSS/custom.css');
import './LaserficheRepositoryAccess.module.scss';
import { ILaserficheRepositoryAccessWebPartProps } from './ILaserficheRepositoryAccessWebPartProps';
import { getRegion } from '../../../Utils/Funcs';
import { ProblemDetails } from '@laserfiche/lf-repository-api-client';

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

  const redirectPage =
    props.context.pageContext.web.absoluteUrl + props.laserficheRedirectPage;

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
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/zone.js@0.11.4/bundles/zone.umd.min.js'
      );
      await SPComponentLoader.loadScript(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ui-components.js'
      );
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/indigo-pink.css'
      );
      SPComponentLoader.loadCss(
        'https://cdn.jsdelivr.net/npm/@laserfiche/lf-ui-components@14/cdn/lf-ms-office-lite.css'
      );
      const loginCompleted: () => Promise<void> = async () => {
        await getAndInitializeRepositoryClientAndServicesAsync();
        setLoggedIn(true);
      };
      const logoutCompleted: () => Promise<void> = async () => {
        setLoggedIn(false);
        window.location.href =
          props.context.pageContext.web.absoluteUrl +
          props.laserficheRedirectPage;
      };

      loginComponent.current.addEventListener('loginCompleted', loginCompleted);
      loginComponent.current.addEventListener(
        'logoutCompleted',
        logoutCompleted
      );
      if (loginComponent.current.authorization_credentials) {
        await getAndInitializeRepositoryClientAndServicesAsync();
        setLoggedIn(true);
      }
    };

    initializeComponentAsync().catch((err: Error | ProblemDetails) => {
      console.warn(
        `Error: ${(err as Error).message ?? (err as ProblemDetails).title}`
      );
    });
  }, []);

  return (
    <React.StrictMode>
      <div style={{ display: 'none' }}>
        <SvgHtmlIcons />
      </div>
      <div
        className='container-fluid p-3'
        style={{ maxWidth: '100%', marginLeft: '-30px' }}
      >
        <div className='btnSignOut'>
          <lf-login
            redirect_uri={redirectPage}
            redirect_behavior='Replace'
            client_id={clientId}
            authorize_url_host_name={region}
            ref={loginComponent}
          />
        </div>
        <RepositoryViewComponent
          webClientUrl={webClientUrl}
          repoClient={repoClient}
          webPartTitle={props.webPartTitle}
          loggedIn={loggedIn}
        />
      </div>
    </React.StrictMode>
  );
}
