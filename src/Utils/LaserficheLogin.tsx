// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent, LoginType } from '@laserfiche/types-lf-ui-components';
import { clientId, repositoryScopes } from '../webparts/constants';
import { getRegion } from './Funcs';

/**
 * The hidden lf-login element every page and dialog signs in through. Its
 * setup lives here once, so each copy signs in to the same Laserfiche Cloud
 * app, with the same scopes, and shares one session.
 *
 * login_type matters beyond the default it matches: lf-login only creates the
 * provider it refreshes tokens with once it is told its login type.
 *
 * @param props.redirectUri Where the Laserfiche sign-in page sends the browser back to
 */
export const LaserficheLogin = React.forwardRef<
  NgElement & WithProperties<LfLoginComponent>,
  { redirectUri: string }
>(function LaserficheLogin(props, ref) {
  return (
    <lf-login
      ref={ref}
      redirect_uri={props.redirectUri}
      authorize_url_host_name={getRegion()}
      redirect_behavior='Replace'
      client_id={clientId}
      scope={repositoryScopes}
      login_type={LoginType.Cloud}
      hidden
    />
  );
});
