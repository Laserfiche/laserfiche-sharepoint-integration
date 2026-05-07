// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import type { RefObject } from 'react';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent, LoginState } from '@laserfiche/types-lf-ui-components';

/**
 * Polls until the lf-login web component has loaded `authorization_credentials`,
 * the user is definitively logged out, the component unmounts, or `timeoutMs`
 * elapses. Returns `true` if credentials are present when polling ends, `false`
 * otherwise (timeout, unmount, or logged out).
 */
export async function waitForLoginCredentialsAsync(
  loginRef: RefObject<NgElement & WithProperties<LfLoginComponent>>,
  timeoutMs = 3000,
  intervalMs = 50
): Promise<boolean> {
  const start = Date.now();
  while (
    loginRef.current &&
    !loginRef.current.authorization_credentials &&
    loginRef.current.state !== LoginState.LoggedOut &&
    Date.now() - start < timeoutMs
  ) {
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
  return !!loginRef.current?.authorization_credentials;
}
