// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent, LoginState } from '@laserfiche/types-lf-ui-components';

export async function waitForLoginCredentialsAsync(
  loginRef: React.RefObject<NgElement & WithProperties<LfLoginComponent>>,
  timeoutMs = 3000,
  intervalMs = 50
): Promise<void> {
  const start = Date.now();
  while (
    loginRef.current &&
    !loginRef.current.authorization_credentials &&
    loginRef.current.state !== LoginState.LoggedOut &&
    Date.now() - start < timeoutMs
  ) {
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}
