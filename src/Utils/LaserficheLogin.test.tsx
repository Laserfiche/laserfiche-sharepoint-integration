// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { render } from '@testing-library/react';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';
import { LaserficheLogin } from './LaserficheLogin';
import { clientId, repositoryScopes } from '../webparts/constants';

const REDIRECT_URI = 'https://contoso.sharepoint.com/SitePages/LaserficheSignIn.aspx?autologin';

afterEach(() => {
  window.localStorage.clear();
});

describe('LaserficheLogin', () => {
  test("renders a hidden lf-login element set up for this app's Laserfiche Cloud sign-in", () => {
    // Act
    const { container } = render(<LaserficheLogin redirectUri={REDIRECT_URI} />);

    // Assert
    const lfLogin = container.querySelector('lf-login');
    expect(lfLogin).toHaveAttribute('client_id', clientId);
    expect(lfLogin).toHaveAttribute('scope', repositoryScopes);
    expect(lfLogin).toHaveAttribute('redirect_uri', REDIRECT_URI);
    expect(lfLogin).toHaveAttribute('redirect_behavior', 'Replace');
    expect(lfLogin).toHaveAttribute('authorize_url_host_name', 'laserfiche.com');
    expect(lfLogin).toHaveAttribute('hidden');
  });

  // lf-login only creates the provider it refreshes tokens with once it knows
  // its login type. Without one, an expired token signs the user out instead.
  test('tells lf-login it is a Laserfiche Cloud sign-in, so it can refresh an expired token', () => {
    // Act
    const { container } = render(<LaserficheLogin redirectUri={REDIRECT_URI} />);

    // Assert
    expect(container.querySelector('lf-login')).toHaveAttribute('login_type', 'Cloud');
  });

  // Laserfiche employees use the clouddev environment when spDevMode is on.
  test('signs in against clouddev in dev mode', () => {
    // Arrange
    window.localStorage.setItem('spDevMode', 'true');

    // Act
    const { container } = render(<LaserficheLogin redirectUri={REDIRECT_URI} />);

    // Assert
    expect(container.querySelector('lf-login')).toHaveAttribute(
      'authorize_url_host_name',
      'a.clouddev.laserfiche.com'
    );
  });

  test('hands its ref the lf-login element', () => {
    // Arrange
    const ref = React.createRef<NgElement & WithProperties<LfLoginComponent>>();

    // Act
    const { container } = render(<LaserficheLogin ref={ref} redirectUri={REDIRECT_URI} />);

    // Assert
    expect(ref.current).toBe(container.querySelector('lf-login'));
  });
});
