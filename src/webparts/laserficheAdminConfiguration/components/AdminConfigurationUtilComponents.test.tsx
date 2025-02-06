// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import '@testing-library/jest-dom';
import { render, screen } from '@testing-library/react';
import { LoggedOutMessageWrapper } from './AdminConfigurationUtilComponents';
import { PLEASE_LOGIN_TO_LASERFICHE } from '../../strings';

const loggedInTestText = 'Logged in';

describe('LoggedOutMessageWrapper', () => {
  test('renders component when logged in', () => {
    // Arrange

    //Act
    render(
      <LoggedOutMessageWrapper loggedIn={true}>
        <span>{loggedInTestText}</span>
      </LoggedOutMessageWrapper>
    );

    // Assert
    const loggedOutElement = screen.queryByText(PLEASE_LOGIN_TO_LASERFICHE);
    expect(loggedOutElement).not.toBeInTheDocument();
    const loggedInElement = screen.getByText(loggedInTestText);
    expect(loggedInElement).toBeInTheDocument();
  });

  test('renders logged out message when logged out', () => {
    // Arrange

    // Act
    render(
      <LoggedOutMessageWrapper loggedIn={false}>
        <span>{loggedInTestText}</span>
      </LoggedOutMessageWrapper>
    );

    // Assert
    // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
    const regex = new RegExp(PLEASE_LOGIN_TO_LASERFICHE);
    const loggedOutElement = screen.getByText(regex);
    expect(loggedOutElement).toBeInTheDocument();
    const loggedInElement = screen.queryByText(loggedInTestText);
    expect(loggedInElement).not.toBeInTheDocument();
  });
});
