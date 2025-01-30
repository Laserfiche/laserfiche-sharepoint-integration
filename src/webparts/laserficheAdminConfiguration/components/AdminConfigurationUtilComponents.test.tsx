// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

jest.mock('@microsoft/sp-loader', () => {
  return {
    SPComponentLoader: {
      loadCss: jest.fn(),
      loadScript: jest.fn(),
    },
  };
});

import * as React from 'react';
import '@testing-library/jest-dom';
import { render, screen } from '@testing-library/react';
import { LoggedOutMessageWrapper } from './AdminConfigurationUtilComponents';

const loggedOutText =
  /Please login to Laserfiche in order to use this web part./i;
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
    const loggedOutElement = screen.queryByText(loggedOutText);
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
    const loggedOutElement = screen.getByText(loggedOutText);
    expect(loggedOutElement).toBeInTheDocument();
    const loggedInElement = screen.queryByText(loggedInTestText);
    expect(loggedInElement).not.toBeInTheDocument();
  });
});
