// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import '@testing-library/jest-dom';
import { render, screen } from '@testing-library/react';
import {
  LoggedOutMessageWrapper,
  LoginComponent,
} from './AdminConfigurationUtilComponents';
import { PLEASE_LOGIN_TO_LASERFICHE } from '../../strings';
// By path, not by package name: jest's moduleNameMapper would resolve the
// bare specifier to this same file, but tsc would resolve the real package.
import mockWebPartContext from '../../../__mocks__/@microsoft/sp-webpart-base';

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

describe('LoginComponent', () => {
  function messageCalls(spy: jest.SpyInstance): unknown[][] {
    return spy.mock.calls.filter((call) => call[0] === 'message');
  }

  function renderLoginComponent(): ReturnType<typeof render> {
    return render(
      <LoginComponent
        loggedIn={false}
        setLoggedIn={jest.fn()}
        setMessageErrorModal={jest.fn()}
        ensureRepoClientInitializedAsync={jest.fn()}
        context={mockWebPartContext}
      />
    );
  }

  test('removes its popup message listener on unmount', () => {
    // Arrange
    const addSpy = jest.spyOn(window, 'addEventListener');
    const removeSpy = jest.spyOn(window, 'removeEventListener');
    const { unmount } = renderLoginComponent();
    const added = messageCalls(addSpy);
    expect(added).toHaveLength(1);

    // Act
    unmount();

    // Assert
    // The same reference must come back off the window, or the handler leaks
    // and a remount stacks another one on top of it.
    const removed = messageCalls(removeSpy);
    expect(removed).toHaveLength(1);
    expect(removed[0][1]).toBe(added[0][1]);

    addSpy.mockRestore();
    removeSpy.mockRestore();
  });

  test('does not accumulate listeners across remounts', () => {
    // Arrange
    const addSpy = jest.spyOn(window, 'addEventListener');
    const removeSpy = jest.spyOn(window, 'removeEventListener');

    // Act
    renderLoginComponent().unmount();
    renderLoginComponent().unmount();

    // Assert
    const added = messageCalls(addSpy);
    const removed = messageCalls(removeSpy);
    expect(added).toHaveLength(2);
    expect(removed).toHaveLength(2);

    addSpy.mockRestore();
    removeSpy.mockRestore();
  });
});
