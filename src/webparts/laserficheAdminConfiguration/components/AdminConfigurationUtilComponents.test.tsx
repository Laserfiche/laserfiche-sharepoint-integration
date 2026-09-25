// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import type { MockInstance } from 'vitest';
import * as React from 'react';
import { act, render, screen } from '@testing-library/react';
import {
  LoggedOutMessageWrapper,
  LoginComponent,
} from './AdminConfigurationUtilComponents';
import { PLEASE_LOGIN_TO_LASERFICHE } from '../../strings';
// By path, not by package name: vitest.config.mts's alias would resolve the
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
  function messageCalls(spy: MockInstance): unknown[][] {
    return spy.mock.calls.filter((call) => call[0] === 'message');
  }

  function renderLoginComponent(): ReturnType<typeof render> {
    return render(
      <LoginComponent
        loggedIn={false}
        setLoggedIn={vi.fn()}
        setMessageErrorModal={vi.fn()}
        ensureRepoClientInitializedAsync={vi.fn()}
        context={mockWebPartContext}
      />
    );
  }

  test('removes its popup message listener on unmount', () => {
    // Arrange
    const addSpy = vi.spyOn(window, 'addEventListener');
    const removeSpy = vi.spyOn(window, 'removeEventListener');
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
    const addSpy = vi.spyOn(window, 'addEventListener');
    const removeSpy = vi.spyOn(window, 'removeEventListener');

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

  // Lets the mount effect get past its (mocked) loadScript await and register
  // its lf-login listeners, and lets a dispatched listener run to completion.
  async function flushAsync(): Promise<void> {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
  }

  test('initializes the repository client on sign-in after a re-render', async () => {
    // Arrange
    const ensureRepoClientInitializedAsync = vi
      .fn()
      .mockResolvedValue(undefined);
    const setLoggedIn = vi.fn();
    const loginComponent = (): JSX.Element => (
      <LoginComponent
        loggedIn={false}
        setLoggedIn={setLoggedIn}
        setMessageErrorModal={vi.fn()}
        ensureRepoClientInitializedAsync={ensureRepoClientInitializedAsync}
        context={mockWebPartContext}
      />
    );
    const { container, rerender } = render(loginComponent());
    await flushAsync();
    // Any parent re-render (an error dialog opening, say) re-renders this
    // component before the user signs in.
    rerender(loginComponent());
    const lfLogin = container.querySelector('lf-login') as HTMLElement & {
      authorization_credentials?: { accessToken: string };
    };
    lfLogin.authorization_credentials = { accessToken: 'token' };

    // Act
    await act(async () => {
      lfLogin.dispatchEvent(new CustomEvent('loginCompleted'));
    });
    await flushAsync();

    // Assert
    expect(ensureRepoClientInitializedAsync).toHaveBeenCalledTimes(1);
    expect(setLoggedIn).toHaveBeenCalledWith(true);
  });

  test('does not log an error when unmounted before lf-ui-components loads', async () => {
    // Arrange
    const consoleErrorSpy = vi
      .spyOn(console, 'error')
      .mockImplementation(() => undefined);

    // Act
    renderLoginComponent().unmount();
    await flushAsync();

    // Assert
    expect(consoleErrorSpy).not.toHaveBeenCalled();

    consoleErrorSpy.mockRestore();
  });
});
