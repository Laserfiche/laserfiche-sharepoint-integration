// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { LoggedOutMessageWrapper, LoginComponent } from './AdminConfigurationUtilComponents';
import { LASERFICHE_SIGNIN_PAGE_NAME, LOGIN_WINDOW_SUCCESS } from '../../constants';
import { PLEASE_LOGIN_TO_LASERFICHE, SIGN_OUT } from '../../strings';
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

  // lf-login only creates the provider it refreshes tokens with once it knows
  // its login type. Without one, an expired token signs the user out instead.
  test('renders lf-login as a Laserfiche Cloud sign-in, so it can refresh an expired token', () => {
    // Act
    const { container } = renderLoginComponent();

    // Assert
    expect(container.querySelector('lf-login')).toHaveAttribute('login_type', 'Cloud');
  });

  // The popup signs out on its own lf-login element, so nothing else tells
  // this page that the user signed out.
  test('a sign-out in the popup signs the page out', async () => {
    // Arrange
    globalThis.fetch = vi.fn().mockResolvedValue({
      json: () => Promise.resolve({ value: [{ Title: LASERFICHE_SIGNIN_PAGE_NAME }] }),
    }) as unknown as typeof fetch;
    const openSpy = vi
      .spyOn(window, 'open')
      .mockReturnValue({ close: vi.fn() } as unknown as Window);
    const setLoggedIn = vi.fn();
    render(
      <LoginComponent
        loggedIn={true}
        setLoggedIn={setLoggedIn}
        setMessageErrorModal={vi.fn()}
        ensureRepoClientInitializedAsync={vi.fn()}
        context={mockWebPartContext}
      />
    );
    fireEvent.click(screen.getByRole('button', { name: SIGN_OUT }));
    await waitFor(() => expect(openSpy).toHaveBeenCalledTimes(1));

    // Act
    await act(async () => {
      window.dispatchEvent(
        new MessageEvent('message', { origin: window.origin, data: LOGIN_WINDOW_SUCCESS })
      );
    });

    // Assert
    expect(setLoggedIn).toHaveBeenCalledWith(false);

    openSpy.mockRestore();
    delete (globalThis as { fetch?: unknown }).fetch;
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
    const ensureRepoClientInitializedAsync = vi.fn().mockResolvedValue(undefined);
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
    const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => undefined);

    // Act
    renderLoginComponent().unmount();
    await flushAsync();

    // Assert
    expect(consoleErrorSpy).not.toHaveBeenCalled();

    consoleErrorSpy.mockRestore();
  });
});
