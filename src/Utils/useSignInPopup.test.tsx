// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import type { Mock } from 'vitest';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import * as React from 'react';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useSignInPopup } from './useSignInPopup';
import { LASERFICHE_SIGNIN_PAGE_NAME, LOGIN_WINDOW_SUCCESS } from '../webparts/constants';
import {
  POPUP_BLOCKED,
  SIGN_IN,
  SIGN_IN_FAILED,
  SIGN_IN_FAILED_PLEASE_TRY_AGAIN_DETAILS,
  SIGN_OUT,
  needLaserficheSignInPage,
} from '../webparts/strings';

const SITE_URL = 'https://contoso.sharepoint.com/sites/team';
const SIGN_IN_PAGE_URL = `${SITE_URL}/SitePages/${LASERFICHE_SIGNIN_PAGE_NAME}.aspx`;

const mockContext = {
  pageContext: { web: { absoluteUrl: SITE_URL } },
} as unknown as BaseComponentContext;

// The smallest page that uses the hook: a Sign in / Sign out button and
// whatever dialog the hook asks it to show.
function SignInButton(props: {
  loggedIn: boolean;
  onSignedOut?: () => void;
  onSignedIn?: () => Promise<void>;
}): JSX.Element {
  const [message, setMessage] = React.useState<JSX.Element | undefined>(undefined);
  const signInOrOutAsync = useSignInPopup({
    context: mockContext,
    loggedIn: props.loggedIn,
    setMessageModal: setMessage,
    onSignedOut: props.onSignedOut ?? (() => undefined),
    onSignedIn: props.onSignedIn,
  });
  return (
    <>
      <button onClick={signInOrOutAsync}>{props.loggedIn ? SIGN_OUT : SIGN_IN}</button>
      {message}
    </>
  );
}

function mockSitePages(titles: string[]): void {
  globalThis.fetch = vi.fn().mockResolvedValue({
    json: () => Promise.resolve({ value: titles.map((Title) => ({ Title })) }),
  }) as unknown as typeof fetch;
}

function mockPopup(): { close: Mock } {
  const popup = { close: vi.fn() };
  vi.spyOn(window, 'open').mockReturnValue(popup as unknown as Window);
  return popup;
}

async function clickAndWaitForPopupAsync(name: string): Promise<void> {
  fireEvent.click(screen.getByRole('button', { name }));
  await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));
}

function popupUrl(): string {
  return (window.open as Mock).mock.calls[0][0];
}

async function postFromPopupAsync(data: unknown, origin: string = window.origin): Promise<void> {
  await act(async () => {
    window.dispatchEvent(new MessageEvent('message', { origin, data }));
  });
}

beforeEach(() => {
  mockSitePages([LASERFICHE_SIGNIN_PAGE_NAME]);
});

afterEach(() => {
  vi.restoreAllMocks();
  window.sessionStorage.clear();
  delete (globalThis as { fetch?: unknown }).fetch;
});

describe('useSignInPopup', () => {
  test('opens the sign-in page in a popup, asking it to sign in', async () => {
    // Arrange
    mockPopup();
    render(<SignInButton loggedIn={false} />);

    // Act
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Assert
    expect(popupUrl()).toBe(`${SIGN_IN_PAGE_URL}?autologin&action=login`);
  });

  test('asks the popup to sign out when the page is signed in', async () => {
    // Arrange
    mockPopup();
    render(<SignInButton loggedIn={true} />);

    // Act
    await clickAndWaitForPopupAsync(SIGN_OUT);

    // Assert
    expect(popupUrl()).toBe(`${SIGN_IN_PAGE_URL}?autologin&action=logout`);
  });

  test('carries the debug manifests into the popup during a local debug session', async () => {
    // Arrange
    mockPopup();
    const manifestsFileUrl = 'https://localhost:4321/temp/build/manifests.js';
    window.sessionStorage.setItem('spfx-debug', JSON.stringify({ manifestsFileUrl }));
    render(<SignInButton loggedIn={false} />);

    // Act
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Assert
    expect(popupUrl()).toBe(
      `${SIGN_IN_PAGE_URL}?autologin&action=login&debugManifestsFile=${encodeURIComponent(
        manifestsFileUrl
      )}&loadSPFX=true&debug=true&noredir=true`
    );
  });

  test('ignores a debug session entry that is not valid JSON', async () => {
    // Arrange
    mockPopup();
    window.sessionStorage.setItem('spfx-debug', '{not valid json');
    render(<SignInButton loggedIn={false} />);

    // Act
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Assert
    expect(popupUrl()).toBe(`${SIGN_IN_PAGE_URL}?autologin&action=login`);
  });

  test('explains a missing sign-in page instead of opening a popup', async () => {
    // Arrange
    mockSitePages(['Home']);
    vi.spyOn(window, 'open');
    render(<SignInButton loggedIn={false} />);

    // Act
    fireEvent.click(screen.getByRole('button', { name: SIGN_IN }));

    // Assert
    expect(await screen.findByText(needLaserficheSignInPage)).toBeInTheDocument();
    expect(screen.getByText(SIGN_IN_FAILED)).toBeInTheDocument();
    expect(window.open).not.toHaveBeenCalled();
  });

  test('closes the dialog when its Okay button is clicked', async () => {
    // Arrange
    mockSitePages(['Home']);
    render(<SignInButton loggedIn={false} />);
    fireEvent.click(screen.getByRole('button', { name: SIGN_IN }));
    await screen.findByText(needLaserficheSignInPage);

    // Act
    fireEvent.click(screen.getByRole('button', { name: 'Okay' }));

    // Assert
    expect(screen.queryByText(needLaserficheSignInPage)).not.toBeInTheDocument();
  });

  test('explains a blocked pop-up instead of throwing', async () => {
    // Arrange
    vi.spyOn(window, 'open').mockReturnValue(null);
    render(<SignInButton loggedIn={false} />);

    // Act
    fireEvent.click(screen.getByRole('button', { name: SIGN_IN }));

    // Assert
    expect(await screen.findByText(POPUP_BLOCKED)).toBeInTheDocument();
  });

  test('closes the popup and reports a sign-in when the popup signs in', async () => {
    // Arrange
    const popup = mockPopup();
    const onSignedIn = vi.fn().mockResolvedValue(undefined);
    const onSignedOut = vi.fn();
    render(<SignInButton loggedIn={false} onSignedIn={onSignedIn} onSignedOut={onSignedOut} />);
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Act
    await postFromPopupAsync(LOGIN_WINDOW_SUCCESS);

    // Assert
    expect(popup.close).toHaveBeenCalledTimes(1);
    expect(onSignedIn).toHaveBeenCalledTimes(1);
    expect(onSignedOut).not.toHaveBeenCalled();
  });

  // The popup signs out on its own lf-login element, so the page's element
  // doesn't necessarily raise logoutCompleted: the hook has to report it.
  test('closes the popup and reports a sign-out when the popup signs out', async () => {
    // Arrange
    const popup = mockPopup();
    const onSignedIn = vi.fn().mockResolvedValue(undefined);
    const onSignedOut = vi.fn();
    render(<SignInButton loggedIn={true} onSignedIn={onSignedIn} onSignedOut={onSignedOut} />);
    await clickAndWaitForPopupAsync(SIGN_OUT);

    // Act
    await postFromPopupAsync(LOGIN_WINDOW_SUCCESS);

    // Assert
    expect(popup.close).toHaveBeenCalledTimes(1);
    expect(onSignedOut).toHaveBeenCalledTimes(1);
    expect(onSignedIn).not.toHaveBeenCalled();
  });

  test('closes the popup and shows the details when sign-in fails', async () => {
    // Arrange
    const popup = mockPopup();
    render(<SignInButton loggedIn={false} />);
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Act
    await postFromPopupAsync({ ErrorType: 'Login Error', ErrorMessage: 'details here' });

    // Assert
    expect(popup.close).toHaveBeenCalledTimes(1);
    expect(
      screen.getByText(`${SIGN_IN_FAILED_PLEASE_TRY_AGAIN_DETAILS} details here`)
    ).toBeInTheDocument();
    expect(screen.getByText(SIGN_IN_FAILED)).toBeInTheDocument();
  });

  test('ignores messages from other origins', async () => {
    // Arrange
    const popup = mockPopup();
    const onSignedIn = vi.fn().mockResolvedValue(undefined);
    render(<SignInButton loggedIn={false} onSignedIn={onSignedIn} />);
    await clickAndWaitForPopupAsync(SIGN_IN);

    // Act
    await postFromPopupAsync(LOGIN_WINDOW_SUCCESS, 'https://evil.example.com');
    await postFromPopupAsync(
      { ErrorType: 'Login Error', ErrorMessage: 'details here' },
      'https://evil.example.com'
    );

    // Assert
    expect(popup.close).not.toHaveBeenCalled();
    expect(onSignedIn).not.toHaveBeenCalled();
    expect(screen.queryByText(SIGN_IN_FAILED)).not.toBeInTheDocument();
  });

  // The message listener is registered once, on mount, so it must not keep
  // calling the callbacks it was first given.
  test('reports to the callbacks of the latest render', async () => {
    // Arrange
    mockPopup();
    const firstOnSignedOut = vi.fn();
    const latestOnSignedOut = vi.fn();
    const { rerender } = render(<SignInButton loggedIn={true} onSignedOut={firstOnSignedOut} />);
    rerender(<SignInButton loggedIn={true} onSignedOut={latestOnSignedOut} />);
    await clickAndWaitForPopupAsync(SIGN_OUT);

    // Act
    await postFromPopupAsync(LOGIN_WINDOW_SUCCESS);

    // Assert
    expect(latestOnSignedOut).toHaveBeenCalledTimes(1);
    expect(firstOnSignedOut).not.toHaveBeenCalled();
  });

  test('removes its message listener on unmount', () => {
    // Arrange
    const addSpy = vi.spyOn(window, 'addEventListener');
    const removeSpy = vi.spyOn(window, 'removeEventListener');
    const { unmount } = render(<SignInButton loggedIn={false} />);
    const added = addSpy.mock.calls.filter(([type]) => type === 'message');
    expect(added).toHaveLength(1);

    // Act
    unmount();

    // Assert
    // The same reference must come back off the window, or the handler leaks
    // and a remount stacks another one on top of it.
    const removed = removeSpy.mock.calls.filter(([type]) => type === 'message');
    expect(removed).toHaveLength(1);
    expect(removed[0][1]).toBe(added[0][1]);
  });
});
