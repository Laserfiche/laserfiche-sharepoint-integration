// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('spfx-navigation', () => ({ Navigation: { navigate: vi.fn() } }));
vi.mock('@laserfiche/types-lf-ui-components', () => ({
  LoginState: { LoggedIn: 'LoggedIn', LoggedOut: 'LoggedOut' },
  LoginType: { Cloud: 'Cloud' },
}));
vi.mock('../../../extensions/savetoLaserfiche/SaveToLaserficheDialog', () => ({
  default: vi.fn().mockImplementation(function () {
    return { show: vi.fn().mockResolvedValue(undefined), successful: true };
  }),
}));
vi.mock('../../../Utils/Funcs', async (importOriginal) => ({
  getEntryWebAccessUrl: vi.fn().mockReturnValue('https://webclient.example.com'),
  getRegion: vi.fn().mockReturnValue('a.clouddev.laserfiche.com'),
  getSPDocumentDataFromLocalStorage: vi.fn(), // configure return value per test
  getSPListURL: vi
    .fn()
    .mockReturnValue("https://contoso.sharepoint.com/_api/web/lists/getbytitle('Site Pages')"),
  // The real one, so the popup tests below can keep asserting on window.open.
  openLoginWindow: (await importOriginal<typeof import('../../../Utils/Funcs')>()).openLoginWindow,
}));

import type { Mock } from 'vitest';
import * as React from 'react';
import { render, screen, fireEvent, waitFor, act } from '@testing-library/react';
import SendToLaserficheLoginComponent from './SendToLaserficheLoginComponent';
import { Navigation } from 'spfx-navigation';
import { getSPDocumentDataFromLocalStorage } from '../../../Utils/Funcs';
import SaveToLaserficheCustomDialog from '../../../extensions/savetoLaserfiche/SaveToLaserficheDialog';
import {
  LASERFICHE_SIGNIN_PAGE_NAME,
  LOGIN_WINDOW_SUCCESS,
  SP_LOCAL_STORAGE_KEY,
} from '../../constants';
import { ISendToLaserficheLoginComponentProps } from './ISendToLaserficheLoginComponentProps';

// Fake replacement for the real Angular-Elements-based lf-login custom
// element. Registered once at module scope: customElements.define throws if
// the same tag name is defined twice in one test process. Each render()
// creates a brand-new instance (and therefore brand-new vi.fn() fields,
// since they are class fields, not statics), so no manual reset is needed
// between tests beyond grabbing a fresh handle to the new element.
class FakeLfLogin extends HTMLElement {
  state = 'LoggedOut';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  authorization_credentials: any = undefined;
  account_endpoints = {
    webClientUrl: 'https://webclient.example.com',
    regionalDomain: 'a.clouddev.laserfiche.com',
  };
  account_id = 'account-1';
  initLoginFlowAsync = vi.fn().mockResolvedValue(undefined);
  refreshTokenAsync = vi.fn();
}
customElements.define('lf-login', FakeLfLogin);

const mockContext = {
  pageContext: {
    web: { absoluteUrl: 'https://contoso.sharepoint.com' },
  },
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
} as any as ISendToLaserficheLoginComponentProps['context'];

function getLfLogin(container: HTMLElement): FakeLfLogin {
  return container.querySelector('lf-login') as unknown as FakeLfLogin;
}

function mockFetchSitePages(hasSignInPage: boolean): Mock {
  const mockFetch = vi.fn().mockResolvedValue({
    json: () =>
      Promise.resolve({
        value: hasSignInPage
          ? [{ Title: LASERFICHE_SIGNIN_PAGE_NAME }]
          : [{ Title: 'SomeOtherPage' }],
      }),
  });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (globalThis as any).fetch = mockFetch;
  return mockFetch;
}

// Flushes the microtask queue a handful of times. Used instead of waitFor()
// in the fake-timer popup tests, since mixing RTL's waitFor polling with
// fake timers is fragile; the awaited work here is plain promise
// microtasks (loadScript/handleLoginOrLogoutInPopupAsync), not timers.
async function flushMicrotasks(times = 6): Promise<void> {
  for (let i = 0; i < times; i++) {
    // eslint-disable-next-line no-await-in-loop
    await act(async () => {
      await Promise.resolve();
    });
  }
}

beforeEach(() => {
  vi.clearAllMocks();
  (getSPDocumentDataFromLocalStorage as Mock).mockReturnValue(undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.useRealTimers();
  window.sessionStorage.clear();
  window.localStorage.clear();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (window as any).opener = null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  delete (globalThis as any).fetch;
});

describe('SendToLaserficheLoginComponent - main window path', () => {
  test('renders not-signed-in copy with a link to laserfiche.com and a "Sign in" button when logged out with no pending file', async () => {
    render(<SendToLaserficheLoginComponent context={mockContext} />);

    expect(await screen.findByRole('button', { name: 'Sign in' })).toBeInTheDocument();
    expect(
      screen.getByText('You are not signed in. You can sign in using the following button.')
    ).toBeInTheDocument();
    const link = screen.getByRole('link', { name: 'laserfiche.com' });
    expect(link).toHaveAttribute('href', 'https://www.laserfiche.com/products/pricing');
  });

  test('renders welcome copy with a repository link and a "Sign out" button when already logged in', async () => {
    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    getLfLogin(container).state = 'LoggedIn';

    expect(await screen.findByRole('button', { name: 'Sign out' })).toBeInTheDocument();
    expect(screen.getByText(/Welcome to Laserfiche/)).toBeInTheDocument();
    const link = screen.getByRole('link', {
      name: 'your Laserfiche repository',
    });
    expect(link).toHaveAttribute('href', 'https://webclient.example.com');
  });

  test('shows pending-file copy and a Cancel button when logged out with a pending file; Cancel clears storage and navigates', async () => {
    (getSPDocumentDataFromLocalStorage as Mock).mockReturnValue({
      fileUrl: '/sites/x/Shared Documents/Report.pdf',
      fileName: 'Report.pdf',
    });
    const removeItemSpy = vi.spyOn(Storage.prototype, 'removeItem');

    render(<SendToLaserficheLoginComponent context={mockContext} />);

    expect(
      await screen.findByText(
        'You are not signed in. Please sign in to continue saving Report.pdf.'
      )
    ).toBeInTheDocument();
    const cancelButton = screen.getByRole('button', { name: 'Cancel' });

    fireEvent.click(cancelButton);

    expect(removeItemSpy).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    expect(Navigation.navigate).toHaveBeenCalledWith(
      'http://localhost/sites/x/Shared Documents/',
      true
    );
  });

  test('constructs SaveToLaserficheDialog and calls show() when logged in with a pending file', async () => {
    (getSPDocumentDataFromLocalStorage as Mock).mockReturnValue({
      fileUrl: '/sites/x/Shared Documents/Report.pdf',
      fileName: 'Report.pdf',
    });
    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    getLfLogin(container).state = 'LoggedIn';

    await waitFor(() => {
      expect(SaveToLaserficheCustomDialog).toHaveBeenCalledTimes(1);
    });
    const dialogInstance = (SaveToLaserficheCustomDialog as unknown as Mock).mock.results[0].value;
    await waitFor(() => {
      expect(dialogInstance.show).toHaveBeenCalledTimes(1);
    });
  });
});

describe('SendToLaserficheLoginComponent - clickLogin', () => {
  test('shows a "Sign In Failed" dialog mentioning the missing page, and does not open a popup, when the SharePoint page is absent', async () => {
    mockFetchSitePages(false);
    window.open = vi.fn();

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);

    // A plain substring-matcher function instead of `new RegExp(...)` from a
    // template string -- LASERFICHE_SIGNIN_PAGE_NAME is a fixed constant, not
    // untrusted input, but @rushstack/security/no-unsafe-regexp can't tell
    // that, and this sidesteps it without a suppression. `content` here is
    // RTL's own direct-text-only extraction (not `element.textContent`,
    // which would also pull in ancestors and risk a multiple-matches error);
    // the rendered message has more text after "SharePoint page" (see
    // needLaserficheSignInPage in the source), so this checks for the
    // substring rather than equality.
    expect(
      await screen.findByText((content) =>
        content.includes(`Missing ${LASERFICHE_SIGNIN_PAGE_NAME} SharePoint page`)
      )
    ).toBeInTheDocument();
    expect(window.open).not.toHaveBeenCalled();
  });

  test('opens the popup with a login URL when the sign-in page exists', async () => {
    mockFetchSitePages(true);
    window.open = vi.fn().mockReturnValue({ close: vi.fn() });

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);

    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));
    const [url, , options] = (window.open as Mock).mock.calls[0];
    expect(url).toContain('/SitePages/LaserficheSignIn.aspx?autologin&action=login');
    expect(options).toContain('popup,width=');
  });

  test('opens the popup with a logout URL when already signed in', async () => {
    mockFetchSitePages(true);
    window.open = vi.fn().mockReturnValue({ close: vi.fn() });

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    getLfLogin(container).state = 'LoggedIn';
    const signOutButton = await screen.findByRole('button', {
      name: 'Sign out',
    });
    fireEvent.click(signOutButton);

    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));
    const [url] = (window.open as Mock).mock.calls[0];
    expect(url).toContain('action=logout');
  });

  test('shows a pop-up-blocked "Sign In Failed" dialog instead of throwing when window.open returns null', async () => {
    mockFetchSitePages(true);
    window.open = vi.fn().mockReturnValue(null);

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);

    expect(await screen.findByText(/pop-ups/i)).toBeInTheDocument();
  });

  test('adds debug manifest query params to the popup URL when spfx-debug session storage holds valid JSON', async () => {
    mockFetchSitePages(true);
    window.open = vi.fn().mockReturnValue({ close: vi.fn() });
    const manifestsFileUrl = 'https://localhost:4321/temp/build/manifests.js';
    sessionStorage.setItem('spfx-debug', JSON.stringify({ manifestsFileUrl }));

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);

    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));
    const [url] = (window.open as Mock).mock.calls[0];
    const expectedDebugParams = `&debugManifestsFile=${encodeURIComponent(
      manifestsFileUrl
    )}&loadSPFX=true&debug=true&noredir=true`;
    expect(url).toContain(expectedDebugParams);
  });

  test('falls back cleanly with no debug params when spfx-debug session storage holds malformed JSON', async () => {
    mockFetchSitePages(true);
    window.open = vi.fn().mockReturnValue({ close: vi.fn() });
    sessionStorage.setItem('spfx-debug', '{not valid json');

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);

    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));
    const [url] = (window.open as Mock).mock.calls[0];
    expect(url).not.toContain('debugManifestsFile');
    expect(url).not.toContain('loadSPFX');
  });
});

describe('SendToLaserficheLoginComponent - handlePopupMessage', () => {
  test('ignores a message from a different origin', async () => {
    render(<SendToLaserficheLoginComponent context={mockContext} />);
    await screen.findByRole('button', { name: 'Sign in' });

    act(() => {
      window.dispatchEvent(
        new MessageEvent('message', {
          origin: 'https://evil.example.com',
          data: LOGIN_WINDOW_SUCCESS,
        })
      );
    });

    expect(screen.getByRole('button', { name: 'Sign in' })).toBeInTheDocument();
    expect(screen.queryByText(/Sign In Failed/)).not.toBeInTheDocument();
    expect(screen.queryByText(/Sign in failed/)).not.toBeInTheDocument();
  });

  test('closes the popup when the real success message arrives after a login click', async () => {
    mockFetchSitePages(true);
    const fakePopup = { close: vi.fn() };
    window.open = vi.fn().mockReturnValue(fakePopup);

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);
    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));

    act(() => {
      window.dispatchEvent(
        new MessageEvent('message', {
          origin: window.origin,
          data: LOGIN_WINDOW_SUCCESS,
        })
      );
    });

    expect(fakePopup.close).toHaveBeenCalledTimes(1);
  });

  test('flips loggedIn to false when the success message arrives after a logout click', async () => {
    mockFetchSitePages(true);
    const fakePopup = { close: vi.fn() };
    window.open = vi.fn().mockReturnValue(fakePopup);

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    getLfLogin(container).state = 'LoggedIn';
    const signOutButton = await screen.findByRole('button', {
      name: 'Sign out',
    });
    fireEvent.click(signOutButton);
    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));

    act(() => {
      window.dispatchEvent(
        new MessageEvent('message', {
          origin: window.origin,
          data: LOGIN_WINDOW_SUCCESS,
        })
      );
    });

    expect(await screen.findByRole('button', { name: 'Sign in' })).toBeInTheDocument();
  });

  test('shows a sign-in-failed dialog with the error details and closes the popup', async () => {
    mockFetchSitePages(true);
    const fakePopup = { close: vi.fn() };
    window.open = vi.fn().mockReturnValue(fakePopup);

    render(<SendToLaserficheLoginComponent context={mockContext} />);
    const signInButton = await screen.findByRole('button', {
      name: 'Sign in',
    });
    fireEvent.click(signInButton);
    await waitFor(() => expect(window.open).toHaveBeenCalledTimes(1));

    act(() => {
      window.dispatchEvent(
        new MessageEvent('message', {
          origin: window.origin,
          data: { ErrorType: 'SomeError', ErrorMessage: 'details here' },
        })
      );
    });

    expect(fakePopup.close).toHaveBeenCalledTimes(1);
    expect(
      await screen.findByText('Sign in failed, please try again. Details: details here')
    ).toBeInTheDocument();
  });
});

describe('SendToLaserficheLoginComponent - popup window path', () => {
  // window.location's href/search/pathname/etc. are spec'd as
  // [LegacyUnforgeable] own properties on the Location instance, and this
  // jsdom version enforces that fully: Object.defineProperty on
  // window.location itself, on window (to replace the whole location), and
  // vi.spyOn on its getters were all tried here and every one throws
  // "Cannot redefine property". The supported way to change the URL a test
  // sees without touching vitest.config.mts's environmentOptions is the
  // History API, which jsdom implements for real (it only refuses actual
  // page navigation, e.g. assigning location.href). pushState requires a
  // same-origin URL, hence building off window.location.origin (always
  // http://localhost/ here) rather than a real SharePoint origin.
  const originalHref = window.location.href;

  afterEach(() => {
    window.history.pushState({}, '', originalHref);
    Object.defineProperty(document, 'referrer', {
      value: '',
      configurable: true,
    });
  });

  function setPopupUrl(pathAndSearch: string): void {
    window.history.pushState({}, '', `${window.location.origin}${pathAndSearch}`);
  }

  test('reports success to the opener without touching the login button when already logged in and not asked to log out', async () => {
    setPopupUrl('/SitePages/LaserficheSignIn.aspx?autologin&action=login');
    const fakeOpener = { postMessage: vi.fn() };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).opener = fakeOpener;

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container);
    lfLogin.state = 'LoggedIn';
    const loginButton = document.createElement('button');
    loginButton.className = 'login-button';
    const clickSpy = vi.fn();
    loginButton.addEventListener('click', clickSpy);
    lfLogin.appendChild(loginButton);

    await waitFor(() => {
      expect(fakeOpener.postMessage).toHaveBeenCalledTimes(1);
    });
    expect(fakeOpener.postMessage).toHaveBeenCalledWith(LOGIN_WINDOW_SUCCESS, window.origin);
    // Regression pin: a sign-in request against an element that is already
    // logged in must not sign the user back out by clicking the button.
    expect(clickSpy).not.toHaveBeenCalled();
  });

  test('clicks the login-button element to sign out when the popup was asked to log out', async () => {
    setPopupUrl('/SitePages/LaserficheSignIn.aspx?autologin&action=logout');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).opener = { postMessage: vi.fn() };

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container);
    lfLogin.state = 'LoggedIn';
    const loginButton = document.createElement('button');
    loginButton.className = 'login-button';
    const clickSpy = vi.fn();
    loginButton.addEventListener('click', clickSpy);
    lfLogin.appendChild(loginButton);

    await waitFor(() => {
      expect(clickSpy).toHaveBeenCalledTimes(1);
    });
  });

  test('starts the login flow when landing on the popup fresh (not returning from the sign-in page)', async () => {
    setPopupUrl('/SitePages/LaserficheSignIn.aspx?autologin&action=login');
    Object.defineProperty(document, 'referrer', {
      value: '',
      configurable: true,
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).opener = { postMessage: vi.fn() };

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container); // state stays default 'LoggedOut'

    await waitFor(() => {
      expect(lfLogin.initLoginFlowAsync).toHaveBeenCalledTimes(1);
    });
  });

  test('waits for the token exchange and reports a timeout to the opener if it never completes', async () => {
    vi.useFakeTimers();
    setPopupUrl('/SitePages/LaserficheSignIn.aspx?autologin&action=login');
    Object.defineProperty(document, 'referrer', {
      value: 'https://accounts.laserfiche.com/signin',
      configurable: true,
    });
    const fakeOpener = { postMessage: vi.fn() };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).opener = fakeOpener;

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container);

    await flushMicrotasks();
    expect(lfLogin.initLoginFlowAsync).not.toHaveBeenCalled();

    act(() => {
      vi.advanceTimersByTime(30000);
    });

    expect(fakeOpener.postMessage).toHaveBeenCalledWith(
      expect.objectContaining({ ErrorType: 'TokenExchangeTimeout' }),
      window.origin
    );
  });

  test('reports success instead of a timeout when loginCompleted fires before the timer, and does not double-post afterward', async () => {
    vi.useFakeTimers();
    setPopupUrl('/SitePages/LaserficheSignIn.aspx?autologin&action=login');
    Object.defineProperty(document, 'referrer', {
      value: 'https://accounts.laserfiche.com/signin',
      configurable: true,
    });
    const fakeOpener = { postMessage: vi.fn() };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).opener = fakeOpener;

    const { container } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container);

    await flushMicrotasks();

    act(() => {
      lfLogin.dispatchEvent(new Event('loginCompleted'));
    });

    expect(fakeOpener.postMessage).toHaveBeenCalledTimes(1);
    expect(fakeOpener.postMessage).toHaveBeenCalledWith(LOGIN_WINDOW_SUCCESS, window.origin);

    act(() => {
      vi.advanceTimersByTime(30000);
    });

    // "Post once" guard: the timeout was cleared by loginCompletedInPopup, so
    // advancing past it must not add a second message.
    expect(fakeOpener.postMessage).toHaveBeenCalledTimes(1);
  });
});

describe('SendToLaserficheLoginComponent - unmount', () => {
  test('removes its lf-login listeners on unmount', async () => {
    const { container, unmount } = render(<SendToLaserficheLoginComponent context={mockContext} />);
    const lfLogin = getLfLogin(container);
    await flushMicrotasks();
    const removeSpy = vi.spyOn(lfLogin, 'removeEventListener');

    unmount();

    // React 17 has already cleared the ref when the cleanup runs, so a cleanup
    // that reads the ref throws before removing anything.
    const removed = removeSpy.mock.calls.map(([type]) => type).sort();
    expect(removed).toEqual([
      'loginCompleted',
      'loginCompleted',
      'logoutCompleted',
      'logoutCompleted',
    ]);
  });
});
