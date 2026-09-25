// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import * as React from 'react';
import { AbortedLoginError } from '@laserfiche/types-lf-ui-components';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { MessageDialog } from '../extensions/savetoLaserfiche/CommonDialogs';
import { LASERFICHE_SIGNIN_PAGE_NAME, LOGIN_WINDOW_SUCCESS } from '../webparts/constants';
import {
  POPUP_BLOCKED,
  SIGN_IN_FAILED,
  SIGN_IN_FAILED_PLEASE_TRY_AGAIN_DETAILS,
  needLaserficheSignInPage,
} from '../webparts/strings';
import { getSPListURL, openLoginWindow } from './Funcs';

type SignInAction = 'login' | 'logout';

export interface SignInPopupOptions {
  context: BaseComponentContext;
  // Decides whether the button signs in or signs out.
  loggedIn: boolean;
  // Shows a dialog, or removes it when passed undefined.
  setMessageModal: (message: JSX.Element | undefined) => void;
  // The popup signs out on its own lf-login element, so the page's element
  // doesn't necessarily raise logoutCompleted: this is how the page hears.
  onSignedOut: () => void;
  onSignedIn?: () => Promise<void>;
}

// SPFx records the gulp serve manifest URL here once a page is loaded with
// debugManifestsFile, and keeps using it for the rest of the session.
const SPFX_DEBUG_SESSION_STORAGE_KEY = 'spfx-debug';

// The sign-in popup gets a URL of our own making, so nothing carries the local
// debug build into it. A popup inherits the opener's session storage, but that
// on its own is not enough: it asks for the localhost manifests and fails with
// "Error loading debug manifests". Repeating the parameters in the popup URL is
// what actually loads them, so forward them whenever the opener is itself
// running against a debug manifest. Returns '' in a normal deployment, where
// the key is absent, leaving the popup URL untouched.
function debugManifestsQueryString(): string {
  try {
    const debugSettings = sessionStorage.getItem(SPFX_DEBUG_SESSION_STORAGE_KEY);
    if (!debugSettings) {
      return '';
    }
    const manifestsFileUrl: string | undefined = JSON.parse(debugSettings).manifestsFileUrl;
    if (!manifestsFileUrl) {
      return '';
    }
    return `&debugManifestsFile=${encodeURIComponent(
      manifestsFileUrl
    )}&loadSPFX=true&debug=true&noredir=true`;
  } catch {
    // Session storage can throw, and the entry is not ours to assume is valid
    // JSON. Debug parameters are a developer convenience, so fall back to the
    // plain URL rather than break sign-in over them.
    return '';
  }
}

async function isSignInPageConfiguredAsync(context: BaseComponentContext): Promise<boolean> {
  try {
    const res = await fetch(`${getSPListURL(context, 'Site Pages')}/items`, {
      method: 'GET',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    });
    const sitePages = await res.json();
    for (let o = 0; o < sitePages.value.length; o++) {
      const pageName = sitePages.value[o].Title;
      if (pageName === LASERFICHE_SIGNIN_PAGE_NAME) {
        return true;
      }
    }
  } catch (error) {
    console.warn(
      `Unable to determine if a SharePoint Page with name ${LASERFICHE_SIGNIN_PAGE_NAME} exists.`,
      error
    );
    return false;
  }
  return false;
}

/**
 * Signs in to or out of Laserfiche through the site's LaserficheSignIn page,
 * opened in a popup window, and reports the outcome the popup posts back.
 *
 * @param options What the calling page needs to know and how to tell it
 * @returns The click handler for the page's Sign in / Sign out button
 */
export function useSignInPopup(options: SignInPopupOptions): () => Promise<void> {
  // The message listener is attached once, on mount, so it reads the options
  // through this ref to reach the latest render's callbacks.
  const latestOptions = React.useRef(options);
  latestOptions.current = options;
  // What the last click asked the popup to do. The popup reports the same
  // success message either way, so this is the only thing that says which.
  const requestedAction = React.useRef<SignInAction>('login');
  const loginWindow = React.useRef<Window | undefined>(undefined);

  function showSignInFailed(message: string): void {
    const { setMessageModal } = latestOptions.current;
    setMessageModal(
      <MessageDialog
        title={SIGN_IN_FAILED}
        message={message}
        clickOkay={() => {
          setMessageModal(undefined);
        }}
      />
    );
  }

  function closeLoginWindow(): void {
    loginWindow.current?.close();
    loginWindow.current = undefined;
  }

  async function handlePopupMessageAsync(event: MessageEvent): Promise<void> {
    if (event.origin !== window.origin) {
      return;
    }
    if (event.data === LOGIN_WINDOW_SUCCESS) {
      closeLoginWindow();
      if (requestedAction.current === 'logout') {
        latestOptions.current.onSignedOut();
      } else {
        await latestOptions.current.onSignedIn?.();
      }
      return;
    }
    const parsedError: AbortedLoginError | undefined = event.data;
    if (parsedError?.ErrorMessage && parsedError?.ErrorType) {
      closeLoginWindow();
      showSignInFailed(`${SIGN_IN_FAILED_PLEASE_TRY_AGAIN_DETAILS} ${parsedError.ErrorMessage}`);
    }
  }

  // Tied to the component's lifetime: registering on click leaked a listener
  // per mount, because nothing removed it when the page went away.
  React.useEffect(() => {
    const handleMessage: (event: MessageEvent) => void = (event) => {
      void handlePopupMessageAsync(event);
    };
    window.addEventListener('message', handleMessage);
    return () => {
      window.removeEventListener('message', handleMessage);
    };
  }, []);

  return async function signInOrOutAsync(): Promise<void> {
    // The popup cannot tell a sign-in apart from a sign-out by looking at its
    // own element state, so say which one this click means. Without it the
    // popup takes the click for a sign-in, reports success and signs nobody
    // out, which is what made "Sign out" do nothing.
    const action: SignInAction = options.loggedIn ? 'logout' : 'login';
    requestedAction.current = action;
    const url =
      `${options.context.pageContext.web.absoluteUrl}/SitePages/${LASERFICHE_SIGNIN_PAGE_NAME}.aspx` +
      `?autologin&action=${action}${debugManifestsQueryString()}`;
    if (!(await isSignInPageConfiguredAsync(options.context))) {
      showSignInFailed(needLaserficheSignInPage);
      return;
    }
    const popup = openLoginWindow(url);
    if (!popup) {
      // A blocked pop-up returns null, so bail out instead of throwing.
      showSignInFailed(POPUP_BLOCKED);
      return;
    }
    loginWindow.current = popup;
  };
}
