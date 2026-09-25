// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// @laserfiche/lf-js-utils is NOT mocked: these tests assert the URL the
// Laserfiche web client actually receives, built by the real UrlUtils.

import { LfLoginComponent, LoginState } from '@laserfiche/types-lf-ui-components';
import {
  getEntryWebAccessUrl,
  getSPDocumentDataFromLocalStorage,
  isLfLoginSignedIn,
  openLoginWindow,
} from './Funcs';
import { SP_LOCAL_STORAGE_KEY } from '../webparts/constants';

const WA_URL = 'https://app.laserfiche.com/laserfiche';
const REPO_ID = 'r-26a7e408';
const CUSTOMER_ID = '491263275';
const ENTRY_ID = '1862863';

describe('getEntryWebAccessUrl', () => {
  test('includes repo and customerId for a document', () => {
    // Act
    const url = getEntryWebAccessUrl(ENTRY_ID, WA_URL, false, REPO_ID, CUSTOMER_ID);

    // Assert
    expect(url).toBe(
      `${WA_URL}/DocView.aspx?repo=${REPO_ID}&customerId=${CUSTOMER_ID}&id=${ENTRY_ID}`
    );
  });

  test('includes repo and customerId for a container', () => {
    // Act
    const url = getEntryWebAccessUrl(ENTRY_ID, WA_URL, true, REPO_ID, CUSTOMER_ID);

    // Assert
    expect(url).toBe(
      `${WA_URL}/Browse.aspx?repo=${REPO_ID}&customerId=${CUSTOMER_ID}#?id=${ENTRY_ID}`
    );
  });

  test('omits customerId when it is not provided', () => {
    // Act
    const documentUrl = getEntryWebAccessUrl(ENTRY_ID, WA_URL, false, REPO_ID);
    const containerUrl = getEntryWebAccessUrl(ENTRY_ID, WA_URL, true, REPO_ID);

    // Assert
    expect(documentUrl).toBe(`${WA_URL}/DocView.aspx?repo=${REPO_ID}&id=${ENTRY_ID}`);
    expect(containerUrl).toBe(`${WA_URL}/Browse.aspx?repo=${REPO_ID}#?id=${ENTRY_ID}`);
  });

  // account_id is '' on the login component until the user signs in.
  test('omits customerId when it is an empty string', () => {
    // Act
    const url = getEntryWebAccessUrl(ENTRY_ID, WA_URL, false, REPO_ID, '');

    // Assert
    expect(url).toBe(`${WA_URL}/DocView.aspx?repo=${REPO_ID}&id=${ENTRY_ID}`);
  });

  test('includes customerId when there is no repo', () => {
    // Act
    const documentUrl = getEntryWebAccessUrl(ENTRY_ID, WA_URL, false, undefined, CUSTOMER_ID);
    const containerUrl = getEntryWebAccessUrl(ENTRY_ID, WA_URL, true, undefined, CUSTOMER_ID);

    // Assert
    expect(documentUrl).toBe(`${WA_URL}/DocView.aspx?customerId=${CUSTOMER_ID}&id=${ENTRY_ID}`);
    expect(containerUrl).toBe(`${WA_URL}/Browse.aspx?customerId=${CUSTOMER_ID}#?id=${ENTRY_ID}`);
  });

  test('omits the query string when there is no repo or customerId', () => {
    // Act
    const containerUrl = getEntryWebAccessUrl(ENTRY_ID, WA_URL, true);

    // Assert
    expect(containerUrl).toBe(`${WA_URL}/Browse.aspx#?id=${ENTRY_ID}`);
  });

  test('encodes the entry id in the container fragment', () => {
    // Act
    const url = getEntryWebAccessUrl('186 2863', WA_URL, true, REPO_ID);

    // Assert
    expect(url).toBe(`${WA_URL}/Browse.aspx?repo=${REPO_ID}#?id=186%202863`);
  });

  test('returns undefined when the entry id or web client URL is missing', () => {
    // Act / Assert
    expect(getEntryWebAccessUrl('', WA_URL, false, REPO_ID, CUSTOMER_ID)).toBeUndefined();
    expect(getEntryWebAccessUrl(ENTRY_ID, '', false, REPO_ID, CUSTOMER_ID)).toBeUndefined();
    expect(getEntryWebAccessUrl(ENTRY_ID, undefined, false, REPO_ID, CUSTOMER_ID)).toBeUndefined();
  });
});

describe('getSPDocumentDataFromLocalStorage', () => {
  afterEach(() => {
    window.localStorage.clear();
    vi.restoreAllMocks();
  });

  test('returns the parsed document data when the stored value is valid', () => {
    // Arrange
    const docData = {
      fileName: 'invoice.pdf',
      documentName: 'invoice',
      action: 'copy',
      fileUrl: 'https://contoso.sharepoint.com/invoice.pdf',
      entryId: '1',
      contextPageAbsoluteUrl: 'https://contoso.sharepoint.com',
    };
    window.localStorage.setItem(SP_LOCAL_STORAGE_KEY, JSON.stringify(docData));

    // Act
    const result = getSPDocumentDataFromLocalStorage();

    // Assert
    expect(result).toEqual(docData);
  });

  test('returns undefined when nothing is stored', () => {
    // Arrange

    // Act
    const result = getSPDocumentDataFromLocalStorage();

    // Assert
    expect(result).toBeUndefined();
  });

  test('returns undefined instead of throwing on malformed JSON', () => {
    // Arrange
    // This is read during render, so a SyntaxError here would crash the tree.
    window.localStorage.setItem(SP_LOCAL_STORAGE_KEY, '{"fileName":');
    const warn = vi.spyOn(console, 'warn');
    warn.mockImplementation(() => undefined);

    // Act
    const result = getSPDocumentDataFromLocalStorage();

    // Assert
    expect(result).toBeUndefined();
    expect(warn).toHaveBeenCalled();
  });
});

describe('openLoginWindow', () => {
  const LOGIN_URL =
    'https://contoso.sharepoint.com/SitePages/LaserficheSignIn.aspx?autologin&action=login';
  const stubbed: [object, string, PropertyDescriptor | undefined][] = [];

  // The window and screen geometry are read-only in the DOM typings, so they
  // are shadowed here and put back after each test.
  function stubGeometry(target: object, values: Record<string, number>): void {
    for (const [key, value] of Object.entries(values)) {
      stubbed.push([target, key, Object.getOwnPropertyDescriptor(target, key)]);
      Object.defineProperty(target, key, { configurable: true, value });
    }
  }

  afterEach(() => {
    for (const [target, key, descriptor] of stubbed.reverse()) {
      if (descriptor) {
        Object.defineProperty(target, key, descriptor);
      } else {
        delete (target as Record<string, unknown>)[key];
      }
    }
    stubbed.length = 0;
    vi.restoreAllMocks();
  });

  test('centers the popup on the browser window that opened it', () => {
    // Arrange
    stubGeometry(window.screen, { width: 1920, height: 1080 });
    stubGeometry(window, {
      screenX: 200,
      screenY: 100,
      outerWidth: 1600,
      outerHeight: 1000,
    });
    const popup = {} as Window;
    const open = vi.spyOn(window, 'open').mockReturnValue(popup);

    // Act
    const result = openLoginWindow(LOGIN_URL);

    // Assert
    expect(result).toBe(popup);
    expect(open).toHaveBeenCalledWith(
      LOGIN_URL,
      'loginWindow',
      'popup,width=1000,height=800,left=500,top=200'
    );
  });

  test('centers on a browser window on a monitor left of the primary one', () => {
    // Arrange
    stubGeometry(window.screen, { width: 1920, height: 1080 });
    stubGeometry(window, {
      screenX: -1920,
      screenY: 0,
      outerWidth: 1920,
      outerHeight: 1040,
    });
    const open = vi.spyOn(window, 'open').mockReturnValue(null);

    // Act
    openLoginWindow(LOGIN_URL);

    // Assert
    expect(open).toHaveBeenCalledWith(
      LOGIN_URL,
      'loginWindow',
      'popup,width=1000,height=800,left=-1460,top=120'
    );
  });

  test('shrinks the popup to fit a screen smaller than the default size', () => {
    // Arrange
    stubGeometry(window.screen, { width: 800, height: 600 });
    stubGeometry(window, {
      screenX: 0,
      screenY: 0,
      outerWidth: 800,
      outerHeight: 600,
    });
    const open = vi.spyOn(window, 'open').mockReturnValue(null);

    // Act
    const result = openLoginWindow(LOGIN_URL);

    // Assert
    expect(result).toBeUndefined();
    expect(open).toHaveBeenCalledWith(
      LOGIN_URL,
      'loginWindow',
      'popup,width=800,height=600,left=0,top=0'
    );
  });
});

describe('isLfLoginSignedIn', () => {
  test('is true when lf-login reports it is signed in', () => {
    // Arrange
    const loginElement = { state: LoginState.LoggedIn } as LfLoginComponent;

    // Act
    const signedIn = isLfLoginSignedIn(loginElement);

    // Assert
    expect(signedIn).toBe(true);
  });

  // lf-login restores credentials from local storage before its state
  // catches up with them.
  test('is true when lf-login holds credentials its state does not reflect yet', () => {
    // Arrange
    const loginElement = {
      state: LoginState.LoggedOut,
      authorization_credentials: { accessToken: 'token' },
    } as unknown as LfLoginComponent;

    // Act
    const signedIn = isLfLoginSignedIn(loginElement);

    // Assert
    expect(signedIn).toBe(true);
  });

  test('is false when lf-login is signed out and holds no credentials', () => {
    // Arrange
    const loginElement = { state: LoginState.LoggedOut } as LfLoginComponent;

    // Act
    const signedIn = isLfLoginSignedIn(loginElement);

    // Assert
    expect(signedIn).toBe(false);
  });

  test('is false before the lf-login element exists', () => {
    // Act
    const signedIn = isLfLoginSignedIn(undefined);

    // Assert
    expect(signedIn).toBe(false);
  });
});
