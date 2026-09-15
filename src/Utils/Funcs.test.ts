// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// The real @laserfiche/lf-js-utils is ESM only, so jest.config.cjs maps it to an
// empty module. Stand in a copy of UrlUtils.combineURLs so these tests assert
// the URL the Laserfiche web client actually receives.
jest.mock('@laserfiche/lf-js-utils', () => ({
  UrlUtils: {
    combineURLs: (
      baseURL: string,
      relativeURL: string,
      queryStringParams?: [string, string | number | boolean][]
    ): string => {
      const separator = baseURL.endsWith('/') ? '' : '/';
      let url = baseURL + separator + relativeURL;
      if (queryStringParams && queryStringParams.length > 0) {
        url +=
          '?' +
          queryStringParams
            .map((param) => `${param[0]}=${encodeURIComponent(param[1])}`)
            .join('&');
      }
      return url;
    },
  },
}));

import { getEntryWebAccessUrl } from './Funcs';

const WA_URL = 'https://app.laserfiche.com/laserfiche';
const REPO_ID = 'r-26a7e408';
const CUSTOMER_ID = '491263275';
const ENTRY_ID = '1862863';

describe('getEntryWebAccessUrl', () => {
  test('includes repo and customerId for a document', () => {
    // Act
    const url = getEntryWebAccessUrl(
      ENTRY_ID,
      WA_URL,
      false,
      REPO_ID,
      CUSTOMER_ID
    );

    // Assert
    expect(url).toBe(
      `${WA_URL}/DocView.aspx?repo=${REPO_ID}&customerId=${CUSTOMER_ID}&docid=${ENTRY_ID}`
    );
  });

  test('includes repo and customerId for a container', () => {
    // Act
    const url = getEntryWebAccessUrl(
      ENTRY_ID,
      WA_URL,
      true,
      REPO_ID,
      CUSTOMER_ID
    );

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
    expect(documentUrl).toBe(
      `${WA_URL}/DocView.aspx?repo=${REPO_ID}&docid=${ENTRY_ID}`
    );
    expect(containerUrl).toBe(
      `${WA_URL}/Browse.aspx?repo=${REPO_ID}#?id=${ENTRY_ID}`
    );
  });

  // account_id is '' on the login component until the user signs in.
  test('omits customerId when it is an empty string', () => {
    // Act
    const url = getEntryWebAccessUrl(ENTRY_ID, WA_URL, false, REPO_ID, '');

    // Assert
    expect(url).toBe(`${WA_URL}/DocView.aspx?repo=${REPO_ID}&docid=${ENTRY_ID}`);
  });

  test('includes customerId when there is no repo', () => {
    // Act
    const documentUrl = getEntryWebAccessUrl(
      ENTRY_ID,
      WA_URL,
      false,
      undefined,
      CUSTOMER_ID
    );
    const containerUrl = getEntryWebAccessUrl(
      ENTRY_ID,
      WA_URL,
      true,
      undefined,
      CUSTOMER_ID
    );

    // Assert
    expect(documentUrl).toBe(
      `${WA_URL}/DocView.aspx?customerId=${CUSTOMER_ID}&docid=${ENTRY_ID}`
    );
    expect(containerUrl).toBe(
      `${WA_URL}/Browse.aspx?customerId=${CUSTOMER_ID}#?id=${ENTRY_ID}`
    );
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
    expect(
      getEntryWebAccessUrl('', WA_URL, false, REPO_ID, CUSTOMER_ID)
    ).toBeUndefined();
    expect(
      getEntryWebAccessUrl(ENTRY_ID, '', false, REPO_ID, CUSTOMER_ID)
    ).toBeUndefined();
    expect(
      getEntryWebAccessUrl(ENTRY_ID, undefined, false, REPO_ID, CUSTOMER_ID)
    ).toBeUndefined();
  });
});
