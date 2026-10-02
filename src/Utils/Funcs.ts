// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { UrlUtils } from '@laserfiche/lf-js-utils';
import { FieldType } from '@laserfiche/lf-repository-api-client-v2';
import { LfLoginComponent, LoginState } from '@laserfiche/types-lf-ui-components';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import {
  LOGIN_WINDOW_HEIGHT_PX,
  LOGIN_WINDOW_WIDTH_PX,
  SPDEVMODE_LOCAL_STORAGE_KEY,
  SP_LOCAL_STORAGE_KEY,
} from '../webparts/constants';
import { ISPDocumentData } from './Types';

export function getEntryWebAccessUrl(
  entryId: string,
  waUrl: string,
  isContainer: boolean,
  repoId?: string,
  customerId?: string
): string | undefined {
  if (!entryId || entryId?.length === 0 || !waUrl || waUrl?.length === 0) {
    return undefined;
  }
  const commonQueryParams: UrlUtils.QueryParameter[] = [];
  if (repoId) {
    commonQueryParams.push(['repo', repoId]);
  }
  if (customerId) {
    commonQueryParams.push(['customerId', customerId]);
  }
  let newUrl: string;
  if (isContainer) {
    newUrl = UrlUtils.combineURLs(waUrl ?? '', 'Browse.aspx', commonQueryParams);
    newUrl += `#?id=${encodeURIComponent(entryId)}`;
  } else {
    const queryParams: UrlUtils.QueryParameter[] = [...commonQueryParams, ['id', entryId]];
    newUrl = UrlUtils.combineURLs(waUrl ?? '', 'DocView.aspx', queryParams);
  }
  return newUrl;
}

export function getSPListURL(context: BaseComponentContext, listName: string): string {
  return context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${listName}')`;
}

export function getRegion(): string {
  const spDevMode = window?.localStorage.getItem(SPDEVMODE_LOCAL_STORAGE_KEY);
  if (!spDevMode) {
    window.localStorage.setItem(SPDEVMODE_LOCAL_STORAGE_KEY, 'false');
  }
  const spDevModeTrue = spDevMode && spDevMode.toLocaleLowerCase() === 'true';
  const region = spDevModeTrue ? 'a.clouddev.laserfiche.com' : 'laserfiche.com';
  return region;
}

/**
 * Opens the Laserfiche sign-in popup centered on the current browser window.
 *
 * Without left/top the browser picks the position itself, which in Edge and
 * Chrome is the top-left corner of the opener (of the screen, when the browser
 * is maximized). Do not follow this with resizeTo: it keeps the top-left corner
 * fixed, so it undoes the centering, and it sets the outer size where the
 * width/height features set the inner one.
 *
 * @param url The sign-in page to load in the popup
 * @returns The popup, or undefined when the browser blocked it
 */
export function openLoginWindow(url: string): Window | undefined {
  // Cap to the screen the same way the add-ins do, so the window still fits
  // on smaller displays.
  const width = Math.min(LOGIN_WINDOW_WIDTH_PX, window.screen.width || LOGIN_WINDOW_WIDTH_PX);
  const height = Math.min(LOGIN_WINDOW_HEIGHT_PX, window.screen.height || LOGIN_WINDOW_HEIGHT_PX);
  const left = Math.round(window.screenX + (window.outerWidth - width) / 2);
  const top = Math.round(window.screenY + (window.outerHeight - height) / 2);
  return (
    window.open(
      url,
      'loginWindow',
      `popup,width=${width},height=${height},left=${left},top=${top}`
    ) ?? undefined
  );
}

/**
 * Whether an lf-login element holds a Laserfiche session. Either sign counts:
 * the element can restore credentials from local storage before its state
 * says LoggedIn.
 *
 * @param loginElement The lf-login element, or undefined before it renders
 * @returns true when the element reports LoggedIn or holds credentials
 */
export function isLfLoginSignedIn(
  loginElement: Pick<LfLoginComponent, 'state' | 'authorization_credentials'> | undefined
): boolean {
  return loginElement?.state === LoginState.LoggedIn || !!loginElement?.authorization_credentials;
}

export function getCorrespondingTypeFieldName(fieldType: FieldType): string {
  switch (fieldType) {
    case FieldType.Date:
    case FieldType.List:
    case FieldType.Time:
    case FieldType.Number:
      return fieldType;
    case FieldType.DateTime:
      return 'Date/Time';
    case FieldType.String:
      return 'Text';
    case FieldType.ShortInteger:
      return 'Integer';
    case FieldType.LongInteger:
      return 'Long Integer';
  }
}

/**
 * Extract a human-readable message from a thrown value of unknown shape.
 *
 * `title` is checked before `message` because the repository API client
 * (`@laserfiche/lf-repository-api-client-v2`) throws RFC 9457 `ProblemDetails`
 * errors, whose `title` is the short, human-readable summary; `message` is
 * used as a fallback for errors that aren't shaped that way (e.g. plain
 * `Error`s or rejected browser API promises).
 *
 * @param error The thrown value, of unknown shape
 * @returns `title`, `message`, or the value itself if it's a string; `undefined` otherwise
 */
export function getErrorDetails(error: unknown): string | undefined {
  if (error && typeof error === 'object') {
    const { title, message } = error as { title?: unknown; message?: unknown };
    if (typeof title === 'string') {
      return title;
    }
    if (typeof message === 'string') {
      return message;
    }
    return undefined;
  }
  return typeof error === 'string' ? error : undefined;
}

/**
 * Formats a thrown value for the browser console. Its details can carry text
 * from a server response or local storage, so they are JSON-encoded: a line
 * break comes out as `\r` or `\n` rather than starting a forged log line
 * (CWE-117). `JSON.stringify` is also a CWE-117 cleanser Veracode recognizes.
 *
 * @param error The thrown value, of unknown shape
 * @returns The value's details as a quoted, single-line JSON string
 */
export function formatErrorForLog(error: unknown): string {
  return JSON.stringify(getErrorDetails(error) ?? String(error));
}

/**
 * Reads the SharePoint document data the Send to Laserfiche flow stashes in
 * local storage. Returns undefined rather than throwing when the slot holds
 * something that is not valid JSON: this is read during render, so a parse
 * error here would take down the component tree.
 */
export function getSPDocumentDataFromLocalStorage(): ISPDocumentData | undefined {
  try {
    const raw = window.localStorage.getItem(SP_LOCAL_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as ISPDocumentData) : undefined;
  } catch (error) {
    console.warn(
      `Unable to read ${SP_LOCAL_STORAGE_KEY} from local storage: ${formatErrorForLog(error)}`
    );
    return undefined;
  }
}
