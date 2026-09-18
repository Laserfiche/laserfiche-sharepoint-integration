// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { UrlUtils } from '@laserfiche/lf-js-utils';
import { FieldType } from '@laserfiche/lf-repository-api-client-v2';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import {
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
    newUrl = UrlUtils.combineURLs(
      waUrl ?? '',
      'Browse.aspx',
      commonQueryParams
    );
    newUrl += `#?id=${encodeURIComponent(entryId)}`;
  } else {
    const queryParams: UrlUtils.QueryParameter[] = [
      ...commonQueryParams,
      ['id', entryId],
    ];
    newUrl = UrlUtils.combineURLs(waUrl ?? '', 'DocView.aspx', queryParams);
  }
  return newUrl;
}

export function getSPListURL(
  context: BaseComponentContext,
  listName: string
): string {
  return (
    context.pageContext.web.absoluteUrl +
    `/_api/web/lists/GetByTitle('${listName}')`
  );
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
 * Reads the SharePoint document data the Send to Laserfiche flow stashes in
 * local storage. Returns undefined rather than throwing when the slot holds
 * something that is not valid JSON: this is read during render, so a parse
 * error here would take down the component tree.
 */
export function getSPDocumentDataFromLocalStorage():
  | ISPDocumentData
  | undefined {
  try {
    const raw = window.localStorage.getItem(SP_LOCAL_STORAGE_KEY);
    return raw ? (JSON.parse(raw) as ISPDocumentData) : undefined;
  } catch (error) {
    console.warn(
      `Unable to read ${SP_LOCAL_STORAGE_KEY} from local storage.`,
      error
    );
    return undefined;
  }
}
