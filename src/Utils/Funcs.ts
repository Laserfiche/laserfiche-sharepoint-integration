import { UrlUtils } from '@laserfiche/lf-js-utils';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { SPDEVMODE_LOCAL_STORAGE_KEY } from '../webparts/constants';

export function getEntryWebAccessUrl(
  nodeId: string,
  waUrl: string,
  isContainer: boolean,
  repoId?: string
): string | undefined {
  if (!nodeId || nodeId?.length === 0 || !waUrl || waUrl?.length === 0) {
    return undefined;
  }
  let newUrl: string;
  if (isContainer) {
    const queryParams: UrlUtils.QueryParameter[] = repoId
      ? [['repo', repoId]]
      : [];
    newUrl = UrlUtils.combineURLs(waUrl ?? '', 'Browse.aspx', queryParams);
    newUrl += `#?id=${encodeURIComponent(nodeId)}`;
  } else {
    const queryParams: UrlUtils.QueryParameter[] = repoId
      ? [
          ['repo', repoId],
          ['docid', nodeId],
        ]
      : [['docid', nodeId]];
    newUrl = UrlUtils.combineURLs(waUrl ?? '', 'DocView.aspx', queryParams);
  }
  return newUrl;
}

export function getSPListURL(context: BaseComponentContext, listName: string): string {
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
