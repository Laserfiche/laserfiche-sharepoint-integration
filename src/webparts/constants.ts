// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

export const clientId = '8ee987ea-a0b1-4ca2-85c4-a79b335cd214';

export const LASERFICHE_ADMIN_CONFIGURATION_NAME = 'LaserficheAdminConfiguration';
export const MANAGE_MAPPING = 'ManageMapping';
export const MANAGE_CONFIGURATIONS = 'ManageConfigurations';

export const LOGIN_WINDOW_SUCCESS = 'loginWindowSuccess';
export const LASERFICHE_SIGNIN_PAGE_NAME = 'LaserficheSignIn';

// Bumped to 'spdocdata2' with the api-client v1 -> v2 migration: the stored
// ISPDocumentData.metadata shape changed from PostEntryWithEdocMetadataRequest
// to ImportEntryRequestMetadata, and the handoff spans a full page navigation
// to the sign-in page. A stale v1 entry would be read as empty metadata rather
// than throwing, so the key changes to make old entries ignored instead.
export const SP_LOCAL_STORAGE_KEY = 'spdocdata2';
export const SPDEVMODE_LOCAL_STORAGE_KEY = 'spDevMode';

export { default as LF_UI_COMPONENTS_URL } from '../Assets/packages/lf-ui-components.js';
export { default as LF_MS_OFFICE_LITE_CSS_URL } from '../Assets/packages/lf-ms-office-lite.cssasset';
export { default as LF_INDIGO_PINK_CSS_URL } from '../Assets/packages/indigo-pink.cssasset';
