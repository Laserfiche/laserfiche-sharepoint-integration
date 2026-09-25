// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { IImportEntryRequestMetadata } from '@laserfiche/lf-repository-api-client-v2';
import { ActionTypes } from '../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';

export interface ISPDocumentData {
  metadata?: IImportEntryRequestMetadata;
  fileName: string;
  documentName: string;
  templateName?: string;
  action: ActionTypes;
  fileUrl: string;
  entryId: string;
  contextPageAbsoluteUrl: string;
  lfProfile?: string;
}

// A document just saved to Laserfiche, with Web Client links to it and to the
// folder it was saved into; either link is undefined when it can't be built.
export interface SavedLaserficheDocument {
  fileName: string;
  fileLink: string;
  folderLink: string;
}

export interface ProfileMappingConfiguration {
  id: string;
  SharePointContentType: string;
  LaserficheContentType: string;
  toggle: boolean;
}
