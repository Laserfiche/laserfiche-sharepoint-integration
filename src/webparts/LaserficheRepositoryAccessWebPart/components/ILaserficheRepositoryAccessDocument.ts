// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

export interface IDocument {
  key: string;
  name: string;
  value: string;
  parentId: number;
  iconName: string;
  fileType: string;
  modifiedBy: string;
  dateModified: string;
  dateModifiedValue: number;
  title: string;
  id: number;
  creationTime: string;
  lastModifiedTime: string;
  entryType: string;
  volumeName: string;
  templateName: string;
  extension: string;
  pageCount: number;
}
