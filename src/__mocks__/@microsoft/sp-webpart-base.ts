// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContext, SPPermission, SPWeb } from '@microsoft/sp-page-context';

const mockWebPartContext: Partial<WebPartContext> = {
  pageContext: {
    web: {
      permissions: new SPPermission({ High: 0, Low: 0 })
    } as SPWeb
  } as PageContext
};

export default mockWebPartContext as WebPartContext;
