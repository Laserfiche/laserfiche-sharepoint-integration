// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRepositoryApiClientExInternal } from '../../../../repository-client/repository-client-types';

export interface IAddNewManageConfigurationProps {
  context: WebPartContext;
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
}
