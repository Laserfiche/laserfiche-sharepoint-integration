// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { IRepositoryApiClientEx } from '@laserfiche/lf-ui-components-services';

export interface IRepositoryApiClientExInternal extends IRepositoryApiClientEx {
  clearCurrentRepo: () => void;
  _repoId?: string;
  _repoName?: string;
}
