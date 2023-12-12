// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IRepositoryApiClientExInternal } from "../../../repository-client/repository-client-types";
import { ProfileConfiguration } from "./ProfileConfigurationComponents";

export interface IManageConfigurationProps {
  header?: JSX.Element;
  extraConfiguration?: JSX.Element;
  repoClient: IRepositoryApiClientExInternal;
  loggedIn: boolean;
  profileConfig: ProfileConfiguration;
  loadingContent: boolean;
  createNew: boolean;
  context: WebPartContext;
  validate: boolean;
  handleProfileConfigUpdate: (profileConfig: ProfileConfiguration) => void;
  saveConfiguration: () => Promise<void>;
}
