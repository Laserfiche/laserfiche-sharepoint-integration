// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import {
  IRepositoryApiClient,
  RepositoryApiClient,
} from '@laserfiche/lf-repository-api-client';
import { IRepositoryApiClientExInternal } from './repository-client-types';
import { NgElement, WithProperties } from '@angular/elements';
import { LfLoginComponent } from '@laserfiche/types-lf-ui-components';

export class RepositoryClientExInternal {
  public repoClient: IRepositoryApiClientExInternal;

  public addAuthorizationHeader(
    request: RequestInit,
    accessToken: string | undefined
  ): void {
    const headers: Headers | undefined = new Headers(request.headers);
    const AUTH = 'Authorization';
    headers.set(AUTH, 'Bearer ' + accessToken);
    request.headers = headers;
  }

  public beforeFetchRequestAsync: (
    url: string,
    request: RequestInit
  ) => Promise<{
    regionalDomain: string;
  }> = async (url: string, request: RequestInit) => {
    // TODO trigger authorization flow if no accessToken
    const lfLogin = document.querySelector('lf-login') as NgElement &
      WithProperties<LfLoginComponent>;
    const accessToken = lfLogin.authorization_credentials?.accessToken;
    if (accessToken) {
      this.addAuthorizationHeader(request, accessToken);
      return {
        regionalDomain: lfLogin.account_endpoints.regionalDomain,
      }; // update this if you are using a different region
    } else {
      throw new Error('No access token');
    }
  };

  public afterFetchResponseAsync: (
    url: string,
    response: ResponseInit,
    request: RequestInit
  ) => Promise<boolean> = async (
    url: string,
    response: ResponseInit,
    request: RequestInit
  ) => {
    if (response.status === 401) {
      const lfLogin = document.querySelector('lf-login') as NgElement &
        WithProperties<LfLoginComponent>;
      const refresh = await lfLogin?.refreshTokenAsync(true);
      if (refresh) {
        const accessToken = lfLogin?.authorization_credentials?.accessToken;
        this.addAuthorizationHeader(request, accessToken);
        return true;
      } else {
        this.repoClient?.clearCurrentRepo();
        return false;
      }
    }
    return false;
  };

  public getCurrentRepo: () => Promise<{
    repoId: string;
    repoName: string;
  }> = async () => {
    if (this.repoClient) {
      const repos = await this.repoClient.repositoriesClient.getRepositoryList(
        {}
      );
      const repo = repos[0];
      if (repo.repoId && repo.repoName) {
        return { repoId: repo.repoId, repoName: repo.repoName };
      } else {
        throw new Error('Current repoId undefined.');
      }
    } else {
      throw new Error('repoClient undefined.');
    }
  };

  public async createRepositoryClientAsync(): Promise<IRepositoryApiClientExInternal> {
    const partialRepoClient: IRepositoryApiClient =
      RepositoryApiClient.createFromHttpRequestHandler({
        beforeFetchRequestAsync: this.beforeFetchRequestAsync,
        afterFetchResponseAsync: this.afterFetchResponseAsync,
      });
    const clearCurrentRepo: () => void = () => {
      if (this.repoClient) {
        this.repoClient._repoId = undefined;
        this.repoClient._repoName = undefined;
      }
    };
    this.repoClient = {
      clearCurrentRepo,
      _repoId: undefined,
      _repoName: undefined,
      getCurrentRepoId: async () => {
        if (this.repoClient?._repoId) {
          console.log('getting id from cache');
          return this.repoClient._repoId;
        } else {
          console.log('getting id from api');
          const repo = (await this.getCurrentRepo()).repoId;
          if (this.repoClient) {
            this.repoClient._repoId = repo;
          }
          return repo;
        }
      },
      getCurrentRepoName: async () => {
        if (this.repoClient?._repoName) {
          return this.repoClient._repoName;
        } else {
          const repo = (await this.getCurrentRepo()).repoName;
          if (this.repoClient) {
            this.repoClient._repoName = repo;
          } else {
            console.debug('Cannot set repoName, repoClient does not exist');
          }
          return repo;
        }
      },
      ...partialRepoClient,
    };
    return this.repoClient;
  }
}
