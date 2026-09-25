// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('@laserfiche/lf-repository-api-client-v2', () => ({
  RepositoryApiClient: {
    createFromHttpRequestHandler: vi.fn(),
  },
}));

import type { Mock } from 'vitest';
import { RepositoryApiClient } from '@laserfiche/lf-repository-api-client-v2';
import { RepositoryClientExInternal } from './repository-client';
import { IRepositoryApiClientExInternal } from './repository-client-types';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type LfLoginNode = any;

describe('RepositoryClientExInternal', () => {
  let instance: RepositoryClientExInternal;
  let lfLogin: LfLoginNode;

  beforeEach(() => {
    vi.clearAllMocks();
    document.body.innerHTML = '<lf-login></lf-login>';
    lfLogin = document.querySelector('lf-login');
    instance = new RepositoryClientExInternal();
  });

  describe('addAuthorizationHeader', () => {
    test('sets Authorization: Bearer <token> on the request', () => {
      // Arrange
      const request: RequestInit = {};

      // Act
      instance.addAuthorizationHeader(request, 'abc123');

      // Assert
      const headers = new Headers(request.headers);
      expect(headers.get('Authorization')).toBe('Bearer abc123');
    });

    test('preserves headers that were already present', () => {
      // Arrange
      const request: RequestInit = { headers: { 'X-Custom': 'value' } };

      // Act
      instance.addAuthorizationHeader(request, 'abc123');

      // Assert
      const headers = new Headers(request.headers);
      expect(headers.get('Authorization')).toBe('Bearer abc123');
      expect(headers.get('X-Custom')).toBe('value');
    });
  });

  describe('beforeFetchRequestAsync', () => {
    test('resolves regionalDomain and sets the Authorization header when an accessToken is present', async () => {
      // Arrange
      lfLogin.authorization_credentials = { accessToken: 'tok-1' };
      lfLogin.account_endpoints = { regionalDomain: 'a.laserfiche.com' };
      const request: RequestInit = {};

      // Act
      const result = await instance.beforeFetchRequestAsync(
        'https://api.example.com',
        request
      );

      // Assert
      expect(result).toEqual({ regionalDomain: 'a.laserfiche.com' });
      const headers = new Headers(request.headers);
      expect(headers.get('Authorization')).toBe('Bearer tok-1');
    });

    test('rejects with "No access token" when there are no authorization_credentials', async () => {
      // Arrange
      lfLogin.authorization_credentials = undefined;
      const request: RequestInit = {};

      // Act / Assert
      await expect(
        instance.beforeFetchRequestAsync('https://api.example.com', request)
      ).rejects.toThrow('No access token');
    });
  });

  describe('afterFetchResponseAsync', () => {
    test('resolves false and does not refresh the token when status is 200', async () => {
      // Arrange
      lfLogin.refreshTokenAsync = vi.fn();
      const response = { status: 200 } as Response;
      const request: RequestInit = {};

      // Act
      const result = await instance.afterFetchResponseAsync(
        'https://api.example.com',
        response,
        request
      );

      // Assert
      expect(result).toBe(false);
      expect(lfLogin.refreshTokenAsync).not.toHaveBeenCalled();
    });

    test('resolves true and reapplies the Authorization header when refresh succeeds after a 401', async () => {
      // Arrange
      lfLogin.authorization_credentials = { accessToken: 'refreshed-tok' };
      lfLogin.refreshTokenAsync = vi.fn().mockResolvedValue(true);
      const response = { status: 401 } as Response;
      const request: RequestInit = {};

      // Act
      const result = await instance.afterFetchResponseAsync(
        'https://api.example.com',
        response,
        request
      );

      // Assert
      expect(result).toBe(true);
      expect(lfLogin.refreshTokenAsync).toHaveBeenCalledWith(true);
      const headers = new Headers(request.headers);
      expect(headers.get('Authorization')).toBe('Bearer refreshed-tok');
    });

    test('clears the current repo and resolves false when refresh fails after a 401', async () => {
      // Arrange
      lfLogin.refreshTokenAsync = vi.fn().mockResolvedValue(false);
      const clearCurrentRepo = vi.fn();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      instance.repoClient = { clearCurrentRepo } as any;
      const response = { status: 401 } as Response;
      const request: RequestInit = {};

      // Act
      const result = await instance.afterFetchResponseAsync(
        'https://api.example.com',
        response,
        request
      );

      // Assert
      expect(result).toBe(false);
      expect(clearCurrentRepo).toHaveBeenCalled();
    });
  });

  describe('getCurrentRepo', () => {
    test('resolves the repoId and repoName from listRepositories', async () => {
      // Arrange
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      instance.repoClient = {
        repositoriesClient: {
          listRepositories: vi
            .fn()
            .mockResolvedValue({ value: [{ id: 'repo-1', name: 'Main Repo' }] }),
        },
      } as any;

      // Act
      const result = await instance.getCurrentRepo();

      // Assert
      expect(result).toEqual({ repoId: 'repo-1', repoName: 'Main Repo' });
    });

    test('falls back to the id as the repoName when the repo has no name', async () => {
      // Arrange
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      instance.repoClient = {
        repositoriesClient: {
          listRepositories: vi
            .fn()
            .mockResolvedValue({ value: [{ id: 'repo-1' }] }),
        },
      } as any;

      // Act
      const result = await instance.getCurrentRepo();

      // Assert
      expect(result).toEqual({ repoId: 'repo-1', repoName: 'repo-1' });
    });

    test('rejects when the value array is empty', async () => {
      // Arrange
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      instance.repoClient = {
        repositoriesClient: {
          listRepositories: vi.fn().mockResolvedValue({ value: [] }),
        },
      } as any;

      // Act / Assert
      await expect(instance.getCurrentRepo()).rejects.toThrow(
        'Current repoId undefined.'
      );
    });

    test('rejects when repoClient is undefined', async () => {
      // Arrange
      instance.repoClient = undefined as unknown as IRepositoryApiClientExInternal;

      // Act / Assert
      await expect(instance.getCurrentRepo()).rejects.toThrow(
        'repoClient undefined.'
      );
    });
  });

  describe('createRepositoryClientAsync', () => {
    test('attaches clearCurrentRepo/getCurrentRepoId/getCurrentRepoName onto the object returned by createFromHttpRequestHandler', async () => {
      // Arrange
      const partialRepoClient = {
        repositoriesClient: { listRepositories: vi.fn() },
      };
      (
        RepositoryApiClient.createFromHttpRequestHandler as Mock
      ).mockReturnValue(partialRepoClient);

      // Act
      const result = await instance.createRepositoryClientAsync();

      // Assert
      expect(typeof result.clearCurrentRepo).toBe('function');
      expect(typeof result.getCurrentRepoId).toBe('function');
      expect(typeof result.getCurrentRepoName).toBe('function');
      expect(result).toBe(partialRepoClient);
      expect(instance.repoClient).toBe(result);
    });
  });

  describe('getCurrentRepoId caching', () => {
    test('only calls listRepositories once across two calls', async () => {
      // Arrange
      const listRepositories = vi
        .fn()
        .mockResolvedValue({ value: [{ id: 'repo-1' }] });
      (
        RepositoryApiClient.createFromHttpRequestHandler as Mock
      ).mockReturnValue({ repositoriesClient: { listRepositories } });
      const client = await instance.createRepositoryClientAsync();

      // Act
      const firstId = await client.getCurrentRepoId();
      const secondId = await client.getCurrentRepoId();

      // Assert
      expect(firstId).toBe('repo-1');
      expect(secondId).toBe('repo-1');
      expect(listRepositories).toHaveBeenCalledTimes(1);
    });

    test('calls listRepositories again after clearCurrentRepo', async () => {
      // Arrange
      const listRepositories = vi
        .fn()
        .mockResolvedValue({ value: [{ id: 'repo-1' }] });
      (
        RepositoryApiClient.createFromHttpRequestHandler as Mock
      ).mockReturnValue({ repositoriesClient: { listRepositories } });
      const client = await instance.createRepositoryClientAsync();
      await client.getCurrentRepoId();

      // Act
      client.clearCurrentRepo();
      await client.getCurrentRepoId();

      // Assert
      expect(listRepositories).toHaveBeenCalledTimes(2);
    });
  });
});
