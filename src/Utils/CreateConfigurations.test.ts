// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
}));

import type { Mock } from 'vitest';
import { CreateConfigurations } from './CreateConfigurations';
import { BaseComponentContext } from '@microsoft/sp-component-base';

const WEB_ABSOLUTE_URL = 'https://contoso.sharepoint.com/sites/Test';
const WEB_TITLE = 'Test Site';

function buildContext(getMock: Mock, postMock: Mock): BaseComponentContext {
  return {
    spHttpClient: {
      get: getMock,
      post: postMock,
    },
    pageContext: {
      web: {
        absoluteUrl: WEB_ABSOLUTE_URL,
        title: WEB_TITLE,
      },
    },
  } as unknown as BaseComponentContext;
}

function mockFetchByUrl(
  handlers: { match: string; response: unknown }[]
): Mock {
  return vi.fn((url: string) => {
    const handler = handlers.find((h) => url.includes(h.match));
    if (handler) {
      return Promise.resolve(handler.response);
    }
    return Promise.reject(new Error(`Unexpected fetch URL: ${url}`));
  });
}

const jsonResponse = (value: unknown): { json: () => Promise<unknown> } => ({
  json: () => Promise.resolve(value),
});

describe('CreateConfigurations.ensureAdminConfigListCreatedAsync', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  test('does nothing when the list already exists (status 200)', async () => {
    // Arrange
    const get = vi.fn().mockResolvedValue({ status: 200 });
    const post = vi.fn();
    const context = buildContext(get, post);
    window.fetch = vi.fn();

    // Act
    await CreateConfigurations.ensureAdminConfigListCreatedAsync(context);

    // Assert
    expect(get).toHaveBeenCalledTimes(1);
    expect(post).not.toHaveBeenCalled();
    expect(window.fetch).not.toHaveBeenCalled();
  });

  test('creates the list, columns and security when the list is missing (status 404)', async () => {
    // Arrange
    const get = vi.fn().mockResolvedValue({ status: 404 });
    const post = vi
      .fn()
      .mockResolvedValue(jsonResponse({ Title: 'LaserficheAdminConfiguration' }));
    const context = buildContext(get, post);

    window.fetch = mockFetchByUrl([
      {
        match: '/_api/contextinfo',
        response: jsonResponse({
          d: { GetContextWebInformation: { FormDigestValue: 'digest-123' } },
        }),
      },
      {
        match: '/fields',
        response: jsonResponse({}),
      },
      {
        match: "/sitegroups/getbyname('Test Site Members')/id",
        response: jsonResponse({ d: { Id: 'group-1' } }),
      },
      {
        match: "/roledefinitions/getbyname('Read')/id",
        response: jsonResponse({ d: { Id: 'role-1' } }),
      },
      {
        match: '/breakroleinheritance(true)',
        response: jsonResponse({}),
      },
      {
        match: '/roleassignments/getbyprincipalid(',
        response: jsonResponse({}),
      },
      {
        match: '/roleassignments/addroleassignment(principalid=',
        response: jsonResponse({}),
      },
    ]);

    // Act
    await CreateConfigurations.ensureAdminConfigListCreatedAsync(context);

    // Assert
    // (1) form digest fetch
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining('/_api/contextinfo'),
      expect.any(Object)
    );

    // (2) spHttpClient.post to create the list
    expect(post).toHaveBeenCalledWith(
      expect.stringMatching(/\/_api\/web\/lists$/),
      expect.anything(),
      expect.objectContaining({
        body: expect.stringContaining('"BaseTemplate":100'),
      })
    );
    expect(post.mock.calls[0][2].body).toContain(
      '"Title":"LaserficheAdminConfiguration"'
    );

    // (3) fields fetch
    const fieldsCall = (window.fetch as Mock).mock.calls.find(([url]) =>
      url.includes('/fields')
    );
    expect(fieldsCall).toBeDefined();
    expect(fieldsCall![1].body).toContain('"FieldTypeKind":3');

    // (4) sitegroups fetch, asserting the exact group name string
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining("/sitegroups/getbyname('Test Site Members')/id"),
      expect.any(Object)
    );

    // (5) roledefinitions fetch
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining("/roledefinitions/getbyname('Read')/id"),
      expect.any(Object)
    );

    // (6) breakroleinheritance
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining('/breakroleinheritance(true)'),
      expect.objectContaining({ method: 'POST' })
    );

    // (7) delete current role assignment for the group
    const deleteRoleCall = (window.fetch as Mock).mock.calls.find(
      ([url]) => url.includes('/roleassignments/getbyprincipalid(')
    );
    expect(deleteRoleCall).toBeDefined();
    expect(deleteRoleCall![1].headers['X-HTTP-Method']).toBe('DELETE');

    // (8) add new role assignment
    expect(window.fetch).toHaveBeenCalledWith(
      expect.stringContaining('/roleassignments/addroleassignment(principalid='),
      expect.any(Object)
    );
  });

  test('does not perform any writes for an unexpected status', async () => {
    // Arrange
    const get = vi.fn().mockResolvedValue({ status: 500 });
    const post = vi.fn();
    const context = buildContext(get, post);
    window.fetch = vi.fn();

    // Act
    await CreateConfigurations.ensureAdminConfigListCreatedAsync(context);

    // Assert
    expect(post).not.toHaveBeenCalled();
    expect(window.fetch).not.toHaveBeenCalled();
  });

  test('swallows an error thrown by spHttpClient.post when creating the list', async () => {
    // Arrange
    const get = vi.fn().mockResolvedValue({ status: 404 });
    const post = vi.fn().mockRejectedValue(new Error('list creation failed'));
    const context = buildContext(get, post);
    const consoleErrorSpy = vi
      .spyOn(console, 'error')
      .mockImplementation(() => undefined);

    window.fetch = mockFetchByUrl([
      {
        match: '/_api/contextinfo',
        response: jsonResponse({
          d: { GetContextWebInformation: { FormDigestValue: 'digest-123' } },
        }),
      },
      {
        match: "/sitegroups/getbyname('Test Site Members')/id",
        response: jsonResponse({ d: { Id: 'group-1' } }),
      },
      {
        match: "/roledefinitions/getbyname('Read')/id",
        response: jsonResponse({ d: { Id: 'role-1' } }),
      },
      { match: '/breakroleinheritance(true)', response: jsonResponse({}) },
      {
        match: '/roleassignments/getbyprincipalid(',
        response: jsonResponse({}),
      },
      {
        match: '/roleassignments/addroleassignment(principalid=',
        response: jsonResponse({}),
      },
    ]);

    // Act / Assert: the public method still resolves rather than rejecting.
    await expect(
      CreateConfigurations.ensureAdminConfigListCreatedAsync(context)
    ).resolves.toBeUndefined();
    expect(consoleErrorSpy).toHaveBeenCalled();
  });

  test('proceeds without throwing when the contextinfo fetch rejects', async () => {
    // Arrange
    const get = vi.fn().mockResolvedValue({ status: 404 });
    const post = vi
      .fn()
      .mockResolvedValue(jsonResponse({ Title: 'LaserficheAdminConfiguration' }));
    const context = buildContext(get, post);

    window.fetch = vi.fn((url: string) => {
      if (url.includes('/_api/contextinfo')) {
        return Promise.reject(new Error('network error'));
      }
      if (url.includes("/sitegroups/getbyname('Test Site Members')/id")) {
        return Promise.resolve(jsonResponse({ d: { Id: 'group-1' } }));
      }
      if (url.includes("/roledefinitions/getbyname('Read')/id")) {
        return Promise.resolve(jsonResponse({ d: { Id: 'role-1' } }));
      }
      return Promise.resolve(jsonResponse({}));
    }) as unknown as typeof fetch;

    // Act / Assert
    await expect(
      CreateConfigurations.ensureAdminConfigListCreatedAsync(context)
    ).resolves.toBeUndefined();
  });
});
