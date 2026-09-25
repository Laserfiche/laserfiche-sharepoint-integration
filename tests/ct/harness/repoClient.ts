// Component-test props are serialized across the Playwright/browser
// boundary, so a live mock object whose async methods must resolve specific
// values (e.g. `getCurrentRepoId()` -> 'repo-1') cannot be passed in as a
// prop. Instead, harness components build this mock *inside* the browser
// bundle from a plain, serializable `scenario` string, and every call is
// recorded onto `window.__repoClientCalls` so test code can assert on it via
// `page.evaluate(() => window.__repoClientCalls)`.
import type { IRepositoryApiClientExInternal } from '../../../src/repository-client/repository-client-types';

export type RepoClientScenario =
  | 'default'
  | 'createEntry-rejects-exists'
  | 'getEntryByPath-exists'
  | 'getEntryByPath-404';

declare global {
  interface Window {
    __repoClientCalls: Array<{ method: string; args: unknown[] }>;
  }
}

function record(method: string, args: unknown[]): void {
  window.__repoClientCalls = window.__repoClientCalls ?? [];
  window.__repoClientCalls.push({ method, args });
}

export function makeRepoClient(
  scenario: RepoClientScenario,
  importDelayMs?: number
): IRepositoryApiClientExInternal {
  const client = {
    getCurrentRepoId: async (): Promise<string> => {
      record('getCurrentRepoId', []);
      return 'repo-1';
    },
    entriesClient: {
      createEntry: async (...args: unknown[]): Promise<unknown> => {
        record('createEntry', args);
        if (scenario === 'createEntry-rejects-exists') {
          throw new Error('Object already exists');
        }
        return { id: '99', name: 'New Entry' };
      },
      getEntryByPath: async (...args: unknown[]): Promise<unknown> => {
        record('getEntryByPath', args);
        if (scenario === 'getEntryByPath-exists') {
          return { entry: { id: '5' } };
        }
        const notFound = new Error('Not found') as Error & { status: number };
        notFound.status = 404;
        throw notFound;
      },
      importEntry: async (...args: unknown[]): Promise<unknown> => {
        record('importEntry', args);
        if (importDelayMs) {
          await new Promise((r) => setTimeout(r, importDelayMs));
        }
        return { id: 100, name: 'imported' };
      },
      setTags: async (...args: unknown[]): Promise<unknown> => {
        record('setTags', args);
        return {};
      },
    },
    tagDefinitionsClient: {
      listTagDefinitionsForEach: async (args: {
        callback: (response: { value: Array<{ name: string }> }) => Promise<boolean>;
        repositoryId: string;
      }): Promise<void> => {
        record('listTagDefinitionsForEach', [{ repositoryId: args.repositoryId }]);
        await args.callback({ value: [{ name: 'Contract' }, { name: 'Reviewed' }] });
      },
    },
  } as unknown as IRepositoryApiClientExInternal;
  return client;
}
