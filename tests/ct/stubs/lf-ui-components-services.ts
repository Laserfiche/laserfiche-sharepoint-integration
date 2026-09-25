// Browser-side stand-in for @laserfiche/lf-ui-components-services. The real
// package is CJS with no `type: module`, sitting next to ESM-only siblings
// (lf-js-utils, lf-repository-api-client-v2) in the same import graph -- the
// combination most likely to trip up Vite's dependency pre-bundling. Since
// component tests never talk to a real repository, a plain stub matching the
// existing unit-test manual mock (src/__mocks__/@laserfiche/lf-ui-components-services.js)
// is both simpler and safer than trying to load the real thing.
import type { IRepositoryApiClientExInternal } from '../../../src/repository-client/repository-client-types';

export class LfRepoTreeNodeService {
  public viewableEntryTypes: string[] = [];
  public columnIds: string[] = [];
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  constructor(_repoClient?: unknown) {}
}

export class LfFieldsService {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  constructor(_repoClient?: unknown) {}

  async getAvailableTemplatesAsync(): Promise<unknown[]> {
    await new Promise((resolve) => setTimeout(resolve, window.__templatesLoadDelayMs ?? 0));
    return [];
  }
}

// Lists tags through the repo client it's given (the CT harness's fake
// tagDefinitionsClient), so a test proves the dialog handed the real client to
// the tags service.
export class LfRepoTagsService {
  constructor(private repoClient: IRepositoryApiClientExInternal) {}

  async getTagDefinitions(): Promise<Array<{ displayName: string }>> {
    const tags: Array<{ displayName: string }> = [];
    await this.repoClient.tagDefinitionsClient.listTagDefinitionsForEach({
      repositoryId: await this.repoClient.getCurrentRepoId(),
      callback: async (response) => {
        tags.push(...(response.value ?? []).map((tag) => ({ displayName: tag.name ?? '' })));
        return true;
      },
    });
    return tags;
  }
}
