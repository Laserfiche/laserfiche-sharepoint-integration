import * as React from 'react';
import ManageMappingsPage from '../../../src/webparts/laserficheAdminConfiguration/components/ManageMappingsPage/ManageMappingsPage';

declare global {
  interface Window {
    __postCalls: Array<{ url: string; body: unknown }>;
  }
}

const CONTENT_TYPES = ['Document', 'Invoice'];
const LF_PROFILES = ['ProfileA', 'ProfileB'];

export type MappingsScenario = 'with-existing-mapping' | 'empty';

function existingMappingsJson(scenario: MappingsScenario): string {
  if (scenario === 'with-existing-mapping') {
    return JSON.stringify([
      {
        id: 'm1',
        SharePointContentType: 'Invoice',
        LaserficheContentType: 'ProfileA',
        toggle: true,
      },
    ]);
  }
  return JSON.stringify([]);
}

export default function ManageMappingsHarness(props: { scenario?: MappingsScenario }): JSX.Element {
  const scenario = props.scenario ?? 'with-existing-mapping';

  const realFetch = window.fetch.bind(window);
  window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = typeof input === 'string' ? input : input.toString();
    if (url.includes('_api/web/contenttypes')) {
      return new Response(JSON.stringify({ value: CONTENT_TYPES.map((name) => ({ Name: name })) }));
    }
    if (
      url.includes(`Title%20eq%20'ManageConfigurations'`) ||
      url.includes(`Title eq 'ManageConfigurations'`)
    ) {
      return new Response(
        JSON.stringify({
          value: [
            {
              Id: 'cfg-1',
              JsonValue: JSON.stringify(LF_PROFILES.map((name) => ({ ConfigurationName: name }))),
            },
          ],
        })
      );
    }
    if (url.includes(`Title%20eq%20'ManageMapping'`) || url.includes(`Title eq 'ManageMapping'`)) {
      return new Response(
        JSON.stringify({ value: [{ Id: 'map-1', JsonValue: existingMappingsJson(scenario) }] })
      );
    }
    return realFetch(input, init);
  };

  const context = {
    pageContext: { web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/Test' } },
    spHttpClient: {
      post: async (
        url: string,
        _config: unknown,
        options: { body: unknown }
      ): Promise<Response> => {
        window.__postCalls = window.__postCalls ?? [];
        window.__postCalls.push({ url, body: options.body });
        return new Response(JSON.stringify({}));
      },
    },
  } as never;

  return <ManageMappingsPage context={context} repoClient={{} as never} isLoggedIn={true} />;
}
