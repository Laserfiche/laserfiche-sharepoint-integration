import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';
import {
  ADMIN_CONFIGURATION_LIST,
} from '../webparts/constants';
import { getSPListURL } from './Funcs';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export class CreateConfigurations {
  public static async ensureAdminConfigListCreatedAsync(
    context: BaseComponentContext
  ): Promise<void> {
    const listUrl: string = getSPListURL(context, ADMIN_CONFIGURATION_LIST);
    const response = await context.spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      return;
    }
    if (response.status === 404) {
      await CreateConfigurations.createAdminConfigListAsync(context);
    }
  }

  private static async createAdminConfigListAsync(
    context: BaseComponentContext
  ): Promise<void> {
    const url: string = context.pageContext.web.absoluteUrl + '/_api/web/lists';
    const listDefinition = {
      Title: ADMIN_CONFIGURATION_LIST,
      Description: 'My description',
      BaseTemplate: 100,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listDefinition),
    };
    const responses: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      spHttpClientOptions
    );
    const adminConfigList = await responses.json();
    const listTitle = adminConfigList.Title;
    const formDigestValue = await this.getFormDigestValueAsync(context);
    await this.createColumnsAsync(context, listTitle, formDigestValue);
  }

  private static async getFormDigestValueAsync(context: BaseComponentContext): Promise<string> {
    try {
      const res = await fetch(
        context.pageContext.web.absoluteUrl + '/_api/contextinfo',
        {
          method: 'POST',
          headers: { accept: 'application/json;odata=verbose' },
        }
      );
      const contextInfo = await res.json();
      const FormDigestValue =
        contextInfo.d.GetContextWebInformation.FormDigestValue;
      return FormDigestValue;
    } catch {
      // TODO handle
    }
  }

  private static async createColumnsAsync(
    context: BaseComponentContext,
    listTitle: string,
    formDigestValue: string
  ): Promise<void> {
    const siteUrl: string = getSPListURL(context, listTitle) + '/fields';
    try {
      await fetch(siteUrl, {
        method: 'POST',
        body: JSON.stringify({
          __metadata: { type: 'SP.Field' },
          Title: 'JsonValue',
          FieldTypeKind: 3,
        }),
        headers: {
          accept: 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
          'X-RequestDigest': formDigestValue,
        },
      });
      console.log('Fields created');
    } catch {
      console.log('Error!');
    }
  }
}
