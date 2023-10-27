// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from '@microsoft/sp-http';
import { LASERFICHE_ADMIN_CONFIGURATION_NAME } from '../webparts/constants';
import { getSPListURL } from './Funcs';
import { BaseComponentContext } from '@microsoft/sp-component-base';

const targetRoleDefinitionName = 'Read';

export class CreateConfigurations {
  public static async ensureAdminConfigListCreatedAsync(
    context: BaseComponentContext
  ): Promise<void> {
    const listUrl: string = getSPListURL(
      context,
      LASERFICHE_ADMIN_CONFIGURATION_NAME
    );
    const response = await context.spHttpClient.get(
      listUrl,
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      return;
    }
    if (response.status === 404) {
      const formDigestValue = await this.getFormDigestValueAsync(context);
      const listTitle = await CreateConfigurations.createAdminConfigListAsync(
        context,
        formDigestValue
      );
      await CreateConfigurations.updateAdminConfigListSecurityAsync(
        context,
        formDigestValue,
        listTitle
      );
    }
  }

  private static async createAdminConfigListAsync(
    context: BaseComponentContext,
    formDigestValue: string
  ): Promise<string> {
    try {
      const url: string = context.pageContext.web.absoluteUrl + '/_api/web/lists';
      const listDefinition = {
        Title: LASERFICHE_ADMIN_CONFIGURATION_NAME,
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
      await this.createColumnsAsync(context, listTitle, formDigestValue);
      return listTitle;
    }
    catch (err) {
      console.error(`Error when creating LaserficheAdminConfiguration List: ${err}`);
    }
  }

  private static async updateAdminConfigListSecurityAsync(
    context: BaseComponentContext,
    formDigestValue: string,
    listTitle: string
  ): Promise<void> {
    const groupId = await CreateConfigurations.getMembersGroupIdAsync(context);

    const targetRoleDefinitionId =
      await CreateConfigurations.getTargetRoleDefinitionIdAsync(context);

    await CreateConfigurations.breakRoleInheritanceOfListAsync(
      context,
      formDigestValue,
      listTitle
    );

    await CreateConfigurations.deleteCurrentRoleForGroupAsync(
      context,
      formDigestValue,
      groupId,
      listTitle
    );

    await CreateConfigurations.setNewPermissionsForGroupAsync(
      context,
      formDigestValue,
      groupId,
      targetRoleDefinitionId,
      listTitle
    );
  }

  private static async getMembersGroupIdAsync(
    context: BaseComponentContext
  ): Promise<string> {
    const membersGroupName = `${context.pageContext.web.title} Members`;

    const res = await fetch(
      context.pageContext.web.absoluteUrl +
        "/_api/web/sitegroups/getbyname('" +
        membersGroupName +
        "')/id",
      {
        method: 'GET',
        headers: { accept: 'application/json;odata=verbose' },
      }
    );

    const siteGroup = await res.json();
    const groupId = siteGroup.d.Id;
    return groupId;
  }

  private static async getTargetRoleDefinitionIdAsync(
    context: BaseComponentContext
  ): Promise<string> {
    const res = await fetch(
      context.pageContext.web.absoluteUrl +
        "/_api/web/roledefinitions/getbyname('" +
        targetRoleDefinitionName +
        "')/id",
      {
        method: 'GET',
        headers: { accept: 'application/json;odata=verbose' },
      }
    );

    const targetRole = await res.json();
    const targetRoleDefinitionId = targetRole.d.Id;
    return targetRoleDefinitionId;
  }

  private static async breakRoleInheritanceOfListAsync(
    context: BaseComponentContext,
    formDigestValue: string,
    listTitle: string
  ): Promise<void> {
    await fetch(
      context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" +
        listTitle +
        "')/breakroleinheritance(true)",
      {
        method: 'POST',
        headers: {
          accept: 'application/json;odata=verbose',
          'X-RequestDigest': formDigestValue,
        },
      }
    );
  }

  private static async deleteCurrentRoleForGroupAsync(
    context: BaseComponentContext,
    formDigestValue: string,
    groupId: string,
    listTitle: string
  ): Promise<void> {
    await fetch(
      context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" +
        listTitle +
        "')/roleassignments/getbyprincipalid(" +
        groupId +
        ')',
      {
        method: 'POST',
        headers: {
          accept: 'application/json;odata=verbose',
          'X-HTTP-Method': 'DELETE',
          'X-RequestDigest': formDigestValue,
        },
      }
    );
  }

  private static async setNewPermissionsForGroupAsync(
    context: BaseComponentContext,
    formDigestValue: string,
    groupId: string,
    targetRoleDefinitionId: string,
    listTitle: string
  ): Promise<void> {
    await fetch(
      context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" +
        listTitle +
        "')/roleassignments/addroleassignment(principalid=" +
        groupId +
        ',roledefid=' +
        targetRoleDefinitionId +
        ')',
      {
        method: 'POST',
        headers: {
          accept: 'application/json;odata=verbose',
          'X-RequestDigest': formDigestValue,
        },
      }
    );
  }

  private static async getFormDigestValueAsync(
    context: BaseComponentContext
  ): Promise<string> {
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
  }
}
