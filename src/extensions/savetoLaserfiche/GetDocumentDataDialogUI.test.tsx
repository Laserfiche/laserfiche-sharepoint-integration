// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

vi.mock('@microsoft/sp-dialog', () => ({
  BaseDialog: class {},
}));
vi.mock('./SaveToLaserficheDialog', () => ({
  SaveToLaserficheDialog: vi.fn(),
}));
vi.mock('@microsoft/sp-http', () => {
  return {
    SPHttpClient: {
      configurations: {
        v1: {},
      },
    },
  };
});
vi.mock('@laserfiche/lf-repository-api-client-v2', () => ({
  TemplateFieldDefinition: vi.fn().mockImplementation(function ({ name }) {
    return { name };
  }),
  FieldToUpdate: vi.fn().mockImplementation(function ({ name, values }) {
    return { name, values };
  }),
}));

import { render, waitFor } from '@testing-library/react';
import * as React from 'react';
import { GetDocumentDialogData } from './GetDocumentDataDialogUI';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import { TemplateFieldDefinition } from '@laserfiche/lf-repository-api-client-v2';
import { ISPDocumentData } from '../../Utils/Types';

describe('GetDocumentDataDialog', () => {
  test('getDocumentDataDialog correctly matches SP field data to Laserfiche fields', async () => {
    const mockContext: BaseComponentContext = {
      spHttpClient: {
        get: vi.fn((url) => {
          if (url.includes('FieldValuesForEdit')) {
            return Promise.resolve({
              ok: true,
              json: () =>
                Promise.resolve({
                  Title: 'Document Title',
                  Author: 'John Doe',
                  Created: '2021-01-01T00:00:00Z',
                }),
            } as SPHttpClientResponse);
          } else if (url.includes('ManageConfigurations')) {
            return Promise.resolve({
              json: () =>
                Promise.resolve({
                  value: [
                    {
                      JsonValue: JSON.stringify([
                        {
                          ConfigurationName: 'TestConfig',
                          DocumentName: '%(DocumentName)',
                          selectedTemplateName: 'Test Template',
                          selectedFolder: { id: '1234', path: '/Test Folder' },
                          Action: ActionTypes.COPY,
                          mappedFields: [
                            {
                              id: 'dd',
                              lfField: new TemplateFieldDefinition({
                                name: 'testField',
                              }),
                              spField: {
                                Title: 'Title',
                                TypeAsString: 'string',
                                InternalName: 'Title',
                                EntityPropertyName:
                                  'Title',
                              },
                            },
                          ],
                        },
                      ]),
                    },
                  ],
                }),
            });
          } else if (url.includes('ManageMapping')) {
            return Promise.resolve({
              json: () =>
                Promise.resolve({
                  value: [
                    {
                      JsonValue: JSON.stringify([
                        {
                          id: 'dd',
                          SharePointContentType: 'TestContentType',
                          LaserficheContentType: 'TestConfig',
                          toggle: true,
                        },
                        {
                          id: 'string',
                          SharePointContentType: 'MockContentType',
                          LaserficheContentType: 'mockConfigMapped',
                          toggle: true,
                        },
                      ]),
                    },
                  ],
                }),
            });
          }
          return Promise.reject(new Error('Unknown URL'));
        }),
      } as unknown as SPHttpClient,
      pageContext: {
        list: {
          title: 'Test List',
        },
        web: {
          absoluteUrl: 'https://test.sharepoint.com',
        },
      },
    } as BaseComponentContext;

    let spDocData: ISPDocumentData;

    window.fetch = vi
      .fn()
      .mockResolvedValue({ json: () => Promise.resolve({ value: [] }) });
    Object.defineProperty(window, 'localStorage', {
      value: {
        setItem: vi.fn().mockImplementation((key, value) => {
          spDocData = JSON.parse(value);
        }),
      },
      writable: true,
    });

    const spFileInfo = {
      fileName: 'Document Name Test',
      spContentType: 'TestContentType',
      spFileUrl: 'https://sharepoint-fileurl.com',
      fileId: '1234',
    };

    render(
      <GetDocumentDialogData
        showSaveToDialog={vi.fn()}
        handleCancelDialog={vi.fn()}
        spFileInfo={spFileInfo}
        context={mockContext}
      />
    );
    await waitFor(() => {
      expect(spDocData.documentName).toEqual('%(DocumentName)');
      expect(spDocData.templateName).toEqual('Test Template');
      expect(spDocData.metadata!.fields).toEqual([
        { name: 'testField', values: ['Document Title'] },
      ]);
    });
  });

  test('getDocumentDataDialog correctly matches SP field data to Laserfiche fields with special characters', async () => {
    const mockContext: BaseComponentContext = {
      spHttpClient: {
        get: vi.fn((url) => {
          if (url.includes('FieldValuesForEdit')) {
            return Promise.resolve({
              ok: true,
              json: () =>
                Promise.resolve({
                  OData__x005f_x0040_x005f__x005f_x0021_x005f_TestSpecial:
                    'Document Title',
                  Author: 'John Doe',
                  Created: '2021-01-01T00:00:00Z',
                }),
            } as SPHttpClientResponse);
          } else if (url.includes('ManageConfigurations')) {
            return Promise.resolve({
              json: () =>
                Promise.resolve({
                  value: [
                    {
                      JsonValue: JSON.stringify([
                        {
                          ConfigurationName: 'TestConfig',
                          DocumentName: '%(DocumentName)',
                          selectedTemplateName: 'Test Template',
                          selectedFolder: { id: '1234', path: '/Test Folder' },
                          Action: ActionTypes.COPY,
                          mappedFields: [
                            {
                              id: 'dd',
                              lfField: new TemplateFieldDefinition({
                                name: 'testField',
                              }),
                              spField: {
                                Title: 'Title',
                                TypeAsString: 'string',
                                InternalName: '_x0040__x0021_TestSpecial',
                                EntityPropertyName:
                                  'OData__x0040__x0021_TestSpecial',
                              },
                            },
                          ],
                        },
                      ]),
                    },
                  ],
                }),
            });
          } else if (url.includes('ManageMapping')) {
            return Promise.resolve({
              json: () =>
                Promise.resolve({
                  value: [
                    {
                      JsonValue: JSON.stringify([
                        {
                          id: 'dd',
                          SharePointContentType: 'TestContentType',
                          LaserficheContentType: 'TestConfig',
                          toggle: true,
                        },
                        {
                          id: 'string',
                          SharePointContentType: 'MockContentType',
                          LaserficheContentType: 'mockConfigMapped',
                          toggle: true,
                        },
                      ]),
                    },
                  ],
                }),
            });
          }
          return Promise.reject(new Error('Unknown URL'));
        }),
      } as unknown as SPHttpClient,
      pageContext: {
        list: {
          title: 'Test List',
        },
        web: {
          absoluteUrl: 'https://test.sharepoint.com',
        },
      },
    } as BaseComponentContext;

    let spDocData: ISPDocumentData;

    window.fetch = vi
      .fn()
      .mockResolvedValue({ json: () => Promise.resolve({ value: [] }) });
    Object.defineProperty(window, 'localStorage', {
      value: {
        setItem: vi.fn().mockImplementation((key, value) => {
          spDocData = JSON.parse(value);
        }),
      },
      writable: true,
    });

    const spFileInfo = {
      fileName: 'Document Name Test',
      spContentType: 'TestContentType',
      spFileUrl: 'https://sharepoint-fileurl.com',
      fileId: '1234',
    };

    render(
      <GetDocumentDialogData
        showSaveToDialog={vi.fn()}
        handleCancelDialog={vi.fn()}
        spFileInfo={spFileInfo}
        context={mockContext}
      />
    );
    await waitFor(() => {
      expect(spDocData.documentName).toEqual('%(DocumentName)');
      expect(spDocData.templateName).toEqual('Test Template');
      expect(spDocData.metadata!.fields).toEqual([
        { name: 'testField', values: ['Document Title'] },
      ]);
    });
  });
});
