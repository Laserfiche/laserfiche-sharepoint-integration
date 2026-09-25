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
import * as React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import ManageMappingsPage from './ManageMappingsPage';
import { IManageMappingsPageProps } from './IManageMappingsPageProps';
import { ProfileMappingConfiguration } from '../../../../Utils/Types';
import { getSPListURL } from '../../../../Utils/Funcs';
import {
  LASERFICHE_ADMIN_CONFIGURATION_NAME,
  MANAGE_CONFIGURATIONS,
  MANAGE_MAPPING,
} from '../../../constants';

const ABSOLUTE_URL = 'https://contoso.sharepoint.com/sites/Test';

interface MockMappingListItem {
  Id: string;
  mappings: ProfileMappingConfiguration[];
}

const SINGLE_MAPPING_ITEM: MockMappingListItem = {
  Id: 'item-1',
  mappings: [
    {
      id: 'm1',
      SharePointContentType: 'Invoice',
      LaserficheContentType: 'ProfileA',
      toggle: true,
    },
  ],
};

const TWO_MAPPING_ITEM: MockMappingListItem = {
  Id: 'item-1',
  mappings: [
    {
      id: 'm1',
      SharePointContentType: 'Invoice',
      LaserficheContentType: 'ProfileA',
      toggle: true,
    },
    {
      id: 'm2',
      SharePointContentType: 'Document',
      LaserficheContentType: 'ProfileB',
      toggle: true,
    },
  ],
};

interface FetchMockOptions {
  contentTypeNames?: string[];
  profileNames?: string[];
  mappingListItem?: MockMappingListItem;
  rejectContentTypes?: Error;
  rejectMappings?: Error;
}

function buildFetchMock(options: FetchMockOptions = {}): Mock {
  const contentTypeNames = options.contentTypeNames ?? ['Document', 'Invoice'];
  const profileNames = options.profileNames ?? ['ProfileA', 'ProfileB'];

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return vi.fn((url: string): Promise<any> => {
    if (url.includes('/_api/web/contenttypes')) {
      if (options.rejectContentTypes) {
        return Promise.reject(options.rejectContentTypes);
      }
      return Promise.resolve({
        json: () =>
          Promise.resolve({
            value: contentTypeNames.map((name) => ({ Name: name })),
          }),
      });
    }
    if (url.includes(`Title eq '${MANAGE_CONFIGURATIONS}'`)) {
      return Promise.resolve({
        json: () =>
          Promise.resolve({
            value: [
              {
                Id: 'config-1',
                JsonValue: JSON.stringify(
                  profileNames.map((name) => ({ ConfigurationName: name }))
                ),
              },
            ],
          }),
      });
    }
    if (url.includes(`Title eq '${MANAGE_MAPPING}'`)) {
      if (options.rejectMappings) {
        return Promise.reject(options.rejectMappings);
      }
      if (!options.mappingListItem) {
        return Promise.resolve({ json: () => Promise.resolve({ value: [] }) });
      }
      return Promise.resolve({
        json: () =>
          Promise.resolve({
            value: [
              {
                Id: options.mappingListItem.Id,
                JsonValue: JSON.stringify(options.mappingListItem.mappings),
              },
            ],
          }),
      });
    }
    return Promise.reject(new Error(`Unhandled fetch URL: ${url}`));
  });
}

function buildContext(
  post: Mock = vi.fn().mockResolvedValue({ ok: true })
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
): any {
  return {
    pageContext: { web: { absoluteUrl: ABSOLUTE_URL } },
    spHttpClient: { post },
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function renderPage(context: any): ReturnType<typeof render> {
  const props = {
    context,
    repoClient: {},
    isLoggedIn: true,
  } as unknown as IManageMappingsPageProps;
  return render(<ManageMappingsPage {...props} />);
}

function getSelects(): HTMLSelectElement[] {
  return screen.getAllByRole('combobox') as HTMLSelectElement[];
}

function queryAllSelects(): HTMLSelectElement[] {
  return screen.queryAllByRole('combobox') as HTMLSelectElement[];
}

const listUrl = getSPListURL(buildContext(), LASERFICHE_ADMIN_CONFIGURATION_NAME);

describe('ManageMappingsPage', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('Initial load', () => {
    test('loads content types, profiles, and renders the existing mapping row', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext());

      await waitFor(() => expect(getSelects()).toHaveLength(2));

      const [spSelect, lfSelect] = getSelects();

      const spOptions = Array.from(spSelect.options).map((o) => ({
        value: o.value,
        text: o.text,
      }));
      expect(spOptions).toEqual([
        { value: 'Select', text: 'Select' },
        { value: 'DEFAULT', text: '[Default]' },
        { value: 'Document', text: 'Document' },
        { value: 'Invoice', text: 'Invoice' },
      ]);

      const lfOptions = Array.from(lfSelect.options).map((o) => ({
        value: o.value,
        text: o.text,
      }));
      expect(lfOptions).toEqual([
        { value: 'Select', text: 'Select' },
        { value: 'ProfileA', text: 'ProfileA' },
        { value: 'ProfileB', text: 'ProfileB' },
      ]);

      expect(spSelect).toBeDisabled();
      expect(lfSelect).toBeDisabled();
      expect(spSelect.value).toBe('Invoice');
      expect(lfSelect.value).toBe('ProfileA');

      expect(screen.getByText('edit')).toBeInTheDocument();
      expect(screen.getByText('delete')).toBeInTheDocument();
      expect(screen.queryByText('save')).not.toBeInTheDocument();
    });

    test('renders zero mapping rows when the ManageMapping list has no items, without erroring', async () => {
      const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
      window.fetch = buildFetchMock({ mappingListItem: undefined });
      renderPage(buildContext());

      await waitFor(() => expect((window.fetch as Mock).mock.calls.length).toBe(3));

      expect(queryAllSelects()).toHaveLength(0);
      expect(screen.getByText('Content Type Mappings Laserfiche')).toBeInTheDocument();
      expect(consoleErrorSpy).not.toHaveBeenCalled();
    });

    test('logs and does not crash when a fetch call rejects during initial load', async () => {
      const consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
      window.fetch = buildFetchMock({
        rejectContentTypes: new Error('network down'),
      });
      renderPage(buildContext());

      await waitFor(() => {
        expect(consoleErrorSpy).toHaveBeenCalledWith(
          expect.stringContaining('Error getting mappings:')
        );
      });
      expect(screen.getByText('Content Type Mappings Laserfiche')).toBeInTheDocument();
    });
  });

  describe('Add / validate / save', () => {
    test('clicking Add adds a new editable row with Select defaults and a save icon', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext());
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('Add'));

      await waitFor(() => expect(getSelects()).toHaveLength(4));
      const [, , newSpSelect, newLfSelect] = getSelects();
      expect(newSpSelect.value).toBe('Select');
      expect(newLfSelect.value).toBe('Select');
      expect(newSpSelect).not.toBeDisabled();
      expect(newLfSelect).not.toBeDisabled();

      expect(screen.getAllByText('save')).toHaveLength(1);
      expect(screen.getAllByText('edit')).toHaveLength(1);
      expect(screen.getAllByText('delete')).toHaveLength(2);
    });

    test('clicking save with no SharePoint content type selected shows validation and does not save', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(4));

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => {
        expect(
          screen.getByText(
            'Please select a content type from the SharePoint Content Type drop down'
          )
        ).toBeInTheDocument();
      });
      expect(post).not.toHaveBeenCalled();
    });

    test('clicking save with SharePoint set but Laserfiche left on Select shows validation and does not save', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(4));
      const [, , newSpSelect] = getSelects();
      fireEvent.change(newSpSelect, { target: { value: 'Document' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => {
        expect(
          screen.getByText('Please select a content type from the Laserfiche Profile dropdown')
        ).toBeInTheDocument();
      });
      expect(post).not.toHaveBeenCalled();
    });

    test('saving a new mapping when no ManageMapping item exists creates a list item without MERGE headers', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: undefined });
      renderPage(buildContext(post));

      await waitFor(() => expect((window.fetch as Mock).mock.calls.length).toBe(3));

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(2));
      const [spSelect, lfSelect] = getSelects();
      fireEvent.change(spSelect, { target: { value: 'Document' } });
      fireEvent.change(lfSelect, { target: { value: 'ProfileA' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => expect(post).toHaveBeenCalledTimes(1));
      const [url, , requestOptions] = post.mock.calls[0];
      expect(url).toBe(`${listUrl}/items`);
      expect(requestOptions.headers['IF-MATCH']).toBeUndefined();
      expect(requestOptions.headers['X-HTTP-Method']).toBeUndefined();
      const body = JSON.parse(requestOptions.body);
      expect(body.Title).toBe(MANAGE_MAPPING);
      const jsonValue = JSON.parse(body.JsonValue);
      expect(jsonValue).toHaveLength(1);
      expect(jsonValue[0].SharePointContentType).toBe('Document');
      expect(jsonValue[0].LaserficheContentType).toBe('ProfileA');

      await waitFor(() => expect(getSelects()[0]).toBeDisabled());
      expect(screen.queryByText(/Please select/)).not.toBeInTheDocument();
    });

    test('saving a new mapping when a ManageMapping item exists merges into the JsonValue array with MERGE headers', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(4));
      const [, , newSpSelect, newLfSelect] = getSelects();
      fireEvent.change(newSpSelect, { target: { value: 'Document' } });
      fireEvent.change(newLfSelect, { target: { value: 'ProfileB' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => expect(post).toHaveBeenCalledTimes(1));
      const [url, , requestOptions] = post.mock.calls[0];
      expect(url).toBe(`${listUrl}/items(${SINGLE_MAPPING_ITEM.Id})`);
      expect(requestOptions.headers['IF-MATCH']).toBe('*');
      expect(requestOptions.headers['X-HTTP-Method']).toBe('MERGE');
      const body = JSON.parse(requestOptions.body);
      const jsonValue = JSON.parse(body.JsonValue) as ProfileMappingConfiguration[];
      expect(jsonValue).toHaveLength(2);
      expect(jsonValue.find((m) => m.SharePointContentType === 'Invoice')).toBeTruthy();
      expect(
        jsonValue.find((m) => m.SharePointContentType === 'Document').LaserficheContentType
      ).toBe('ProfileB');
    });

    test('saving a new mapping that duplicates an already-saved content type shows validation and skips save', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(4));
      const [, , newSpSelect, newLfSelect] = getSelects();
      fireEvent.change(newSpSelect, { target: { value: 'Invoice' } });
      fireEvent.change(newLfSelect, { target: { value: 'ProfileB' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => {
        expect(
          screen.getByText('Mapping already exists for this SharePoint content type')
        ).toBeInTheDocument();
      });
      expect(post).not.toHaveBeenCalled();
    });
  });

  describe('Edit', () => {
    test('clicking edit enables the selects for that row', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext());
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('edit'));

      const [spSelect, lfSelect] = getSelects();
      expect(spSelect).not.toBeDisabled();
      expect(lfSelect).not.toBeDisabled();
    });

    test('editing the Laserfiche profile and saving posts the updated pair with MERGE headers and clears validation', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('edit'));
      const [, lfSelect] = getSelects();
      fireEvent.change(lfSelect, { target: { value: 'ProfileB' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => expect(post).toHaveBeenCalledTimes(1));
      const [url, , requestOptions] = post.mock.calls[0];
      expect(url).toBe(`${listUrl}/items(${SINGLE_MAPPING_ITEM.Id})`);
      expect(requestOptions.headers['IF-MATCH']).toBe('*');
      expect(requestOptions.headers['X-HTTP-Method']).toBe('MERGE');
      const body = JSON.parse(requestOptions.body);
      const jsonValue = JSON.parse(body.JsonValue) as ProfileMappingConfiguration[];
      expect(jsonValue).toHaveLength(1);
      expect(jsonValue[0]).toMatchObject({
        SharePointContentType: 'Invoice',
        LaserficheContentType: 'ProfileB',
      });

      expect(screen.queryByText(/Please select/)).not.toBeInTheDocument();
      await waitFor(() => expect(getSelects()[0]).toBeDisabled());
    });

    test('editing a SharePoint content type to match a different existing mapping shows duplicate validation and skips save', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: TWO_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(4));

      fireEvent.click(screen.getAllByText('edit')[0]);
      const [spSelect] = getSelects();
      fireEvent.change(spSelect, { target: { value: 'Document' } });

      fireEvent.click(screen.getAllByText('save')[0]);

      await waitFor(() => {
        expect(
          screen.getByText('Mapping already exists for this SharePoint content type')
        ).toBeInTheDocument();
      });
      expect(post).not.toHaveBeenCalled();
    });

    test('a rejecting spHttpClient.post shows the creating-mapping error message', async () => {
      const post = vi.fn().mockRejectedValue(new Error('Save failed'));
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('edit'));
      const [, lfSelect] = getSelects();
      fireEvent.change(lfSelect, { target: { value: 'ProfileB' } });

      fireEvent.click(screen.getByText('save'));

      await waitFor(() => {
        expect(screen.getByText('Error creating mapping: Save failed')).toBeInTheDocument();
      });
    });
  });

  describe('Delete', () => {
    test('clicking delete shows a confirmation modal naming the row content type', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      const { container } = renderPage(buildContext());
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('delete'));

      expect(screen.getByText('Delete Confirmation')).toBeInTheDocument();
      expect(container.textContent).toContain('Do you want to permanently delete "Invoice"?');
    });

    test('confirming delete removes the row and posts the mapping array minus that entry', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('delete'));
      fireEvent.click(screen.getByText('Ok'));

      await waitFor(() => expect(post).toHaveBeenCalledTimes(1));
      const [url, , requestOptions] = post.mock.calls[0];
      expect(url).toBe(`${listUrl}/items(${SINGLE_MAPPING_ITEM.Id})`);
      expect(requestOptions.headers['IF-MATCH']).toBe('*');
      expect(requestOptions.headers['X-HTTP-Method']).toBe('MERGE');
      const body = JSON.parse(requestOptions.body);
      expect(JSON.parse(body.JsonValue)).toEqual([]);

      await waitFor(() => expect(queryAllSelects()).toHaveLength(0));
    });

    test('cancelling delete keeps the row and does not call post', async () => {
      const post = vi.fn().mockResolvedValue({ ok: true });
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('delete'));
      fireEvent.click(screen.getByText('Cancel'));

      expect(screen.queryByText('Delete Confirmation')).not.toBeInTheDocument();
      expect(getSelects()).toHaveLength(2);
      expect(post).not.toHaveBeenCalled();
    });

    test('a rejecting delete post shows the deleting-mapping error message', async () => {
      const post = vi.fn().mockRejectedValue(new Error('Delete failed'));
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext(post));
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      fireEvent.click(screen.getByText('delete'));
      fireEvent.click(screen.getByText('Ok'));

      await waitFor(() => {
        expect(screen.getByText('Error deleting mapping: Delete failed')).toBeInTheDocument();
      });
    });
  });

  describe('Reset / misc', () => {
    test('clicking Reset re-fetches all endpoints, reverts unsaved changes, and clears validation', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext());
      await waitFor(() => expect(getSelects()).toHaveLength(2));
      const fetchCallsAfterLoad = (window.fetch as Mock).mock.calls.length;

      fireEvent.click(screen.getByText('Add'));
      await waitFor(() => expect(getSelects()).toHaveLength(4));

      fireEvent.click(screen.getByText('edit'));
      let selects = getSelects();
      fireEvent.change(selects[1], { target: { value: 'ProfileB' } });

      fireEvent.click(screen.getAllByText('save')[1]);
      await waitFor(() => {
        expect(
          screen.getByText(
            'Please select a content type from the SharePoint Content Type drop down'
          )
        ).toBeInTheDocument();
      });

      fireEvent.click(screen.getByText('Reset'));

      await waitFor(() => {
        expect((window.fetch as Mock).mock.calls.length).toBe(fetchCallsAfterLoad + 3);
      });
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      selects = getSelects();
      expect(selects[0].value).toBe('Invoice');
      expect(selects[1].value).toBe('ProfileA');
      expect(selects[0]).toBeDisabled();
      expect(screen.queryByText(/Please select/)).not.toBeInTheDocument();
    });

    test('clicking the View SharePoint Content Types link opens the content types page', async () => {
      window.fetch = buildFetchMock({ mappingListItem: SINGLE_MAPPING_ITEM });
      renderPage(buildContext());
      await waitFor(() => expect(getSelects()).toHaveLength(2));

      const openSpy = vi.spyOn(window, 'open').mockImplementation(() => null);

      fireEvent.click(screen.getByText('View SharePoint Content Types'));

      expect(openSpy).toHaveBeenCalledWith(`${ABSOLUTE_URL}/_layouts/15/mngctype.aspx`);
    });
  });
});
