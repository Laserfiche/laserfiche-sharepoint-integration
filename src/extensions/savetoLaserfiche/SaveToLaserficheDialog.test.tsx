// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// SaveToLaserficheDialog (the function component that renders the dialog UI)
// is not exported from ./SaveToLaserficheDialog -- only the SPFx
// SaveToLaserficheCustomDialog class is. So, like production code, this file
// drives it indirectly: build the class, give it a detached domElement, and
// call render(). See `renderDialog` below for how the class's hardcoded
// `isSuccessfulLoggedIn`/`closeClick` props are still made observable as
// vi.fn() mocks despite not being directly injectable, and `setLoginToken` for
// how the unregistered <lf-login> ref is populated before the component's
// mount effect reads it.

vi.mock('@microsoft/sp-dialog', () => ({
  // The real BaseDialog implements `close()` (tears down the SPFx dialog
  // chrome), which SaveToLaserficheCustomDialog#closeClick awaits before
  // calling closeParent. A bare `class {}` -- the mock used for this package
  // elsewhere in this project, e.g. GetDocumentDataDialogUI.test.tsx -- would
  // leave `close` undefined, so `await this.close()` throws before
  // closeParent ever runs. Adding a no-op resolved-promise `close()` keeps
  // the mock just as minimal while letting closeClick actually complete.
  BaseDialog: class {
    close(): Promise<void> {
      return Promise.resolve();
    }
  },
}));

vi.mock('../../repository-client/repository-client');

vi.mock('./SaveDocumentToLaserfiche');

vi.mock('@laserfiche/lf-js-utils', () => ({
  PathUtils: {
    combinePaths: (a: string, b: string) => `${a}/${b}`,
    removeFileExtension: (name: string) =>
      name.includes('.') ? name.slice(0, name.lastIndexOf('.')) : name,
  },
}));

import type { Mock, MockInstance } from 'vitest';
import * as ReactDOM from 'react-dom';
import { act, fireEvent, waitFor, within } from '@testing-library/react';
import SaveToLaserficheCustomDialog from './SaveToLaserficheDialog';
import { RepositoryClientExInternal } from '../../repository-client/repository-client';
import { SaveDocumentToLaserfiche } from './SaveDocumentToLaserfiche';
import type { SavedToLaserficheDocumentData } from './SaveDocumentToLaserfiche';
import type { ISPDocumentData } from '../../Utils/Types';
import { ActionTypes } from '../../webparts/laserficheAdminConfiguration/components/ProfileConfigurationComponents';
import {
  CANCEL,
  CONTINUE,
  DOCUMENT_ALREADY_EXISTS,
  ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME,
  SAVED_A_COPY_TO_LASERFICHE,
} from '../../webparts/strings';
import { SP_LOCAL_STORAGE_KEY } from '../../webparts/constants';

const spFileMetadata: ISPDocumentData = {
  fileName: 'Contract.pdf',
  documentName: 'Contract.pdf',
  action: ActionTypes.COPY,
  fileUrl: '/sites/test/Shared Documents/Contract.pdf',
  entryId: '42',
  contextPageAbsoluteUrl: 'https://sharepoint.example.com/sites/test',
};

function buildRepoClient(): {
  repositoriesClient: { listRepositories: Mock };
  entriesClient: { getEntry: Mock; getEntryByPath: Mock };
  getCurrentRepoId: Mock;
} {
  return {
    repositoriesClient: { listRepositories: vi.fn() },
    entriesClient: { getEntry: vi.fn(), getEntryByPath: vi.fn() },
    getCurrentRepoId: vi.fn().mockResolvedValue('repo-1'),
  };
}

function mockRepoClient(repoClient: ReturnType<typeof buildRepoClient>): void {
  (RepositoryClientExInternal as unknown as Mock).mockImplementation(function () {
    return {
      createRepositoryClientAsync: vi.fn().mockResolvedValue(repoClient),
    };
  });
}

function mockSaveDocumentToLaserfiche(trySaveDocumentToLaserficheAsync: Mock): void {
  (SaveDocumentToLaserfiche as unknown as Mock).mockImplementation(function () {
    return { trySaveDocumentToLaserficheAsync };
  });
}

function buildSuccess(
  overrides: Partial<SavedToLaserficheDocumentData> = {}
): SavedToLaserficheDocumentData {
  return {
    fileLink: 'https://laserfiche.example.com/Contract.pdf',
    folderLink: 'https://laserfiche.example.com/Folder',
    pathBack: 'https://sharepoint.example.com',
    fileName: 'Contract.pdf',
    action: undefined,
    ...overrides,
  };
}

let currentDomElement: HTMLDivElement | undefined;

// Builds SaveToLaserficheCustomDialog and mounts it, standing in for
// rendering the inner (unexported) component with injected props:
//  - `isSuccessfulLoggedIn` is observed by spying on the instance's own
//    `handleSuccessSave` field *before* render() so the JSX (built inside
//    render()) captures the spied function. vi.spyOn calls through to the
//    original implementation by default, so `dialog.successful` still gets
//    set normally.
//  - `closeClick` is observed via the `closeParent` constructor argument:
//    SaveToLaserficheCustomDialog#closeClick always awaits the (mocked,
//    no-op) `this.close()` and then `this.closeParent(success)`, so
//    `closeParent` sees every call the inner component makes through its
//    `closeClick` prop.
function renderDialog(fileMetadata: ISPDocumentData = spFileMetadata): {
  dialog: SaveToLaserficheCustomDialog;
  domElement: HTMLDivElement;
  closeParent: Mock;
  isSuccessfulLoggedIn: MockInstance;
} {
  const closeParent = vi.fn().mockResolvedValue(undefined);
  const dialog = new SaveToLaserficheCustomDialog(fileMetadata, closeParent);
  const isSuccessfulLoggedIn = vi.spyOn(dialog, 'handleSuccessSave');

  const domElement = document.createElement('div');
  document.body.appendChild(domElement);
  currentDomElement = domElement;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (dialog as any).domElement = domElement;
  dialog.render();

  return { dialog, domElement, closeParent, isSuccessfulLoggedIn };
}

// Sets the properties SaveToLaserficheDialog's mount effect (and, on a real
// save, SaveDocumentToLaserfiche/RepositoryClientExInternal) read off the
// <lf-login> element. <lf-login> is never registered as a custom element in
// this suite, so React's ref to it is just a plain HTMLElement -- arbitrary
// JS properties can be set on it like any other object.
//
// Timing is what makes this deterministic: ReactDOM.render (React 17) mounts
// the DOM synchronously, but useEffect only flushes on a later tick. As long
// as nothing here `await`s or yields between renderDialog() and this call,
// it always runs before the effect reads authorization_credentials.
function setLoginToken(
  domElement: HTMLElement,
  authorization_credentials: { accessToken: string } | undefined
): void {
  const loginEl = domElement.querySelector('lf-login') as unknown as {
    authorization_credentials?: { accessToken: string };
    account_endpoints?: { webClientUrl: string; regionalDomain: string };
    account_id?: string;
  };
  loginEl.authorization_credentials = authorization_credentials;
  loginEl.account_endpoints = {
    webClientUrl: 'https://webclient.example.com',
    regionalDomain: 'region.example.com',
  };
  loginEl.account_id = 'customer-1';
}

// The confirmation flow (useConfirm's own extra setState calls layered on
// top of the dialog's own async effect) needs an actual macrotask tick to
// commit to the DOM in this environment -- a chain of already-resolved
// mock promises alone isn't enough, and plain `waitFor`/`findByText`
// polling doesn't reliably force that tick since this component was never
// mounted through RTL's own `render()`. Explicitly yielding one via
// `setTimeout` (wrapped in `act` so the resulting state updates are
// committed cleanly) makes the subsequent assertions deterministic instead
// of depending on incidental timing.
async function flushEffects(): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 0));
  });
}

let consoleErrorSpy: MockInstance;
let removeItemSpy: MockInstance;

beforeEach(() => {
  vi.resetAllMocks();
  consoleErrorSpy = vi.spyOn(console, 'error').mockImplementation(() => undefined);
  removeItemSpy = vi.spyOn(Storage.prototype, 'removeItem');
});

afterEach(() => {
  if (currentDomElement) {
    // This pairs with the ReactDOM.render() call inside the production
    // SaveToLaserficheCustomDialog#render() method (invoked indirectly via
    // dialog.render() in renderDialog() above) -- the lint rule only
    // recognizes a literal ReactDOM.render(...) call in the same file, so it
    // can't see the pairing and flags this as unmatched.
    // eslint-disable-next-line @rushstack/pair-react-dom-render-unmount
    ReactDOM.unmountComponentAtNode(currentDomElement);
    currentDomElement.remove();
    currentDomElement = undefined;
  }
  consoleErrorSpy.mockRestore();
  removeItemSpy.mockRestore();
});

describe('SaveToLaserficheDialog', () => {
  // lf-login only creates the provider it refreshes tokens with once it knows
  // its login type. Without one, a token that expires mid-save signs the user
  // out instead of being refreshed.
  test('renders lf-login as a Laserfiche Cloud sign-in, so it can refresh an expired token', async () => {
    const { domElement } = renderDialog();
    setLoginToken(domElement, undefined);

    expect(domElement.querySelector('lf-login')).toHaveAttribute('login_type', 'Cloud');
    await flushEffects();
  });

  test('missing authorization_credentials reports failed login and closes without touching the repository client', async () => {
    const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, undefined);

    await waitFor(() => {
      expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(false);
    });
    expect(closeParent).toHaveBeenCalled();
    expect(RepositoryClientExInternal).not.toHaveBeenCalled();
    expect(SaveDocumentToLaserfiche).not.toHaveBeenCalled();
  });

  test('stale authorization_credentials (token probe rejects) reports failed login and closes, instead of erroring on an undefined repository client', async () => {
    // lf-login restores authorization_credentials from local storage, so an
    // expired or revoked token is still "present" -- only the listRepositories
    // probe reveals it is unusable.
    const repoClient = buildRepoClient();
    repoClient.repositoriesClient.listRepositories.mockRejectedValue({
      status: 401,
    });
    mockRepoClient(repoClient);

    const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'expired-token' });

    await waitFor(() => {
      expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(false);
    });
    expect(closeParent).toHaveBeenCalled();
    expect(isSuccessfulLoggedIn).not.toHaveBeenCalledWith(true);
    expect(repoClient.getCurrentRepoId).not.toHaveBeenCalled();
    expect(SaveDocumentToLaserfiche).not.toHaveBeenCalled();
    expect(within(domElement).queryByText(/error saving/i)).not.toBeInTheDocument();
  });

  test('existing entry at the target path shows the duplicate-name confirmation; Continue proceeds to save, and the header close button then relays the success state', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockResolvedValue({
      entry: { id: 5 },
    });
    mockRepoClient(repoClient);
    const success = buildSuccess();
    const trySave = vi.fn().mockResolvedValue(success);
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await flushEffects();
    expect(
      within(domElement).getByText(ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME)
    ).toBeInTheDocument();
    expect(within(domElement).getByText(DOCUMENT_ALREADY_EXISTS)).toBeInTheDocument();
    expect(trySave).not.toHaveBeenCalled();

    fireEvent.click(within(domElement).getByText(CONTINUE));
    await flushEffects();

    expect(trySave).toHaveBeenCalled();
    expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    expect(within(domElement).getByText(SAVED_A_COPY_TO_LASERFICHE)).toBeInTheDocument();

    fireEvent.click(within(domElement).getByTitle('close'));
    await waitFor(() => {
      expect(closeParent).toHaveBeenCalledWith(success);
    });
  });

  test('existing entry at the target path: Cancel skips the save entirely', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockResolvedValue({
      entry: { id: 5 },
    });
    mockRepoClient(repoClient);
    const trySave = vi.fn().mockResolvedValue(buildSuccess());
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await flushEffects();
    expect(
      within(domElement).getByText(ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME)
    ).toBeInTheDocument();

    fireEvent.click(within(domElement).getByText(CANCEL));
    await flushEffects();

    expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    expect(removeItemSpy).toHaveBeenCalledWith(SP_LOCAL_STORAGE_KEY);
    expect(closeParent).toHaveBeenCalledWith(undefined);
    expect(trySave).not.toHaveBeenCalled();
  });

  test('no entry at the target path (falsy entry field) proceeds straight to save without confirmation', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockResolvedValue({});
    mockRepoClient(repoClient);
    const trySave = vi.fn().mockResolvedValue(buildSuccess());
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await waitFor(() => expect(trySave).toHaveBeenCalled());
    expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    expect(
      within(domElement).queryByText(
        ENTRY_WITH_SAME_NAME_EXISTS_IN_FOLDER_IF_CONTINUE_LF_WILL_RENAME
      )
    ).not.toBeInTheDocument();
  });

  test('getEntryByPath rejecting with 404 (no entry found) is treated the same as a falsy entry field', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockRejectedValue({ status: 404 });
    mockRepoClient(repoClient);
    const trySave = vi.fn().mockResolvedValue(buildSuccess());
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await waitFor(() => expect(trySave).toHaveBeenCalled());
    expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    expect(within(domElement).queryByText(/error saving/i)).not.toBeInTheDocument();
  });

  test.each([401, 403])(
    'getEntryByPath rejecting with %i reports failed login and closes, without showing an error',
    async (status) => {
      const repoClient = buildRepoClient();
      repoClient.entriesClient.getEntry.mockResolvedValue({
        fullPath: '/Folder',
      });
      repoClient.entriesClient.getEntryByPath.mockRejectedValue({ status });
      mockRepoClient(repoClient);

      const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
      setLoginToken(domElement, { accessToken: 'token-1' });

      await waitFor(() => {
        expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(false);
      });
      expect(closeParent).toHaveBeenCalled();
      expect(within(domElement).queryByText(/error saving/i)).not.toBeInTheDocument();
    }
  );

  test('a 404 from getEntry itself falls through to a save attempt; when that save also fails with 404, a verify-entry-id error is shown', async () => {
    // SaveToLaserficheDialog.tsx wraps BOTH the getEntry and getEntryByPath
    // calls in one inner try/catch, nested inside an outer try/catch. A 404
    // from either call is treated by the inner catch as "no entry at that
    // path" and triggers a save attempt anyway (same as the
    // getEntryByPath-404 case above) via `continueSavingDocumentAsync`,
    // called directly from *inside* the inner catch block. If that save
    // itself then also rejects (e.g. because the configured parent entryId
    // genuinely does not exist), the rejection is no longer covered by the
    // inner try, so it propagates to the outer catch -- which is what
    // actually produces the "Verify that an entry ... exists" message.
    const repoClient = buildRepoClient();
    const notFoundError = { status: 404, message: 'not found' };
    repoClient.entriesClient.getEntry.mockRejectedValue(notFoundError);
    mockRepoClient(repoClient);
    const trySave = vi.fn().mockRejectedValue(notFoundError);
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await waitFor(() => {
      expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    });
    expect(repoClient.entriesClient.getEntryByPath).not.toHaveBeenCalled();
    expect(await within(domElement).findByText(/not found/)).toBeInTheDocument();
    expect(
      within(domElement).getByText(
        `Verify that an entry with ID "${spFileMetadata.entryId}" exists and that you have access to it.`
      )
    ).toBeInTheDocument();
  });

  test('any other error surfaces its message as-is', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockRejectedValue({
      status: 500,
      message: 'server error',
    });
    mockRepoClient(repoClient);

    const { domElement, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await waitFor(() => {
      expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(true);
    });
    expect(await within(domElement).findByText('server error')).toBeInTheDocument();
  });

  test('a save that resolves undefined (lf-login lost its token mid-flow) reports failed login instead of hanging on the loading spinner', async () => {
    const repoClient = buildRepoClient();
    repoClient.entriesClient.getEntry.mockResolvedValue({
      fullPath: '/Folder',
    });
    repoClient.entriesClient.getEntryByPath.mockResolvedValue({});
    mockRepoClient(repoClient);
    const trySave = vi.fn().mockResolvedValue(undefined);
    mockSaveDocumentToLaserfiche(trySave);

    const { domElement, closeParent, isSuccessfulLoggedIn } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    await waitFor(() => {
      expect(isSuccessfulLoggedIn).toHaveBeenCalledWith(false);
    });
    expect(closeParent).toHaveBeenCalled();
    expect(within(domElement).queryByText(SAVED_A_COPY_TO_LASERFICHE)).not.toBeInTheDocument();
  });

  test('the header close button relays the current (still-undefined) success state while a save is pending', async () => {
    const repoClient = buildRepoClient();
    // Never resolves, so the component stays on its initial loading state
    // (success/error both undefined) for the life of the test.
    repoClient.entriesClient.getEntry.mockReturnValue(new Promise(() => {}));
    mockRepoClient(repoClient);

    const { domElement, closeParent } = renderDialog();
    setLoginToken(domElement, { accessToken: 'token-1' });

    fireEvent.click(within(domElement).getByTitle('close'));

    await waitFor(() => {
      expect(closeParent).toHaveBeenCalledWith(undefined);
    });
  });
});
