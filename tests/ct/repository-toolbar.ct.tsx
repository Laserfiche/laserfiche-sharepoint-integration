import { test, expect } from '@playwright/experimental-ct-react17';
import * as React from 'react';
import RepositoryToolbarHarness from './harness/RepositoryToolbarHarness';

// Targets the gap in src/webparts/LaserficheRepositoryAccessWebPart/components/
// RepositoryViewWebPart.test.tsx (Vitest): that suite renders the toolbar but
// never opens either modal. These are real-browser cases jsdom cannot cover
// well: a genuine `File` object via setInputFiles, uncontrolled-input DOM
// state, and CSS-driven visibility.
//
// Locators below query via `page.*` rather than the `component` locator
// `mount()` returns -- @playwright/experimental-ct-react17's
// `internal:control=component` root-scoping selector intermittently hangs or
// reports elements missing that are demonstrably present (see
// common-dialogs.ct.tsx for the full explanation). Each test here mounts
// exactly one component tree, so page-level locators are equally precise.

test.describe('CreateFolderModal', () => {
  test('empty name shows a validation message and does not create an entry', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.getByRole('button', { name: 'Submit' }).click();

    await expect(page.getByText('Please provide a folder name')).toBeVisible();
    const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    expect(calls.some((c) => c.method === 'createEntry')).toBe(false);
  });

  test('backslash in name shows a validation message', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.locator('#folderName').fill('My\\Folder');
    await page.getByRole('button', { name: 'Submit' }).click();

    await expect(page.getByText('Entry names cannot contain backslash')).toBeVisible();
  });

  test('a name that already exists keeps the modal open with an error', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness scenario='createEntry-rejects-exists' />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.locator('#folderName').fill('Invoices');
    await page.getByRole('button', { name: 'Submit' }).click();

    await expect(page.getByText('Object already exists')).toBeVisible();
    // Modal is still open -- the name input is still present in the DOM.
    await expect(page.locator('#folderName')).toBeVisible();
  });

  test('a valid unique name creates the folder and closes the modal', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.locator('#folderName').fill('Contracts');
    await page.getByRole('button', { name: 'Submit' }).click();

    await expect(page.locator('#folderName')).toBeHidden();
    const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    const createCall = calls.find((c) => c.method === 'createEntry');
    expect(createCall).toBeTruthy();
    expect(calls.some((c) => c.method === 'refreshFolderBrowserAsync')).toBe(true);
  });

  // The folder already exists by then: reporting "Object already exists" would
  // be wrong, and submitting again would really hit that error.
  test('a failed folder list refresh after creating the folder still closes the modal', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness refreshShouldFail={true} />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.locator('#folderName').fill('Contracts');
    await page.getByRole('button', { name: 'Submit' }).click();

    await expect(page.locator('#folderName')).toBeHidden();
    await expect(page.getByText('Object already exists')).toHaveCount(0);
  });

  test('closing and reopening starts with an empty name field', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Create folder in Laserfiche').click();
    await page.locator('#folderName').fill('Draft Name');
    // Two elements match role=button name="Close" here: the header's `x`
    // (aria-label="Close") and the footer's text button -- target the
    // footer one specifically by its class.
    await page.locator('button.sec-button', { hasText: 'Close' }).click();

    await page.getByTitle('Create folder in Laserfiche').click();
    await expect(page.locator('#folderName')).toHaveValue('');
  });
});

test.describe('ImportFileModal', () => {
  test('clicking OK with no file selected shows a validation message', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText('Please select the file to upload')).toBeVisible();
  });

  test('picking a real file auto-populates the name field without its extension', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('%PDF-1.4 test'),
    });

    await expect(page.locator('#uploadFileID')).toHaveValue('contract');
  });

  test('a filename containing a backslash is rejected', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.locator('#uploadFileID').fill('a\\b');
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText('Please provide a valid filename without backslash')).toBeVisible();
  });

  test('clearing the name field is rejected as an invalid filename', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.locator('#uploadFileID').fill('');
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText('Please provide a valid filename')).toBeVisible();
  });

  test('an existing entry at the target path shows a rename confirmation before importing', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness scenario='getEntryByPath-exists' />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText(/already exists in the specified folder/i)).toBeVisible();

    let calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    expect(calls.some((c) => c.method === 'importEntry')).toBe(false);

    await page.getByRole('button', { name: 'Go back' }).click();
    calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    expect(calls.some((c) => c.method === 'importEntry')).toBe(false);
  });

  test('continuing past the rename confirmation proceeds to import', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness scenario='getEntryByPath-exists' />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();
    await page.getByRole('button', { name: 'Continue' }).click();

    await expect
      .poll(async () => {
        const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
        return calls.some((c) => c.method === 'importEntry');
      })
      .toBe(true);
  });

  test('OK is disabled and the progress bar is visible while an import is in flight', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness importDelayMs={500} />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByRole('button', { name: 'OK' })).toBeDisabled();
    await expect(page.getByText('Uploading')).toBeVisible();
  });

  test('OK is disabled while a selected template/field is invalid, and re-enables once valid', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();

    // The modal's effect attaches lf-field-container's event listeners
    // asynchronously after mount; dispatching a custom event before that
    // runs is silently lost (no listener yet to receive it), which is a
    // real race, not just a slow assertion -- toBeDisabled() below would
    // still time out even with a longer wait, since the missed event never
    // gets redelivered. initAsync is recorded (see tests/ct/playwright/
    // index.ts) right after the listeners are attached, so waiting for that
    // call to show up guarantees they're in place before we dispatch.
    await expect
      .poll(async () =>
        (await page.evaluate(() => window.__ctCalls ?? [])).some(
          (c) => c.element === 'lf-field-container' && c.method === 'initAsync'
        )
      )
      .toBe(true);

    await page
      .locator('lf-field-container')
      .evaluate((el: HTMLElement & { emitFieldValuesChanged: (v: boolean) => void }) =>
        el.emitFieldValuesChanged(false)
      );
    await expect(page.getByRole('button', { name: 'OK' })).toBeDisabled();

    await page
      .locator('lf-field-container')
      .evaluate((el: HTMLElement & { emitFieldValuesChanged: (v: boolean) => void }) =>
        el.emitFieldValuesChanged(true)
      );
    await expect(page.getByRole('button', { name: 'OK' })).toBeEnabled();
  });

  test('submitting while required fields are invalid shows a message instead of leaving OK stuck disabled', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    // Simulate forceValidation() reporting invalid without the proactive
    // fieldValuesChanged event having fired, so OK is still clickable.
    await page
      .locator('lf-field-container')
      .evaluate((el: HTMLElement & { forceValidationResult: boolean }) => {
        el.forceValidationResult = false;
      });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText('Please provide values for all required fields')).toBeVisible();
    await expect(page.getByRole('button', { name: 'OK' })).toBeEnabled();
  });

  test('selected template and field values are included in the saved metadata', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.locator('lf-field-container').evaluate(
      (
        el: HTMLElement & {
          templateValue: { name: string };
          fieldValues: Record<string, { values: Array<{ value: unknown }> }>;
        }
      ) => {
        el.templateValue = { name: 'Board Meetings' };
        el.fieldValues = { 'Meeting Date': { values: [{ value: '2026-01-01' }] } };
      }
    );
    await page.getByRole('button', { name: 'OK' }).click();

    await expect
      .poll(async () => {
        const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
        return calls.some((c) => c.method === 'importEntry');
      })
      .toBe(true);

    const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    const importCall = calls.find((c) => c.method === 'importEntry');
    const metadata = (
      importCall.args[0] as { request: { metadata: { templateName: string; fields: unknown[] } } }
    ).request.metadata;
    expect(metadata.templateName).toBe('Board Meetings');
    expect(metadata.fields).toEqual([{ name: 'Meeting Date', values: ['2026-01-01'] }]);
  });

  test('a successful import refreshes the folder browser', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect
      .poll(async () => {
        const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
        return calls.some((c) => c.method === 'refreshFolderBrowserAsync');
      })
      .toBe(true);
  });

  test('a successful import replaces the modal with a saved-copy dialog linking to the document and its folder', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(
      page.getByText(
        'Saved a copy to Laserfiche. To keep editing, open or check out from Web Client.'
      )
    ).toBeVisible();
    await expect(page.locator('#importFile')).toHaveCount(0);
    await expect(page.getByRole('link', { name: 'imported' })).toHaveAttribute(
      'href',
      /DocView\.aspx\?.*id=100/
    );
    await expect(page.getByRole('link', { name: 'Show in folder' })).toHaveAttribute(
      'href',
      /Browse\.aspx.*#\?id=1$/
    );

    await page.getByRole('button', { name: 'Close' }).click();

    await expect(page.getByText(/Saved a copy to Laserfiche/)).toHaveCount(0);
  });

  // The document is already in Laserfiche by then: reporting the upload as
  // failed would invite a retry, and that retry would create a duplicate.
  test('a failed folder list refresh after a successful import still shows the saved-copy dialog', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness refreshShouldFail={true} />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect(page.getByText(/Saved a copy to Laserfiche/)).toBeVisible();
    await expect(page.getByText(/Error uploading/)).toHaveCount(0);
  });

  test("tag picker lists the repository's tags", async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();

    const tags = page.locator('lf-tags');
    await expect(tags.getByText('Contract')).toBeVisible();
    await expect(tags.getByText('Reviewed')).toBeVisible();
    await expect(tags.getByText('No tags available')).toBeHidden();
  });

  test('selected tags are applied to the imported entry', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    // See the "OK is disabled..." test above -- <lf-tags>'s listener is
    // attached by LfTagsPicker's effect, and React runs a child's effects
    // before its parent's, so once the modal's effect has called
    // <lf-field-container>'s initAsync the tags listener is in place too.
    await expect
      .poll(async () =>
        (await page.evaluate(() => window.__ctCalls ?? [])).some(
          (c) => c.element === 'lf-field-container' && c.method === 'initAsync'
        )
      )
      .toBe(true);
    await page.locator('lf-tags').evaluate(
      (
        el: HTMLElement & {
          emitSelectedTagsChanged: (tags: Array<{ displayName: string }>) => void;
        }
      ) => el.emitSelectedTagsChanged([{ displayName: 'Contract' }, { displayName: 'Reviewed' }])
    );
    await page.getByRole('button', { name: 'OK' }).click();

    await expect
      .poll(async () => {
        const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
        return calls.some((c) => c.method === 'setTags');
      })
      .toBe(true);

    const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    const setTagsCall = calls.find((c) => c.method === 'setTags');
    const args = setTagsCall.args[0] as {
      repositoryId: string;
      entryId: number;
      request: { tags: string[] };
    };
    expect(args.repositoryId).toBe('repo-1');
    expect(args.entryId).toBe(100);
    expect(args.request.tags).toEqual(['Contract', 'Reviewed']);
  });

  test('no tags selected means setTags is never called', async ({ mount, page }) => {
    await mount(<RepositoryToolbarHarness />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await page.locator('#importFile').setInputFiles({
      name: 'contract.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('data'),
    });
    await page.getByRole('button', { name: 'OK' }).click();

    await expect
      .poll(async () => {
        const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
        return calls.some((c) => c.method === 'importEntry');
      })
      .toBe(true);

    const calls = await page.evaluate(() => window.__repoClientCalls ?? []);
    expect(calls.some((c) => c.method === 'setTags')).toBe(false);
  });

  test('a slow template list load shows a Loading... spinner beside the Template header until the list arrives', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness templatesLoadDelayMs={1500} />);
    await page.getByTitle('Upload file to Laserfiche').click();
    await expect
      .poll(async () =>
        (await page.evaluate(() => window.__ctCalls ?? [])).some(
          (c) => c.element === 'lf-field-container' && c.method === 'initAsync'
        )
      )
      .toBe(true);

    const templatesLoaded = page
      .locator('lf-field-container')
      .evaluate((el: HTMLElement & { openTemplatesDropdown: () => Promise<void> }) =>
        el.openTemplatesDropdown()
      );
    const spinner = page.locator('lf-field-container mat-panel-title').getByRole('status');
    await expect(spinner).toBeVisible();
    await expect(spinner).toHaveText('Loading...');
    await expect(page.getByText('Loading templates...')).toHaveCount(0);
    await templatesLoaded;

    await expect(spinner).toBeHidden();
  });

  test('a failed template/field load surfaces an error message instead of failing silently', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness fieldContainerShouldFailInit={true} />);
    await page.getByTitle('Upload file to Laserfiche').click();

    await expect(page.getByText(/Unable to load templates and fields/i)).toBeVisible();
  });
});

test.describe('Toolbar', () => {
  test('clicking Open with no selection and no parent folder shows an alert modal', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness hasParent={false} />);
    await page.getByTitle('Open entry in Laserfiche').click();

    await expect(page.getByText('Please select file/folder to open')).toBeVisible();
    await page.getByRole('button', { name: 'OK' }).click();
    await expect(page.getByText('Please select file/folder to open')).toBeHidden();
  });

  test('import button is disabled when the open folder is a record series', async ({
    mount,
    page,
  }) => {
    await mount(<RepositoryToolbarHarness parentEntryType='RecordSeries' />);
    await expect(page.getByTitle('Cannot import into a Record Series')).toBeDisabled();
  });
});
