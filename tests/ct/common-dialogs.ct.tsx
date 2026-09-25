import { test, expect } from '@playwright/experimental-ct-react17';
import * as React from 'react';
import {
  Collapsible,
  MessageDialog,
  SavedToLaserficheSuccessDialogButtons,
} from '../../src/extensions/savetoLaserfiche/CommonDialogs';
import ConfirmHarness from './harness/ConfirmHarness';

// NOTE: every assertion here queries via `page.*` rather than the `component`
// locator `mount()` returns. @playwright/experimental-ct-react17's
// `internal:control=component` root-scoping selector intermittently hangs or
// reports "element(s) not found" for elements that are demonstrably present
// and clickable (confirmed via page-level locators, which always resolved
// immediately) -- reproduced here for `useConfirm`'s dialog (whose
// `Confirmation` is a fresh inline component recreated on every render,
// forcing React to remount that subtree) and, less predictably, for a later
// `mount()` call reusing the same worker's page. This is exactly the kind of
// experimental-package rough edge the test plan flagged as a risk. Page-level
// locators sidestep it entirely and are just as precise here, since each
// test mounts exactly one component tree.
test.describe('useConfirm', () => {
  test('shows nothing until getConfirmation is called', async ({ mount, page }) => {
    await mount(
      <ConfirmHarness label="Are you sure?" cancelButtonText="Go back" headerText="Please Confirm" />
    );
    await expect(page.getByRole('button', { name: 'Ask' })).toBeVisible();
    await expect(page.getByText('Please Confirm')).toHaveCount(0);
  });

  test('Continue resolves the confirmation promise true and hides the dialog', async ({ mount, page }) => {
    await mount(
      <ConfirmHarness label="Are you sure?" cancelButtonText="Go back" headerText="Please Confirm" />
    );
    await page.getByRole('button', { name: 'Ask' }).click();
    await expect(page.getByText('Please Confirm')).toBeVisible();
    await expect(page.getByText('Are you sure?')).toBeVisible();

    await page.getByRole('button', { name: 'Continue' }).click();

    await expect(page.getByText('Please Confirm')).toHaveCount(0);
    const results = await page.evaluate(() => window.__confirmResults);
    expect(results[results.length - 1]).toBe(true);
  });

  test('the cancel button resolves the confirmation promise false', async ({ mount, page }) => {
    await mount(
      <ConfirmHarness label="Are you sure?" cancelButtonText="Go back" headerText="Please Confirm" />
    );
    await page.getByRole('button', { name: 'Ask' }).click();
    await page.getByRole('button', { name: 'Go back' }).click();

    const results = await page.evaluate(() => window.__confirmResults);
    expect(results[results.length - 1]).toBe(false);
  });

  test('two sequential confirmations resolve independently', async ({ mount, page }) => {
    await mount(
      <ConfirmHarness label="Are you sure?" cancelButtonText="Go back" headerText="Please Confirm" />
    );
    await page.getByRole('button', { name: 'Ask' }).click();
    await page.getByRole('button', { name: 'Continue' }).click();

    await page.getByRole('button', { name: 'Ask' }).click();
    await page.getByRole('button', { name: 'Go back' }).click();

    const results = await page.evaluate(() => window.__confirmResults);
    expect(results).toEqual([true, false]);
  });
});

test.describe('Collapsible', () => {
  test('starts collapsed: children are absent from the DOM', async ({ mount, page }) => {
    await mount(
      <Collapsible title="Section">
        <span>Hidden content</span>
      </Collapsible>
    );
    await expect(page.getByText('Hidden content')).toHaveCount(0);
    await expect(page.getByText('chevron_right')).toBeVisible();
  });

  test('open={true} shows children immediately', async ({ mount, page }) => {
    await mount(
      <Collapsible title="Section" open={true}>
        <span>Visible content</span>
      </Collapsible>
    );
    await expect(page.getByText('Visible content')).toBeVisible();
    await expect(page.getByText('expand_less')).toBeVisible();
  });

  test('clicking the toggle reveals children and flips the icon', async ({ mount, page }) => {
    await mount(
      <Collapsible title="Section">
        <span>Toggled content</span>
      </Collapsible>
    );
    await page.getByText('chevron_right').click();
    await expect(page.getByText('Toggled content')).toBeVisible();
    await expect(page.getByText('expand_less')).toBeVisible();
  });
});

test.describe('MessageDialog', () => {
  test('renders title/message and fires clickOkay', async ({ mount, page }) => {
    let clicked = false;
    await mount(
      <MessageDialog
        title="Sign In Failed"
        message="Something went wrong"
        clickOkay={() => {
          clicked = true;
        }}
      />
    );
    await expect(page.getByText('Sign In Failed')).toBeVisible();
    await expect(page.getByText('Something went wrong')).toBeVisible();
    await page.getByRole('button', { name: 'Okay' }).click();
    await page.waitForTimeout(50);
    expect(clicked).toBe(true);
  });
});

test.describe('SavedToLaserficheSuccessDialogButtons', () => {
  test('Close closes the dialog without opening anything', async ({ mount, page }) => {
    let closeCount = 0;
    await mount(
      <SavedToLaserficheSuccessDialogButtons
        closeClick={async () => {
          closeCount++;
        }}
      />
    );
    // The call log is module-level, so clear it in case this worker's page
    // still holds an earlier test's window.open entry.
    await page.evaluate(() => {
      window.__ctCalls = [];
    });
    await page.getByRole('button', { name: 'Close' }).click();
    await expect.poll(() => closeCount).toBe(1);
    const calls = await page.evaluate(() => window.__ctCalls);
    expect(calls.some((c) => c.element === 'window' && c.method === 'open')).toBe(false);
  });

  test('only Close renders', async ({ mount, page }) => {
    await mount(<SavedToLaserficheSuccessDialogButtons closeClick={async () => {}} />);
    await expect(page.getByRole('button', { name: 'View file in Laserfiche' })).toHaveCount(0);
    await expect(page.getByRole('button')).toHaveCount(1);
    await expect(page.getByRole('button', { name: 'Close' })).toBeVisible();
  });
});
