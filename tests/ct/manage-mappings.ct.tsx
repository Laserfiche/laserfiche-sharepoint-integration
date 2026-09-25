import { test, expect } from '@playwright/experimental-ct-react17';
import * as React from 'react';
import ManageMappingsHarness from './harness/ManageMappingsHarness';

// The Vitest suite (ManageMappingsPage.test.tsx) already covers this
// component's branching logic thoroughly with mocked fetch/DOM. What it
// can't cover is real <select> interaction (native browser behavior jsdom
// doesn't implement) -- that's the focus here.

test('loads existing mappings into real, interactable selects', async ({ mount, page }) => {
  await mount(<ManageMappingsHarness />);

  await expect(page.getByText('Content Type Mappings Laserfiche')).toBeVisible();
  const spSelect = page.locator('select[name="SharePointContentType"]').first();
  const lfSelect = page.locator('select[name="LaserficheContentType"]').first();
  await expect(spSelect).toHaveValue('Invoice');
  await expect(lfSelect).toHaveValue('ProfileA');
  await expect(spSelect).toBeDisabled();
});

test('adding a row, picking real select options, and saving posts the new mapping', async ({ mount, page }) => {
  await mount(<ManageMappingsHarness scenario="empty" />);
  await expect(page.getByText('Content Type Mappings Laserfiche')).toBeVisible();

  await page.getByRole('button', { name: 'Add' }).click();

  const spSelect = page.locator('select[name="SharePointContentType"]').last();
  const lfSelect = page.locator('select[name="LaserficheContentType"]').last();
  // Native <select> option selection -- real browser behavior.
  await spSelect.selectOption('Invoice');
  await lfSelect.selectOption('ProfileB');
  await expect(spSelect).toHaveValue('Invoice');
  await expect(lfSelect).toHaveValue('ProfileB');

  await page.locator('button:has(span:text("save"))').click();

  await expect
    .poll(async () => page.evaluate(() => window.__postCalls?.length ?? 0))
    .toBeGreaterThan(0);
  const body = await page.evaluate(() => JSON.stringify(window.__postCalls));
  expect(body).toContain('Invoice');
  expect(body).toContain('ProfileB');
});

test('saving with a selection left at "Select" shows a validation message', async ({ mount, page }) => {
  await mount(<ManageMappingsHarness scenario="empty" />);
  await expect(page.getByText('Content Type Mappings Laserfiche')).toBeVisible();

  await page.getByRole('button', { name: 'Add' }).click();
  await page.locator('button:has(span:text("save"))').click();

  await expect(
    page.getByText('Please select a content type from the SharePoint Content Type drop down')
  ).toBeVisible();
});
