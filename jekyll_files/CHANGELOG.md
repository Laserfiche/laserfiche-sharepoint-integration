---
layout: default
title: Changelog
nav_order: 4
---

# Laserfiche SharePoint Online Integration Changelog

## Unreleased

### Features

- Uploading a file from the repository explorer now lets you add tags, shows when templates are loading, and keeps OK disabled until required fields are filled in. Afterwards, a dialog links to the new document and to its folder.
- After saving a document to Laserfiche, the confirmation links to the saved document and to its folder.

### Maintenance

- Updated to SharePoint Framework 1.23.2, the Laserfiche Repository API v2 (704606) and Laserfiche UI Components 21 (710113).
- Updated the test framework and added component tests. (710120)
- After updating, users need to sign in to Laserfiche again.

### Fixes

- The repository explorer now shows new folders and uploaded documents without a manual refresh. (709485)
- Signing in and out from the Laserfiche Sign In page, repository explorer and administrator configuration page now works reliably: Sign in no longer signs you out, and Sign out now signs you out.
- If your browser blocks the sign-in pop-up, a message now explains how to allow it.

## 1.0.0.560

### Maintenance

- Added support for [Content Security Policy (CSP) in SharePoint Online](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/content-securty-policy-trusted-script-sources).
- Security updates

## 1.0.0.471

### Features

- Added support for record series to the repository explorer and configuration pages. (545562)

### Maintenance

- Updated dependencies

### Fixes

- The administration configuration page now refreshes when user is signed out. (Previous Known Issue, 488097)
- The profile configuration page can now display a warning if a profile references a Laserfiche template that has been deleted. (Previous Known Issue, 487980)
- The integration can now detect the content type for a document when the Content Type column is not shown. (Previous Known Issue, 490833)
- The integration can now map columns with names containing special characters. (504592)

## 1.0.0.360

### Features

- Initial release
- Ability to save documents in a SharePoint document library to Laserfiche
- Ability to view your Laserfiche repository from within SharePoint Online

### Known Issues

- **488127:** Importing file from repository explorer hits a dead end when invalid metadata is entered
  - **Workaround:** Press the cancel button and restart import process.
- **492278:** Image files are saved to Laserfiche as electronic documents
- [FIXED] **488097:** Admin Configuration page appears interactive when user is logged out
  - Fixed in version 1.0.0.471
- [FIXED] **487980:** If configured Laserfiche template is deleted, profile configuration page does not show a good error.
  - Fixed in version 1.0.0.471
- [FIXED] **490833:** Content Type can't be determined unless column is shown
  - Fixed in version 1.0.0.471
