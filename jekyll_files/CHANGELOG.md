---
layout: default
title: Changelog
nav_order: 4
---

# Laserfiche SharePoint Online Integration Changelog

## Unreleased

### Fixes

- Resolved Microsoft SharePoint Online Content Security Policy (CSP) enforcement issue: vendored `zone.js`, `lf-ui-components.js`, `indigo-pink.css`, and `lf-ms-office-lite.css` directly into the solution package instead of fetching from `lfxstatic.com` at runtime. This unblocks tenants whose CSP no longer permits external script sources after Microsoft's March 1, 2026 enforcement (90-day grace period ends June 1, 2026). (651728)

### Maintenance

- The deployed `.sppkg` size grows from ~0.73 MB to ~2.76 MB (~3.8x) due to the vendored assets above. SharePoint admins should be aware of the upload-size change.

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
