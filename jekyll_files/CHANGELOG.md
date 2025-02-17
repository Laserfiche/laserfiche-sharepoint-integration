---
layout: default
title: Changelog
nav_order: 4
---

# Laserfiche SharePoint Online Integration Changelog

## 1.0.0.471

### Maintenance

- Updated dependencies to resolve some vulnerabilities
- Fixed some known issues:
  - **488097:** Admin Configuration page appears interactive when user is logged out
  - **487980:** If configured Laserfiche template is deleted, profile configuration page does not show a good error.
  - **490833:** Content Type can't be determined unless column is shown
- Fixed issue where certain columns were not mapped correctly if they had special characters in the name.

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
