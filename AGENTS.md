# AGENTS.md — laserfiche-sharepoint-integration

Instructions for AI coding agents (Copilot CLI, Claude Code, Cursor, etc.) working in this repo. Humans should read [`README.md`](./README.md) and [`CONTRIBUTING.md`](./CONTRIBUTING.md) first.

---

## What this repo is

A SharePoint Framework (SPFx) solution that integrates Laserfiche Cloud with SharePoint Online. Built with React. Ships as a single `.sppkg` uploaded to a tenant App Catalog. Public, MIT‑licensed.

Three web parts + one command set:

| Component | Path |
|---|---|
| Laserfiche Administrator Configuration (web part) | `src/webparts/laserficheAdminConfiguration/` |
| Laserfiche Repository Explorer (web part) | `src/webparts/LaserficheRepositoryAccessWebPart/` |
| Laserfiche Sign In (web part) | `src/webparts/sendToLaserficheLoginComponent/` |
| Save to Laserfiche (command set) | `src/extensions/savetoLaserfiche/` |

Documentation site (Jekyll) lives under `jekyll_files/` and is published to GitHub Pages.

---

## Branches you will see

- `main` — historical/abandoned. **Do NOT branch off `main`.**
- `1.x` — **active release branch.** Tagged builds (`1.0.0.NNN`) are published from here.
- `2.x` would be the next major (does not exist as of writing).
- `jonathan/...`, `dependabot/...` — topic branches.

**Branch off `1.x`. Open PRs targeting `1.x`.** The Integrations wiki release runbook is built around this convention.

---

## Toolchain (must match)

| Tool | Version | Source |
|---|---|---|
| Node.js | **22.x** (`nvm use 22`; pinned in `.nvmrc`) | locally |
| Node.js (CI) | **22.x** | both workflows use `actions/setup-node@v4` with `node-version-file: .nvmrc` |
| Node.js 24 | works, but **opt-in only** — set `SPFX_OVERRIDE_NODE_VERSION_CHECK=true` | unsupported by Microsoft; never set in CI |
| SPFx | **1.23.2** (all `@microsoft/sp-*` and `@microsoft/eslint-*-spfx` pinned to `1.23.2`) | `package.json` |
| TypeScript | **5.3.3** + Rush stack compiler 5.3 | `tsconfig.json` extends `@microsoft/rush-stack-compiler-5.3` |
| React | **17.0.1** (NOT 18) | `package.json`; React 18 only lands in the SPFx 1.24 *preview* |
| Angular Elements | **21.2.x** | for `lf-ui-components` wrapper; types only (`NgElement`/`WithProperties`) |
| `@laserfiche/lf-ui-components` | **21.1.x** | zoneless — no `zone.js`; the `cdn/` bundle self-registers the custom elements |
| Jest | **30.x** (plain `jest`; `--experimental-vm-modules` no longer needed) | `package.json` `test` script |
| ESLint | **8.57.1** with `@microsoft/eslint-config-spfx` | `.eslintrc.js` |
| Prettier | repo config | `.prettierrc` |
| Gulp | 4.0.2 (SPFx build wrapper) | `gulpfile.js` |

Run `nvm use 22` before building; that is the version CI uses and the only one Microsoft supports for SPFx 1.23.2.

**On Node 24:** the SPFx rig (`@microsoft/sp-build-web` → `SPBuildRig.js`) checks `process.version` against `>=18.17.1 <19 || >=20.11.0 <21 || >=22.14.0 <23` and throws otherwise. `SPFX_OVERRIDE_NODE_VERSION_CHECK=true` bypasses it — this is SPFx's own hook for testing unreleased Node versions. Verified on Node 24.15.0: `build`, `bundle --ship`, `package-solution --ship` and all 29 tests pass, and the `.sppkg` payload is byte-for-byte identical to the Node 22 build. The only file that differs between any two packaging runs is the auto-generated `Client Side Assets` feature GUID in `ClientSideAssets.xml`, which SPFx regenerates every run on *any* Node version — so it is not a Node-version artifact.

**Why this repo stays on gulp:** SPFx 1.22+ scaffolds new projects with Heft, and `m365 spfx project upgrade` will tell you to migrate. Do **not** follow that advice here. `@microsoft/sp-build-web` is still published and patched on the gulp rig (1.23.2, June 2026), and Microsoft supports existing gulp projects. `gulpfile.js` carries load-bearing custom logic — the `copy-vendored-packages` pre-build task plus the `file-loader` rules — that vendors `lf-ui-components` into the `.sppkg` to avoid a runtime CDN fetch from `lfxstatic.com` that SharePoint's CSP blocks. A Heft migration would require rewriting exactly that code.

---

## Repo layout

```
src/
  webparts/                       three web parts + shared constants.ts
    constants.ts                  app-wide IDs, URLs (now resolved via webpack file-loader, see CSP section)
    laserficheAdminConfiguration/
    LaserficheRepositoryAccessWebPart/
    sendToLaserficheLoginComponent/
  extensions/savetoLaserfiche/    command-set (dialogs invoked from list items)
  repository-client/              repository API wrapper
  Assets/
    CSS/         bootstrap.min.css, commonStyles.css
    Images/      logo, icons
    packages-assets.d.ts          TypeScript module declarations for the vendored CDN assets
  __mocks__/                      Jest mocks for @microsoft/sp-*, @laserfiche/*, .png/.svg/.cssasset

config/                           SPFx config (package-solution.json, config.json, etc.)
sharepoint/                       SPFx build outputs (.sppkg lands in sharepoint/solution/)
jekyll_files/                     User & admin documentation site (GitHub Pages)
  docs/assets/                    Versioned .sppkg files for sideload distribution
.github/workflows/                main.yml (SPFx CI/CD), jekyll_gh_pages.yml (docs deploy)
gulpfile.js                       SPFx build + vendored-packages pre-build task
package-solution.json             SPFx solution descriptor (id, version, features)
jest.config.cjs                   Jest config with module mappings for SPFx
```

`lib/` and `temp/` are build outputs — never commit them. They’re gitignored.

---

## Common commands

```sh
# Setup
npm ci                               # use ci, not install, to lock to package-lock.json

# Dev — local SharePoint Workbench
gulp trust-dev-cert                  # one-time per machine
# Edit serve.json: replace REPLACE_WITH_YOUR_SHAREPOINT_SITE
npm run serve

# Build & package (DEV - includeClientSideAssets unbundled)
npm run build                        # gulp bundle (no --ship)
npm run package                      # gulp package-solution (no --ship)

# Build & package (RELEASE - same as CI)
npm run bundle                       # gulp bundle --ship
npm run package-solution             # gulp package-solution --ship
# .sppkg lands at sharepoint/solution/LaserficheSharePointOnlineIntegration.sppkg

# Tests
npm test                             # jest (v30)
# 6 test suites, 29 tests as of 1.0.0.566. All must pass for PR merge.

# Clean
npm run clean                        # gulp clean (wipes lib/, temp/, sharepoint/solution/)
```

**Do NOT build the release `.sppkg` locally for distribution.** Use the [SPFx CI/CD GitHub Action](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml) so the version number is consistent (`1.0.0.${{github.run_number}}`).

---

## CI/CD

### Build pipeline — `.github/workflows/main.yml`

Triggers on push/PR to `\d+.x` branches (so `1.x`, future `2.x`) and `workflow_dispatch`.

- Checkout → setup Node from `.nvmrc` (22.x)
- `if 1.x`: auto-tag the commit `1.0.0.${run_number}`
- `sed -i` replaces the literal `"1.0.0.0"` in `config/package-solution.json` and `package.json` with the run-number version
- `npm ci`
- `gulp build`
- `gulp bundle --ship && gulp package-solution --ship`
- Renames artifact to `LaserficheSharePointOnlineIntegration.1.0.0.${run_number}.sppkg`
- `npm test`
- Uploads the `.sppkg` + `UserDocuments/` as a workflow artifact

### Docs pipeline — `.github/workflows/jekyll_gh_pages.yml`

Builds the Jekyll site from `jekyll_files/` and deploys to GitHub Pages. Requires ACT-team approval.

### Release process

See the [Integrations wiki](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration) for the full runbook. Short version:

1. Merge to `1.x` → CI builds & tags `1.0.0.NNN`
2. Download workflow artifact → grab `solution/LaserficheSharePointOnlineIntegration.1.0.0.NNN.sppkg`
3. Smoke test in a SharePoint test tenant (upload to App Catalog, add web parts)
4. Copy the `.sppkg` into `jekyll_files/docs/assets/` (keep prior versions for rollback) and update the download link in `jekyll_files/docs/admin-documentation/adding-app-organization.md`
5. Open a sideload PR → on merge, run the Jekyll deploy workflow
6. **Microsoft Marketplace publish** — see [Cloud wiki #342030](https://tfs/DefaultCollection/Cloud/_wiki/wikis/Cloud.wiki/342030/Managing-Laserfiche-Apps-the-Microsoft-Marketplace)

---

## ⚠️ SharePoint Online CSP — DO NOT load scripts from external CDNs

Microsoft activated SharePoint Online Content Security Policy (CSP) enforcement on **March 1, 2026** ([MS guidance](https://techcommunity.microsoft.com/blog/spblog/sharepoint-online-content-security-policy-csp-enforcement-dates-and-guidance/4472662)). The default CSP **blocks `script-src` to external origins** including `lfxstatic.com`.

This was fixed in PR #116 (tag `1.0.0.528`) by **vendoring** `lf-ui-components` and the two CSS files into the `.sppkg`. (`zone.js` was vendored too until the v21 upgrade made it unnecessary.) The pattern:

1. **`gulpfile.js`** has a pre-build task `copy-vendored-packages` that copies three files from `node_modules` into `lib/Assets/packages/` at build time:
   - `@laserfiche/lf-ui-components/cdn/lf-ui-components.js` → `lf-ui-components.js`
   - `@laserfiche/lf-ui-components/cdn/indigo-pink.css` → `indigo-pink.cssasset` (rename!)
   - `@laserfiche/lf-ui-components/cdn/lf-ms-office-lite.css` → `lf-ms-office-lite.cssasset` (rename!)

2. The `.cssasset` rename is critical — it stops SPFx’s own CSS pipeline from claiming those files. A custom webpack rule (`module.rules.push(...)` in `gulpfile.js`) catches `\.(js|cssasset)$` inside `lib/Assets/packages/` and runs them through **`file-loader`**, which emits content-hashed copies (e.g. `lf-ui-components.abc12345.js`) and replaces the import with the emitted URL.

3. **`src/webparts/constants.ts`** then re-exports those imports as if they were the old URL constants:
   ```ts
   export { default as LF_UI_COMPONENTS_URL } from '../Assets/packages/lf-ui-components.js';
   export { default as LF_MS_OFFICE_LITE_CSS_URL } from '../Assets/packages/lf-ms-office-lite.cssasset';
   export { default as LF_INDIGO_PINK_CSS_URL } from '../Assets/packages/indigo-pink.cssasset';
   ```

4. **`src/Assets/packages-assets.d.ts`** declares those modules so TypeScript accepts the imports as `string` (the emitted URL).

5. Consumer files (`AdminConfigurationUtilComponents.tsx`, `LaserficheAdminConfiguration.tsx`, `SendToLaserficheLoginComponent.tsx`, `LaserficheRepositoryAccessWebPart.tsx`, `SaveToLaserficheDialog.tsx`, `CommonDialogs.tsx`, `GetDocumentDataDialogUI.tsx`) still call `SPComponentLoader.loadScript(LF_UI_COMPONENTS_URL)` and `SPComponentLoader.loadCss(LF_INDIGO_PINK_CSS_URL)`, but the URL is now a `same-origin` SharePoint URL — CSP-safe.

6. **`config/package-solution.json`** has `"includeClientSideAssets": true` so webpack’s emitted files are physically embedded inside the `.sppkg`.

7. **Jest** can’t process those imports the same way, so `jest.config.cjs` maps them to `genericFileMock.js`:
   ```js
   '\\.cssasset$':                    '<rootDir>/src/__mocks__/genericFileMock.js',
   '/Assets/packages/.*\\.js$':       '<rootDir>/src/__mocks__/genericFileMock.js',
   ```

### Rules

- ❌ **NEVER add a new `SPComponentLoader.loadScript(<external URL>)` or `loadCss(<external URL>)` call.** Any external script CDN will be CSP-blocked.
- ❌ **NEVER hard-code `https://lfxstatic.com/...` URLs in `.ts`/`.tsx` source.** (Some passive `font-src` / `img-src` references inside `.scss` and inside the vendored `lf-ui-components.js` itself remain — those are CSS-level `url()` requests which the default SharePoint CSP currently permits. Hardened-CSP tenants would need follow-up vendoring of those sub-resources too.)
- ✅ If you need a new third-party library, add it as an npm dependency, extend `gulpfile.js`'s `VENDORED_FILES` array, declare it in `src/Assets/packages-assets.d.ts`, and re-export the URL constant from `src/webparts/constants.ts`.

### What’s still in `.scss` (acceptable today, watch for the future)

`src/webparts/laserficheAdminConfiguration/components/LaserficheAdminConfiguration.module.scss` and `LaserficheRepositoryAccessWebPart/components/LaserficheRepositoryAccess.module.scss` still reference `https://lfxstatic.com/libs/Open_Sans/OpenSans.woff` etc. via CSS `url()`. These are `font-src`, not `script-src`, so default SP CSP allows them. If Microsoft tightens `font-src` later, vendor these too.

---

## Architecture notes

### Bundling

`includeClientSideAssets: true` in `config/package-solution.json` means webpack-emitted files (CSS, fonts, the vendored packages) get physically copied into the `.sppkg` and served from the SharePoint tenant. Without this, the vendoring above wouldn’t work — the file-loader-emitted URLs would 404.

### Authentication & API

- Uses `@laserfiche/lf-repository-api-client-v2` for Laserfiche Cloud API calls. Request arguments are named objects keyed `repositoryId` (not v1's `repoId`), collection responses expose `.value`, and `importEntry` is atomic: metadata rides on `ImportEntryRequest.metadata` and a metadata failure fails the whole import (v1's partial-success `CreateEntryResult.operations` envelope is gone).
- OAuth login is handled by `<lf-login>` web component from `lf-ui-components`. `connectedCallback` restores credentials from local storage asynchronously — see `loadLfUiComponentsLib` callers and any logic that reads `authorization_credentials` immediately after the script tag is added.
- Multi-tenant SP app: the SP app registration ID is `clientId = '8ee987ea-a0b1-4ca2-85c4-a79b335cd214'` (in `constants.ts`).

### Solution identity

- Solution ID: `9d9d4fa7-bbf4-489a-acd5-0e1709d40e8b` (`config/package-solution.json`)
- Feature ID: `9c2b7b96-fb6e-4c42-9676-9d434992110f`
- mpnId: `2783806`
- Marketplace shortDescription / longDescription / categories live in the same file — keep them in sync with the Marketplace listing.

---

## Testing

- `npm test` runs Jest with `--experimental-vm-modules` (required for ESM in some deps).
- Tests live next to source as `*.test.ts(x)` or `*.spec.ts(x)`.
- All `@microsoft/sp-*` and `@laserfiche/*` imports are mocked under `src/__mocks__/`. If you add a new SP API import in a tested code path, you may need to extend the mock.
- `npm run serve` for SharePoint Workbench (interactive verification with `spDevMode` localStorage flag to point at `a.clouddev.laserfiche.com`).
- Manual verification in a SharePoint tenant: upload `.sppkg` to `https://<tenant>.sharepoint.com/sites/appcatalog`, install on a site, add the web parts to a page, **check DevTools Console for CSP errors** before merging anything that changes script loads.

---

## Style & PR conventions

- **Format with Prettier** before committing (`.prettierrc` is the source of truth).
- **Lint with ESLint** (`@microsoft/eslint-config-spfx`); CI runs it as part of `gulp build`.
- 2-space indent, single quotes (per Prettier config).
- Functional React components; class components only where the legacy `@microsoft/sp-webpart-base` API requires them (the web part wrapper class).
- Test changes locally: `npm ci && npm test && npm run bundle && npm run package-solution`. CI runs the same.
- PR title pattern: short, imperative, mention the affected component (`Vendor lf-ui-components ... to satisfy SP CSP`, `Resize images`, `Update jws`).
- Link related Issues / TFS work items in the PR description.
- Squash-merge is fine; multi-commit merges (PR #116) are also used for traceability of major refactors.

---

## Things to avoid

- ❌ Branching off / PRing to `main` (use `1.x`).
- ❌ Adding `SPComponentLoader.loadScript(<external URL>)` calls — see the CSP section.
- ❌ Hard-coded `https://lfxstatic.com/...` URLs in `.ts`/`.tsx` source.
- ❌ Building the release `.sppkg` locally for distribution — always use the GitHub Action so the version is auto-stamped.
- ❌ Committing `lib/`, `temp/`, `sharepoint/solution/*.sppkg` (only the versioned `jekyll_files/docs/assets/*.sppkg` belong in git, and only via the sideload PR).
- ❌ Bumping `package-solution.json` `version` or `package.json` `version` by hand — CI does this via `sed`.
- ❌ Upgrading React past 17 without verifying SPFx 1.23 still bundles it (SPFx 1.23 supports React 17; React 18 arrives with SPFx 1.24, still in preview).
- ❌ Setting `SPFX_OVERRIDE_NODE_VERSION_CHECK` in CI, or building release packages on Node 24 — local Node 24 dev is fine, releases build on supported Node 22.
- ❌ Bypassing the SP Workbench — the SPFx serve workflow is the cheapest way to catch render bugs before tenant upload.
- ❌ Putting a caret on `@laserfiche/types-lf-ui-components` while it is on a `--preview-` version. `lf-ui-components-services` declares its peer as `^21.1.0`, and semver caret ranges do **not** match prerelease versions, so `npm i` fails with `ERESOLVE`. Pin it **exactly** and add a matching `overrides` entry (a caret plus an override is rejected outright with `EOVERRIDE`). Keep the pinned version in step with `lf-ui-components`.
- ❌ Dropping the explicit `@angular/*` block from `devDependencies`. Angular is not bundled — it lives inside the vendored `cdn/lf-ui-components.js` — but without those entries `@angular/cdk` (peers `^21 || ^22`) floats `@angular/common` to 22.x and breaks `lf-ui-components`'s `^21.2.8` peer.

---

## Useful links

- [Repo on GitHub](https://github.com/Laserfiche/laserfiche-sharepoint-integration)
- [Admin documentation (GitHub Pages)](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation)
- [SPFx CI/CD workflow](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml)
- [Integrations release runbook](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration)
- [Microsoft Marketplace publish runbook](https://tfs/DefaultCollection/Cloud/_wiki/wikis/Cloud.wiki/342030/Managing-Laserfiche-Apps-the-Microsoft-Marketplace)
- [Microsoft SP CSP enforcement guidance](https://techcommunity.microsoft.com/blog/spblog/sharepoint-online-content-security-policy-csp-enforcement-dates-and-guidance/4472662)
- [SPFx documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
