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
| Node.js | **18.x** (`>=18.0.0 <19.0.0` per `package.json` engines) | locally |
| Node.js (CI) | **20.x** | `.github/workflows/main.yml` (`actions/setup-node@v4` `node-version: 20.x`) |
| SPFx | **1.20.x** (`@microsoft/sp-build-web` `1.20.2`, all `@microsoft/sp-*` `1.20.0`) | `package.json` |
| TypeScript | **4.9.5** + Rush stack compiler 4.5 | `tsconfig.json` extends `@microsoft/rush-stack-compiler-4.5` |
| React | **17.0.1** (NOT 18) | `package.json` |
| Angular Elements | **16.2.x** | for `lf-ui-components` wrapper |
| `@laserfiche/lf-ui-components` | **16.x** | NOT v21 yet (v21 is zoneless) |
| Jest | **29.x** with `--experimental-vm-modules` | `package.json` `test` script |
| ESLint | **8.57.1** with `@microsoft/eslint-config-spfx` | `.eslintrc.js` |
| Prettier | repo config | `.prettierrc` |
| Gulp | 4.0.2 (SPFx build wrapper) | `gulpfile.js` |

If you accidentally use Node 22/Node 21 etc., builds may pass locally but produce a `.sppkg` that differs from CI’s output — always run final verification on Node 20.x to match the GitHub Action.

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
npm test                             # jest with --experimental-vm-modules
# 6 test suites, 29 tests as of 1.0.0.566. All must pass for PR merge.

# Clean
npm run clean                        # gulp clean (wipes lib/, temp/, sharepoint/solution/)
```

**Do NOT build the release `.sppkg` locally for distribution.** Use the [SPFx CI/CD GitHub Action](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml) so the version number is consistent (`1.0.0.${{github.run_number}}`).

---

## CI/CD

### Build pipeline — `.github/workflows/main.yml`

Triggers on push/PR to `\d+.x` branches (so `1.x`, future `2.x`) and `workflow_dispatch`.

- Checkout → setup Node 20.x
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

This was fixed in PR #116 (tag `1.0.0.528`) by **vendoring** `lf-ui-components`, the two CSS files, and `zone.js` into the `.sppkg`. The pattern:

1. **`gulpfile.js`** has a pre-build task `copy-vendored-packages` that copies four files from `node_modules` into `lib/Assets/packages/` at build time:
   - `@laserfiche/lf-ui-components/cdn/lf-ui-components.js` → `lf-ui-components.js`
   - `@laserfiche/lf-ui-components/cdn/indigo-pink.css` → `indigo-pink.cssasset` (rename!)
   - `@laserfiche/lf-ui-components/cdn/lf-ms-office-lite.css` → `lf-ms-office-lite.cssasset` (rename!)
   - `zone.js/bundles/zone.umd.min.js` → `zone.umd.min.js`

2. The `.cssasset` rename is critical — it stops SPFx’s own CSS pipeline from claiming those files. A custom webpack rule (`module.rules.push(...)` in `gulpfile.js`) catches `\.(js|cssasset)$` inside `lib/Assets/packages/` and runs them through **`file-loader`**, which emits content-hashed copies (e.g. `lf-ui-components.abc12345.js`) and replaces the import with the emitted URL.

3. **`src/webparts/constants.ts`** then re-exports those imports as if they were the old URL constants:
   ```ts
   export { default as LF_UI_COMPONENTS_URL } from '../Assets/packages/lf-ui-components.js';
   export { default as LF_MS_OFFICE_LITE_CSS_URL } from '../Assets/packages/lf-ms-office-lite.cssasset';
   export { default as LF_INDIGO_PINK_CSS_URL } from '../Assets/packages/indigo-pink.cssasset';
   export { default as ZONE_JS_URL } from '../Assets/packages/zone.umd.min.js';
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

- Uses `@laserfiche/lf-repository-api-client` v1 (NOT v2) for Laserfiche Cloud API calls.
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
- ❌ Upgrading React past 17 without verifying SPFx 1.20 still bundles it (SPFx 1.20 supports React 17; React 18+ may need SPFx 1.21+).
- ❌ Upgrading Node to >=19 locally without verifying CI’s Node 20 still produces an equivalent `.sppkg`.
- ❌ Removing `zone.js` until `lf-ui-components` is upgraded to v21 (zoneless) — `@angular/elements` 15/16 still needs Zone.js.
- ❌ Bypassing the SP Workbench — the SPFx serve workflow is the cheapest way to catch render bugs before tenant upload.

---

## Useful links

- [Repo on GitHub](https://github.com/Laserfiche/laserfiche-sharepoint-integration)
- [Admin documentation (GitHub Pages)](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation)
- [SPFx CI/CD workflow](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml)
- [Integrations release runbook](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration)
- [Microsoft Marketplace publish runbook](https://tfs/DefaultCollection/Cloud/_wiki/wikis/Cloud.wiki/342030/Managing-Laserfiche-Apps-the-Microsoft-Marketplace)
- [Microsoft SP CSP enforcement guidance](https://techcommunity.microsoft.com/blog/spblog/sharepoint-online-content-security-policy-csp-enforcement-dates-and-guidance/4472662)
- [SPFx documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
