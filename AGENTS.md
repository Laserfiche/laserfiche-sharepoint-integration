# AGENTS.md — laserfiche-sharepoint-integration

Instructions for AI coding agents working in this repo. Humans: start with [`README.md`](./README.md) and [`CONTRIBUTING.md`](./CONTRIBUTING.md).

A SharePoint Framework (SPFx) solution, written in React, that integrates Laserfiche Cloud with SharePoint Online. It ships as one `.sppkg`:

| Component                                         | Path                                              |
| ------------------------------------------------- | ------------------------------------------------- |
| Laserfiche Administrator Configuration (web part) | `src/webparts/laserficheAdminConfiguration/`      |
| Laserfiche Repository Explorer (web part)         | `src/webparts/LaserficheRepositoryAccessWebPart/` |
| Laserfiche Sign In (web part)                     | `src/webparts/sendToLaserficheLoginComponent/`    |
| Save to Laserfiche (command set)                  | `src/extensions/savetoLaserfiche/`                |

---

## How to work here

### 1. Reuse before you add (DRY)

Before writing markup, a style, a string, an icon or a helper, search for one that already exists and reuse it. If something close exists but isn't shared, extract it to a shared place and move **every** copy onto it in the same change. Never paste a second copy of a component, style block, data URI or asset.

Where the shared pieces live:

| What                                                                                                                                            | Where                                                          |
| ----------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------- |
| Dialog building blocks: `LaserficheDialogTitle` (icon + title), `MessageDialog`, `useConfirm`, `Collapsible`, `LoadingDialog`, `DelayedSpinner` | `src/extensions/savetoLaserfiche/CommonDialogs.tsx`            |
| Dialog styles (`wrapper`, `header`, `contentBox`, `footer`, `logoHeader`, …)                                                                    | `src/extensions/savetoLaserfiche/SendToLaserFiche.module.scss` |
| App-wide constants and asset URLs (`LASERFICHE_ICON_URL`, vendored package URLs, client ID, scopes)                                             | `src/webparts/constants.ts`                                    |
| User-facing strings                                                                                                                             | `src/webparts/strings.ts`                                      |
| Helpers                                                                                                                                         | `src/Utils/Funcs.ts`                                           |
| Sign in / Sign out through the LaserficheSignIn popup (`useSignInPopup`)                                                                        | `src/Utils/useSignInPopup.tsx`                                 |
| The `<lf-login>` element every page and dialog signs in through (`LaserficheLogin`); never render a bare `<lf-login>`                           | `src/Utils/LaserficheLogin.tsx`                                |
| JSX typings for the `lf-ui-components` custom elements (never redeclare them inside a component)                                                | `src/lf-ui-components-elements.d.ts`                           |
| Images                                                                                                                                          | `src/Assets/Images/`                                           |

When you add something reusable, put it in one of these places, not next to its first caller.

### 2. Clean code

- Small, single-purpose functions and components with descriptive names. No dead or commented-out code.
- Refactor as you go, leaving code you touch cleaner than you found it. Keep refactors behavior-preserving and covered by tests, and keep them separate in intent from feature changes.
- Match the surrounding code. Prettier (`.prettierrc`) and ESLint (`@microsoft/eslint-config-spfx`) are the source of truth: 2-space indent, single quotes, 100-column lines. Run `npm run format` before committing; CI fails on unformatted files. `.prettierignore` keeps Prettier off `jekyll_files/`, whose Liquid templates it would break.
- Comments explain _why_, not _what_.
- Functional React components and hooks. Use class components only where the SPFx base classes require them.
- Accessibility: decorative images get `alt=''`, and icon-only buttons get a `title` or `aria-label`.

### 3. Test-driven development

- **Red → green → refactor.** First write a test that fails for the right reason, then make it pass with the simplest change, then refactor while it stays green. For a bug, reproduce it with a failing test before fixing it.
- Test behavior through what the user sees (text, roles, titles, enabled/disabled), not implementation details. Prefer Testing Library's role/text/title queries.
- Test at the lowest level that can prove the behavior:
  - **Vitest + jsdom** (`src/**/*.test.ts(x)`, next to the source): logic, DOM structure and text.
  - **Playwright component tests** (`tests/ct/*.ct.tsx`): real `File` objects, CSS visibility, focus/keyboard, and `lf-ui-components` custom-element events. Don't add CT coverage just to have it.
- Keep tests current with the code. Changing a component's markup or flow means updating its tests in the same change. Never skip or delete a failing test to get a change through.
- **Definition of done:** `npm run format:check`, `npm test`, `npm run test:ct` **and** `npm run bundle` all pass. The bundle (`gulp bundle --ship`) is the stricter SHIP build CI runs: it fails on any lint warning, and it is where TypeScript errors in the Vitest test files surface (Vitest strips types without checking them). `npm run test:ct` type-checks the component tests against `tsconfig.ct.json` before running them.
- Delete scratch or diagnostic tests before finishing, and check `git status` for stray files.

---

## Project essentials

### Branches

Branch off **`1.x`** and open PRs against it. `main` is abandoned.

### Toolchain (pinned; don't bump casually)

| Tool                           | Version                                          | Note                                                                                                                                                                                                     |
| ------------------------------ | ------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Node.js                        | **22.x** (`.nvmrc`)                              | Node 24 works locally only with `SPFX_OVERRIDE_NODE_VERSION_CHECK=true`; never in CI or for releases                                                                                                     |
| SPFx                           | **1.23.2**                                       | all `@microsoft/sp-*` pinned together; stays on the **gulp** rig. Don't migrate to Heft: `gulpfile.js` carries the vendoring logic below                                                                 |
| TypeScript                     | 5.3.3                                            | `strict`, but `strictNullChecks: false`                                                                                                                                                                  |
| React                          | **17.0.1**                                       | React 18 waits for SPFx 1.24                                                                                                                                                                             |
| `@laserfiche/lf-ui-components` | 21.1.x                                           | Angular Elements 21.2.x; keep the explicit `@angular/*` devDependencies. `@laserfiche/types-lf-ui-components` is on a `--preview-` build: pin it exactly, with a matching `overrides` entry and no caret |
| Vitest                         | 5.x                                              | jsdom **26.1.0** pinned                                                                                                                                                                                  |
| Testing Library                | RTL **12.1.5**, jest-dom **6.9.1**               | exact pins: the last React-17-compatible versions                                                                                                                                                        |
| Playwright CT                  | `@playwright/experimental-ct-react17` **1.62.1** | exact pin (experimental package)                                                                                                                                                                         |
| ESLint                         | 8.57.1                                           |                                                                                                                                                                                                          |
| Prettier                       | **3.9.9**                                        | exact pin, so the CLI, CI and the VS Code extension (which uses the project's copy) format identically                                                                                                   |

### Commands

```sh
npm ci                   # install (never `npm install` for setup)
npm test                 # Vitest
npm run test:ct          # type-check (tsconfig.ct.json), then Playwright component tests
npm run format           # Prettier: rewrite files in place
npm run format:check     # Prettier: fail on unformatted files (what CI runs)
npm run build            # gulp build (DEBUG, lenient)
npm run bundle           # gulp bundle --ship (what CI runs)
npm run package-solution # .sppkg -> sharepoint/solution/
npm run serve            # hosted workbench only; there is no local workbench
```

`.npmrc`'s `min-release-age=3` keeps `npm install` and `npm update` from resolving a version published less than 3 days ago, but only on npm 11.10 or later. Node 22's bundled npm 10.9 ignores it, and `npm ci` never resolves versions, so don't count on it in CI.

For manual checks against the hosted workbench (`https://<site>/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true`), use a real browser such as Claude in Chrome. Sandboxed agent browser panes block the tenant from reaching `localhost`.

### SharePoint CSP: never load external scripts

SharePoint Online blocks `script-src` to external origins, `lfxstatic.com` included. `lf-ui-components` and its CSS are therefore **vendored**. `gulpfile.js`'s `copy-vendored-packages` task copies them into `lib/Assets/packages/` (CSS renamed to `.cssasset`), webpack's `file-loader` emits them, and `src/webparts/constants.ts` re-exports the resulting same-origin URLs (`LF_UI_COMPONENTS_URL`, …).

- ❌ No `SPComponentLoader.loadScript/loadCss(<external URL>)` and no hard-coded `https://lfxstatic.com/...` in `.ts`/`.tsx`.
- ✅ A new third-party file: add the npm dependency, then extend `VENDORED_FILES` in `gulpfile.js`, declare the module in `src/Assets/packages-assets.d.ts`, and re-export the URL from `constants.ts`.
- Bootstrap **5.3.8** CSS is vendored the same way, into `lib/Assets/CSS/`. Never commit `bootstrap.min.css` under `src/`. Bootstrap's JS isn't bundled.

### Markup and imports

- Bootstrap **5** classes only (`form-select`, `mb-3`, `btn-close`, `data-bs-*`). None of the v4 ones (`custom-select`, `form-group`, `.close`, `data-dismiss`).
- Load CSS and images with ES `import`, never `require()`. The Vite-based test runners have no `require()`.
- `React.useRef()`, not `React.createRef()`, inside function components.

### Test infrastructure gotchas

- `@microsoft/sp-*` imports resolve to **one stub per package** under `src/__mocks__/@microsoft/`. Never alias two packages to one file, because their `vi.mock` factories would collide. The `@laserfiche/*` API clients load for real; don't hand-copy their helpers into mocks.
- A `vi.fn()` that source code calls with `new` needs a `function`/`class` implementation.
- CT `mount()` props are serialized, so build mocks inside a harness component (`tests/ct/harness/`) from a serializable `scenario` prop. Omit a key rather than setting it to `undefined`. Prefer `page.*` locators over `component.*`.
- `*.module.scss` is an identity stub in both runners (CT prefixes class names with `ct-scss-`), so real styles aren't applied there.

### CI and release

- `.github/workflows/main.yml` runs on `\d+.x` branches: `npm ci` → `npm run format:check` → `gulp build` → `gulp bundle --ship && gulp package-solution --ship` → `npm test`, plus a parallel `npm run test:ct` job.
- `.github/workflows/jekyll_gh_pages.yml` builds the docs site (`jekyll_files/`) on push to `1.x`, but it deploys to GitHub Pages only on a re-run (`run_attempt != 1`).
- CI stamps the version `1.0.0.<run_number>`. Never bump `version` in `package.json` or `config/package-solution.json` by hand, and never build a release `.sppkg` locally.
- Never commit `lib/`, `temp/` or `sharepoint/solution/*.sppkg`.
- Runbook: [Integrations wiki](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration).

### PRs

Short imperative title that names the component. Link the related issue or TFS work item. If a change alters a convention documented here, update this file in the same PR.

---

## Links

- [Repo](https://github.com/Laserfiche/laserfiche-sharepoint-integration) · [CI workflow](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml) · [Admin docs](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation)
- [Microsoft Marketplace runbook](https://tfs/DefaultCollection/Cloud/_wiki/wikis/Cloud.wiki/342030/Managing-Laserfiche-Apps-the-Microsoft-Marketplace) · [SharePoint CSP guidance](https://techcommunity.microsoft.com/blog/spblog/sharepoint-online-content-security-policy-csp-enforcement-dates-and-guidance/4472662) · [SPFx docs](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
