<!--Copyright (c) Laserfiche.
Licensed under the MIT License. See LICENSE.md in the project root for license information.-->

# laserfiche-sharepoint-integration

A SharePoint Framework (SPFx) solution, built with React, that connects SharePoint Online to Laserfiche Cloud. It ships as one package, `LaserficheSharePointOnlineIntegration.sppkg`, containing:

- **Web parts**, the UI components you place on SharePoint pages:
  - [Laserfiche Administrator Configuration](./src/webparts/laserficheAdminConfiguration/): configure metadata mappings from SharePoint to Laserfiche
  - [Laserfiche Repository Explorer](./src/webparts/LaserficheRepositoryAccessWebPart/): browse, open and upload Laserfiche repository documents
  - [Laserfiche Sign In](./src/webparts/sendToLaserficheLoginComponent/): the sign-in page used when saving documents to Laserfiche
- **A command set**, an action on items in lists and libraries:
  - [Save to Laserfiche](./src/extensions/savetoLaserfiche/)

Admin and user documentation: [laserfiche.github.io/laserfiche-sharepoint-integration](https://laserfiche.github.io/laserfiche-sharepoint-integration/). Microsoft's documentation covers [using web parts](https://support.microsoft.com/en-us/office/using-web-parts-on-sharepoint-pages-336e8e92-3e2d-4298-ae01-d404bbe751e0) and [building them](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part).

## Prerequisites

- **Node.js 22 LTS.** The solution targets SPFx 1.23.2, which supports Node 22. The version is pinned in [`.nvmrc`](./.nvmrc), and CI builds on it:

  ```bash
  nvm use 22
  ```

- A **SharePoint Online** site to run the workbench on. Testing a packaged build also needs access to an App Catalog.

<details>
<summary>Developing on Node.js 24 (unsupported)</summary>

No SPFx release supports Node 24, and the SPFx build refuses to run on it by default. To opt in locally, set SPFx's override flag:

```bash
SPFX_OVERRIDE_NODE_VERSION_CHECK=true npx gulp build
```

```powershell
$env:SPFX_OVERRIDE_NODE_VERSION_CHECK = "true"; npx gulp build
```

Builds, packages and tests have been verified on Node 24.15.0, but never set this flag in CI. Release packages are built on Node 22.

</details>

## Getting started

```bash
npm ci                       # install exactly what package-lock.json pins
npm run gulp-trust-dev-cert  # one time per machine: trusts the local HTTPS dev certificate
```

## Build and test

| Command                                        | What it does                                                                                                             |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| `npm test`                                     | Unit tests (Vitest + jsdom)                                                                                              |
| `npm run test:ct`                              | Component tests (Playwright, real Chromium), type-checked first. Before the first run: `npx playwright install chromium` |
| `npm run build`                                | Debug build (`gulp bundle`)                                                                                              |
| `npm run bundle`                               | Release build (`gulp bundle --ship`). This is what CI runs, and it fails on any lint warning                             |
| `npm run package` / `npm run package-solution` | Debug / release `.sppkg`, written to `sharepoint/solution/`                                                              |
| `npm run clean`                                | Removes the build output                                                                                                 |

Before opening a PR, run `npm test`, `npm run test:ct` and `npm run bundle`. CI runs all three.

## Run locally

### In the SharePoint workbench

There is no local workbench (Microsoft removed it in SPFx 1.13), so use the hosted one on your SharePoint site.

1. In [`config/serve.json`](./config/serve.json), replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in `initialPage` with your site. Don't commit that change.
2. Start the dev server:

   ```bash
   npm run serve -- --nobrowser
   ```

3. Open the workbench with the debug parameters, then choose **Load debug scripts**:

   ```text
   https://<your-site>/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true
   ```

   `serve.json` binds the dev server to every address (`"ipAddress": "[::]"`), so the URL `gulp serve` prints or opens contains `[::]:4321`, which a browser can't load. Use `localhost:4321` as shown above.

Troubleshooting:

- **Your changes don't appear:** if the solution is installed on the same site, the installed bundle can load instead of your local code. Uninstall it from that test site.
- **Laserfiche internal only:** to sign in against `a.clouddev.laserfiche.com`, set the `spDevMode` Local Storage key to `true` in the browser's dev tools for the SharePoint site.

### In a SharePoint site, serving code from localhost

1. Build a debug package, whose code loads from `https://localhost:4321`:

   ```bash
   npm run build
   npm run package
   ```

2. Add `sharepoint/solution/LaserficheSharePointOnlineIntegration.sppkg` to your App Catalog and site, following the [Admin Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation).
3. Run `npm run serve -- --nobrowser` while you test.

### In a SharePoint site, with code hosted by SharePoint

```bash
npm run bundle
npm run package-solution
```

This produces a self-contained `sharepoint/solution/LaserficheSharePointOnlineIntegration.sppkg`. Upload it as described in the [Admin Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation). Use it for testing only: release packages come from CI (below).

## Release packages (CI/CD)

Don't build release packages locally. The [SPFx CI CD workflow](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml) runs on every push and pull request to a release branch (`1.x`), and can also be started manually. It:

- stamps the version `1.0.0.<run number>` into the package, and tags the commit with it on `1.x`
- runs the release build, unit tests and component tests
- publishes `LaserficheSharePointOnlineIntegration.1.0.0.<run number>.sppkg` and `UserDocuments/` as a workflow artifact

To publish a sideload release:

1. Download the artifact from the workflow run and take the versioned `.sppkg`.
2. Copy it into [`jekyll_files/docs/assets/`](./jekyll_files/docs/assets/) (keep older versions for rollback), and point the download link in [`adding-app-organization.md`](./jekyll_files/docs/admin-documentation/adding-app-organization.md) at it.
3. Merge that change to `1.x`. This runs the [Deploy SharePoint Integration Documentation](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/jekyll_gh_pages.yml) workflow, which can also be started manually. Its first attempt only builds the site: the `deploy` job is skipped on attempt 1. Re-run the workflow run (**Re-run all jobs**) to publish it. The `github-pages` deployment may also need approval.
4. Check that the new package downloads from the [published page](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-organization.html).

The full procedure, including the Microsoft Marketplace publish, is in the [Integrations wiki release runbook](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration).

## Build the documentation site locally

The documentation is a Jekyll site in [`jekyll_files/`](./jekyll_files/).

1. [Install Ruby and Bundler](https://jekyllrb.com/docs/installation/). CI uses Ruby 3.4.
2. Install the dependencies and serve the site:

   ```bash
   cd jekyll_files
   bundle install
   bundle exec jekyll serve
   ```

3. Open <http://localhost:4000>.

Behind a TLS-inspecting proxy such as Zscaler, `gem`/`bundle` downloads can fail with certificate errors or `Gem::RemoteFetcher::FetchError ... 403`:

- Add the proxy's root certificate (e.g. `zscaler.crt` renamed to `.pem`) to RubyGems' certificate folder, e.g. `C:\Ruby33-x64\lib\ruby\3.3.0\rubygems\ssl_certs\rubygems.org`, and append its contents to `C:\Ruby33-x64\bin\etc\ssl\cert.pem`. Verify with [check.rb](https://github.com/rubygems/ruby-ssl-check/blob/master/check.rb).
- If one gem still won't download, download its `.gem` file manually and run `gem install <file>` (for example `ffi` or `google-protobuf`) from the folder you saved it in, then re-run `bundle install` in `jekyll_files/`.

## Change log

See [CHANGELOG](./jekyll_files/CHANGELOG.md).

## Contributing

We welcome contributions and feedback. See the [contributing guidelines](./CONTRIBUTING.md).

## License

[MIT](./LICENSE.md)
