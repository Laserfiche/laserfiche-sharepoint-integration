<!--Copyright (c) Laserfiche.
Licensed under the MIT License. See LICENSE.md in the project root for license information.-->

# laserfiche-sharepoint-integration

## Summary

This project, built with React, contains 3 SharePoint web parts and a command set that can be used to communicate with Laserfiche. To learn more about web parts, consult Microsoft's documentation for [using them](https://support.microsoft.com/en-us/office/using-web-parts-on-sharepoint-pages-336e8e92-3e2d-4298-ae01-d404bbe751e0) and [building them](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/build-a-hello-world-web-part).

Project contains:

- 3 Web Parts (UI Components that are placed on SharePoint pages)
  - [Laserfiche Administrator Configuration](./src/webparts/laserficheAdminConfiguration/)
  - [Laserfiche Repository Explorer](./src/webparts/LaserficheRepositoryAccessWebPart/)
  - [Laserfiche Sign In](./src/webparts/sendToLaserficheLoginComponent/)
- 1 Command Set (Action available on items in lists and libraries)
  - [Send to Laserfiche](./src/extensions/savetoLaserfiche/)

Admin and User Documentation is available on [GitHub pages](https://laserfiche.github.io/laserfiche-sharepoint-integration/)

## Prerequisites

See .github/workflows/main.yml for Node and NPM version used.

## Change Log

See CHANGELOG [here](./jekyll_files/CHANGELOG.md).

## Contribution

We welcome contributions and feedback. Please follow our [contributing guidelines](./CONTRIBUTING.md).

---

## To run locally in your SharePoint Workbench

- Ensure that you are at the solution folder
  - run **npm install**
  - run **npm run gulp-trust-dev-cert**
  - Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
  - run **npm run serve**
    - this should open up a window in the browser called a SharePoint workbench.
  - To use a.clouddev.laserfiche.com: Open browser dev tools and go to site Local Storage: set 'spDevMode' to true

## To test in a SharePoint site using localhost

- run **npm install**
- **npm run build**
- **npm run package**
- this should result in the creation of a file with the path `/sharepoint/solution/LaserficheSharePointOnlineIntegration.sppkg` from the root folder.
- Navigate to the solution folder
- run **npm run gulp-trust-dev-cert** (one-time only)
- Replace `REPLACE_WITH_YOUR_SHAREPOINT_SITE` in serve.json with your sharepoint site
- run **npm run serve** to host the code for the integration
- reference the [Admin Documentation](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation) for instructions on how to use the solution file you generated in the first steps to test your changes to the web parts in SharePoint sites.

## To test in a SharePoint site with files hosted by SharePoint

- run **npm install**
- **npm run bundle**
- **npm run package-solution**
- This will create the solution file at /sharepoint/solution/LaserficheSharePointOnlineIntegration.sppkg.
- Once you've built and packaged the solution file, you can use it as a production package and upload it in the SharePoint admin center (see [Admin Docs](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation))

## To produce a release-ready package via CI/CD

For official releases, do **not** build the `.sppkg` locally. Use the [SPFx CI/CD GitHub Action](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/main.yml) instead:

1. Run the workflow off the release branch (e.g. `1.x`).
2. The workflow auto-versions the artifact as `LaserficheSharePointOnlineIntegration.[VERSION_NUMBER].sppkg` (version derived from `package.json` + build number).
3. Download the workflow artifact → use the `.sppkg` inside `artifact/solution/`.
4. For sideload distribution: copy that file into `jekyll_files/docs/assets/` (keep historical versions for rollback), update the download link in `jekyll_files/docs/admin-documentation/adding-app-organization.md` to point at the new version, then run the [Deploy SharePoint Integration Documentation](https://github.com/Laserfiche/laserfiche-sharepoint-integration/actions/workflows/jekyll_gh_pages.yml) action.
5. Verify the new package is reachable from the [download link](https://laserfiche.github.io/laserfiche-sharepoint-integration/docs/admin-documentation/adding-app-organization.html).

See the [Integrations wiki release runbook](https://v-dev-tfs.laserfiche.com/DefaultCollection/Integrations/_wiki/wikis/Integrations.wiki/134901/Laserfiche-SharePoint-Online-Integration) for the full release procedure (including Microsoft Marketplace publish).

## To build documentation locally

1. [Install Ruby and bundler](https://jekyllrb.com/docs/installation/)
   - See the instructions [here](https://jekyllrb.com/docs/) to install jekyll and bundler.
     - You may need to add the zscaler.crt (rename extension to .pem) to your rubygems certificate folder, e.g.: `C:\Ruby33-x64\lib\ruby\3.3.0\rubygems\ssl_certs\rubygems.org`
     - You may also need to paste the zscaler.crt content into `C:\Ruby33-x64\bin\etc\ssl\cert.pem`
   - Verify installation running [check.rb](https://github.com/rubygems/ruby-ssl-check/blob/master/check.rb)
   - run `gem install bundler`
   - run `gem install jekyll`
     - If fails to download a package, you may need to manually download the gem and install it. E.g. switch to the download folder and run:
       - `gem install ffi`
       - `gem install google-protobuf`
1. navigate to the src directory
1. remove the Gemfile.lock, if it exists
1. run `bundle install`
1. run `bundle exec jekyll serve`

- Check installations and versions:

  ```sh
  gem --version
  bundle --version
  bundle exec jekyll --version
  ```

1. Navigate to directory `<project_path>/jekyll_files`.
1. Run `bundle install` to install all the dependencies needed to serve.
   - Troubleshoot:
     - If you see error `Gem::RemoteFetcher::FetchError bad response Forbidden 403`, stay under directory `/jekyll_files` and follow [the steps](#fetch_error) in the dependency issues in step 1 to install the dependency from a local .gem file.
1. Run `bundle exec jekyll serve` to serve the documentation.
1. Open `localhost:4000` in a browser.
