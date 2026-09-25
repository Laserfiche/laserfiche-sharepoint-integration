// Component-test config for @playwright/experimental-ct-react17. Mounts the
// real React components in a real Chromium browser with the lf-ui-components
// custom elements stubbed (tests/ct/playwright/index.ts) and SPFx context
// passed as plain mock props -- no live SharePoint/Laserfiche tenant.
//
// Vite (used to bundle components here) resolves from src/, unlike webpack
// which resolves the *build output* under lib/. Three repo-specific
// frictions and their fixes:
//  1. src/webparts/constants.ts re-exports URLs from ../Assets/packages/*,
//     which gulpfile.js's copy-vendored-packages task only creates under
//     lib/, never under src/. -> aliased to tests/ct/stubs/asset-url.ts.
//  2. AdminMainPage.tsx / ManageMappingsPage.tsx import a bootstrap.min.css
//     that likewise only exists under lib/ (vendored at build time) -> aliased
//     to tests/ct/stubs/empty-css.ts (real Bootstrap is loaded once globally
//     from tests/ct/playwright/index.ts instead).
//  3. `*.module.scss` imports resolve to gitignored generated .module.scss.ts
//     files for TypeScript (tests/ct/scss-modules.d.ts covers typecheck:ct
//     when gulp hasn't generated them); Vite would try to compile the real .scss,
//     and the hoisted `sass` package is pinned by SPFx below what Vite's
//     default Sass API needs -> aliased to an identity Proxy stub, matching
//     the unit tests' identity-obj-proxy alias in vitest.config.mts.
// Source must load CSS and images with ES `import`, never `require()`: Vite
// leaves a bare require() untouched, so it throws "require is not defined" in
// the browser.
import { defineConfig, devices } from '@playwright/experimental-ct-react17';
import { resolve } from 'path';

export default defineConfig({
  testDir: './tests/ct',
  testMatch: '**/*.ct.tsx',
  timeout: 15 * 1000,
  // job-summary-reporter.ts only writes when GITHUB_STEP_SUMMARY is set (CI).
  reporter: [['html'], ['./tests/ct/job-summary-reporter.ts']],
  use: {
    trace: 'on-first-retry',
    ctTemplateDir: './tests/ct/playwright',
    ctViteConfig: {
      // Playwright CT always builds with `build.sourcemap: true`, but Vite
      // only emits a real CSS sourcemap when `devSourcemap` is on (it honors
      // it in build mode too) AND the stylesheet actually goes through
      // PostCSS. With no PostCSS config, Vite skips PostCSS for any file with
      // no url()/@import -- e.g. src/Assets/CSS/commonStyles.css -- and
      // returns it without a map, which Rolldown reports as SOURCEMAP_BROKEN.
      // A no-op plugin makes Vite run PostCSS (and generate a map) for every
      // stylesheet.
      css: {
        devSourcemap: true,
        postcss: { plugins: [{ postcssPlugin: 'ct-force-css-sourcemaps' }] },
      },
      resolve: {
        alias: [
          // NOTE: Vite/Rollup's alias plugin substitutes a RegExp `find`
          // the way `String.prototype.replace` does -- only the MATCHED
          // substring is swapped for `replacement`, not the whole specifier.
          // Every pattern below must therefore anchor `^.*` so the match
          // spans the entire import path, or the replacement gets appended
          // to the unmatched remainder instead of replacing it (this bit us
          // during the initial spike: paths came out as
          // "./SendToLaserFicheC:\...\scss-module.ts").
          {
            find: /^.*Assets[\\/]packages[\\/].*$/,
            replacement: resolve(__dirname, './tests/ct/stubs/asset-url.ts'),
          },
          {
            find: /^.*Assets[\\/]CSS[\\/]bootstrap\.min\.css$/,
            replacement: resolve(__dirname, './tests/ct/stubs/empty-css.ts'),
          },
          {
            find: /^.*\.module\.scss$/,
            replacement: resolve(__dirname, './tests/ct/stubs/scss-module.ts'),
          },
          {
            find: '@microsoft/sp-loader',
            replacement: resolve(__dirname, './tests/ct/stubs/sp-loader.ts'),
          },
          {
            find: '@microsoft/sp-http',
            replacement: resolve(__dirname, './tests/ct/stubs/sp-http.ts'),
          },
          {
            find: '@laserfiche/lf-ui-components-services',
            replacement: resolve(__dirname, './tests/ct/stubs/lf-ui-components-services.ts'),
          },
        ],
      },
    },
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
});
