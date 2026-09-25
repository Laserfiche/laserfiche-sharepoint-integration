// Unit-test config for Vitest (jsdom). Tests live next to source as
// *.test.ts(x) / *.spec.ts(x); the Playwright component tests under tests/ct/
// have their own config (playwright-ct.config.ts) and are not run here.
//
// Vite resolves from src/, not from the webpack build output in lib/, so the
// same frictions playwright-ct.config.ts documents apply here too:
//  - Source loads CSS and images with ES `import`, never `require()` (which
//    Vite does not rewrite), so they resolve through the aliases below.
//  - Stylesheets resolve to identity-obj-proxy (`styles.foo === 'foo'`), and
//    the gulp-vendored Assets/packages/* files plus image assets resolve to an
//    empty module.
//  - Every `@microsoft/sp-*` import resolves to its OWN stub file under
//    src/__mocks__/@microsoft/, never to one shared file: vi.mock() keys its
//    registry by resolved path, so two packages aliased to the same file would
//    share (and clobber) each other's vi.mock factories. Add a stub there when
//    a tested code path imports a new SPFx package.
import { defineConfig } from 'vitest/config';
import { readFileSync } from 'fs';
import { resolve } from 'path';
import { fileURLToPath } from 'url';

const rootDir = fileURLToPath(new URL('.', import.meta.url));
const genericFileMock = resolve(rootDir, './src/__mocks__/genericFileMock.js');
const laserficheDistFile = /[\\/]node_modules[\\/]@laserfiche[\\/].*\.js$/;
const sourceMappingUrlComment = /^\/\/# sourceMappingURL=.*$/m;

export default defineConfig({
  plugins: [
    {
      // The inlined @laserfiche/* builds (see server.deps.inline below) end in
      // `//# sourceMappingURL=` comments whose maps are missing or point at
      // unpublished .ts sources, so Vite warns about each file on every run.
      // Loading them with the comment stripped skips that lookup.
      name: 'strip-laserfiche-sourcemap-comments',
      load(id) {
        const file = id.split('?')[0];
        if (!laserficheDistFile.test(file)) return null;
        return readFileSync(file, 'utf-8').replace(sourceMappingUrlComment, '');
      },
    },
  ],
  resolve: {
    alias: [
      // NOTE: a RegExp `find` substitutes only the MATCHED substring (the way
      // `String.prototype.replace` does), so each pattern anchors `^.*` to span
      // the whole import path -- see the same note in playwright-ct.config.ts.
      { find: /^.*\.(css|less|scss|sass)$/, replacement: 'identity-obj-proxy' },
      { find: /^.*\.(cssasset|resx|png|svg)$/, replacement: genericFileMock },
      {
        find: /^.*[\\/]Assets[\\/]packages[\\/].*\.js$/,
        replacement: genericFileMock,
      },
      { find: /^@ms\/.*$/, replacement: genericFileMock },
      {
        find: /^@microsoft\/(sp-[a-z-]+)$/,
        replacement: resolve(rootDir, './src/__mocks__/@microsoft/$1'),
      },
      // The other @laserfiche/* packages load for real (see server.deps.inline);
      // this one stays mocked because tests configure the shared
      // LfRepoTreeNodeService vi.fn() its manual mock exports.
      {
        find: /^@laserfiche\/lf-ui-components-services$/,
        replacement: resolve(
          rootDir,
          './src/__mocks__/@laserfiche/lf-ui-components-services.js'
        ),
      },
    ],
  },
  test: {
    environment: 'jsdom',
    // Vitest's jsdom defaults to http://localhost:3000; the tests assert on an
    // http://localhost/ origin (Jest's default, which they were written against).
    environmentOptions: { jsdom: { url: 'http://localhost/' } },
    globals: true,
    setupFiles: ['./setupTests.ts'],
    include: ['src/**/*.{test,spec}.{js,jsx,ts,tsx}'],
    server: {
      deps: {
        // The @laserfiche/* API clients are "type": "module" but import their
        // own files without extensions (`./ApiException`), which Node's native
        // ESM loader rejects. Inlining runs them through Vite's resolver instead.
        inline: [/[\\/]node_modules[\\/]@laserfiche[\\/]/],
      },
    },
  },
});
