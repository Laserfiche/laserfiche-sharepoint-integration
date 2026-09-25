// Browser-side stand-in for @microsoft/sp-loader, matching the shape of the
// existing unit-test manual mock (src/__mocks__/@microsoft/sp-loader.ts). The real
// package expects an SPFx-hosted page and isn't meaningful in a bare
// component-test browser context.
export const SPComponentLoader = {
  loadCss: (): void => {
    /* no-op in component tests */
  },
  loadScript: async (): Promise<void> => {
    /* no-op in component tests */
  },
};
