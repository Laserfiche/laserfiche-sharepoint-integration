// Browser-side stand-in for @microsoft/sp-http, matching the shape used by
// the existing unit tests (vi.mock('@microsoft/sp-http', () => ({
// SPHttpClient: { configurations: { v1: {} } } }))).
export const SPHttpClient = {
  configurations: {
    v1: {},
  },
};
