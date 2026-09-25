// Mirrors the unit tests' identity-obj-proxy handling of `*.module.scss` imports
// (vitest.config.mts aliases the same specifiers to identity-obj-proxy), with one
// deliberate difference: keys are prefixed rather than returned verbatim.
// Real Bootstrap CSS is loaded globally for these component tests (see
// tests/ct/playwright/index.ts), and this repo's markup mixes CSS-module
// classes with plain Bootstrap ones (e.g. `className={`modal ${styles.modal}`}`-
// style patterns and bare `.modal-dialog`/`.modal-body`). In production the
// SCSS build emits hashed class names (e.g. `modal_a1b2c3`) that can't
// collide with Bootstrap's own global `.modal` selector; returning the bare
// key here would collide with it directly and pick up Bootstrap's default
// `display: none`, hiding elements that are actually visible in production.
// A prefixed identity value keeps class names assertable in tests
// (`toHaveClass(/ct-scss-modalContent/)`) without that collision. This does
// mean a CT assertion can't verify a real `.module.scss` rule (e.g. that
// `styles.hideImport` really hides something) -- that's an accepted gap,
// not something these tests attempt to cover.
const identityObjProxy: Record<string, string> = new Proxy(
  {},
  {
    get: (_target, key: string) => (key === '__esModule' ? false : `ct-scss-${key}`),
  }
);

export default identityObjProxy;
