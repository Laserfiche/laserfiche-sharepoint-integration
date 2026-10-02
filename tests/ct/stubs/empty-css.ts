// Stands in for `Assets/CSS/bootstrap.min.css`, which only exists under
// lib/ after gulpfile.js's copy-vendored-packages task runs (not under src/).
// Real Bootstrap CSS is loaded once, globally, from tests/ct/playwright/index.ts
// instead, so this require just needs to resolve to something harmless.
export default {};
