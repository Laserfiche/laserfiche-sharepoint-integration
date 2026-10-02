// Stands in for the gulp-vendored files that `src/webparts/constants.ts`
// re-exports from `../Assets/packages/*` (created by gulpfile.js's
// copy-vendored-packages task into lib/Assets/packages/, not present under
// src/ at all). Component tests never load the real lf-ui-components bundle,
// so a plain placeholder URL is enough to satisfy the import.
export default 'https://localhost/mock-vendored-asset.js';
