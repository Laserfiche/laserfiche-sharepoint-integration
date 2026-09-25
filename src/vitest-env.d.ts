// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// Vitest's globals (describe/test/expect/vi, enabled by `globals: true` in
// vitest.config.mts) and jest-dom's matchers, typed for the colocated
// *.test.ts(x) files that `gulp build` type-checks. Referenced here rather than
// in tsconfig.json's `types`: tsconfig sets custom `typeRoots`, and TypeScript
// skips the node_modules lookup for `types` entries when it does.
/// <reference types="vitest/globals" />
/// <reference types="@testing-library/jest-dom/vitest" />
