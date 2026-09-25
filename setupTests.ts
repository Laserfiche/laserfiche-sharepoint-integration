// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import '@testing-library/jest-dom/vitest';

// Mock the lf-repository-browser custom element
class LfRepositoryBrowser extends HTMLElement {
  initAsync = vi.fn();
}

customElements.define('lf-repository-browser', LfRepositoryBrowser);
