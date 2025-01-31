// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

import '@testing-library/jest-dom';

// Mock the lf-repository-browser custom element
class LfRepositoryBrowser extends HTMLElement {
  initAsync = jest.fn();
}

customElements.define('lf-repository-browser', LfRepositoryBrowser);
