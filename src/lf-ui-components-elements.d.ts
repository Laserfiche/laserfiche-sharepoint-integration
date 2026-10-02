// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// The @laserfiche/lf-ui-components custom elements rendered as JSX. They're
// driven through refs (initAsync, properties, events), so their attributes
// stay untyped. Declared once here, rather than in each component, so every
// program that type-checks those components sees them -- gulp's tsc and
// tsconfig.ct.json alike.
declare namespace JSX {
  interface IntrinsicElements {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ['lf-field-container']: any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ['lf-login']: any;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ['lf-repository-browser']: any;
  }
}
